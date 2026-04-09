from __future__ import annotations

import argparse
import csv
import json
import mimetypes
import os
import re
import subprocess
import tempfile
import webbrowser
from dataclasses import dataclass
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parent.parent
APP_ROOT = REPO_ROOT / "web" / "etabs-control-center"
SCRIPTS_ROOT = REPO_ROOT / "scripts"
POWERSHELL = "powershell"
DISPLAY_LENGTH_UNIT = "Feet"


class AppError(Exception):
    pass


@dataclass
class StaticAsset:
    path: Path
    content_type: str


def run_powershell_script(script_name: str, extra_args: list[str] | None = None) -> Any:
    command = [
        POWERSHELL,
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPTS_ROOT / script_name),
        "-AsJson",
    ]
    if extra_args:
        command.extend(extra_args)

    completed = subprocess.run(
        command,
        cwd=str(REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )

    stdout = completed.stdout.strip()
    stderr = completed.stderr.strip()
    if completed.returncode != 0:
        message = stderr or stdout or f"{script_name} exited with {completed.returncode}."
        raise AppError(message)

    if not stdout:
        raise AppError(f"{script_name} did not return JSON.")

    try:
        return json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise AppError(f"{script_name} returned invalid JSON: {exc}") from exc


def try_parse_float(raw_value: str) -> float:
    candidate = raw_value.strip().replace(",", "")
    if not candidate:
        raise ValueError("blank numeric value")
    return float(candidate)


def is_numeric_value(raw_value: str) -> bool:
    try:
        try_parse_float(raw_value)
        return True
    except ValueError:
        return False


def is_base_elevation_line(cells: list[str]) -> bool:
    normalized = " ".join(cell.strip().lower() for cell in cells)
    return "base" in normalized and "elev" in normalized


def parse_base_elevation_line(cells: list[str], line_number: int) -> float:
    for index in range(len(cells) - 1, -1, -1):
        if is_numeric_value(cells[index]):
            return try_parse_float(cells[index])

    raise AppError(f"Line {line_number} looks like a base elevation row but does not include a numeric value.")


def looks_like_header(cells: list[str]) -> bool:
    lowered = [cell.strip().lower() for cell in cells]
    joined = " ".join(lowered)
    return ("story" in joined or "name" in joined) and ("elev" in joined or "height" in joined)


def get_header_map(cells: list[str]) -> dict[str, int]:
    normalized = [cell.strip().lower() for cell in cells]
    header_map: dict[str, int] = {}

    for index, cell in enumerate(normalized):
        if "story" in cell or cell == "name":
            header_map.setdefault("name", index)
        elif "elev" in cell:
            header_map.setdefault("elevation", index)
        elif "height" in cell:
            header_map.setdefault("height", index)
        elif cell in {"#", "index", "row", "no.", "no"}:
            header_map.setdefault("index", index)

    return header_map


def parse_story_line(cells: list[str], line_number: int, header_map: dict[str, int] | None) -> tuple[str, float]:
    if header_map and "name" in header_map and "elevation" in header_map:
        try:
            name = cells[header_map["name"]].strip()
            elevation = try_parse_float(cells[header_map["elevation"]])
        except (IndexError, ValueError) as exc:
            raise AppError(f"Line {line_number} could not be parsed from the detected header columns.") from exc

        if not name:
            raise AppError(f"Line {line_number} has a blank story name.")
        return name, elevation

    if len(cells) >= 4 and is_numeric_value(cells[0]) and not is_numeric_value(cells[1]) and is_numeric_value(cells[2]):
        return cells[1].strip(), try_parse_float(cells[2])

    if len(cells) >= 3 and not is_numeric_value(cells[0]) and is_numeric_value(cells[1]) and is_numeric_value(cells[2]):
        return cells[0].strip(), try_parse_float(cells[1])

    name: str | None = None
    elevation: float | None = None
    if len(cells) >= 2:
        for index in range(len(cells) - 1, 0, -1):
            try:
                numeric_value = try_parse_float(cells[index])
                name_candidate = " ".join(cell for cell in cells[:index] if cell).strip()
                if name_candidate:
                    name = name_candidate
                    elevation = numeric_value
                    break
            except ValueError:
                continue

    if not name or elevation is None:
        raise AppError(f"Line {line_number} could not be parsed. Expected a story name and a numeric elevation.")

    return name, elevation


def split_story_line(line: str) -> list[str]:
    stripped = line.strip()
    if not stripped:
        return []

    if "\t" in stripped:
        return [cell.strip() for cell in stripped.split("\t") if cell.strip()]

    if "," in stripped:
        return [cell.strip() for cell in next(csv.reader([stripped])) if cell.strip()]

    multi_space = [cell.strip() for cell in re.split(r"\s{2,}", stripped) if cell.strip()]
    if len(multi_space) >= 2:
        return multi_space

    return [cell.strip() for cell in stripped.split() if cell.strip()]


def parse_story_rows(raw_text: str, base_elevation: float | None = None) -> dict[str, Any]:
    if not raw_text or not raw_text.strip():
        raise AppError("Paste at least one story row before parsing.")

    parsed_rows: list[dict[str, Any]] = []
    warnings: list[str] = []
    seen_names: set[str] = set()
    header_map: dict[str, int] | None = None
    resolved_base_elevation = base_elevation

    for line_number, line in enumerate(raw_text.splitlines(), start=1):
        stripped = line.strip()
        if not stripped:
            continue

        cells = split_story_line(stripped)
        if not cells:
            continue

        if is_base_elevation_line(cells):
            resolved_base_elevation = parse_base_elevation_line(cells, line_number)
            continue

        if looks_like_header(cells):
            header_map = get_header_map(cells)
            continue

        name, elevation = parse_story_line(cells, line_number, header_map)

        normalized_key = name.casefold()
        if normalized_key in seen_names:
            raise AppError(f"Duplicate story name '{name}' was found in the pasted rows.")
        seen_names.add(normalized_key)

        parsed_rows.append(
            {
                "lineNumber": line_number,
                "name": name,
                "elevation": elevation,
            }
        )

    if not parsed_rows:
        raise AppError("No story rows were parsed from the pasted text.")

    sorted_rows = sorted(parsed_rows, key=lambda row: (row["elevation"], row["lineNumber"]))
    order_changed = [row["name"] for row in parsed_rows] != [row["name"] for row in sorted_rows]

    for index in range(1, len(sorted_rows)):
        if sorted_rows[index]["elevation"] <= sorted_rows[index - 1]["elevation"]:
            raise AppError("Story elevations must be strictly increasing after sorting.")

    resolved_base = resolved_base_elevation
    preview_rows: list[dict[str, Any]] = []
    if resolved_base is not None:
        first_height = sorted_rows[0]["elevation"] - resolved_base
        if first_height <= 0:
            raise AppError(
                f"The first story elevation {sorted_rows[0]['elevation']} must be above the base elevation {resolved_base}."
            )

    for index, row in enumerate(sorted_rows):
        height = None
        if resolved_base is not None:
            if index == 0:
                height = row["elevation"] - resolved_base
            else:
                height = row["elevation"] - sorted_rows[index - 1]["elevation"]

            if height <= 0:
                raise AppError(f"Computed story height for '{row['name']}' is not positive.")

        preview_rows.append(
            {
                "index": index,
                "name": row["name"],
                "elevation": row["elevation"],
                "height": height,
                "sourceLine": row["lineNumber"],
            }
        )

    if order_changed:
        warnings.append("Rows were reordered from lowest to highest elevation before preview.")

    return {
        "storyCount": len(preview_rows),
        "lengthUnit": "ft",
        "baseElevation": resolved_base,
        "orderChanged": order_changed,
        "warnings": warnings,
        "stories": preview_rows,
    }


def build_static_asset(request_path: str) -> StaticAsset:
    relative_path = request_path.lstrip("/") or "index.html"
    asset_path = (APP_ROOT / relative_path).resolve()
    if APP_ROOT not in asset_path.parents and asset_path != APP_ROOT:
        raise AppError("Invalid asset path.")
    if not asset_path.exists() or not asset_path.is_file():
        raise FileNotFoundError(str(asset_path))

    content_type, _ = mimetypes.guess_type(asset_path.name)
    return StaticAsset(path=asset_path, content_type=content_type or "application/octet-stream")


class ControlCenterHandler(BaseHTTPRequestHandler):
    server_version = "EtabsControlCenter/0.1"

    def do_GET(self) -> None:
        try:
            if self.path.startswith("/api/stories/current"):
                payload = run_powershell_script("get-etabs-stories.ps1", ["-LengthUnit", DISPLAY_LENGTH_UNIT])
                self.respond_json(payload)
                return

            asset = build_static_asset("/index.html" if self.path == "/" else self.path)
            self.respond_file(asset)
        except FileNotFoundError:
            self.respond_error(HTTPStatus.NOT_FOUND, "Not found.")
        except AppError as exc:
            self.respond_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self.respond_json({"error": str(exc)}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_POST(self) -> None:
        try:
            payload = self.read_json()
            if self.path == "/api/stories/parse":
                text = str(payload.get("text", ""))
                base_elevation = payload.get("baseElevation")
                result = parse_story_rows(
                    text,
                    base_elevation=None if base_elevation in (None, "") else float(base_elevation),
                )
                self.respond_json(result)
                return

            if self.path == "/api/stories/apply":
                stories = payload.get("stories")
                if not isinstance(stories, list) or not stories:
                    raise AppError("The apply request must include at least one parsed story.")

                script_payload = {
                    "BaseElevation": payload.get("baseElevation"),
                    "Stories": [{"Name": row["name"], "Elevation": row["elevation"]} for row in stories],
                }

                temp_path = None
                try:
                    with tempfile.NamedTemporaryFile(
                        mode="w",
                        encoding="utf-8",
                        suffix=".json",
                        delete=False,
                    ) as handle:
                        json.dump(script_payload, handle)
                        temp_path = handle.name

                    args = ["-StoryJsonPath", temp_path, "-LengthUnit", DISPLAY_LENGTH_UNIT]
                    if payload.get("save", True):
                        args.append("-Save")
                    if payload.get("unlockIfLocked", True):
                        args.append("-UnlockIfLocked")
                    if payload.get("skipBackup", False):
                        args.append("-SkipBackup")
                    if payload.get("baseElevation") not in (None, ""):
                        args.extend(["-BaseElevation", str(payload["baseElevation"])])

                    result = run_powershell_script("set-etabs-stories.ps1", args)
                finally:
                    if temp_path and os.path.exists(temp_path):
                        os.remove(temp_path)

                self.respond_json(result)
                return

            self.respond_error(HTTPStatus.NOT_FOUND, "Unknown API route.")
        except AppError as exc:
            self.respond_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self.respond_json({"error": str(exc)}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def log_message(self, format: str, *args: Any) -> None:
        return

    def read_json(self) -> dict[str, Any]:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            return {}
        raw = self.rfile.read(content_length).decode("utf-8")
        if not raw:
            return {}
        return json.loads(raw)

    def respond_file(self, asset: StaticAsset) -> None:
        data = asset.path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", asset.content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def respond_json(self, payload: Any, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def respond_error(self, status: HTTPStatus, message: str) -> None:
        self.respond_json({"error": message}, status=status)


def main() -> None:
    parser = argparse.ArgumentParser(description="Serve the local ETABS control center.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--open-browser", action="store_true")
    args = parser.parse_args()

    if not APP_ROOT.exists():
        raise SystemExit(f"App directory not found: {APP_ROOT}")

    server = ThreadingHTTPServer((args.host, args.port), ControlCenterHandler)
    url = f"http://{args.host}:{args.port}/"
    print(f"ETABS Control Center running at {url}")
    if args.open_browser:
        webbrowser.open(url)
    server.serve_forever()


if __name__ == "__main__":
    main()
