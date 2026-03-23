import csv
import os
import re
import zipfile
from collections import defaultdict
from datetime import datetime, timezone
from xml.sax.saxutils import escape


DEFAULT_INPUT = r"C:\Users\gkafka\Documents\358-flatbush-wall-smax-envelope.csv"
DEFAULT_OUTPUT = r"C:\Users\gkafka\Documents\358-flatbush-wall-fr-report.xlsx"
DEFAULT_THRESHOLD_PSI = 520.0


def column_name(index):
    result = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def sanitize_sheet_name(name):
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", "_", name)
    return cleaned[:31] or "Sheet"


def story_sort_key(story_name):
    lower_name = story_name.lower()
    if lower_name == "cellar":
        return (0, 0)
    if lower_name.startswith("story"):
        suffix = lower_name.replace("story", "", 1)
        if suffix.isdigit():
            return (1, int(suffix))
    if lower_name == "roof":
        return (2, 0)
    return (1, 9999)


def wall_sort_key(label):
    match = re.match(r"([A-Za-z]+)(\d+)$", label)
    if match:
        return (match.group(1), int(match.group(2)))
    return (label, 0)


def xml_cell(cell_ref, value, style_id):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{cell_ref}" s="{style_id}"><v>{value}</v></c>'

    text = "" if value is None else str(value)
    return (
        f'<c r="{cell_ref}" t="inlineStr" s="{style_id}">'
        f"<is><t>{escape(text)}</t></is></c>"
    )


def worksheet_xml(headers, rows):
    widths = []
    for column_index, header in enumerate(headers):
        max_length = len(str(header))
        for row in rows:
            value = row[column_index]
            if isinstance(value, float):
                text = f"{value:.4f}"
            else:
                text = str(value)
            max_length = max(max_length, len(text))
        widths.append(min(max_length + 2, 40))

    column_xml = "".join(
        f'<col min="{index}" max="{index}" width="{widths[index - 1]}" customWidth="1"/>'
        for index in range(1, len(headers) + 1)
    )

    xml_rows = []
    header_cells = []
    for column_index, header in enumerate(headers, start=1):
        header_cells.append(xml_cell(f"{column_name(column_index)}1", header, 1))
    xml_rows.append(f'<row r="1">{"".join(header_cells)}</row>')

    for row_index, row in enumerate(rows, start=2):
        fail = bool(row[-1])
        style_id = 2 if fail else 0
        cells = []
        for column_index, value in enumerate(row[:-1], start=1):
            cells.append(xml_cell(f"{column_name(column_index)}{row_index}", value, style_id))
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<cols>{column_xml}</cols>"
        f"<sheetData>{''.join(xml_rows)}</sheetData>"
        "</worksheet>"
    )


def workbook_xml(sheet_names):
    sheets_xml = []
    for index, name in enumerate(sheet_names, start=1):
        sheets_xml.append(
            f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{''.join(sheets_xml)}</sheets>"
        "</workbook>"
    )


def workbook_rels_xml(sheet_count):
    relationships = []
    for index in range(1, sheet_count + 1):
        relationships.append(
            '<Relationship '
            f'Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )

    relationships.append(
        '<Relationship '
        f'Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{''.join(relationships)}"
        "</Relationships>"
    )


def content_types_xml(sheet_count):
    overrides = [
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{''.join(overrides)}"
        "</Types>"
    )


def root_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def styles_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        "</fonts>"
        '<fills count="4">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>'
        "</fills>"
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>'
        "</cellXfs>"
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )


def core_xml():
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>Codex</dc:creator>"
        "<cp:lastModifiedBy>Codex</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def app_xml(sheet_names):
    titles = "".join(f"<vt:lpstr>{escape(name)}</vt:lpstr>" for name in sheet_names)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>Codex</Application>"
        f"<TitlesOfParts><vt:vector size=\"{len(sheet_names)}\" baseType=\"lpstr\">{titles}</vt:vector></TitlesOfParts>"
        f"<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\">"
        "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>"
        f"<vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>"
        "</vt:vector></HeadingPairs>"
        "</Properties>"
    )


def load_rows(input_path):
    with open(input_path, newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


def collapse_to_story_max(rows):
    grouped = {}
    for row in rows:
        key = (row["Label"], row["Story"])
        governing_smax = float(row["GoverningSMax"])
        if key not in grouped or governing_smax > float(grouped[key]["GoverningSMax"]):
            grouped[key] = row
    return grouped


def build_sheet_rows(story_rows, threshold_psi):
    sheet_rows = []
    for row in story_rows:
        smax_kip_in2 = float(row["GoverningSMax"])
        smax_psi = smax_kip_in2 * 1000.0
        fail = smax_psi > threshold_psi
        sheet_rows.append([
            row["Story"],
            row["ObjectName"],
            row["ElementName"],
            row["PointElement"],
            row["GoverningSurface"],
            round(smax_kip_in2, 6),
            round(smax_psi, 1),
            threshold_psi,
            "FAIL" if fail else "PASS",
            row["GoverningOutputName"],
            row["GoverningStepType"],
            float(row["GoverningStepNumber"]),
            fail,
        ])
    return sheet_rows


def build_summary_rows(grouped_rows, threshold_psi):
    by_wall = defaultdict(list)
    for (label, _story), row in grouped_rows.items():
        by_wall[label].append(row)

    summary_rows = []
    for label in sorted(by_wall, key=wall_sort_key):
        wall_rows = by_wall[label]
        worst_row = max(wall_rows, key=lambda item: float(item["GoverningSMax"]))
        fail_count = sum((float(item["GoverningSMax"]) * 1000.0) > threshold_psi for item in wall_rows)
        worst_psi = float(worst_row["GoverningSMax"]) * 1000.0
        fail = fail_count > 0
        summary_rows.append([
            label,
            len(wall_rows),
            worst_row["Story"],
            round(float(worst_row["GoverningSMax"]), 6),
            round(worst_psi, 1),
            threshold_psi,
            fail_count,
            "FAIL" if fail else "PASS",
            worst_row["GoverningOutputName"],
            fail,
        ])
    return summary_rows


def create_workbook(input_path, output_path, threshold_psi):
    rows = load_rows(input_path)
    grouped_rows = collapse_to_story_max(rows)

    wall_labels = sorted({label for label, _story in grouped_rows}, key=wall_sort_key)
    sheet_definitions = []

    summary_headers = [
        "Wall",
        "FloorCount",
        "WorstStory",
        "WorstSMax_kip_in2",
        "WorstSMax_psi",
        "fr_psi",
        "FailFloorCount",
        "Status",
        "WorstCombo",
    ]
    summary_rows = build_summary_rows(grouped_rows, threshold_psi)
    sheet_definitions.append(("Summary", summary_headers, summary_rows))

    wall_headers = [
        "Story",
        "ObjectName",
        "ElementName",
        "PointElement",
        "GoverningSurface",
        "SMax_kip_in2",
        "SMax_psi",
        "fr_psi",
        "Status",
        "GoverningCombo",
        "GoverningStepType",
        "GoverningStepNumber",
    ]

    for label in wall_labels:
        story_rows = [
            grouped_rows[(candidate_label, story)]
            for candidate_label, story in grouped_rows
            if candidate_label == label
        ]
        story_rows.sort(key=lambda item: story_sort_key(item["Story"]))
        sheet_definitions.append((label, wall_headers, build_sheet_rows(story_rows, threshold_psi)))

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        sheet_names = [sanitize_sheet_name(name) for name, _headers, _rows in sheet_definitions]

        archive.writestr("[Content_Types].xml", content_types_xml(len(sheet_names)))
        archive.writestr("_rels/.rels", root_rels_xml())
        archive.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheet_names)))
        archive.writestr("xl/styles.xml", styles_xml())
        archive.writestr("docProps/core.xml", core_xml())
        archive.writestr("docProps/app.xml", app_xml(sheet_names))

        for sheet_index, (_name, headers, sheet_rows) in enumerate(sheet_definitions, start=1):
            archive.writestr(
                f"xl/worksheets/sheet{sheet_index}.xml",
                worksheet_xml(headers, sheet_rows),
            )


def main():
    input_path = os.environ.get("WALL_FR_INPUT", DEFAULT_INPUT)
    output_path = os.environ.get("WALL_FR_OUTPUT", DEFAULT_OUTPUT)
    threshold_psi = float(os.environ.get("WALL_FR_THRESHOLD_PSI", DEFAULT_THRESHOLD_PSI))

    create_workbook(input_path, output_path, threshold_psi)
    print(output_path)


if __name__ == "__main__":
    main()
