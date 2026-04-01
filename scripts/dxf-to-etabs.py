"""DXF-to-ETABS import debugger.

Catches what ETABS will choke on or silently corrupt before you import.
Every check maps to a specific ETABS import behavior:

  - LWPOLYLINE entities   → ETABS silently ignores (the #1 gotcha)
  - Duplicate entities    → doubled stiffness (ETABS imports both)
  - Unexploded blocks     → ETABS errors on INSERT entities
  - Stray Z coordinates   → joint merge failures, wrong elevations
  - Origin distance       → float precision loss breaks connectivity
  - Open polylines        → ETABS reads as lines, not area elements
  - Layer 0 / Defpoints   → ETABS silently drops these on import
  - Near-miss endpoints   → disconnected joints, load path breaks
  - Unsupported entities  → SPLINE, ELLIPSE, etc. vanish on import
  - Zero-length lines     → meshing failures, division by zero
  - Fragmented curves     → dozens of tiny beam elements per curve

Usage
-----
    python dxf-to-etabs.py inspect  plan.dxf
    python dxf-to-etabs.py validate plan.dxf
    python dxf-to-etabs.py classify plan.dxf
    python dxf-to-etabs.py floors   plan.dxf
    python dxf-to-etabs.py align    plan.dxf -o clouded.dxf
    python dxf-to-etabs.py clean    plan.dxf clean.dxf --all
    python dxf-to-etabs.py split    plan.dxf output_dir/

Dependencies: pip install ezdxf
"""

import argparse
import json
import math
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import ezdxf
from ezdxf.entities import DXFGraphic


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SKIP_LAYERS = {"defpoints"}  # layer "0" handled separately
AREA_LAYER_HINTS = {"wall", "slab", "floor", "boundary", "area", "core"}
FLOOR_LABEL_RE = [
    re.compile(
        r"(?i)\b(level|floor|flr|lev|l)\s*[-:]?\s*"
        r"(\d+|[a-z]+|roof|ground|grd|basement|bsmt|cellar|mech|mezzanine|mezz|penthouse|ph)\b"
    ),
    re.compile(r"(?i)\b(\d+)\s*(st|nd|rd|th)\s*(floor|flr|fl|level)\b"),
    # "1ST FL", "2ND FL", "ROOF FL", "BULKHEAD", etc.
    re.compile(r"(?i)\b(roof|bulkhead|penthouse|ph|mech|mechanical|cellar|basement)\s*(fl|floor|flr|level)?\b"),
    re.compile(r"(?i)\b\d+(st|nd|rd|th)\s*(fl|floor|flr)\b"),
]


# ---------------------------------------------------------------------------
# DXF loading helpers
# ---------------------------------------------------------------------------

def load_dxf(path):
    """Read a DXF file, falling back to recovery mode for corrupt files."""
    try:
        return ezdxf.readfile(path)
    except ezdxf.DXFStructureError:
        try:
            doc, auditor = ezdxf.recover.readfile(path)
            return doc
        except Exception as exc:
            raise SystemExit(f"Cannot read {path}: {exc}")


def iter_model_entities(doc, layers=None):
    """Yield modelspace entities, optionally filtered to specific layers."""
    msp = doc.modelspace()
    for entity in msp:
        if entity.dxftype() == "VIEWPORT":
            continue
        if layers is not None:
            if entity.dxf.layer not in layers:
                continue
        yield entity


# ---------------------------------------------------------------------------
# Entity geometry utilities
# ---------------------------------------------------------------------------

def entity_endpoints(entity):
    """Return representative (x, y, z) points for any entity type."""
    dxftype = entity.dxftype()
    if dxftype == "LINE":
        s = entity.dxf.start
        e = entity.dxf.end
        return [(s.x, s.y, s.z), (e.x, e.y, e.z)]
    if dxftype in ("LWPOLYLINE",):
        pts = []
        for x, y, *rest in entity.get_points(format="xyz"):
            z = rest[0] if rest else 0.0
            pts.append((x, y, z))
        return pts
    if dxftype == "POLYLINE":
        return [(v.dxf.location.x, v.dxf.location.y, v.dxf.location.z)
                for v in entity.vertices]
    if dxftype == "CIRCLE":
        c = entity.dxf.center
        return [(c.x, c.y, c.z)]
    if dxftype == "ARC":
        c = entity.dxf.center
        return [(c.x, c.y, c.z)]
    if dxftype == "INSERT":
        p = entity.dxf.insert
        return [(p.x, p.y, p.z)]
    if dxftype in ("TEXT", "MTEXT"):
        p = entity.dxf.insert
        return [(p.x, p.y, p.z)]
    if dxftype == "POINT":
        p = entity.dxf.location
        return [(p.x, p.y, p.z)]
    if dxftype == "3DFACE":
        pts = []
        for attr in ("vtx0", "vtx1", "vtx2", "vtx3"):
            if entity.dxf.hasattr(attr):
                v = getattr(entity.dxf, attr)
                pts.append((v.x, v.y, v.z))
        return pts
    # Fallback: try to get any point data
    return []


def entity_bbox(entity):
    """Return (xmin, ymin, xmax, ymax) or None."""
    pts = entity_endpoints(entity)
    if not pts:
        return None
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))


def entity_fingerprint(entity, tolerance):
    """Hash-key for duplicate detection: type + rounded coords + layer."""
    dxftype = entity.dxftype()
    layer = entity.dxf.layer
    pts = entity_endpoints(entity)
    if not pts:
        return None
    factor = 1.0 / tolerance if tolerance else 1.0
    rounded = tuple(
        (round(p[0] * factor), round(p[1] * factor), round(p[2] * factor))
        for p in sorted(pts)
    )
    return (dxftype, layer, rounded)


def entity_text(entity):
    """Extract text content from TEXT or MTEXT."""
    dxftype = entity.dxftype()
    if dxftype == "TEXT":
        return entity.dxf.text
    if dxftype == "MTEXT":
        return entity.text  # MTEXT uses .text property for plain content
    return ""


# ---------------------------------------------------------------------------
# Spatial index
# ---------------------------------------------------------------------------

def build_point_grid(points, cell_size):
    """points: list of ((x,y,z), entity_handle). Returns grid dict."""
    grid = defaultdict(list)
    for pt, handle in points:
        gx = int(pt[0] // cell_size) if cell_size else 0
        gy = int(pt[1] // cell_size) if cell_size else 0
        grid[(gx, gy)].append((pt, handle))
    return grid


def query_nearby(grid, point, cell_size, radius):
    """Return all (point, handle) within radius of point."""
    gx = int(point[0] // cell_size) if cell_size else 0
    gy = int(point[1] // cell_size) if cell_size else 0
    results = []
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for pt, handle in grid.get((gx + dx, gy + dy), []):
                d = math.dist((point[0], point[1]), (pt[0], pt[1]))
                if d <= radius:
                    results.append((pt, handle, d))
    return results


# ---------------------------------------------------------------------------
# Validation checks
# ---------------------------------------------------------------------------

def issue(check, severity, entity, message, location=None, **details):
    """Build a single issue dict."""
    loc = location
    if loc is None:
        pts = entity_endpoints(entity)
        loc = list(pts[0]) if pts else None
    return {
        "check": check,
        "severity": severity,
        "entity_type": entity.dxftype(),
        "handle": entity.dxf.handle,
        "layer": entity.dxf.layer,
        "location": loc,
        "message": message,
        "details": details,
    }


def check_duplicates(entities, tolerance, doc):
    """Find entities with identical geometry (within tolerance)."""
    issues = []
    seen = {}  # fingerprint -> first entity
    for ent in entities:
        fp = entity_fingerprint(ent, tolerance)
        if fp is None:
            continue
        if fp in seen:
            first = seen[fp]
            issues.append(issue(
                "duplicate_entity", "error", ent,
                f"Duplicate {ent.dxftype()} overlaps handle {first.dxf.handle}",
                other_handle=first.dxf.handle,
            ))
        else:
            seen[fp] = ent
    return issues


def check_near_miss_endpoints(entities, tolerance, doc):
    """Find endpoints that almost meet but don't (disconnected joints)."""
    issues = []
    all_pts = []
    ent_map = {}
    for ent in entities:
        dxftype = ent.dxftype()
        if dxftype not in ("LINE", "LWPOLYLINE", "POLYLINE", "ARC"):
            continue
        pts = entity_endpoints(ent)
        handle = ent.dxf.handle
        ent_map[handle] = ent
        # Use start/end only for near-miss check
        if pts:
            all_pts.append((pts[0], handle))
            if len(pts) > 1:
                all_pts.append((pts[-1], handle))

    if not all_pts:
        return issues

    grid = build_point_grid(all_pts, tolerance * 4)
    checked = set()

    for pt, handle in all_pts:
        nearby = query_nearby(grid, pt, tolerance * 4, tolerance)
        for npt, nhandle, dist in nearby:
            if nhandle == handle:
                continue
            pair = tuple(sorted((handle, nhandle)))
            if pair in checked:
                continue
            checked.add(pair)
            if 0 < dist <= tolerance:
                ent = ent_map[handle]
                issues.append(issue(
                    "near_miss_endpoint", "warning", ent,
                    f"Endpoint within {dist:.4f} of entity {nhandle} "
                    f"(tolerance {tolerance})",
                    location=list(pt),
                    other_handle=nhandle,
                    distance=round(dist, 6),
                ))
    return issues


def check_stray_z(entities, tolerance, doc):
    """Find entities with non-zero Z coordinates."""
    issues = []
    z_tol = tolerance * 0.1 if tolerance else 0.001
    for ent in entities:
        pts = entity_endpoints(ent)
        for pt in pts:
            if abs(pt[2]) > z_tol:
                issues.append(issue(
                    "stray_z_coordinate", "error", ent,
                    f"Z = {pt[2]:.6f} (expected 0)",
                    location=list(pt),
                    z_value=pt[2],
                ))
                break  # one issue per entity
    return issues


def check_open_polylines(entities, tolerance, doc):
    """Find polylines that should be closed but aren't."""
    issues = []
    for ent in entities:
        dxftype = ent.dxftype()
        if dxftype not in ("LWPOLYLINE", "POLYLINE"):
            continue
        if ent.is_closed:
            continue
        layer_lower = ent.dxf.layer.lower()
        hits_area_hint = any(h in layer_lower for h in AREA_LAYER_HINTS)
        severity = "error" if hits_area_hint else "warning"
        issues.append(issue(
            "open_polyline", severity, ent,
            f"Open polyline on layer '{ent.dxf.layer}'"
            + (" (area-element layer)" if hits_area_hint else ""),
            layer_hint=hits_area_hint,
        ))
    return issues


def check_layer_zero(entities, tolerance, doc):
    """Flag geometry on Layer 0 or Defpoints (ETABS silently skips these)."""
    issues = []
    for ent in entities:
        layer = ent.dxf.layer
        if layer == "0" or layer.lower() in SKIP_LAYERS:
            issues.append(issue(
                "layer_zero", "warning", ent,
                f"Entity on layer '{layer}' — ETABS will skip this on import",
            ))
    return issues


def check_unexploded_blocks(entities, tolerance, doc):
    """Flag INSERT (block reference) entities.

    ETABS returns an import error on INSERT entities. All blocks must be
    exploded to primitive geometry (LINE, ARC, 3DFACE) before import.
    """
    issues = []
    for ent in entities:
        if ent.dxftype() != "INSERT":
            continue
        block_name = ent.dxf.name
        issues.append(issue(
            "unexploded_block", "error", ent,
            f"Block '{block_name}' — ETABS errors on block references. "
            f"Must explode to LINE/ARC/3DFACE before import",
            block_name=block_name,
        ))
    return issues


def check_zero_length(entities, tolerance, doc):
    """Find zero-length line entities."""
    issues = []
    tol = tolerance if tolerance else 0.001
    for ent in entities:
        if ent.dxftype() != "LINE":
            continue
        s = ent.dxf.start
        e = ent.dxf.end
        length = math.dist((s.x, s.y), (e.x, e.y))
        if length < tol:
            issues.append(issue(
                "zero_length_entity", "error", ent,
                f"Zero-length LINE ({length:.6f})",
                length=round(length, 6),
            ))
    return issues


def check_fragmented_curves(entities, tolerance, doc):
    """Find chains of very short connected line segments (curve approximations)."""
    issues = []
    short_threshold = tolerance * 3 if tolerance else 0.75

    lines = []
    for ent in entities:
        if ent.dxftype() != "LINE":
            continue
        s = ent.dxf.start
        e = ent.dxf.end
        length = math.dist((s.x, s.y), (e.x, e.y))
        if length < short_threshold and length > 0:
            lines.append(ent)

    if len(lines) < 5:
        return issues

    # Build endpoint connectivity
    pt_to_ents = defaultdict(list)
    snap = tolerance * 0.5 if tolerance else 0.01
    for ent in lines:
        s = ent.dxf.start
        e = ent.dxf.end
        sk = (round(s.x / snap) * snap, round(s.y / snap) * snap)
        ek = (round(e.x / snap) * snap, round(e.y / snap) * snap)
        pt_to_ents[sk].append(ent)
        pt_to_ents[ek].append(ent)

    # Find chains of 5+ connected short segments
    visited = set()
    for ent in lines:
        if ent.dxf.handle in visited:
            continue
        chain = []
        stack = [ent]
        while stack:
            current = stack.pop()
            if current.dxf.handle in visited:
                continue
            visited.add(current.dxf.handle)
            chain.append(current)
            s = current.dxf.start
            e = current.dxf.end
            for pt in (s, e):
                pk = (round(pt.x / snap) * snap, round(pt.y / snap) * snap)
                for neighbor in pt_to_ents.get(pk, []):
                    if neighbor.dxf.handle not in visited:
                        stack.append(neighbor)

        if len(chain) >= 5:
            first = chain[0]
            issues.append(issue(
                "fragmented_curve", "info", first,
                f"Chain of {len(chain)} short segments "
                f"(likely curve approximation)",
                chain_length=len(chain),
                handles=[e.dxf.handle for e in chain[:10]],
            ))
    return issues


def check_lwpolyline(entities, tolerance, doc):
    """ETABS silently ignores LWPOLYLINE entities.

    Modern AutoCAD creates LWPOLYLINE by default. ETABS only reads classic
    POLYLINE, LINE, ARC, and 3DFACE. Any LWPOLYLINE in the file will vanish
    on import with no warning — walls, slabs, and column outlines disappear.
    """
    issues = []
    for ent in entities:
        if ent.dxftype() != "LWPOLYLINE":
            continue
        layer = ent.dxf.layer
        pts = list(ent.get_points(format="xy"))
        closed = ent.is_closed
        desc = "closed" if closed else "open"
        # Closed LWPOLYLINE on structural layers = lost area element
        layer_lower = layer.lower()
        is_structural = any(h in layer_lower for h in
                           AREA_LAYER_HINTS | {"beam", "col", "brace", "etabs"})
        severity = "error" if is_structural else "warning"
        issues.append(issue(
            "lwpolyline_ignored", severity, ent,
            f"LWPOLYLINE ({desc}, {len(pts)} verts) on '{layer}' — "
            f"ETABS silently skips LWPOLYLINE. Must explode to LINEs "
            f"or convert to classic POLYLINE/3DFACE",
            vertex_count=len(pts),
            closed=closed,
        ))
    return issues


def check_origin_distance(entities, tolerance, doc):
    """Flag geometry far from origin — floating-point precision loss.

    CSI docs: ratio of max coordinate to model dimension should be < 10,000.
    Large absolute coords (survey coordinates) cause joints that should be
    coincident to drift apart, breaking connectivity after import.
    """
    issues = []
    all_pts = []
    for ent in entities:
        all_pts.extend(entity_endpoints(ent))
    if not all_pts:
        return issues

    xs = [abs(p[0]) for p in all_pts]
    ys = [abs(p[1]) for p in all_pts]
    max_coord = max(max(xs), max(ys))

    # Model extent
    x_range = max(p[0] for p in all_pts) - min(p[0] for p in all_pts)
    y_range = max(p[1] for p in all_pts) - min(p[1] for p in all_pts)
    model_dim = max(x_range, y_range, 1.0)

    ratio = max_coord / model_dim
    if ratio > 5000:
        # Create a synthetic issue — no single entity is the problem
        severity = "error" if ratio > 10000 else "warning"
        # Use the first entity as the issue anchor
        anchor = entities[0] if entities else None
        if anchor:
            issues.append(issue(
                "origin_distance", severity, anchor,
                f"Geometry far from origin: max coord = {max_coord:.0f}, "
                f"model extent = {model_dim:.0f}, ratio = {ratio:.0f}. "
                f"CSI limit is 10,000. Move geometry to origin before import "
                f"to prevent floating-point joint merge failures",
                max_coordinate=round(max_coord, 2),
                model_dimension=round(model_dim, 2),
                ratio=round(ratio, 1),
            ))
    return issues


def check_unsupported_entities(entities, tolerance, doc):
    """Flag entity types that ETABS cannot import.

    ETABS reads: LINE, ARC, POLYLINE (classic), 3DFACE, CIRCLE (limited).
    Everything else is silently skipped. If structural geometry uses
    SPLINE, ELLIPSE, HATCH, etc., it will vanish on import.
    """
    ETABS_SUPPORTED = {"LINE", "ARC", "CIRCLE", "POLYLINE", "3DFACE", "SOLID"}
    ETABS_ANNOTATION = {"TEXT", "MTEXT", "DIMENSION", "LEADER", "MLEADER",
                        "TABLE", "VIEWPORT", "HATCH", "IMAGE", "WIPEOUT"}
    issues = []
    for ent in entities:
        dxftype = ent.dxftype()
        if dxftype in ETABS_SUPPORTED or dxftype == "LWPOLYLINE":
            continue  # LWPOLYLINE handled by its own check
        if dxftype in ETABS_ANNOTATION:
            continue  # Annotation is expected to be skipped
        if dxftype == "INSERT":
            continue  # Handled by unexploded_blocks check
        # Anything else is unsupported structural geometry
        layer_lower = ent.dxf.layer.lower()
        is_structural = any(h in layer_lower for h in
                           AREA_LAYER_HINTS | {"beam", "col", "brace", "etabs"})
        if is_structural:
            issues.append(issue(
                "unsupported_entity", "error", ent,
                f"{dxftype} on structural layer '{ent.dxf.layer}' — "
                f"ETABS cannot import {dxftype}. Convert to LINE/ARC/3DFACE",
            ))
        else:
            issues.append(issue(
                "unsupported_entity", "info", ent,
                f"{dxftype} on '{ent.dxf.layer}' — ETABS will skip this",
            ))
    return issues


ALL_CHECKS = [
    check_duplicates,
    check_near_miss_endpoints,
    check_stray_z,
    check_lwpolyline,
    check_open_polylines,
    check_layer_zero,
    check_unexploded_blocks,
    check_unsupported_entities,
    check_origin_distance,
    check_zero_length,
    check_fragmented_curves,
]


# ---------------------------------------------------------------------------
# Report formatting
# ---------------------------------------------------------------------------

def build_report(filepath, entities_count, all_issues, tolerance):
    """Build the full validation report dict."""
    by_check = defaultdict(list)
    for iss in all_issues:
        by_check[iss["check"]].append(iss)

    summary = {"total_entities": entities_count, "error": 0, "warning": 0, "info": 0}
    checks = {}
    for check_name, check_issues in by_check.items():
        # A single check can produce mixed severities (e.g. open_polyline)
        sev_counts = defaultdict(int)
        for ci in check_issues:
            sev_counts[ci["severity"]] += 1
            summary[ci["severity"]] += 1
        worst = "error" if sev_counts.get("error") else (
            "warning" if sev_counts.get("warning") else "info")
        checks[check_name] = {"count": len(check_issues), "severity": worst}

    return {
        "file": str(filepath),
        "timestamp": datetime.now(timezone.utc).replace(microsecond=0)
                     .isoformat().replace("+00:00", "Z"),
        "tolerance": tolerance,
        "summary": summary,
        "checks": checks,
        "issues": all_issues,
    }


def format_report_text(report):
    """Human-readable summary."""
    lines = []
    s = report["summary"]
    lines.append(f"dxf-to-etabs: {report['file']}")
    lines.append(f"  {s['total_entities']:,} entities  "
                 f"tolerance={report['tolerance']}")
    lines.append("")

    if s["error"]:
        lines.append(f"  ERRORS ({s['error']})")
        for name, info in sorted(report["checks"].items()):
            if info["severity"] == "error":
                lines.append(f"    {name:30s} {info['count']}")
        lines.append("")

    if s["warning"]:
        lines.append(f"  WARNINGS ({s['warning']})")
        for name, info in sorted(report["checks"].items()):
            if info["severity"] == "warning":
                lines.append(f"    {name:30s} {info['count']}")
        lines.append("")

    if s["info"]:
        lines.append(f"  INFO ({s['info']})")
        for name, info in sorted(report["checks"].items()):
            if info["severity"] == "info":
                lines.append(f"    {name:30s} {info['count']}")
        lines.append("")

    if not any((s["error"], s["warning"], s["info"])):
        lines.append("  No issues found.")
        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Subcommand: inspect
# ---------------------------------------------------------------------------

def cmd_inspect(args):
    """Full structural survey of a DXF file — layers, entity types, blocks,
    text labels, coordinate extents.  Designed for Claude to read before
    running validate or floors."""
    filepath = args.path
    doc = load_dxf(filepath)
    msp = doc.modelspace()
    entities = list(msp)

    # Layer breakdown
    layer_counts = defaultdict(lambda: defaultdict(int))
    for ent in entities:
        layer_counts[ent.dxf.layer][ent.dxftype()] += 1

    layers = {}
    for layer_name, type_counts in sorted(layer_counts.items()):
        layers[layer_name] = {
            "total": sum(type_counts.values()),
            "types": dict(type_counts),
        }

    # Block definitions
    block_defs = {}
    for block in doc.blocks:
        name = block.name
        if name.startswith("*"):
            continue
        ents = list(block)
        if not ents:
            continue
        type_counts = defaultdict(int)
        for e in ents:
            type_counts[e.dxftype()] += 1
        block_defs[name] = {
            "entity_count": len(ents),
            "types": dict(type_counts),
        }

    # Block insertions (which blocks are used, how many times, at what locations)
    block_usage = defaultdict(list)
    for ent in entities:
        if ent.dxftype() == "INSERT":
            p = ent.dxf.insert
            block_usage[ent.dxf.name].append({
                "layer": ent.dxf.layer,
                "location": [round(p.x, 2), round(p.y, 2), round(p.z, 2)],
            })

    # Text entities
    texts = []
    for ent in entities:
        if ent.dxftype() in ("TEXT", "MTEXT"):
            text = entity_text(ent)
            pts = entity_endpoints(ent)
            if pts:
                texts.append({
                    "text": text.strip(),
                    "layer": ent.dxf.layer,
                    "location": [round(pts[0][0], 2), round(pts[0][1], 2)],
                })

    # Coordinate extents
    all_pts = []
    for ent in entities:
        all_pts.extend(entity_endpoints(ent))
    extents = None
    if all_pts:
        xs = [p[0] for p in all_pts]
        ys = [p[1] for p in all_pts]
        zs = [p[2] for p in all_pts]
        extents = {
            "x": [round(min(xs), 2), round(max(xs), 2)],
            "y": [round(min(ys), 2), round(max(ys), 2)],
            "z": [round(min(zs), 2), round(max(zs), 2)],
        }

    result = {
        "file": str(filepath),
        "total_entities": len(entities),
        "layers": layers,
        "block_definitions": block_defs,
        "block_insertions": {k: {"count": len(v), "instances": v}
                             for k, v in block_usage.items()},
        "text_entities": texts,
        "extents": extents,
    }

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"inspect: {filepath}")
        print(f"  {len(entities)} entities across {len(layers)} layers")
        print()
        print("  LAYERS")
        for name, info in sorted(layers.items(), key=lambda x: -x[1]["total"]):
            types_str = ", ".join(f"{t}:{c}" for t, c in info["types"].items())
            print(f"    {name:25s}  {info['total']:4d}  ({types_str})")
        print()
        print("  BLOCKS USED")
        for name, instances in sorted(block_usage.items(), key=lambda x: -len(x[1])):
            print(f"    {name:25s}  {len(instances)}x")
        if texts:
            print()
            print("  TEXT LABELS")
            for t in texts:
                print(f"    \"{t['text']}\"  at ({t['location'][0]:.0f}, "
                      f"{t['location'][1]:.0f})  [{t['layer']}]")
        if extents:
            print()
            print(f"  EXTENTS  X=[{extents['x'][0]:.0f}, {extents['x'][1]:.0f}]"
                  f"  Y=[{extents['y'][0]:.0f}, {extents['y'][1]:.0f}]"
                  f"  Z=[{extents['z'][0]:.0f}, {extents['z'][1]:.0f}]")

    return 0


# ---------------------------------------------------------------------------
# Subcommand: validate
# ---------------------------------------------------------------------------

def cmd_validate(args):
    """Run all validation checks on a DXF file or directory."""
    paths = _resolve_dxf_paths(args.path)
    if not paths:
        raise SystemExit(f"No DXF files found: {args.path}")

    layer_set = set(args.layers.split(",")) if args.layers else None
    exit_code = 0

    for filepath in paths:
        doc = load_dxf(filepath)
        entities = list(iter_model_entities(doc, layer_set))
        all_issues = []
        for check_fn in ALL_CHECKS:
            all_issues.extend(check_fn(entities, args.tolerance, doc))

        report = build_report(filepath, len(entities), all_issues, args.tolerance)

        if args.json:
            print(json.dumps(report, indent=2))
        else:
            print(format_report_text(report))

        if report["summary"]["error"] > 0:
            exit_code = 1

    return exit_code


# ---------------------------------------------------------------------------
# Subcommand: clean
# ---------------------------------------------------------------------------

def _flatten_z(doc):
    """Set all Z coordinates to 0."""
    count = 0
    msp = doc.modelspace()
    for ent in msp:
        dxftype = ent.dxftype()
        if dxftype == "LINE":
            s = ent.dxf.start
            e = ent.dxf.end
            if s.z != 0 or e.z != 0:
                ent.dxf.start = (s.x, s.y, 0)
                ent.dxf.end = (e.x, e.y, 0)
                count += 1
        elif dxftype == "LWPOLYLINE":
            if ent.dxf.hasattr("elevation") and ent.dxf.elevation != 0:
                ent.dxf.elevation = 0
                count += 1
        elif dxftype == "POLYLINE":
            changed = False
            for v in ent.vertices:
                loc = v.dxf.location
                if loc.z != 0:
                    v.dxf.location = (loc.x, loc.y, 0)
                    changed = True
            if changed:
                count += 1
        elif dxftype in ("CIRCLE", "ARC"):
            c = ent.dxf.center
            if c.z != 0:
                ent.dxf.center = (c.x, c.y, 0)
                count += 1
        elif dxftype == "INSERT":
            p = ent.dxf.insert
            if p.z != 0:
                ent.dxf.insert = (p.x, p.y, 0)
                count += 1
        elif dxftype in ("TEXT", "MTEXT"):
            p = ent.dxf.insert
            if p.z != 0:
                ent.dxf.insert = (p.x, p.y, 0)
                count += 1
        elif dxftype == "POINT":
            p = ent.dxf.location
            if p.z != 0:
                ent.dxf.location = (p.x, p.y, 0)
                count += 1
    return count


def _remove_duplicates(doc, tolerance):
    """Remove duplicate entities, keeping the first occurrence."""
    msp = doc.modelspace()
    entities = list(msp)
    seen = {}
    to_delete = []
    for ent in entities:
        fp = entity_fingerprint(ent, tolerance)
        if fp is None:
            continue
        if fp in seen:
            to_delete.append(ent)
        else:
            seen[fp] = ent
    for ent in to_delete:
        msp.delete_entity(ent)
    return len(to_delete)


def _explode_blocks(doc):
    """Explode all INSERT entities into primitive geometry."""
    msp = doc.modelspace()
    inserts = [e for e in msp if e.dxftype() == "INSERT"]
    count = 0
    for ins in inserts:
        try:
            ins.explode()
            count += 1
        except Exception:
            pass  # Skip blocks that fail to explode (e.g., empty blocks)
    return count


def _close_polylines(doc, tolerance):
    """Close polylines whose start and end points are within tolerance."""
    msp = doc.modelspace()
    count = 0
    for ent in msp:
        dxftype = ent.dxftype()
        if dxftype not in ("LWPOLYLINE", "POLYLINE"):
            continue
        if ent.is_closed:
            continue
        pts = entity_endpoints(ent)
        if len(pts) < 3:
            continue
        dist = math.dist((pts[0][0], pts[0][1]), (pts[-1][0], pts[-1][1]))
        if dist <= tolerance:
            ent.close()
            count += 1
    return count


def _convert_lwpolylines(doc):
    """Convert LWPOLYLINE entities to LINE segments.

    ETABS silently ignores LWPOLYLINE (the modern AutoCAD default).
    Convert each LWPOLYLINE to a series of LINE entities on the same layer.
    For closed LWPOLYLINEs, also create a 3DFACE if it has 3-4 vertices
    (ETABS reads 3DFACE as area elements).
    """
    msp = doc.modelspace()
    lwpolys = [e for e in msp if e.dxftype() == "LWPOLYLINE"]
    count = 0

    for lw in lwpolys:
        layer = lw.dxf.layer
        pts = list(lw.get_points(format="xy"))
        z = lw.dxf.elevation if lw.dxf.hasattr("elevation") else 0
        closed = lw.is_closed

        if len(pts) < 2:
            msp.delete_entity(lw)
            count += 1
            continue

        # Create LINE segments for each edge
        for i in range(len(pts) - 1):
            msp.add_line(
                (pts[i][0], pts[i][1], z),
                (pts[i + 1][0], pts[i + 1][1], z),
                dxfattribs={"layer": layer},
            )

        # Close the polygon with a final segment
        if closed and len(pts) >= 3:
            msp.add_line(
                (pts[-1][0], pts[-1][1], z),
                (pts[0][0], pts[0][1], z),
                dxfattribs={"layer": layer},
            )

            # For 3-4 vertex closed shapes, also create a 3DFACE
            # (ETABS reads 3DFACE as floor/wall area elements)
            if len(pts) <= 4:
                face_pts = [(p[0], p[1], z) for p in pts]
                while len(face_pts) < 4:
                    face_pts.append(face_pts[-1])  # 3DFACE needs 4 vertices
                msp.add_3dface(face_pts, dxfattribs={"layer": layer})

        msp.delete_entity(lw)
        count += 1

    return count


def _move_to_origin(doc):
    """Translate all geometry so the model centroid is near the origin.

    Prevents floating-point precision loss from large absolute coordinates.
    Returns the offset applied (dx, dy) so the user knows the shift.
    """
    msp = doc.modelspace()
    entities = list(msp)

    # Compute centroid of all geometry
    all_pts = []
    for ent in entities:
        all_pts.extend(entity_endpoints(ent))
    if not all_pts:
        return (0, 0)

    cx = (min(p[0] for p in all_pts) + max(p[0] for p in all_pts)) / 2
    cy = (min(p[1] for p in all_pts) + max(p[1] for p in all_pts)) / 2

    # Only shift if far from origin
    max_coord = max(abs(cx), abs(cy))
    x_range = max(p[0] for p in all_pts) - min(p[0] for p in all_pts)
    y_range = max(p[1] for p in all_pts) - min(p[1] for p in all_pts)
    model_dim = max(x_range, y_range, 1.0)

    if max_coord / model_dim < 5000:
        return (0, 0)

    # Shift everything
    dx, dy = -cx, -cy
    for ent in entities:
        _shift_entity_y(ent, dy)
        _shift_entity_x(ent, dx)

    return (round(dx, 2), round(dy, 2))


def _shift_entity_x(entity, dx):
    """Shift an entity's X coordinates by dx."""
    dxftype = entity.dxftype()
    if dxftype == "LINE":
        s = entity.dxf.start
        e = entity.dxf.end
        entity.dxf.start = (s.x + dx, s.y, s.z)
        entity.dxf.end = (e.x + dx, e.y, e.z)
    elif dxftype == "LWPOLYLINE":
        pts = list(entity.get_points(format="xyseb"))
        shifted = [(p[0] + dx, *p[1:]) for p in pts]
        entity.set_points(shifted, format="xyseb")
    elif dxftype == "POLYLINE":
        for v in entity.vertices:
            loc = v.dxf.location
            v.dxf.location = (loc.x + dx, loc.y, loc.z)
    elif dxftype in ("CIRCLE", "ARC"):
        c = entity.dxf.center
        entity.dxf.center = (c.x + dx, c.y, c.z)
    elif dxftype == "INSERT":
        p = entity.dxf.insert
        entity.dxf.insert = (p.x + dx, p.y, p.z)
    elif dxftype in ("TEXT", "MTEXT"):
        if entity.dxf.hasattr("insert"):
            p = entity.dxf.insert
            entity.dxf.insert = (p.x + dx, p.y, p.z)
    elif dxftype == "POINT":
        p = entity.dxf.location
        entity.dxf.location = (p.x + dx, p.y, p.z)
    elif dxftype == "3DFACE":
        for attr in ("vtx0", "vtx1", "vtx2", "vtx3"):
            if entity.dxf.hasattr(attr):
                v = getattr(entity.dxf, attr)
                setattr(entity.dxf, attr, (v.x + dx, v.y, v.z))


def cmd_clean(args):
    """Apply geometry fixes and write an ETABS-ready DXF.

    Fix order matters:
    1. Explode blocks (creates new entities for subsequent passes)
    2. Convert LWPOLYLINE → LINE + 3DFACE (ETABS ignores LWPOLYLINE)
    3. Flatten Z coordinates
    4. Move to origin (if far from 0,0)
    5. Remove duplicates (after all conversions, more dupes may emerge)
    6. Close near-closed polylines
    """
    doc = load_dxf(args.input)
    do_all = args.all

    results = {}

    if do_all or args.explode_blocks:
        results["blocks_exploded"] = _explode_blocks(doc)

    if do_all or args.convert_lwpoly:
        results["lwpolylines_converted"] = _convert_lwpolylines(doc)

    if do_all or args.flatten_z:
        results["z_flattened"] = _flatten_z(doc)

    if do_all or args.move_origin:
        offset = _move_to_origin(doc)
        if offset != (0, 0):
            results["origin_shifted"] = {"dx": offset[0], "dy": offset[1]}

    if do_all or args.remove_dupes:
        results["duplicates_removed"] = _remove_duplicates(doc, args.tolerance)

    if do_all or args.close_polylines:
        results["polylines_closed"] = _close_polylines(doc, args.tolerance)

    # Save as R2000 format for max ETABS compatibility
    doc.saveas(args.output)

    if args.json:
        results["input"] = args.input
        results["output"] = args.output
        print(json.dumps(results, indent=2))
    else:
        print(f"clean: {args.input} -> {args.output}")
        for k, v in results.items():
            print(f"  {k}: {v}")

    return 0


# ---------------------------------------------------------------------------
# Subcommand: floors
# ---------------------------------------------------------------------------

def _cluster_by_gaps(entities, gap_threshold):
    """Split entities into spatial clusters by detecting large gaps.

    Projects entity centroids onto X and Y axes independently, finds gaps
    larger than gap_threshold, and partitions into rectangular clusters.
    """
    items = []  # (cx, cy, entity)
    for ent in entities:
        bb = entity_bbox(ent)
        if bb is None:
            continue
        cx = (bb[0] + bb[2]) / 2
        cy = (bb[1] + bb[3]) / 2
        items.append((cx, cy, ent))

    if not items:
        return []

    # Find X splits
    x_sorted = sorted(set(round(item[0], 2) for item in items))
    x_splits = _find_gaps(x_sorted, gap_threshold)

    # Find Y splits
    y_sorted = sorted(set(round(item[1], 2) for item in items))
    y_splits = _find_gaps(y_sorted, gap_threshold)

    # Assign each entity to a (x_bin, y_bin) cluster
    clusters = defaultdict(list)
    for cx, cy, ent in items:
        xbin = _assign_bin(cx, x_splits)
        ybin = _assign_bin(cy, y_splits)
        clusters[(xbin, ybin)].append(ent)

    # Sort clusters left-to-right, bottom-to-top
    sorted_keys = sorted(clusters.keys())
    return [clusters[k] for k in sorted_keys]


def _find_gaps(sorted_values, threshold):
    """Return split points where consecutive values differ by > threshold."""
    splits = []
    for i in range(1, len(sorted_values)):
        gap = sorted_values[i] - sorted_values[i - 1]
        if gap > threshold:
            splits.append((sorted_values[i - 1] + sorted_values[i]) / 2)
    return splits


def _assign_bin(value, splits):
    """Return which bin a value falls in given a sorted list of split points."""
    for i, s in enumerate(splits):
        if value < s:
            return i
    return len(splits)


def _detect_gap_threshold(entities):
    """Auto-detect a reasonable gap threshold from entity bounding boxes."""
    sizes = []
    for ent in entities:
        bb = entity_bbox(ent)
        if bb is None:
            continue
        diag = math.dist((bb[0], bb[1]), (bb[2], bb[3]))
        if diag > 0:
            sizes.append(diag)
    if not sizes:
        return 100.0
    sizes.sort()
    median = sizes[len(sizes) // 2]
    return median * 10


def _find_floor_label(cluster_entities, all_text_entities, cluster_bbox):
    """Search for floor-label text near a cluster."""
    cx = (cluster_bbox[0] + cluster_bbox[2]) / 2
    cy = (cluster_bbox[1] + cluster_bbox[3]) / 2
    width = cluster_bbox[2] - cluster_bbox[0]
    height = cluster_bbox[3] - cluster_bbox[1]
    search_radius = max(width, height) * 0.6

    best_match = None
    best_dist = float("inf")

    for tent in all_text_entities:
        text = entity_text(tent)
        if not text:
            continue
        # Check against floor label patterns
        matched = False
        for pattern in FLOOR_LABEL_RE:
            if pattern.search(text):
                matched = True
                break
        if not matched:
            continue

        pts = entity_endpoints(tent)
        if not pts:
            continue
        tx, ty = pts[0][0], pts[0][1]
        dist = math.dist((tx, ty), (cx, cy))
        if dist < search_radius and dist < best_dist:
            best_dist = dist
            best_match = text.strip()

    return best_match


def _cluster_bbox(entities):
    """Compute bounding box of a list of entities."""
    xmin = ymin = float("inf")
    xmax = ymax = float("-inf")
    for ent in entities:
        bb = entity_bbox(ent)
        if bb is None:
            continue
        xmin = min(xmin, bb[0])
        ymin = min(ymin, bb[1])
        xmax = max(xmax, bb[2])
        ymax = max(ymax, bb[3])
    if xmin == float("inf"):
        return None
    return (xmin, ymin, xmax, ymax)


def _detect_floor_labels(entities):
    """Find text entities that match floor label patterns.

    Returns list of (label_text, x, y) sorted by Y ascending.
    """
    labels = []
    for ent in entities:
        if ent.dxftype() not in ("TEXT", "MTEXT"):
            continue
        text = entity_text(ent).strip()
        if not text:
            continue
        matched = False
        for pattern in FLOOR_LABEL_RE:
            if pattern.search(text):
                matched = True
                break
        # Also match bare floor numbers/ranges like "26", "29-35", "44-50"
        if not matched and re.match(r"^\d+(-\d+)?$", text):
            matched = True
        if matched:
            pts = entity_endpoints(ent)
            if pts:
                labels.append((text, pts[0][0], pts[0][1]))
    labels.sort(key=lambda t: t[2])  # sort by Y
    return labels


def _labels_to_bands(labels, all_entities):
    """Convert sorted floor labels into Y-bands and assign entities.

    Given labels at known Y positions, compute band boundaries as the
    midpoints between consecutive labels.  Assign every non-text entity
    to the band whose label Y is closest.
    """
    if not labels:
        return []

    # Compute band boundaries as midpoints between consecutive label Y positions
    boundaries = []
    for i in range(1, len(labels)):
        mid = (labels[i - 1][2] + labels[i][2]) / 2
        boundaries.append(mid)

    # Also need a lower and upper bound
    # Use the label spacing to extend below first and above last
    if len(labels) >= 2:
        typical_spacing = labels[1][2] - labels[0][2]
    else:
        typical_spacing = 5000  # fallback
    lower = labels[0][2] - typical_spacing / 2
    upper = labels[-1][2] + typical_spacing / 2

    # Build bands: [(label, y_min, y_max), ...]
    bands = []
    for i, (text, lx, ly) in enumerate(labels):
        y_min = lower if i == 0 else boundaries[i - 1]
        y_max = upper if i == len(labels) - 1 else boundaries[i]
        bands.append((text, y_min, y_max))

    # Assign entities to bands by centroid Y
    band_entities = [[] for _ in bands]
    for ent in all_entities:
        if ent.dxftype() in ("TEXT", "MTEXT"):
            continue  # skip label text itself
        bb = entity_bbox(ent)
        if bb is None:
            continue
        cy = (bb[1] + bb[3]) / 2
        # Find best band
        best_idx = 0
        best_dist = float("inf")
        for i, (_, y_min, y_max) in enumerate(bands):
            if y_min <= cy <= y_max:
                best_idx = i
                best_dist = 0
                break
            dist = min(abs(cy - y_min), abs(cy - y_max))
            if dist < best_dist:
                best_dist = dist
                best_idx = i
        band_entities[best_idx].append(ent)

    return list(zip(bands, band_entities))


def cmd_floors(args):
    """Identify and label floor plans within a DXF."""
    # Manual mapping mode
    if args.mapping:
        return _cmd_floors_manual(args)

    paths = _resolve_dxf_paths(args.path)
    if not paths:
        raise SystemExit(f"No DXF files found: {args.path}")

    # If directory with multiple files, treat each file as one floor
    if len(paths) > 1:
        return _cmd_floors_multi_file(args, paths)

    # Single file
    filepath = paths[0]
    doc = load_dxf(filepath)
    entities = list(iter_model_entities(doc))

    # Strategy 1: text-label-driven floor detection
    # If we find evenly-spaced floor labels, use them to define Y-bands
    labels = _detect_floor_labels(entities)

    if len(labels) >= 2:
        return _cmd_floors_label_driven(args, filepath, entities, labels)

    # Strategy 2: spatial gap clustering (fallback)
    gap_threshold = args.gap_threshold
    if gap_threshold is None:
        gap_threshold = _detect_gap_threshold(entities)

    clusters = _cluster_by_gaps(entities, gap_threshold)
    text_ents = [e for e in entities if e.dxftype() in ("TEXT", "MTEXT")]

    floors = []
    for i, cluster in enumerate(clusters):
        bbox = _cluster_bbox(cluster)
        if bbox is None:
            continue
        label = _find_floor_label(cluster, text_ents, bbox)
        if label is None:
            label = f"FLOOR_{i + 1}"
        cx = (bbox[0] + bbox[2]) / 2
        cy = (bbox[1] + bbox[3]) / 2
        floors.append({
            "floor_id": label,
            "label_source": "text_proximity" if label != f"FLOOR_{i + 1}" else "auto_index",
            "bbox": [round(v, 2) for v in bbox],
            "entity_count": len(cluster),
            "centroid": [round(cx, 2), round(cy, 2)],
        })

    result = {
        "file": str(filepath),
        "mode": "gap_cluster",
        "gap_threshold": round(gap_threshold, 2),
        "floor_count": len(floors),
        "floors": floors,
    }

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"floors: {filepath}")
        print(f"  gap_threshold={gap_threshold:.1f}  clusters={len(floors)}")
        for f in floors:
            src = f["label_source"]
            print(f"  {f['floor_id']:20s}  {f['entity_count']:5d} entities  "
                  f"[{src}]")

    return 0


def _cmd_floors_label_driven(args, filepath, entities, labels):
    """Use detected text labels to define floor bands and assign entities."""
    band_results = _labels_to_bands(labels, entities)

    floors = []
    for (label_text, y_min, y_max), band_ents in band_results:
        bbox = _cluster_bbox(band_ents)
        cx = cy = None
        if bbox:
            cx = round((bbox[0] + bbox[2]) / 2, 2)
            cy = round((bbox[1] + bbox[3]) / 2, 2)
        floors.append({
            "floor_id": label_text,
            "label_source": "text_label",
            "y_band": [round(y_min, 2), round(y_max, 2)],
            "bbox": [round(v, 2) for v in bbox] if bbox else None,
            "entity_count": len(band_ents),
            "centroid": [cx, cy] if cx is not None else None,
        })

    result = {
        "file": str(filepath),
        "mode": "label_driven",
        "labels_found": len(labels),
        "floor_count": len(floors),
        "floors": floors,
    }

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"floors: {filepath}")
        print(f"  mode=label_driven  labels={len(labels)}")
        for f in floors:
            band = f["y_band"]
            print(f"  {f['floor_id']:20s}  {f['entity_count']:5d} entities  "
                  f"Y=[{band[0]:.0f}, {band[1]:.0f}]")

    return 0


def _cmd_floors_manual(args):
    """Handle --mapping flag: user-provided floor-to-file assignments."""
    pairs = [p.strip() for p in args.mapping.split(",")]
    floors = []
    for pair in pairs:
        if "=" not in pair:
            raise SystemExit(f"Bad mapping format '{pair}': expected FLOOR=file.dxf")
        floor_id, filepath = pair.split("=", 1)
        doc = load_dxf(filepath)
        entities = list(iter_model_entities(doc))
        bbox = _cluster_bbox(entities)
        floors.append({
            "floor_id": floor_id.strip(),
            "label_source": "manual",
            "file": filepath.strip(),
            "bbox": [round(v, 2) for v in bbox] if bbox else None,
            "entity_count": len(entities),
        })

    result = {"mode": "manual", "floor_count": len(floors), "floors": floors}
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        for f in floors:
            print(f"  {f['floor_id']:20s}  {f['entity_count']:5d} entities  "
                  f"[{f['file']}]")
    return 0


def _cmd_floors_multi_file(args, paths):
    """Each file = one floor. Infer labels from filenames."""
    floors = []
    for filepath in paths:
        doc = load_dxf(filepath)
        entities = list(iter_model_entities(doc))
        bbox = _cluster_bbox(entities)

        # Try to extract floor label from filename
        stem = Path(filepath).stem
        label = None
        for pattern in FLOOR_LABEL_RE:
            m = pattern.search(stem)
            if m:
                label = m.group(0).strip()
                break
        if label is None:
            label = stem

        cx = cy = None
        if bbox:
            cx = round((bbox[0] + bbox[2]) / 2, 2)
            cy = round((bbox[1] + bbox[3]) / 2, 2)

        floors.append({
            "floor_id": label,
            "label_source": "filename",
            "file": str(filepath),
            "bbox": [round(v, 2) for v in bbox] if bbox else None,
            "entity_count": len(entities),
            "centroid": [cx, cy] if cx is not None else None,
        })

    result = {"mode": "multi_file", "floor_count": len(floors), "floors": floors}
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"floors: {len(paths)} files")
        for f in floors:
            print(f"  {f['floor_id']:20s}  {f['entity_count']:5d} entities  "
                  f"[{f['file']}]")
    return 0


# ---------------------------------------------------------------------------
# Subcommand: classify
# ---------------------------------------------------------------------------

# Heuristic layer-to-ETABS-element mapping
LAYER_CLASS_HINTS = {
    "wall": "wall",
    "sw": "wall",
    "shear": "wall",
    "core": "wall",
    "col": "column",
    "column": "column",
    "beam": "beam",
    "bm": "beam",
    "framing": "beam",
    "slab": "slab",
    "floor": "slab",
    "deck": "slab",
    "grid": "grid",
    "anno": "skip",
    "text": "skip",
    "dim": "skip",
    "hatch": "skip",
    "defpoints": "skip",
}

BLOCK_CLASS_HINTS = {
    "wall": "wall",
    "sw": "wall",
    "shear": "wall",
    "col": "column",
    "column": "column",
    "beam": "beam",
    "brace": "brace",
}


def _guess_layer_class(layer_name):
    """Heuristic: guess ETABS element type from layer name."""
    lower = layer_name.lower()
    if lower == "0":
        return "skip"
    for hint, cls in LAYER_CLASS_HINTS.items():
        if hint in lower:
            return cls
    return "unknown"


def _guess_block_class(block_name, layer_name):
    """Heuristic: guess what a block represents from its name and layer."""
    lower = block_name.lower()
    for hint, cls in BLOCK_CLASS_HINTS.items():
        if hint in lower:
            return cls
    # Fall back to layer classification
    return _guess_layer_class(layer_name)


def cmd_classify(args):
    """Propose layer and block classifications for ETABS import.

    Claude reads the output, presents it to the user, and collects corrections.
    """
    filepath = args.path
    doc = load_dxf(filepath)
    msp = doc.modelspace()
    entities = list(msp)

    # Layer classification proposals
    layer_entities = defaultdict(lambda: defaultdict(int))
    for ent in entities:
        layer_entities[ent.dxf.layer][ent.dxftype()] += 1

    layer_proposals = {}
    for layer_name, type_counts in sorted(layer_entities.items()):
        layer_proposals[layer_name] = {
            "proposed_class": _guess_layer_class(layer_name),
            "entity_count": sum(type_counts.values()),
            "entity_types": dict(type_counts),
        }

    # Block classification proposals
    block_usage = defaultdict(list)
    for ent in entities:
        if ent.dxftype() == "INSERT":
            p = ent.dxf.insert
            block_usage[ent.dxf.name].append({
                "layer": ent.dxf.layer,
                "y": round(p.y, 2),
            })

    block_proposals = {}
    for block_name, instances in sorted(block_usage.items()):
        layers_used = list(set(inst["layer"] for inst in instances))
        primary_layer = layers_used[0] if layers_used else ""
        y_positions = sorted(set(inst["y"] for inst in instances))

        block_proposals[block_name] = {
            "proposed_class": _guess_block_class(block_name, primary_layer),
            "insert_count": len(instances),
            "layers": layers_used,
            "y_positions": y_positions,
        }

    # Detect floor labels and their Y positions for cross-referencing
    floor_labels = _detect_floor_labels(entities)
    label_y_map = {text: round(y, 2) for text, x, y in floor_labels}

    # Try to match block Y positions to floor labels
    for block_name, info in block_proposals.items():
        floor_matches = []
        for y_pos in info["y_positions"]:
            best_label = None
            best_dist = float("inf")
            for label_text, label_y in label_y_map.items():
                dist = abs(y_pos - label_y)
                if dist < best_dist:
                    best_dist = dist
                    best_label = label_text
            if best_label and best_dist < 5000:  # within one floor spacing
                floor_matches.append({"floor": best_label, "y": y_pos})
        info["floor_matches"] = floor_matches

    result = {
        "file": str(filepath),
        "layer_proposals": layer_proposals,
        "block_proposals": block_proposals,
        "floor_labels": label_y_map,
    }

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"classify: {filepath}")
        print()
        print("  LAYER PROPOSALS")
        for name, info in layer_proposals.items():
            cls = info["proposed_class"]
            count = info["entity_count"]
            print(f"    {name:25s}  → {cls:10s}  ({count} entities)")
        print()
        print("  BLOCK PROPOSALS")
        for name, info in block_proposals.items():
            cls = info["proposed_class"]
            n = info["insert_count"]
            floors = ", ".join(m["floor"] for m in info.get("floor_matches", []))
            floor_str = f"  floors: {floors}" if floors else ""
            print(f"    {name:25s}  → {cls:10s}  ({n}x){floor_str}")

    return 0


# ---------------------------------------------------------------------------
# Subcommand: split
# ---------------------------------------------------------------------------

def cmd_split(args):
    """Export individual DXF files per detected floor."""
    filepath = args.input
    doc = load_dxf(filepath)
    entities = list(iter_model_entities(doc))

    labels = _detect_floor_labels(entities)
    if len(labels) < 2:
        raise SystemExit("Cannot split: fewer than 2 floor labels detected. "
                         "Use 'floors' with --mapping for manual assignment.")

    band_results = _labels_to_bands(labels, entities)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    results = []
    for (label_text, y_min, y_max), band_ents in band_results:
        if not band_ents:
            results.append({
                "floor_id": label_text,
                "entity_count": 0,
                "file": None,
            })
            continue

        # Create a new DXF with just this floor's entities
        new_doc = ezdxf.new(doc.dxfversion)

        # Copy layer definitions
        for layer in doc.layers:
            if layer.dxf.name not in new_doc.layers:
                new_doc.layers.add(layer.dxf.name, color=layer.color)

        # Copy block definitions used by this floor's entities
        block_names_needed = set()
        for ent in band_ents:
            if ent.dxftype() == "INSERT":
                block_names_needed.add(ent.dxf.name)

        for bname in block_names_needed:
            if bname in doc.blocks and bname not in new_doc.blocks:
                src_block = doc.blocks[bname]
                new_block = new_doc.blocks.new(bname)
                for src_ent in src_block:
                    new_doc.entitydb.add(src_ent.copy())
                    new_block.add_entity(src_ent.copy())

        # Copy entities, shifting Y so each floor starts near Y=0
        new_msp = new_doc.modelspace()
        y_offset = y_min
        for ent in band_ents:
            try:
                copied = ent.copy()
                # Shift Y coordinates
                _shift_entity_y(copied, -y_offset)
                new_msp.add_entity(copied)
            except Exception:
                pass  # Skip entities that can't be copied

        safe_name = re.sub(r"[^\w\-]", "_", label_text).strip("_")
        out_path = output_dir / f"{safe_name}.dxf"
        new_doc.saveas(str(out_path))

        results.append({
            "floor_id": label_text,
            "entity_count": len(band_ents),
            "file": str(out_path),
            "y_offset_applied": round(y_offset, 2),
        })

    result = {
        "input": str(filepath),
        "output_dir": str(output_dir),
        "floors_exported": len([r for r in results if r["file"]]),
        "floors": results,
    }

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"split: {filepath} → {output_dir}/")
        for r in results:
            if r["file"]:
                print(f"  {r['floor_id']:20s}  {r['entity_count']:5d} entities  "
                      f"→ {r['file']}")
            else:
                print(f"  {r['floor_id']:20s}  (empty)")

    return 0


def _shift_entity_y(entity, dy):
    """Shift an entity's Y coordinates by dy."""
    dxftype = entity.dxftype()
    if dxftype == "LINE":
        s = entity.dxf.start
        e = entity.dxf.end
        entity.dxf.start = (s.x, s.y + dy, s.z)
        entity.dxf.end = (e.x, e.y + dy, e.z)
    elif dxftype == "LWPOLYLINE":
        pts = list(entity.get_points(format="xyseb"))
        shifted = [(p[0], p[1] + dy, *p[2:]) for p in pts]
        entity.set_points(shifted, format="xyseb")
    elif dxftype == "POLYLINE":
        for v in entity.vertices:
            loc = v.dxf.location
            v.dxf.location = (loc.x, loc.y + dy, loc.z)
    elif dxftype in ("CIRCLE", "ARC"):
        c = entity.dxf.center
        entity.dxf.center = (c.x, c.y + dy, c.z)
    elif dxftype == "INSERT":
        p = entity.dxf.insert
        entity.dxf.insert = (p.x, p.y + dy, p.z)
    elif dxftype in ("TEXT", "MTEXT"):
        if entity.dxf.hasattr("insert"):
            p = entity.dxf.insert
            entity.dxf.insert = (p.x, p.y + dy, p.z)
    elif dxftype == "POINT":
        p = entity.dxf.location
        entity.dxf.location = (p.x, p.y + dy, p.z)


# ---------------------------------------------------------------------------
# Subcommand: align
# ---------------------------------------------------------------------------

def _extract_column_centroids_for_block(doc, block_name, insert_point):
    """Get global (x, y) centroids for all column polylines in a block,
    given the block's insertion point."""
    if block_name not in doc.blocks:
        return []
    block = doc.blocks[block_name]
    centroids = []
    ix, iy = insert_point[0], insert_point[1]
    for ent in block:
        if ent.dxftype() != "LWPOLYLINE":
            continue
        if not ent.is_closed:
            continue
        pts = list(ent.get_points(format="xy"))
        if len(pts) < 3:
            continue
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        cx = (min(xs) + max(xs)) / 2 + ix
        cy = (min(ys) + max(ys)) / 2 + iy
        w = max(xs) - min(xs)
        h = max(ys) - min(ys)
        centroids.append({
            "x": round(cx, 2),
            "y": round(cy, 2),
            "size": f"{round(w)}x{round(h)}",
        })
    return centroids


def _extract_direct_column_centroids(entities):
    """Get centroids from non-block column geometry (direct LWPOLYLINE on column layers)."""
    centroids = []
    for ent in entities:
        if ent.dxftype() != "LWPOLYLINE":
            continue
        if not ent.is_closed:
            continue
        layer_lower = ent.dxf.layer.lower()
        if "col" not in layer_lower:
            continue
        pts = list(ent.get_points(format="xy"))
        if len(pts) < 3:
            continue
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        cx = (min(xs) + max(xs)) / 2
        cy = (min(ys) + max(ys)) / 2
        w = max(xs) - min(xs)
        h = max(ys) - min(ys)
        centroids.append({
            "x": round(cx, 2),
            "y": round(cy, 2),
            "size": f"{round(w)}x{round(h)}",
        })
    return centroids


def _normalize_columns_to_plan(columns, floor_y_min):
    """Shift column Y coords so floor band starts at Y=0 (plan-relative)."""
    return [
        {"x": c["x"], "y": round(c["y"] - floor_y_min, 2), "size": c["size"]}
        for c in columns
    ]


def _match_columns(cols_a, cols_b, tolerance):
    """Find matched, dropped, and added columns between two floors.

    Returns (matched, dropped, added) where:
    - matched: list of (col_a, col_b) pairs within tolerance
    - dropped: cols in A with no match in B (column removed going up)
    - added: cols in B with no match in A (new column appearing)
    """
    used_b = set()
    matched = []
    dropped = []

    for ca in cols_a:
        best_idx = None
        best_dist = float("inf")
        for j, cb in enumerate(cols_b):
            if j in used_b:
                continue
            d = math.dist((ca["x"], ca["y"]), (cb["x"], cb["y"]))
            if d < best_dist:
                best_dist = d
                best_idx = j
        if best_idx is not None and best_dist <= tolerance:
            matched.append((ca, cols_b[best_idx]))
            used_b.add(best_idx)
        else:
            dropped.append(ca)

    added = [cols_b[j] for j in range(len(cols_b)) if j not in used_b]
    return matched, dropped, added


def _draw_revision_cloud(msp, cx, cy, radius, layer="ALIGN-CLOUD"):
    """Draw a revision cloud (bumpy circle) around a point."""
    n_bumps = 12
    bump_r = radius * 0.35
    pts = []
    for i in range(n_bumps * 2):
        angle = 2 * math.pi * i / (n_bumps * 2)
        if i % 2 == 0:
            r = radius
        else:
            r = radius + bump_r
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        pts.append((x, y))
    msp.add_lwpolyline(pts, close=True, dxfattribs={
        "layer": layer,
        "color": 1,  # red
    })


def cmd_align(args):
    """Check column vertical alignment across floors.

    Finds columns that appear on one floor but not the adjacent floor —
    these are transfer conditions that need beams. Outputs a DXF with
    revision clouds marking discontinuities.
    """
    filepath = args.path
    doc = load_dxf(filepath)
    msp_ents = list(doc.modelspace())

    # Detect floors
    labels = _detect_floor_labels(msp_ents)
    if len(labels) < 2:
        raise SystemExit("Need at least 2 floor labels for alignment check.")

    # Build floor bands (just for Y-band boundaries)
    band_results = _labels_to_bands(labels, msp_ents)
    bands = [(label, y_min, y_max) for (label, y_min, y_max), _ in band_results]

    tolerance = args.tolerance

    # Collect ALL column positions globally from every INSERT in the file,
    # then assign each individual column to a floor band by its actual Y.
    all_global_columns = []  # (global_x, global_y, size_str)

    for ent in msp_ents:
        if ent.dxftype() == "INSERT":
            layer_lower = ent.dxf.layer.lower()
            block_lower = ent.dxf.name.lower()
            is_col = "col" in layer_lower or "col" in block_lower
            if is_col:
                p = ent.dxf.insert
                cols = _extract_column_centroids_for_block(
                    doc, ent.dxf.name, (p.x, p.y))
                all_global_columns.extend(cols)

    # Also pick up direct (non-block) column polylines
    direct = _extract_direct_column_centroids(msp_ents)
    all_global_columns.extend(direct)

    # Assign each column to a floor band by its global Y
    floor_columns = []  # (label, y_min, y_max, [columns])
    for label, y_min, y_max in bands:
        cols_in_band = []
        for col in all_global_columns:
            if y_min <= col["y"] <= y_max:
                cols_in_band.append(col)
        # Normalize Y to plan-relative (subtract y_min)
        cols_in_band = _normalize_columns_to_plan(cols_in_band, y_min)
        floor_columns.append((label, y_min, y_max, cols_in_band))

    # Compare adjacent floors
    discontinuities = []
    for i in range(len(floor_columns) - 1):
        below_label, below_ymin, below_ymax, below_cols = floor_columns[i]
        above_label, above_ymin, above_ymax, above_cols = floor_columns[i + 1]

        if not below_cols and not above_cols:
            continue

        matched, dropped, added = _match_columns(below_cols, above_cols, tolerance)

        if dropped or added:
            discontinuities.append({
                "below": below_label,
                "above": above_label,
                "below_count": len(below_cols),
                "above_count": len(above_cols),
                "matched": len(matched),
                "dropped": [{"x": c["x"], "y": c["y"], "size": c["size"]}
                            for c in dropped],
                "added": [{"x": c["x"], "y": c["y"], "size": c["size"]}
                          for c in added],
            })

    # Generate output DXF with clouds
    if args.output:
        out_doc = ezdxf.readfile(filepath)  # fresh copy
        out_msp = out_doc.modelspace()

        # Add cloud layer
        if "ALIGN-CLOUD" not in out_doc.layers:
            out_doc.layers.add("ALIGN-CLOUD", color=1)
        if "ALIGN-LABEL" not in out_doc.layers:
            out_doc.layers.add("ALIGN-LABEL", color=1)

        cloud_radius = tolerance * 3

        for disc in discontinuities:
            # Find the floor band Y offset for the "below" floor
            below_ymin = None
            for label, ymin, ymax, _ in floor_columns:
                if label == disc["below"]:
                    below_ymin = ymin
                    break

            above_ymin = None
            for label, ymin, ymax, _ in floor_columns:
                if label == disc["above"]:
                    above_ymin = ymin
                    break

            # Cloud dropped columns on the lower floor
            for col in disc["dropped"]:
                gx = col["x"]
                gy = col["y"] + (below_ymin or 0)
                _draw_revision_cloud(out_msp, gx, gy, cloud_radius,
                                     layer="ALIGN-CLOUD")
                out_msp.add_text(
                    f"DROP @ {disc['above']}",
                    dxfattribs={
                        "layer": "ALIGN-LABEL",
                        "height": cloud_radius * 0.6,
                        "color": 1,
                        "insert": (gx + cloud_radius * 1.2,
                                   gy + cloud_radius * 0.5),
                    },
                )

            # Cloud added columns on the upper floor
            for col in disc["added"]:
                gx = col["x"]
                gy = col["y"] + (above_ymin or 0)
                _draw_revision_cloud(out_msp, gx, gy, cloud_radius,
                                     layer="ALIGN-CLOUD")
                out_msp.add_text(
                    f"NEW @ {disc['above']}",
                    dxfattribs={
                        "layer": "ALIGN-LABEL",
                        "height": cloud_radius * 0.6,
                        "color": 1,
                        "insert": (gx + cloud_radius * 1.2,
                                   gy + cloud_radius * 0.5),
                    },
                )

        out_doc.saveas(args.output)

    # Report
    result = {
        "file": str(filepath),
        "floor_count": len(floor_columns),
        "floors": [
            {"floor": label, "column_count": len(cols)}
            for label, _, _, cols in floor_columns
        ],
        "discontinuities": discontinuities,
        "output": args.output,
    }

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"align: {filepath}")
        print()
        print("  COLUMNS PER FLOOR")
        for label, _, _, cols in floor_columns:
            print(f"    {label:20s}  {len(cols):3d} columns")
        print()
        if discontinuities:
            print(f"  DISCONTINUITIES ({len(discontinuities)} transitions)")
            for d in discontinuities:
                n_drop = len(d["dropped"])
                n_add = len(d["added"])
                print(f"    {d['below']:10s} → {d['above']:10s}  "
                      f"{d['matched']} aligned, "
                      f"{n_drop} dropped, {n_add} added")
                for col in d["dropped"]:
                    print(f"      DROP  ({col['x']:.0f}, {col['y']:.0f})  "
                          f"{col['size']}")
                for col in d["added"]:
                    print(f"      NEW   ({col['x']:.0f}, {col['y']:.0f})  "
                          f"{col['size']}")
        else:
            print("  All columns align across all floors.")
        if args.output:
            print(f"\n  Clouded DXF: {args.output}")

    return 0


# ---------------------------------------------------------------------------
# Path resolution
# ---------------------------------------------------------------------------

def _resolve_dxf_paths(path_arg):
    """Resolve a path argument to a list of DXF file paths."""
    p = Path(path_arg)
    if p.is_file():
        return [str(p)]
    if p.is_dir():
        files = sorted(p.glob("*.dxf")) + sorted(p.glob("*.DXF"))
        return [str(f) for f in files]
    # Glob pattern
    parent = p.parent
    if parent.is_dir():
        files = sorted(parent.glob(p.name))
        return [str(f) for f in files]
    return []


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        prog="dxf-to-etabs",
        description="DXF-to-ETABS import debugger",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # -- inspect --
    p_insp = sub.add_parser("inspect",
                            help="Full structural survey of a DXF file")
    p_insp.add_argument("path", help="DXF file")
    p_insp.add_argument("--json", action="store_true")

    # -- validate --
    p_val = sub.add_parser("validate", help="Run validation checks on a DXF")
    p_val.add_argument("path", help="DXF file or directory")
    p_val.add_argument("-t", "--tolerance", type=float, default=0.25,
                       help="Snap tolerance in drawing units (default 0.25)")
    p_val.add_argument("--json", action="store_true",
                       help="Output full JSON report")
    p_val.add_argument("--layers", default=None,
                       help="Comma-separated layer filter")

    # -- clean --
    p_clean = sub.add_parser("clean",
                             help="Apply fixes and write an ETABS-ready DXF")
    p_clean.add_argument("input", help="Input DXF file")
    p_clean.add_argument("output", help="Output DXF file")
    p_clean.add_argument("-t", "--tolerance", type=float, default=0.25)
    p_clean.add_argument("--flatten-z", action="store_true")
    p_clean.add_argument("--remove-dupes", action="store_true")
    p_clean.add_argument("--explode-blocks", action="store_true")
    p_clean.add_argument("--convert-lwpoly", action="store_true",
                         help="Convert LWPOLYLINE to LINE + 3DFACE "
                              "(ETABS ignores LWPOLYLINE)")
    p_clean.add_argument("--move-origin", action="store_true",
                         help="Shift geometry to origin if far away "
                              "(prevents float precision loss)")
    p_clean.add_argument("--close-polylines", action="store_true")
    p_clean.add_argument("--all", action="store_true",
                         help="Apply all fixes")
    p_clean.add_argument("--json", action="store_true")

    # -- floors --
    p_floors = sub.add_parser("floors",
                              help="Identify floor plans in a DXF")
    p_floors.add_argument("path", help="DXF file or directory")
    p_floors.add_argument("--gap-threshold", type=float, default=None,
                          help="Min gap between floor plans (auto-detected "
                               "if omitted)")
    p_floors.add_argument("--mapping", default=None,
                          help="Manual floor mapping: L1=file1.dxf,L2=file2.dxf")
    p_floors.add_argument("--json", action="store_true")

    # -- classify --
    p_class = sub.add_parser("classify",
                             help="Propose layer/block classifications")
    p_class.add_argument("path", help="DXF file")
    p_class.add_argument("--json", action="store_true")

    # -- split --
    p_split = sub.add_parser("split",
                             help="Export per-floor DXF files")
    p_split.add_argument("input", help="Input DXF file")
    p_split.add_argument("output_dir", help="Output directory")
    p_split.add_argument("--json", action="store_true")

    # -- align --
    p_align = sub.add_parser("align",
                             help="Check column vertical alignment across floors")
    p_align.add_argument("path", help="DXF file")
    p_align.add_argument("-o", "--output", default=None,
                         help="Output DXF with revision clouds on discontinuities")
    p_align.add_argument("-t", "--tolerance", type=float, default=24.0,
                         help="Match tolerance in drawing units (default 24 — "
                              "half a typical column width)")
    p_align.add_argument("--json", action="store_true")

    args = parser.parse_args()
    commands = {
        "inspect": cmd_inspect,
        "validate": cmd_validate,
        "clean": cmd_clean,
        "floors": cmd_floors,
        "classify": cmd_classify,
        "split": cmd_split,
        "align": cmd_align,
    }
    sys.exit(commands[args.command](args))


if __name__ == "__main__":
    main()
