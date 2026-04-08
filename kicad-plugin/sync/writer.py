"""Export KiCad board data to the sync directory (ecad_to_mcad)."""

import hashlib
import json
import math
import shutil
from pathlib import Path

import pcbnew


def _md5(path: str) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def push(board: pcbnew.BOARD, sync_dir: Path) -> list:
    """
    Export layout JSON, board outline JSON, and per-component STEPs to sync_dir/ecad_to_mcad/.
    Returns a list of change records.
    """
    out_dir = sync_dir / "ecad_to_mcad"
    out_dir.mkdir(parents=True, exist_ok=True)

    changes = []

    # Export per-component STEP models first; returns {ref: model_name} so that
    # layout.json uses the exact same filename that was copied to components/.
    ref_model_map = _export_component_models(board, sync_dir)

    # Export layout JSON (component positions, origins, model offsets)
    layout = _build_layout(board, ref_model_map)
    with open(out_dir / "layout.json", "w") as f:
        json.dump(layout, f, indent=2)

    # Export board outline JSON (outer boundary, holes, slots, cutouts)
    outline = _build_board_outline(board)
    with open(out_dir / "board_outline.json", "w") as f:
        json.dump(outline, f, indent=2)
    changes.append({"type": "board_outline_updated"})
    changes.append({"type": "3d_model_updated"})

    return changes


def _export_component_models(board: pcbnew.BOARD, sync_dir: Path) -> dict:
    """
    Copy per-component STEP files to ecad_to_mcad/components/.
    Returns {ref: model_name} mapping the primary (first readable) STEP model
    for each footprint — used by _build_layout to write the model_name field.
    """
    components_dir = sync_dir / "ecad_to_mcad" / "components"
    components_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = components_dir / "manifest.json"

    if manifest_path.exists():
        try:
            with open(manifest_path) as f:
                manifest = json.load(f)
        except Exception:
            manifest = {"components": {}}
    else:
        manifest = {"components": {}}

    ref_model_map = {}  # ref → model_name (stem of first valid STEP)
    synced = 0

    for fp in board.GetFootprints():
        if _is_mounting_hole(fp):
            continue
        ref = fp.GetReference()
        for model in fp.Models():
            name = Path(model.m_Filename).stem
            if not name:
                continue
            try:
                src_path = pcbnew.ExpandEnvVarSubstitutions(model.m_Filename, None)
            except Exception:
                src_path = None
            if not src_path:
                continue
            src_lower = src_path.lower()
            if not (src_lower.endswith(".step") or src_lower.endswith(".stp")):
                continue
            dest = components_dir / f"{name}.step"
            try:
                src_hash = _md5(src_path)
            except Exception:
                continue  # file not readable — skip (path may exist but file absent)

            # Record this as the primary model for the footprint (first readable STEP)
            if ref not in ref_model_map:
                ref_model_map[ref] = name

            # Copy only if new or hash changed
            entry = manifest["components"].get(name)
            existing_hash = entry["hash"] if isinstance(entry, dict) else ""
            if entry is not None and src_hash == existing_hash:
                continue

            try:
                shutil.copy(src_path, str(dest))
                manifest["components"][name] = {"path": f"components/{name}.step", "hash": src_hash}
                synced += 1
            except Exception:
                pass

    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2)

    if synced:
        print(f"[KiCad→SW] Synced {synced} component model(s) in {components_dir}")

    return ref_model_map


def check_drill_origin(board: pcbnew.BOARD) -> str | None:
    """
    Check if the drill origin is at the default (0, 0) position.
    Returns a warning message if so, or None if it looks intentionally set.
    """
    settings = board.GetDesignSettings()
    origin = settings.GetAuxOrigin()
    if origin.x == 0 and origin.y == 0:
        return (
            "The drill/auxiliary origin is at (0, 0) — the top-left corner of the page.\n\n"
            "This origin is used as the coordinate reference for SolidWorks sync.\n"
            "If you haven't set it intentionally, go to:\n"
            "  Place → Drill/Place File Origin\n"
            "and place it at a known location on your board (e.g. bottom-left corner)\n"
            "before pushing."
        )
    return None


# ── Layout export ──────────────────────────────────────────────────────────


def _build_layout(board: pcbnew.BOARD, ref_model_map: dict) -> dict:
    settings = board.GetDesignSettings()
    origin = settings.GetAuxOrigin()
    origin_x = pcbnew.ToMM(origin.x)
    origin_y = pcbnew.ToMM(origin.y)

    board_path = board.GetFileName()
    board_filename = Path(board_path).stem if board_path else ""

    components = []
    for fp in board.GetFootprints():
        pos = fp.GetPosition()
        ref = fp.GetReference()
        model_name = ref_model_map.get(ref)  # exact name of the copied STEP file

        comp = {
            "ref": ref,
            "value": fp.GetValue(),
            "footprint": fp.GetFPID().GetUniStringLibItemName(),
            "layer": "F.Cu" if fp.GetLayer() == pcbnew.F_Cu else "B.Cu",
            "position": {
                "x_mm": pcbnew.ToMM(pos.x) - origin_x,
                "y_mm": pcbnew.ToMM(pos.y) - origin_y,
                "rotation_deg": fp.GetOrientationDegrees(),
            },
            "has_3d_model": model_name is not None,
        }

        if model_name is not None:
            # Locate the matching model object to get its offset/rotation
            for m in fp.Models():
                if Path(m.m_Filename).stem == model_name:
                    comp["model_offset"] = {
                        "x_mm": m.m_Offset.x,
                        "y_mm": m.m_Offset.y,
                        "z_mm": m.m_Offset.z,
                    }
                    comp["model_rotation"] = {
                        "x_deg": m.m_Rotation.x,
                        "y_deg": m.m_Rotation.y,
                        "z_deg": m.m_Rotation.z,
                    }
                    comp["model_name"] = model_name
                    break

        components.append(comp)

    return {
        "schema_version": "1.0",
        "origin": {
            "type": "drill",
            "x_mm": origin_x,
            "y_mm": origin_y,
        },
        "board": {
            "thickness_mm": pcbnew.ToMM(settings.GetBoardThickness()),
            "file_name": board_filename,
        },
        "components": components,
    }


# ── Board outline export ──────────────────────────────────────────────────


def _build_board_outline(board: pcbnew.BOARD) -> dict:
    """
    Extract all Edge_Cuts geometry, build closed loops, and classify them
    as outer boundary, holes, slots, or cutouts.
    Also extracts pad drill holes and footprint-level Edge_Cuts.
    """
    settings = board.GetDesignSettings()
    origin = settings.GetAuxOrigin()
    origin_x = pcbnew.ToMM(origin.x)
    origin_y = pcbnew.ToMM(origin.y)

    # Gather all Edge_Cuts segments (board-level and footprint-level)
    raw_segments = []
    for drawing in board.GetDrawings():
        if drawing.GetLayer() != pcbnew.Edge_Cuts:
            continue
        raw_segments.extend(_drawing_to_segments(drawing, origin_x, origin_y))

    for fp in board.GetFootprints():
        for item in fp.GraphicalItems():
            if item.GetLayer() == pcbnew.Edge_Cuts:
                raw_segments.extend(_drawing_to_segments(item, origin_x, origin_y))

    # Build closed loops from the segments
    loops = _build_loops(raw_segments)
    if not loops:
        return {"schema_version": "1.0", "outer_boundary": [], "holes": [],
                "cutouts": [], "drills": []}

    # The outer boundary is the loop with the largest bounding area
    loops.sort(key=_loop_area, reverse=True)
    outer = loops[0]
    inner_loops = loops[1:]

    # Classify inner loops
    holes = []
    cutouts = []
    for loop in inner_loops:
        classified = _classify_inner_loop(loop)
        if classified is not None:
            holes.append(classified)
        else:
            cutouts.append(loop)

    # Extract pad drill holes and mounting holes
    drills         = _extract_drills(board, origin_x, origin_y)
    mounting_holes = _extract_mounting_holes(board, origin_x, origin_y)

    return {
        "schema_version": "1.0",
        "outer_boundary": outer,
        "holes": holes,
        "cutouts": cutouts,
        "drills": drills,
        "mounting_holes": mounting_holes,
    }


def _is_mounting_hole(fp) -> bool:
    """Return True if this footprint is a mounting hole (not a component pad)."""
    lib  = fp.GetFPID().GetLibNickname()
    name = fp.GetFPID().GetUniStringLibItemName()
    return lib == "MountingHole" or "MountingHole" in name


def _extract_mounting_holes(board: pcbnew.BOARD, origin_x: float, origin_y: float) -> list:
    """Extract mounting hole footprints as Hole Wizard candidates."""
    import re
    holes = []
    for fp in board.GetFootprints():
        if not _is_mounting_hole(fp):
            continue
        name = fp.GetFPID().GetUniStringLibItemName()
        m = re.search(r'_(M\d+(?:\.\d+)?)(?:\b|_)', name)
        screw_size = m.group(1) if m else None
        for pad in fp.Pads():
            drill = pad.GetDrillSize()
            if drill.x <= 0:
                continue
            pos = pad.GetPosition()
            holes.append({
                "center": {
                    "x_mm": pcbnew.ToMM(pos.x) - origin_x,
                    "y_mm": pcbnew.ToMM(pos.y) - origin_y,
                },
                "diameter_mm": pcbnew.ToMM(drill.x),
                "screw_size":  screw_size,  # e.g. "M3", or None
            })
            break  # one drill pad per mounting hole footprint
    return holes


def _extract_drills(board: pcbnew.BOARD, origin_x: float, origin_y: float) -> list:
    """Extract drill holes grouped by footprint reference (excludes mounting holes)."""
    footprint_drills = []
    for fp in board.GetFootprints():
        if _is_mounting_hole(fp):
            continue  # handled separately as Hole Wizard features
        pads = []
        for pad in fp.Pads():
            drill = pad.GetDrillSize()
            if drill.x <= 0:
                continue

            pos = pad.GetPosition()
            x = pcbnew.ToMM(pos.x) - origin_x
            y = pcbnew.ToMM(pos.y) - origin_y
            dx = pcbnew.ToMM(drill.x)
            dy = pcbnew.ToMM(drill.y)

            if dx == dy:
                pads.append({
                    "type": "round",
                    "center": {"x_mm": x, "y_mm": y},
                    "diameter_mm": dx,
                })
            else:
                rot_deg = pad.GetOrientationDegrees()
                pads.append({
                    "type": "oval",
                    "center": {"x_mm": x, "y_mm": y},
                    "width_mm": min(dx, dy),
                    "height_mm": max(dx, dy),
                    "angle_deg": rot_deg,
                })

        if pads:
            footprint_drills.append({
                "ref": fp.GetReference(),
                "pads": pads,
            })

    return footprint_drills


def _point(pt, origin_x: float, origin_y: float) -> dict:
    return {"x_mm": pcbnew.ToMM(pt.x) - origin_x, "y_mm": pcbnew.ToMM(pt.y) - origin_y}


def _drawing_to_segments(drawing, origin_x: float, origin_y: float) -> list:
    """Convert a PCB_SHAPE on Edge_Cuts to one or more segment dicts."""
    shape = drawing.GetShape()

    if shape == pcbnew.SHAPE_T_SEGMENT:
        return [{
            "type": "line",
            "start": _point(drawing.GetStart(), origin_x, origin_y),
            "end":   _point(drawing.GetEnd(), origin_x, origin_y),
        }]

    if shape == pcbnew.SHAPE_T_ARC:
        start = _point(drawing.GetStart(), origin_x, origin_y)
        mid   = _point(drawing.GetArcMid(), origin_x, origin_y)
        end   = _point(drawing.GetEnd(), origin_x, origin_y)
        # Skip degenerate arcs where all three points are identical (zero-size pad markers)
        if (abs(start["x_mm"] - end["x_mm"]) < _TOLERANCE and
                abs(start["y_mm"] - end["y_mm"]) < _TOLERANCE and
                abs(start["x_mm"] - mid["x_mm"]) < _TOLERANCE and
                abs(start["y_mm"] - mid["y_mm"]) < _TOLERANCE):
            return []
        return [{"type": "arc", "start": start, "mid": mid, "end": end}]

    if shape == pcbnew.SHAPE_T_CIRCLE:
        center = drawing.GetCenter()
        return [{
            "type": "circle",
            "center": _point(center, origin_x, origin_y),
            "radius_mm": pcbnew.ToMM(drawing.GetRadius()),
        }]

    if shape == pcbnew.SHAPE_T_RECT:
        # Rectangle → 4 line segments
        corners = [_point(c, origin_x, origin_y) for c in drawing.GetCorners()]
        segs = []
        for i in range(len(corners)):
            segs.append({
                "type": "line",
                "start": corners[i],
                "end": corners[(i + 1) % len(corners)],
            })
        return segs

    if shape == pcbnew.SHAPE_T_POLY:
        # Polygon → N line segments
        outline = drawing.GetPolyShape().Outline(0)
        count = outline.PointCount()
        segs = []
        for i in range(count):
            pt_a = outline.CPoint(i)
            pt_b = outline.CPoint((i + 1) % count)
            segs.append({
                "type": "line",
                "start": _point(pt_a, origin_x, origin_y),
                "end": _point(pt_b, origin_x, origin_y),
            })
        return segs

    return []


def _seg_start(seg: dict) -> tuple:
    if seg["type"] == "circle":
        return (seg["center"]["x_mm"], seg["center"]["y_mm"])
    return (seg["start"]["x_mm"], seg["start"]["y_mm"])


def _seg_end(seg: dict) -> tuple:
    if seg["type"] == "circle":
        return (seg["center"]["x_mm"], seg["center"]["y_mm"])
    return (seg["end"]["x_mm"], seg["end"]["y_mm"])


_TOLERANCE = 0.001  # mm


def _points_equal(a: tuple, b: tuple) -> bool:
    return abs(a[0] - b[0]) < _TOLERANCE and abs(a[1] - b[1]) < _TOLERANCE


def _build_loops(segments: list) -> list:
    """
    Build closed loops from unordered segments.
    Circles are their own single-element loops.
    Lines and arcs are chained endpoint-to-endpoint.
    """
    loops = []
    remaining = []

    # Circles are standalone loops
    for seg in segments:
        if seg["type"] == "circle":
            loops.append([seg])
        else:
            remaining.append(seg)

    # Chain non-circle segments into loops
    while remaining:
        loop = [remaining.pop(0)]
        changed = True
        while changed:
            changed = False
            loop_end = _seg_end(loop[-1])
            loop_start = _seg_start(loop[0])

            # Check if loop is already closed
            if len(loop) > 1 and _points_equal(loop_end, loop_start):
                break

            for i, seg in enumerate(remaining):
                seg_s = _seg_start(seg)
                seg_e = _seg_end(seg)

                if _points_equal(loop_end, seg_s):
                    loop.append(remaining.pop(i))
                    changed = True
                    break
                elif _points_equal(loop_end, seg_e):
                    # Reverse the segment
                    loop.append(_reverse_segment(seg))
                    remaining.pop(i)
                    changed = True
                    break

        loops.append(loop)

    return loops


def _reverse_segment(seg: dict) -> dict:
    if seg["type"] == "line":
        return {"type": "line", "start": seg["end"], "end": seg["start"]}
    if seg["type"] == "arc":
        return {"type": "arc", "start": seg["end"], "mid": seg["mid"], "end": seg["start"]}
    return seg


def _loop_area(loop: list) -> float:
    """Approximate bounding-box area for sorting loops by size."""
    points = []
    for seg in loop:
        if seg["type"] == "circle":
            c = seg["center"]
            r = seg["radius_mm"]
            return math.pi * r * r  # exact area for circles
        points.append(_seg_start(seg))
        points.append(_seg_end(seg))
        if seg["type"] == "arc" and "mid" in seg:
            points.append((seg["mid"]["x_mm"], seg["mid"]["y_mm"]))

    if not points:
        return 0.0
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return (max(xs) - min(xs)) * (max(ys) - min(ys))


def _classify_inner_loop(loop: list) -> dict | None:
    """
    Classify an inner loop as a hole or slot.
    Returns a hole/slot dict, or None if it's an arbitrary cutout.
    """
    # Single circle → round hole
    if len(loop) == 1 and loop[0]["type"] == "circle":
        seg = loop[0]
        return {
            "type": "round",
            "center": seg["center"],
            "diameter_mm": seg["radius_mm"] * 2,
        }

    # Slot detection: exactly 2 arcs + 2 lines (pill shape)
    if len(loop) == 4:
        arcs = [s for s in loop if s["type"] == "arc"]
        lines = [s for s in loop if s["type"] == "line"]
        if len(arcs) == 2 and len(lines) == 2:
            return _classify_as_slot(arcs, lines)

    return None


def _classify_as_slot(arcs: list, lines: list) -> dict | None:
    """
    Try to classify 2 arcs + 2 lines as a slot (pill/stadium shape).
    Returns a slot dict or None if the geometry doesn't match.
    """
    # Both arcs should have similar radius (the slot end radii)
    def arc_radius(arc):
        sx, sy = arc["start"]["x_mm"], arc["start"]["y_mm"]
        mx, my = arc["mid"]["x_mm"], arc["mid"]["y_mm"]
        ex, ey = arc["end"]["x_mm"], arc["end"]["y_mm"]
        # Radius from 3 points: circumscribed circle
        ax, ay = sx - mx, sy - my
        bx, by = ex - mx, ey - my
        cross = ax * by - ay * bx
        if abs(cross) < 1e-9:
            return None
        a2 = ax * ax + ay * ay
        b2 = bx * bx + by * by
        cx = (a2 * by - b2 * ay) / (2 * cross)
        cy = (b2 * ax - a2 * bx) / (2 * cross)
        return math.sqrt(cx * cx + cy * cy)

    r1 = arc_radius(arcs[0])
    r2 = arc_radius(arcs[1])
    if r1 is None or r2 is None:
        return None
    if abs(r1 - r2) > _TOLERANCE:
        return None

    width = r1 * 2

    # Slot center is midpoint of the two arc centers
    def arc_center(arc):
        sx, sy = arc["start"]["x_mm"], arc["start"]["y_mm"]
        mx, my = arc["mid"]["x_mm"], arc["mid"]["y_mm"]
        ex, ey = arc["end"]["x_mm"], arc["end"]["y_mm"]
        ax, ay = sx - mx, sy - my
        bx, by = ex - mx, ey - my
        cross = ax * by - ay * bx
        a2 = ax * ax + ay * ay
        b2 = bx * bx + by * by
        cx = mx + (a2 * by - b2 * ay) / (2 * cross)
        cy = my + (b2 * ax - a2 * bx) / (2 * cross)
        return cx, cy

    c1x, c1y = arc_center(arcs[0])
    c2x, c2y = arc_center(arcs[1])

    center_x = (c1x + c2x) / 2
    center_y = (c1y + c2y) / 2

    # Length is distance between arc centers + width
    dist = math.sqrt((c2x - c1x) ** 2 + (c2y - c1y) ** 2)
    length = dist + width

    # Angle of the long axis
    angle = math.degrees(math.atan2(c2y - c1y, c2x - c1x))

    return {
        "type": "slot",
        "center": {"x_mm": center_x, "y_mm": center_y},
        "width_mm": width,
        "length_mm": length,
        "angle_deg": angle,
    }
