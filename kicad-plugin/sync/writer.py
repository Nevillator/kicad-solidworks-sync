"""Export KiCad board data to the sync directory (ecad_to_mcad)."""

import json
import math
import subprocess
from pathlib import Path

import pcbnew


def push(board: pcbnew.BOARD, sync_dir: Path) -> list:
    """
    Export board STEP, layout JSON, and board outline JSON to sync_dir/ecad_to_mcad/.
    Returns a list of change records.
    """
    out_dir = sync_dir / "ecad_to_mcad"
    out_dir.mkdir(parents=True, exist_ok=True)

    changes = []

    # Export STEP (for component 3D models)
    step_path = out_dir / "board.step"
    _export_step(board, step_path)
    changes.append({"type": "3d_model_updated"})

    # Export layout JSON (component positions, origins, model offsets)
    layout = _build_layout(board)
    with open(out_dir / "layout.json", "w") as f:
        json.dump(layout, f, indent=2)

    # Export board outline JSON (outer boundary, holes, slots, cutouts)
    outline = _build_board_outline(board)
    with open(out_dir / "board_outline.json", "w") as f:
        json.dump(outline, f, indent=2)
    changes.append({"type": "board_outline_updated"})

    return changes


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


# ── STEP export ────────────────────────────────────────────────────────────


def _export_step(board: pcbnew.BOARD, out_path: Path):
    board_path = board.GetFileName()
    if not board_path:
        raise RuntimeError("Board must be saved to disk before exporting STEP.")
    subprocess.run([
        "kicad-cli", "pcb", "export", "step",
        "--output", str(out_path),
        "--drill-origin",
        "--subst-models",
        "--no-dnp",
        str(board_path),
    ], check=True)


# ── Layout export ──────────────────────────────────────────────────────────


def _build_layout(board: pcbnew.BOARD) -> dict:
    settings = board.GetDesignSettings()
    origin = settings.GetAuxOrigin()
    origin_x = pcbnew.ToMM(origin.x)
    origin_y = pcbnew.ToMM(origin.y)

    board_path = board.GetFileName()
    board_filename = Path(board_path).stem if board_path else ""

    components = []
    for fp in board.GetFootprints():
        pos = fp.GetPosition()

        comp = {
            "ref": fp.GetReference(),
            "value": fp.GetValue(),
            "footprint": fp.GetFPID().GetUniStringLibItemName(),
            "layer": "F.Cu" if fp.GetLayer() == pcbnew.F_Cu else "B.Cu",
            "position": {
                "x_mm": pcbnew.ToMM(pos.x) - origin_x,
                "y_mm": pcbnew.ToMM(pos.y) - origin_y,
                "rotation_deg": fp.GetOrientationDegrees(),
            },
            "has_3d_model": len(fp.Models()) > 0,
        }

        # Add 3D model offset and rotation if a model exists
        models = fp.Models()
        if len(models) > 0:
            model = models[0]  # primary 3D model
            comp["model_offset"] = {
                "x_mm": model.m_Offset.x,
                "y_mm": model.m_Offset.y,
                "z_mm": model.m_Offset.z,
            }
            comp["model_rotation"] = {
                "x_deg": model.m_Rotation.x,
                "y_deg": model.m_Rotation.y,
                "z_deg": model.m_Rotation.z,
            }

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
    """
    settings = board.GetDesignSettings()
    origin = settings.GetAuxOrigin()
    origin_x = pcbnew.ToMM(origin.x)
    origin_y = pcbnew.ToMM(origin.y)

    # Gather all Edge_Cuts segments
    raw_segments = []
    for drawing in board.GetDrawings():
        if drawing.GetLayer() != pcbnew.Edge_Cuts:
            continue
        seg = _drawing_to_segment(drawing, origin_x, origin_y)
        if seg is not None:
            raw_segments.append(seg)

    # Build closed loops from the segments
    loops = _build_loops(raw_segments)
    if not loops:
        return {"schema_version": "1.0", "outer_boundary": [], "holes": [], "cutouts": []}

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

    return {
        "schema_version": "1.0",
        "outer_boundary": outer,
        "holes": holes,
        "cutouts": cutouts,
    }


def _drawing_to_segment(drawing, origin_x: float, origin_y: float) -> dict | None:
    """Convert a PCB_SHAPE on Edge_Cuts to a segment dict."""
    shape = drawing.GetShape()

    if shape == pcbnew.SHAPE_T_SEGMENT:
        start = drawing.GetStart()
        end = drawing.GetEnd()
        return {
            "type": "line",
            "start": {"x_mm": pcbnew.ToMM(start.x) - origin_x,
                       "y_mm": pcbnew.ToMM(start.y) - origin_y},
            "end":   {"x_mm": pcbnew.ToMM(end.x) - origin_x,
                       "y_mm": pcbnew.ToMM(end.y) - origin_y},
        }

    if shape == pcbnew.SHAPE_T_ARC:
        start = drawing.GetStart()
        mid = drawing.GetArcMid()
        end = drawing.GetEnd()
        return {
            "type": "arc",
            "start": {"x_mm": pcbnew.ToMM(start.x) - origin_x,
                       "y_mm": pcbnew.ToMM(start.y) - origin_y},
            "mid":   {"x_mm": pcbnew.ToMM(mid.x) - origin_x,
                       "y_mm": pcbnew.ToMM(mid.y) - origin_y},
            "end":   {"x_mm": pcbnew.ToMM(end.x) - origin_x,
                       "y_mm": pcbnew.ToMM(end.y) - origin_y},
        }

    if shape == pcbnew.SHAPE_T_CIRCLE:
        center = drawing.GetCenter()
        radius = pcbnew.ToMM(drawing.GetRadius())
        return {
            "type": "circle",
            "center": {"x_mm": pcbnew.ToMM(center.x) - origin_x,
                        "y_mm": pcbnew.ToMM(center.y) - origin_y},
            "radius_mm": radius,
        }

    return None


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
