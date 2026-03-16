"""Read sync data from mcad_to_ecad/ and apply changes to the KiCad board."""

import json
import math
from pathlib import Path

import pcbnew


def get_pending_changes(sync_dir: Path) -> list | None:
    """
    Read the mcad_to_ecad directory and return a summary of pending changes,
    or None if there's nothing new.
    """
    in_dir = sync_dir / "mcad_to_ecad"
    outline_path = in_dir / "board_outline.json"
    moves_path = in_dir / "component_moves.json"

    if not outline_path.exists() and not moves_path.exists():
        return None

    changes = []

    if outline_path.exists():
        changes.append({"type": "board_outline_updated"})

    if moves_path.exists():
        with open(moves_path) as f:
            moves = json.load(f)
        for m in moves.get("components", []):
            changes.append({
                "type": "component_moved",
                "ref": m["ref"],
                "to": m["position"],
            })

    return changes if changes else None


def apply(board: pcbnew.BOARD, sync_dir: Path, selected_changes: list) -> list:
    """
    Apply selected changes from mcad_to_ecad/ to the board.
    Returns list of applied changes.
    """
    in_dir = sync_dir / "mcad_to_ecad"

    # Read origin from the last push's layout.json (ecad_to_mcad) so we can
    # convert coordinates back to board-absolute positions.
    origin_x, origin_y = _read_origin(sync_dir)

    applied = []
    change_types = {c["type"] for c in selected_changes}

    if "board_outline_updated" in change_types:
        outline_path = in_dir / "board_outline.json"
        if outline_path.exists():
            with open(outline_path) as f:
                outline_data = json.load(f)
            _apply_board_outline(board, outline_data, origin_x, origin_y)
            applied.append({"type": "board_outline_updated"})

    moves = {c["ref"]: c for c in selected_changes if c["type"] == "component_moved"}
    if moves:
        moves_path = in_dir / "component_moves.json"
        if moves_path.exists():
            with open(moves_path) as f:
                moves_data = json.load(f)
            for m in moves_data.get("components", []):
                if m["ref"] in moves:
                    _apply_component_move(board, m, origin_x, origin_y)
                    applied.append({"type": "component_moved", "ref": m["ref"]})

    pcbnew.Refresh()
    return applied


def _read_origin(sync_dir: Path) -> tuple:
    """Read the drill origin from the last KiCad push's layout.json."""
    layout_path = sync_dir / "ecad_to_mcad" / "layout.json"
    if layout_path.exists():
        with open(layout_path) as f:
            layout = json.load(f)
        origin = layout.get("origin", {})
        return origin.get("x_mm", 0.0), origin.get("y_mm", 0.0)
    return 0.0, 0.0


# ── Board outline application ─────────────────────────────────────────────


def _apply_board_outline(board: pcbnew.BOARD, outline_data: dict,
                         origin_x: float, origin_y: float):
    """Replace all Edge_Cuts geometry with the incoming outline data."""
    # Remove existing Edge_Cuts
    for drawing in list(board.GetDrawings()):
        if drawing.GetLayer() == pcbnew.Edge_Cuts:
            board.Remove(drawing)

    # Draw outer boundary
    for seg in outline_data.get("outer_boundary", []):
        shape = _seg_to_pcb_shape(board, seg, origin_x, origin_y)
        if shape:
            board.Add(shape)

    # Draw holes
    for hole in outline_data.get("holes", []):
        if hole["type"] == "round":
            _add_circle(board, hole["center"], hole["diameter_mm"] / 2, origin_x, origin_y)
        elif hole["type"] == "slot":
            for seg in _slot_to_segments(hole):
                shape = _seg_to_pcb_shape(board, seg, origin_x, origin_y)
                if shape:
                    board.Add(shape)

    # Draw cutouts
    for cutout in outline_data.get("cutouts", []):
        for seg in cutout:
            shape = _seg_to_pcb_shape(board, seg, origin_x, origin_y)
            if shape:
                board.Add(shape)


def _seg_to_pcb_shape(board: pcbnew.BOARD, seg: dict,
                      origin_x: float, origin_y: float) -> pcbnew.PCB_SHAPE | None:
    """Convert a segment dict to a KiCad PCB_SHAPE on Edge_Cuts."""
    if seg["type"] == "line":
        shape = pcbnew.PCB_SHAPE(board)
        shape.SetShape(pcbnew.SHAPE_T_SEGMENT)
        shape.SetLayer(pcbnew.Edge_Cuts)
        shape.SetStart(_to_vec(seg["start"], origin_x, origin_y))
        shape.SetEnd(_to_vec(seg["end"], origin_x, origin_y))
        return shape

    if seg["type"] == "arc":
        shape = pcbnew.PCB_SHAPE(board)
        shape.SetShape(pcbnew.SHAPE_T_ARC)
        shape.SetLayer(pcbnew.Edge_Cuts)
        shape.SetArcGeometry(
            _to_vec(seg["start"], origin_x, origin_y),
            _to_vec(seg["mid"], origin_x, origin_y),
            _to_vec(seg["end"], origin_x, origin_y),
        )
        return shape

    if seg["type"] == "circle":
        shape = pcbnew.PCB_SHAPE(board)
        shape.SetShape(pcbnew.SHAPE_T_CIRCLE)
        shape.SetLayer(pcbnew.Edge_Cuts)
        center = _to_vec(seg["center"], origin_x, origin_y)
        shape.SetCenter(center)
        shape.SetRadius(pcbnew.FromMM(seg["radius_mm"]))
        return shape

    return None


def _add_circle(board: pcbnew.BOARD, center: dict, radius_mm: float,
                origin_x: float, origin_y: float):
    shape = pcbnew.PCB_SHAPE(board)
    shape.SetShape(pcbnew.SHAPE_T_CIRCLE)
    shape.SetLayer(pcbnew.Edge_Cuts)
    shape.SetCenter(_to_vec(center, origin_x, origin_y))
    shape.SetRadius(pcbnew.FromMM(radius_mm))
    board.Add(shape)


def _slot_to_segments(slot: dict) -> list:
    """Expand a slot definition into 2 arcs + 2 lines (pill shape)."""
    cx = slot["center"]["x_mm"]
    cy = slot["center"]["y_mm"]
    w = slot["width_mm"]
    l = slot["length_mm"]
    angle = math.radians(slot["angle_deg"])
    r = w / 2

    # Half-length between arc centers
    half = (l - w) / 2

    # Direction vectors
    dx = math.cos(angle)
    dy = math.sin(angle)
    nx = -dy  # normal
    ny = dx

    # Arc centers
    c1x, c1y = cx - half * dx, cy - half * dy
    c2x, c2y = cx + half * dx, cy + half * dy

    # Four corner points (where lines meet arcs)
    p1 = {"x_mm": c1x + r * nx, "y_mm": c1y + r * ny}
    p2 = {"x_mm": c2x + r * nx, "y_mm": c2y + r * ny}
    p3 = {"x_mm": c2x - r * nx, "y_mm": c2y - r * ny}
    p4 = {"x_mm": c1x - r * nx, "y_mm": c1y - r * ny}

    # Midpoints of arcs
    m1 = {"x_mm": c1x - half * dx / max(half, 1e-9) * r if half > 1e-9 else c1x - r * dx,
           "y_mm": c1y - half * dy / max(half, 1e-9) * r if half > 1e-9 else c1y - r * dy}
    m2 = {"x_mm": c2x + half * dx / max(half, 1e-9) * r if half > 1e-9 else c2x + r * dx,
           "y_mm": c2y + half * dy / max(half, 1e-9) * r if half > 1e-9 else c2y + r * dy}

    # Simpler: arc midpoints are at the ends of the slot axis
    m1 = {"x_mm": c1x - r * dx, "y_mm": c1y - r * dy}
    m2 = {"x_mm": c2x + r * dx, "y_mm": c2y + r * dy}

    return [
        {"type": "line", "start": p1, "end": p2},
        {"type": "arc", "start": p2, "mid": m2, "end": p3},
        {"type": "line", "start": p3, "end": p4},
        {"type": "arc", "start": p4, "mid": m1, "end": p1},
    ]


def _to_vec(pt: dict, origin_x: float, origin_y: float) -> pcbnew.VECTOR2I:
    """Convert a point dict (relative to drill origin) to a KiCad VECTOR2I (board-absolute)."""
    return pcbnew.VECTOR2I(
        pcbnew.FromMM(pt["x_mm"] + origin_x),
        pcbnew.FromMM(pt["y_mm"] + origin_y),
    )


# ── Component position application ────────────────────────────────────────


def _apply_component_move(board: pcbnew.BOARD, move: dict,
                          origin_x: float, origin_y: float):
    for fp in board.GetFootprints():
        if fp.GetReference() == move["ref"]:
            pos = move["position"]
            fp.SetPosition(pcbnew.VECTOR2I(
                pcbnew.FromMM(pos["x_mm"] + origin_x),
                pcbnew.FromMM(pos["y_mm"] + origin_y),
            ))
            fp.SetOrientationDegrees(pos["rotation_deg"])
            break
