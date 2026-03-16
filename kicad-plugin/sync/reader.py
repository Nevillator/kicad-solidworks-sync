"""Read sync data from mcad_to_ecad/ and apply changes to the KiCad board."""

import json
import math
from pathlib import Path

import pcbnew


def get_pending_changes(sync_dir: Path) -> dict | None:
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
    selected_changes: list of change dicts from get_pending_changes(), filtered by user.
    Returns list of applied changes.
    """
    in_dir = sync_dir / "mcad_to_ecad"
    applied = []

    change_types = {c["type"] for c in selected_changes}

    if "board_outline_updated" in change_types:
        outline_path = in_dir / "board_outline.json"
        if outline_path.exists():
            with open(outline_path) as f:
                outline_data = json.load(f)
            _apply_board_outline(board, outline_data)
            applied.append({"type": "board_outline_updated"})

    moves = {c["ref"]: c for c in selected_changes if c["type"] == "component_moved"}
    if moves:
        moves_path = in_dir / "component_moves.json"
        if moves_path.exists():
            with open(moves_path) as f:
                moves_data = json.load(f)
            for m in moves_data.get("components", []):
                if m["ref"] in moves:
                    _apply_component_move(board, m)
                    applied.append({"type": "component_moved", "ref": m["ref"]})

    pcbnew.Refresh()
    return applied


def _apply_board_outline(board: pcbnew.BOARD, outline_data: dict):
    # Remove existing edge cuts
    for drawing in list(board.GetDrawings()):
        if drawing.GetLayer() == pcbnew.Edge_Cuts:
            board.Remove(drawing)

    for seg in outline_data.get("segments", []):
        if seg["type"] == "line":
            line = pcbnew.PCB_SHAPE(board)
            line.SetShape(pcbnew.SHAPE_T_SEGMENT)
            line.SetLayer(pcbnew.Edge_Cuts)
            line.SetStart(pcbnew.FromMM(seg["start"]["x_mm"]), pcbnew.FromMM(seg["start"]["y_mm"]))
            line.SetEnd(pcbnew.FromMM(seg["end"]["x_mm"]), pcbnew.FromMM(seg["end"]["y_mm"]))
            board.Add(line)

        elif seg["type"] == "arc":
            arc = pcbnew.PCB_SHAPE(board)
            arc.SetShape(pcbnew.SHAPE_T_ARC)
            arc.SetLayer(pcbnew.Edge_Cuts)
            arc.SetCenter(pcbnew.FromMM(seg["center"]["x_mm"]), pcbnew.FromMM(seg["center"]["y_mm"]))
            arc.SetRadius(pcbnew.FromMM(seg["radius_mm"]))
            arc.SetArcAngleAndEnd(seg["start_angle_deg"] * 10, seg["end_angle_deg"] * 10)
            board.Add(arc)

    for hole in outline_data.get("holes", []):
        circle = pcbnew.PCB_SHAPE(board)
        circle.SetShape(pcbnew.SHAPE_T_CIRCLE)
        circle.SetLayer(pcbnew.Edge_Cuts)
        circle.SetCenter(pcbnew.FromMM(hole["center"]["x_mm"]), pcbnew.FromMM(hole["center"]["y_mm"]))
        circle.SetRadius(pcbnew.FromMM(hole["diameter_mm"] / 2))
        board.Add(circle)


def _apply_component_move(board: pcbnew.BOARD, move: dict):
    for fp in board.GetFootprints():
        if fp.GetReference() == move["ref"]:
            pos = move["position"]
            fp.SetPosition(pcbnew.VECTOR2I(
                pcbnew.FromMM(pos["x_mm"]),
                pcbnew.FromMM(pos["y_mm"]),
            ))
            fp.SetOrientationDegrees(pos["rotation_deg"])
            break
