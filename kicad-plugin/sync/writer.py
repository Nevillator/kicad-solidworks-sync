"""Export KiCad board data to the sync directory (ecad_to_mcad)."""

import json
import subprocess
from pathlib import Path

import pcbnew


def push(board: pcbnew.BOARD, sync_dir: Path) -> list:
    """
    Export board STEP and layout JSON to sync_dir/ecad_to_mcad/.
    Returns a list of change records.
    """
    out_dir = sync_dir / "ecad_to_mcad"
    out_dir.mkdir(parents=True, exist_ok=True)

    changes = []

    # Export STEP
    step_path = out_dir / "board.step"
    _export_step(board, step_path)
    changes.append({"type": "3d_model_updated"})

    # Export layout JSON
    layout = _build_layout(board)
    layout_path = out_dir / "layout.json"
    with open(layout_path, "w") as f:
        json.dump(layout, f, indent=2)

    return changes


def _export_step(board: pcbnew.BOARD, out_path: Path):
    exporter = pcbnew.STEP_PCB_MODEL(board)
    exporter.SetBoardOnly(False)  # include components
    exporter.SetIncludeUnspecified(False)
    exporter.SetIncludeDNP(False)
    exporter.Export(str(out_path))


def _build_layout(board: pcbnew.BOARD) -> dict:
    components = []
    for fp in board.GetFootprints():
        pos = fp.GetPosition()
        components.append({
            "ref": fp.GetReference(),
            "value": fp.GetValue(),
            "footprint": fp.GetFPID().GetUniStringLibItemName(),
            "layer": "F.Cu" if fp.GetLayer() == pcbnew.F_Cu else "B.Cu",
            "position": {
                "x_mm": pcbnew.ToMM(pos.x),
                "y_mm": pcbnew.ToMM(pos.y),
                "rotation_deg": fp.GetOrientationDegrees(),
            },
            "has_3d_model": len(fp.Models()) > 0,
        })

    return {
        "schema_version": "1.0",
        "board": {
            "thickness_mm": pcbnew.ToMM(board.GetDesignSettings().GetBoardThickness()),
        },
        "components": components,
    }
