# kicad-solidworks-sync

Bidirectional ECAD/MCAD synchronization between KiCad and SolidWorks, inspired by Altium CoDesigner.

## What it does

**KiCad → SolidWorks**
- Exports the board and all components with 3D models as a STEP assembly
- Exports component positions, references, and layer assignments

**SolidWorks → KiCad**
- Exports the board outline (including holes, cutouts, features added in SW)
- Exports updated component positions/rotations after mechanical layout

## Architecture

Two plugins communicate through a shared **sync directory** (local folder, shared drive, or git repo):

```
sync-dir/
  manifest.json          ← change log, timestamps, comments
  ecad_to_mcad/
    board.step           ← board + components (KiCad STEP export)
    layout.json          ← component positions, refs, rotations
  mcad_to_ecad/
    board_outline.json   ← outline geometry from SolidWorks
    component_moves.json ← updated component positions
```

## Components

- `kicad-plugin/` — Python plugin for KiCad (push/pull buttons in PCB editor)
- `solidworks-addin/` — C# SolidWorks add-in (push/pull panel in SolidWorks)
- `sync-schema/` — JSON schemas defining the sync data formats
- `docs/` — Architecture and developer notes

## Setup

See `docs/setup.md`.
