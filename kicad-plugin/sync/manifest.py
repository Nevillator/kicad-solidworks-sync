"""Read and write the sync manifest (change log)."""

import json
import uuid
from datetime import datetime, timezone
from pathlib import Path


def load(sync_dir: Path) -> dict:
    path = sync_dir / "manifest.json"
    if not path.exists():
        return {"schema_version": "1.0", "history": []}
    with open(path) as f:
        return json.load(f)


def save(sync_dir: Path, manifest: dict):
    path = sync_dir / "manifest.json"
    with open(path, "w") as f:
        json.dump(manifest, f, indent=2)


def record_push(manifest: dict, direction: str, author: str, comment: str, changes: list) -> dict:
    """Append a push record and update last_push. Returns the updated manifest."""
    record = {
        "id": str(uuid.uuid4()),
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "direction": direction,
        "author": author,
        "comment": comment,
        "changes": changes,
    }
    manifest["last_push"] = record
    manifest.setdefault("history", []).append(record)
    return manifest
