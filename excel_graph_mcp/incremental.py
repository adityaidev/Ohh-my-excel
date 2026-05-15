from pathlib import Path
import hashlib
import json

from excel_graph_mcp.graph import GraphStore
from excel_graph_mcp.constants import get_graph_dir


class IncrementalUpdater:
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.graph_dir = get_graph_dir(file_path)
        self.hash_file = self.graph_dir / "hashes.json"
        self.graph_dir.mkdir(parents=True, exist_ok=True)
        self._hashes = self._load_hashes()

    def _load_hashes(self) -> dict:
        if self.hash_file.exists():
            return json.loads(self.hash_file.read_text())
        return {}

    def _save_hashes(self):
        self.hash_file.write_text(json.dumps(self._hashes, indent=2))

    def _hash_file(self) -> str:
        return hashlib.sha256(self.file_path.read_bytes()).hexdigest()

    def needs_update(self) -> bool:
        current = self._hash_file()
        previous = self._hashes.get("file_hash")
        return previous != current

    def update(self) -> dict:
        from excel_graph_mcp.dependency import build_dependency_graph
        current = self._hash_file()
        previous = self._hashes.get("file_hash")
        if previous == current:
            return {"status": "unchanged", "file": str(self.file_path)}
        store, builder = build_dependency_graph(self.file_path)
        stats = store.stats()
        self._hashes["file_hash"] = current
        self._save_hashes()
        store.close()
        return {
            "status": "updated",
            "file": str(self.file_path),
            "nodes": stats["nodes"],
            "edges": stats["edges"],
            "was_full_rebuild": previous is None,
        }


def watch_file(file_path: str, callback=None):
    import time
    p = Path(file_path)
    updater = IncrementalUpdater(p)
    last_mtime = p.stat().st_mtime
    while True:
        current_mtime = p.stat().st_mtime
        if current_mtime != last_mtime:
            time.sleep(0.5)
            result = updater.update()
            if callback:
                callback(result)
            last_mtime = current_mtime
        time.sleep(1)
