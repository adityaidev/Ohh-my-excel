import sqlite3
import threading
from pathlib import Path
from typing import Any, Optional

import networkx as nx

from excel_graph_mcp.constants import (
    SCHEMA_VERSION,
    get_db_path,
    get_graph_dir,
)


class GraphStore:
    def __init__(self, workbook_path: Path):
        self.workbook_path = workbook_path
        self.db_path = get_db_path(workbook_path)
        self._local = threading.local()
        self._cache: Optional[nx.DiGraph] = None
        self._lock = threading.RLock()
        self._ensure_dir()
        self._init_db()

    def _ensure_dir(self):
        get_graph_dir(self.workbook_path).mkdir(parents=True, exist_ok=True)

    def _conn(self) -> sqlite3.Connection:
        if not getattr(self._local, "conn", None):
            self._local.conn = sqlite3.connect(str(self.db_path))
            self._local.conn.execute("PRAGMA journal_mode=WAL")
            self._local.conn.execute("PRAGMA synchronous=NORMAL")
            self._local.conn.row_factory = sqlite3.Row
        return self._local.conn

    def _init_db(self):
        conn = self._conn()
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS meta (
                key TEXT PRIMARY KEY,
                value TEXT
            );
            CREATE TABLE IF NOT EXISTS nodes (
                id TEXT PRIMARY KEY,
                type TEXT NOT NULL,
                sheet TEXT,
                data TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                updated_at TEXT DEFAULT (datetime('now'))
            );
            CREATE TABLE IF NOT EXISTS edges (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_id TEXT NOT NULL,
                target_id TEXT NOT NULL,
                edge_type TEXT NOT NULL,
                confidence TEXT DEFAULT 'EXTRACTED',
                data TEXT,
                FOREIGN KEY (source_id) REFERENCES nodes(id),
                FOREIGN KEY (target_id) REFERENCES nodes(id)
            );
            CREATE INDEX IF NOT EXISTS idx_edges_source ON edges(source_id);
            CREATE INDEX IF NOT EXISTS idx_edges_target ON edges(target_id);
            CREATE INDEX IF NOT EXISTS idx_nodes_type ON nodes(type);
            CREATE INDEX IF NOT EXISTS idx_nodes_sheet ON nodes(sheet);
            CREATE VIRTUAL TABLE IF NOT EXISTS nodes_fts USING fts5(
                id, type, sheet, data,
                content='nodes', content_rowid='rowid'
            );
        """)
        conn.execute(
            "INSERT OR IGNORE INTO meta (key, value) VALUES (?, ?)",
            ("schema_version", str(SCHEMA_VERSION)),
        )
        conn.commit()

    def add_node(self, node_id: str, node_type: str, sheet: Optional[str] = None, data: Optional[dict] = None):
        with self._lock:
            self._conn().execute(
                "INSERT OR REPLACE INTO nodes (id, type, sheet, data) VALUES (?, ?, ?, ?)",
                (node_id, node_type, sheet, json_dumps(data)),
            )
            self._conn().commit()
            self._cache = None

    def add_edge(self, source_id: str, target_id: str, edge_type: str, confidence: str = "EXTRACTED", data: Optional[dict] = None):
        with self._lock:
            self._conn().execute(
                "INSERT INTO edges (source_id, target_id, edge_type, confidence, data) VALUES (?, ?, ?, ?, ?)",
                (source_id, target_id, edge_type, confidence, json_dumps(data)),
            )
            self._conn().commit()
            self._cache = None

    def get_node(self, node_id: str) -> Optional[dict]:
        row = self._conn().execute("SELECT * FROM nodes WHERE id = ?", (node_id,)).fetchone()
        return dict(row) if row else None

    def get_edges(self, node_id: str, direction: str = "outgoing") -> list[dict]:
        if direction == "outgoing":
            rows = self._conn().execute(
                "SELECT * FROM edges WHERE source_id = ?", (node_id,)
            ).fetchall()
        else:
            rows = self._conn().execute(
                "SELECT * FROM edges WHERE target_id = ?", (node_id,)
            ).fetchall()
        return [dict(r) for r in rows]

    def search(self, query: str, limit: int = 20) -> list[dict]:
        rows = self._conn().execute(
            "SELECT n.* FROM nodes n JOIN nodes_fts f ON n.rowid = f.rowid "
            "WHERE nodes_fts MATCH ? LIMIT ?",
            (query, limit),
        ).fetchall()
        return [dict(r) for r in rows]

    def bfs(self, start_id: str, max_depth: int = 2, direction: str = "forward") -> list[dict]:
        result = []
        visited = {start_id}
        queue = [(start_id, 0)]
        while queue:
            current, depth = queue.pop(0)
            if depth >= max_depth:
                continue
            edges = self.get_edges(current, direction)
            for edge in edges:
                neighbor = edge["target_id"] if direction == "outgoing" else edge["source_id"]
                if neighbor not in visited:
                    visited.add(neighbor)
                    node = self.get_node(neighbor)
                    if node:
                        result.append({"node": node, "edge": edge, "depth": depth + 1})
                    queue.append((neighbor, depth + 1))
        return result

    def stats(self) -> dict:
        conn = self._conn()
        nodes = conn.execute("SELECT COUNT(*) FROM nodes").fetchone()[0]
        edges = conn.execute("SELECT COUNT(*) FROM edges").fetchone()[0]
        type_counts = dict(
            conn.execute("SELECT type, COUNT(*) FROM nodes GROUP BY type").fetchall()
        )
        return {"nodes": nodes, "edges": edges, "by_type": type_counts}

    def to_networkx(self) -> nx.DiGraph:
        if self._cache is not None:
            return self._cache
        G = nx.DiGraph()
        for row in self._conn().execute("SELECT * FROM nodes").fetchall():
            G.add_node(row["id"], type=row["type"], sheet=row["sheet"], data=row["data"])
        for row in self._conn().execute("SELECT * FROM edges").fetchall():
            G.add_edge(
                row["source_id"], row["target_id"],
                edge_type=row["edge_type"],
                confidence=row["confidence"],
            )
        self._cache = G
        return G

    def clear(self):
        with self._lock:
            self._conn().executescript("DELETE FROM nodes; DELETE FROM edges; DELETE FROM nodes_fts;")
            self._conn().commit()
            self._cache = None

    def close(self):
        if getattr(self._local, "conn", None):
            self._local.conn.close()
            self._local.conn = None


def json_dumps(obj: Any) -> Optional[str]:
    import json
    return json.dumps(obj) if obj is not None else None
