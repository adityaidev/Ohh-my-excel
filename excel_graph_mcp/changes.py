from excel_graph_mcp.graph import GraphStore


class ChangeAnalyzer:
    def __init__(self, old_store: GraphStore, new_store: GraphStore):
        self.old_store = old_store
        self.new_store = new_store

    def detect_changes(self) -> dict:
        old_edges = set()
        for edge in self._get_all_edges(self.old_store):
            old_edges.add((edge["source_id"], edge["target_id"], edge["edge_type"]))

        new_edges = set()
        for edge in self._get_all_edges(self.new_store):
            new_edges.add((edge["source_id"], edge["target_id"], edge["edge_type"]))

        added = new_edges - old_edges
        removed = old_edges - new_edges

        old_nodes = {n["id"] for n in self._get_all_nodes(self.old_store)}
        new_nodes = {n["id"] for n in self._get_all_nodes(self.new_store)}

        cells_added = len(new_nodes - old_nodes)
        cells_removed = len(old_nodes - new_nodes)

        risk_score = len(added) + len(removed) + cells_added + cells_removed

        return {
            "risk_score": risk_score,
            "risk_level": "high" if risk_score > 20 else "medium" if risk_score > 5 else "low",
            "edges_added": len(added),
            "edges_removed": len(removed),
            "cells_added": cells_added,
            "cells_removed": cells_removed,
        }

    def _get_all_edges(self, store):
        conn = store._conn()
        return [dict(r) for r in conn.execute("SELECT * FROM edges").fetchall()]

    def _get_all_nodes(self, store):
        conn = store._conn()
        return [dict(r) for r in conn.execute("SELECT id FROM nodes").fetchall()]
