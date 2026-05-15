from excel_graph_mcp.graph import GraphStore
import networkx as nx


class CommunityDetector:
    def __init__(self, store: GraphStore):
        self.store = store
        self.G = store.to_networkx()

    def detect_communities(self) -> list[dict]:
        sheets = [n for n, d in self.G.nodes(data=True) if d.get("type") == "Sheet"]
        communities = {}
        for sheet in sheets:
            cross_refs = self._count_cross_refs(sheet)
            assigned = False
            for comm_id in communities:
                if any(self._are_connected(sheet, m) for m in communities[comm_id]):
                    communities[comm_id].append(sheet)
                    assigned = True
                    break
            if not assigned:
                communities[f"comm_{len(communities)}"] = [sheet]
        result = []
        for comm_id, members in communities.items():
            result.append({
                "id": comm_id,
                "members": members,
                "size": len(members),
                "cohesion": round(len(members) / len(sheets), 2) if sheets else 0,
            })
        return result

    def _count_cross_refs(self, sheet_id: str) -> int:
        count = 0
        for cell_id in self.G.successors(sheet_id):
            for _, tgt, data in self.G.out_edges(cell_id, data=True):
                if data.get("edge_type") == "CROSS_SHEET_REF":
                    count += 1
        return count

    def _are_connected(self, s1: str, s2: str) -> bool:
        for n in self.G.successors(s1):
            for _, tgt, data in self.G.out_edges(n, data=True):
                if data.get("edge_type") == "CROSS_SHEET_REF" and tgt.startswith(f"cell:{s2.split(':')[1]}"):
                    return True
        return False

    def get_architecture_overview(self) -> dict:
        communities = self.detect_communities()
        return {
            "total_sheets": len([n for n, d in self.G.nodes(data=True) if d.get("type") == "Sheet"]),
            "total_cells": len([n for n, d in self.G.nodes(data=True) if d.get("type") == "Cell"]),
            "total_formulas": len([n for n, d in self.G.nodes(data=True) if d.get("type") == "Formula"]),
            "communities": communities,
        }
