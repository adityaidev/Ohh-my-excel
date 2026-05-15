import networkx as nx

from excel_graph_mcp.graph import GraphStore


class FlowDetector:
    def __init__(self, store: GraphStore):
        self.store = store
        self.G = store.to_networkx()

    def detect_flows(self) -> list[dict]:
        sheets = [n for n, d in self.G.nodes(data=True) if d.get("type") == "Sheet"]
        flows = []
        for sheet in sheets:
            flow = self._analyze_sheet_flow(sheet)
            if flow:
                flows.append(flow)
        return flows

    def _analyze_sheet_flow(self, sheet_id: str) -> dict:
        cells = [n for n in self.G.successors(sheet_id) if self.G.nodes[n].get("type") == "Cell"]
        inputs = []
        calculations = []
        outputs = []
        for cell in cells:
            successors = list(self.G.successors(cell))
            formulas = [s for s in successors if self.G.nodes[s].get("type") == "Formula"]
            if formulas:
                calculations.append(cell)
            elif any(self.G.nodes[s].get("type") == "Cell" for s in successors if self.G.has_edge(s, cell)):
                outputs.append(cell)
            else:
                inputs.append(cell)
        return {
            "sheet": sheet_id,
            "inputs": len(inputs),
            "calculations": len(calculations),
            "outputs": len(outputs),
            "total_cells": len(cells),
        }

    def get_affected_flows(self, cell_ref: str) -> list[dict]:
        target = f"cell:{cell_ref}"
        if target not in self.G:
            return []
        impacted = []
        for successor in nx.descendants(self.G, target):
            node_data = self.G.nodes[successor]
            if node_data.get("type") == "Formula":
                impacted.append({"cell": successor, "type": "calculation"})
            elif node_data.get("type") == "Cell":
                impacted.append({"cell": successor, "type": "dependent"})
        return impacted
