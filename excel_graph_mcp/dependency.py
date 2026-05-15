from pathlib import Path
from typing import Optional

import networkx as nx

from excel_graph_mcp.constants import EDGE_TYPES, CONFIDENCE_LEVELS
from excel_graph_mcp.formula_parser import parse_formula
from excel_graph_mcp.graph import GraphStore
from excel_graph_mcp.parser import ParsedWorkbook


class DependencyBuilder:
    def __init__(self, graph_store: GraphStore):
        self.graph = graph_store
        self.G: nx.DiGraph = nx.DiGraph()

    def build(self, workbook: ParsedWorkbook):
        self._add_workbook(workbook)
        for sheet in workbook.sheets:
            self._add_sheet(sheet.name, workbook.name)
            for cell in sheet.cells:
                self._add_cell(cell.address, sheet.name)
                if cell.analysis:
                    self._add_formula(cell.address, cell.formula, sheet.name, cell.analysis)
            for table in sheet.tables:
                self._add_table(table)
            for chart in sheet.charts:
                self._add_chart(chart)
        self._sync_to_store()

    def _add_workbook(self, wb: ParsedWorkbook):
        wb_id = f"wb:{wb.name}"
        self.G.add_node(wb_id, type="Workbook", label=wb.name)
        self.graph.add_node(wb_id, "Workbook", data={"name": wb.name, "path": str(wb.path)})

    def _add_sheet(self, sheet_name: str, wb_name: str):
        sheet_id = f"sheet:{sheet_name}"
        wb_id = f"wb:{wb_name}"
        self.G.add_node(sheet_id, type="Sheet", label=sheet_name)
        self.G.add_edge(wb_id, sheet_id, edge_type="CONTAINS")
        self.graph.add_node(sheet_id, "Sheet", data={"name": sheet_name})
        self.graph.add_edge(wb_id, sheet_id, "CONTAINS")

    def _add_cell(self, address: str, sheet_name: str):
        cell_id = f"cell:{sheet_name}!{address}"
        sheet_id = f"sheet:{sheet_name}"
        self.G.add_node(cell_id, type="Cell", label=f"{sheet_name}!{address}")
        self.G.add_edge(sheet_id, cell_id, edge_type="CONTAINS")
        self.graph.add_node(cell_id, "Cell", data={"address": address, "sheet": sheet_name})
        self.graph.add_edge(sheet_id, cell_id, "CONTAINS")

    def _add_formula(self, address: str, formula_text: str, sheet_name: str, analysis):
        formula_id = f"formula:{sheet_name}!{address}"
        cell_id = f"cell:{sheet_name}!{address}"
        self.G.add_node(
            formula_id,
            type="Formula",
            label=f"{address}: {formula_text[:50]}",
            functions=analysis.functions_used,
            nesting=analysis.nesting_depth,
        )
        self.G.add_edge(cell_id, formula_id, edge_type="CONTAINS")
        self.graph.add_node(
            formula_id, "Formula",
            data={
                "formula_text": formula_text,
                "functions_used": analysis.functions_used,
                "nesting_depth": analysis.nesting_depth,
                "cell_count": analysis.cell_count,
            },
        )
        self.graph.add_edge(cell_id, formula_id, "CONTAINS")

        for ref in analysis.cell_references:
            confidence = "AMBIGUOUS" if analysis.is_ambiguous else "EXTRACTED"
            ref_cell_id = self._resolve_ref(ref, sheet_name)
            if ref_cell_id:
                self.G.add_edge(formula_id, ref_cell_id, edge_type="REFERENCES", confidence=confidence)
                self.graph.add_edge(formula_id, ref_cell_id, "REFERENCES", confidence)

        for ref in analysis.range_references:
            confidence = "AMBIGUOUS" if analysis.is_ambiguous else "EXTRACTED"
            range_id = self._add_range(ref, sheet_name)
            if range_id:
                self.G.add_edge(formula_id, range_id, edge_type="RANGE_REF", confidence=confidence)
                self.graph.add_edge(formula_id, range_id, "RANGE_REF", confidence)

    def _add_range(self, ref: dict, sheet_name: str) -> Optional[str]:
        range_sheet = ref.get("sheet") or sheet_name
        range_id = f"range:{range_sheet}!{ref['ref']}"
        self.G.add_node(range_id, type="Range", label=ref["ref"])
        self.graph.add_node(
            range_id, "Range",
            data={"start": ref["start"], "end": ref["end"], "sheet": range_sheet},
        )
        return range_id

    def _add_table(self, table):
        table_id = f"table:{table.sheet}!{table.name}"
        sheet_id = f"sheet:{table.sheet}"
        self.G.add_node(table_id, type="Table", label=table.name)
        self.G.add_edge(sheet_id, table_id, edge_type="CONTAINS")
        self.graph.add_node(table_id, "Table", data={"name": table.name, "range": table.range, "columns": table.columns})
        self.graph.add_edge(sheet_id, table_id, "CONTAINS")

    def _add_chart(self, chart):
        chart_id = f"chart:{chart.sheet}!{chart.name}"
        sheet_id = f"sheet:{chart.sheet}"
        self.G.add_node(chart_id, type="Chart", label=chart.name)
        self.G.add_edge(sheet_id, chart_id, edge_type="CONTAINS")
        self.graph.add_node(chart_id, "Chart", data={"name": chart.name, "type": chart.type, "data_source": chart.data_source})
        self.graph.add_edge(sheet_id, chart_id, "CONTAINS")

    def _resolve_ref(self, ref: dict, current_sheet: str) -> Optional[str]:
        ref_sheet = ref.get("sheet") or current_sheet
        ref_addr = ref["ref"]
        return f"cell:{ref_sheet}!{ref_addr}"

    def find_circular_references(self) -> list[list[str]]:
        try:
            cycles = list(nx.simple_cycles(self.G))
            return cycles
        except nx.NetworkXNoCycle:
            return []

    def find_broken_references(self) -> list[str]:
        broken = []
        for _, target, data in self.G.edges(data=True):
            if data.get("edge_type") in ("REFERENCES", "RANGE_REF"):
                if target not in self.G:
                    broken.append(target)
        return broken

    def _sync_to_store(self):
        pass


def build_dependency_graph(workbook_path: Path) -> tuple[GraphStore, DependencyBuilder]:
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(workbook_path)
    store.clear()
    builder = DependencyBuilder(store)
    wb = ParsedWorkbook(workbook_path)
    builder.build(wb)
    return store, builder
