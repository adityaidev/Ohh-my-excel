from pathlib import Path
import networkx as nx

from excel_graph_mcp.graph import GraphStore
from excel_graph_mcp.parser import ParsedWorkbook


def _p(fp) -> Path:
    return fp if isinstance(fp, Path) else Path(fp)


def get_sheet_info(file_path: str, sheet_name: str) -> dict:
    wb = ParsedWorkbook(_p(file_path))
    for sheet in wb.sheets:
        if sheet.name == sheet_name:
            return sheet.to_dict()
    return {"error": f"Sheet '{sheet_name}' not found"}


def find_cross_sheet_references(file_path: str) -> list[dict]:
    wb = ParsedWorkbook(_p(file_path))
    refs = []
    sheet_names = {s.name for s in wb.sheets}
    for sheet in wb.sheets:
        for cell in sheet.cells:
            if cell.analysis:
                for ref in cell.analysis.all_references:
                    for sheet_name in sheet_names:
                        if sheet_name in ref and sheet_name != sheet.name:
                            refs.append({
                                "from_sheet": sheet.name,
                                "from_cell": cell.address,
                                "to_sheet": sheet_name,
                                "reference": ref,
                            })
    return refs


def find_broken_references(file_path: str) -> list[dict]:
    store = GraphStore(_p(file_path))
    G = store.to_networkx()
    broken = []
    for src, tgt, data in G.edges(data=True):
        if data.get("edge_type") in ("REFERENCES", "RANGE_REF"):
            if tgt not in G:
                broken.append({"source": src, "target": tgt, "type": data.get("edge_type")})
    store.close()
    return broken


def get_formula_complexity(file_path: str, cell_ref: str) -> dict:
    store = GraphStore(_p(file_path))
    formula_id = f"formula:{cell_ref}"
    node = store.get_node(formula_id)
    if not node:
        store.close()
        return {"error": f"Formula not found at {cell_ref}"}
    import json
    data = json.loads(node["data"]) if node["data"] else {}
    store.close()
    return {
        "cell": cell_ref,
        "nesting_depth": data.get("nesting_depth", 0),
        "functions_used": data.get("functions_used", []),
        "cell_count": data.get("cell_count", 0),
    }


def get_formula_dependents(file_path: str, cell_ref: str) -> dict:
    store = GraphStore(_p(file_path))
    target = f"cell:{cell_ref}"
    results = store.bfs(target, max_depth=2, direction="incoming")
    store.close()
    return {"cell_ref": cell_ref, "dependents": results}


def get_hub_cells(file_path: str, top_n: int = 10) -> list[dict]:
    store = GraphStore(_p(file_path))
    G = store.to_networkx()
    degrees = [(n, G.in_degree(n)) for n in G.nodes() if G.nodes[n].get("type") == "Cell"]
    degrees.sort(key=lambda x: x[1], reverse=True)
    store.close()
    return [{"cell": cell, "dependents": deg} for cell, deg in degrees[:top_n]]


def get_bridge_cells(file_path: str, top_n: int = 10) -> list[dict]:
    store = GraphStore(_p(file_path))
    G = store.to_networkx()
    try:
        centrality = nx.betweenness_centrality(G)
    except Exception:
        centrality = {}
    cells = [(n, c) for n, c in centrality.items() if G.nodes[n].get("type") == "Cell"]
    cells.sort(key=lambda x: x[1], reverse=True)
    store.close()
    return [{"cell": cell, "centrality": round(c, 4)} for cell, c in cells[:top_n]]


def get_knowledge_gaps(file_path: str) -> dict:
    store = GraphStore(_p(file_path))
    G = store.to_networkx()
    sheets = [n for n, d in G.nodes(data=True) if d.get("type") == "Sheet"]
    orphan_sheets = []
    for sheet in sheets:
        successors = list(G.successors(sheet))
        has_formulas = any(G.nodes[s].get("type") == "Formula" for s in successors)
        if not has_formulas:
            orphan_sheets.append(sheet)
    cells = [n for n, d in G.nodes(data=True) if d.get("type") == "Cell"]
    inputs = sum(1 for cell in cells if not any(G.nodes[s].get("type") == "Formula" for s in G.successors(cell)))
    formulas = sum(1 for cell in cells if any(G.nodes[s].get("type") == "Formula" for s in G.successors(cell)))
    store.close()
    return {
        "orphan_sheets": orphan_sheets,
        "orphan_count": len(orphan_sheets),
        "input_cells": inputs,
        "formula_cells": formulas,
        "total_sheets": len(sheets),
        "total_cells": len(cells),
    }


def get_suggested_questions(file_path: str) -> list[str]:
    store = GraphStore(_p(file_path))
    G = store.to_networkx()
    questions = []
    try:
        cycles = list(nx.simple_cycles(G))
        if cycles:
            questions.append(f"Found {len(cycles)} circular references. Use find_circular_references_tool.")
    except nx.NetworkXNoCycle:
        pass
    sheets = len([n for n, d in G.nodes(data=True) if d.get("type") == "Sheet"])
    cells = len([n for n, d in G.nodes(data=True) if d.get("type") == "Cell"])
    if sheets > 5:
        questions.append(f"Workbook has {sheets} sheets. Use get_architecture_overview_tool.")
    if cells > 100:
        questions.append(f"Large workbook with {cells} cells. Use get_minimal_context_tool.")
    store.close()
    return questions or ["Use query_graph_tool to explore cell dependencies."]
