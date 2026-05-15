from pathlib import Path

from fastmcp import FastMCP

from excel_graph_mcp.constants import APP_NAME, VERSION
from excel_graph_mcp.tools.build import build_or_update_graph, list_graph_stats, run_postprocess
from excel_graph_mcp.tools.context import get_minimal_context
from excel_graph_mcp.tools.query import query_graph, semantic_search, traverse_graph

server = FastMCP(APP_NAME, version=VERSION)


@server.tool()
def build_or_update_graph_tool(file_path: str, full_rebuild: bool = False, detail_level: str = "standard") -> dict:
    """Build or incrementally update the Excel knowledge graph."""
    return build_or_update_graph(file_path, full_rebuild, detail_level)


@server.tool()
def run_postprocess_tool(file_path: str, flows: bool = True, communities: bool = True, fts: bool = True) -> dict:
    """Run post-processing: flow detection, community detection, FTS indexing."""
    return run_postprocess(file_path, flows, communities, fts)


@server.tool()
def list_graph_stats_tool(file_path: str) -> dict:
    """Get aggregate statistics about the knowledge graph."""
    return list_graph_stats(file_path)


@server.tool()
def get_minimal_context_tool(file_path: str, task: str = "") -> dict:
    """Ultra-compact ~100 token summary with next-tool suggestions."""
    return get_minimal_context(file_path, task)


@server.tool()
def semantic_search_nodes_tool(file_path: str, query: str, kind: str = "", limit: int = 20) -> dict:
    """Search for sheets, cells, tables, or formulas by keyword."""
    return semantic_search(file_path, query, kind, limit)


@server.tool()
def query_graph_tool(file_path: str, pattern: str, target: str, detail_level: str = "standard") -> dict:
    """Query graph: precedents_of, dependents_of, cross_sheet_refs, contains."""
    return query_graph(file_path, pattern, target, detail_level)


@server.tool()
def traverse_graph_tool(file_path: str, query: str, mode: str = "bfs", depth: int = 2, token_budget: int = 2000) -> dict:
    """BFS/DFS traversal from any cell/sheet with token budget."""
    return traverse_graph(file_path, query, mode, depth, token_budget)


@server.tool()
def get_impact_radius_tool(file_path: str, cell_ref: str, max_depth: int = 2) -> dict:
    """Blast radius analysis — find all cells/formulas affected when a cell changes."""
    p = _resolve(file_path)
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(p)
    results = store.bfs(f"cell:{cell_ref}", max_depth=max_depth, direction="incoming")
    store.close()
    return {"cell_ref": cell_ref, "max_depth": max_depth, "impacted": results}


@server.tool()
def detect_changes_tool(old_file: str, new_file: str, detail_level: str = "standard") -> dict:
    """Risk-scored change analysis between two workbook versions."""
    from excel_graph_mcp.changes import ChangeAnalyzer
    from excel_graph_mcp.dependency import build_dependency_graph
    from excel_graph_mcp.graph import GraphStore
    build_dependency_graph(_resolve(old_file))
    build_dependency_graph(_resolve(new_file))
    old_store = GraphStore(_resolve(old_file))
    new_store = GraphStore(_resolve(new_file))
    analyzer = ChangeAnalyzer(old_store, new_store)
    result = analyzer.detect_changes()
    old_store.close()
    new_store.close()
    return result


@server.tool()
def get_affected_flows_tool(file_path: str, cell_ref: str) -> dict:
    """Which data flows are impacted by a cell change."""
    from excel_graph_mcp.flows import FlowDetector
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(_resolve(file_path))
    detector = FlowDetector(store)
    result = detector.get_affected_flows(cell_ref)
    store.close()
    return {"cell_ref": cell_ref, "impacted_flows": result}


@server.tool()
def get_formula_dependencies_tool(file_path: str, cell_ref: str, include_ranges: bool = True) -> dict:
    """Get all precedent cells a formula depends on."""
    p = _resolve(file_path)
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(p)
    result = store.bfs(f"formula:{cell_ref}", max_depth=1, direction="outgoing")
    store.close()
    return {"cell_ref": cell_ref, "dependencies": result}


@server.tool()
def get_formula_dependents_tool(file_path: str, cell_ref: str) -> dict:
    """Get all cells that depend on a given cell."""
    from excel_graph_mcp.analysis import get_formula_dependents
    return get_formula_dependents(file_path, cell_ref)


@server.tool()
def find_circular_references_tool(file_path: str) -> dict:
    """Detect circular formula dependencies."""
    p = _resolve(file_path)
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(p)
    G = store.to_networkx()
    import networkx as nx
    cycles = []
    try:
        cycles = list(nx.simple_cycles(G))
    except nx.NetworkXNoCycle:
        pass
    store.close()
    return {"circular_references": cycles, "count": len(cycles)}


@server.tool()
def find_broken_references_tool(file_path: str) -> dict:
    """Find formulas referencing deleted or missing cells."""
    from excel_graph_mcp.analysis import find_broken_references
    return {"broken_references": find_broken_references(file_path)}


@server.tool()
def get_formula_complexity_tool(file_path: str, cell_ref: str) -> dict:
    """Analyze formula complexity: nesting depth, function count, cell count."""
    from excel_graph_mcp.analysis import get_formula_complexity
    return get_formula_complexity(file_path, cell_ref)


@server.tool()
def list_sheets_tool(file_path: str) -> dict:
    """List all sheets with metadata (cell count, formula count)."""
    p = _resolve(file_path)
    from excel_graph_mcp.parser import ParsedWorkbook
    wb = ParsedWorkbook(p)
    return {"sheets": [s.to_dict() for s in wb.sheets]}


@server.tool()
def get_sheet_info_tool(file_path: str, sheet_name: str) -> dict:
    """Detailed info about a specific sheet."""
    p = _resolve(file_path)
    from excel_graph_mcp.parser import ParsedWorkbook
    wb = ParsedWorkbook(p)
    for s in wb.sheets:
        if s.name == sheet_name:
            return s.to_dict()
    return {"error": f"Sheet '{sheet_name}' not found"}


@server.tool()
def find_cross_sheet_references_tool(file_path: str) -> dict:
    """Find all cross-sheet formula references."""
    from excel_graph_mcp.analysis import find_cross_sheet_references
    return {"cross_sheet_references": find_cross_sheet_references(file_path)}


@server.tool()
def list_tables_tool(file_path: str) -> dict:
    """List all Excel tables with ranges and columns."""
    p = _resolve(file_path)
    from excel_graph_mcp.parser import ParsedWorkbook
    wb = ParsedWorkbook(p)
    tables = []
    for s in wb.sheets:
        for t in s.tables:
            tables.append(t.to_dict())
    return {"tables": tables}


@server.tool()
def list_named_ranges_tool(file_path: str) -> dict:
    """List all named ranges and their definitions."""
    p = _resolve(file_path)
    from excel_graph_mcp.parser import ParsedWorkbook
    wb = ParsedWorkbook(p)
    return {"named_ranges": wb.named_ranges}


@server.tool()
def list_flows_tool(file_path: str, sort_by: str = "criticality", limit: int = 50) -> dict:
    """List data flows (input -> calculation -> output chains)."""
    from excel_graph_mcp.flows import FlowDetector
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(_resolve(file_path))
    detector = FlowDetector(store)
    flows = detector.detect_flows()
    store.close()
    return {"flows": flows[:limit]}


@server.tool()
def get_flow_tool(file_path: str, flow_id: str = "", flow_name: str = "") -> dict:
    """Details of a specific data flow."""
    from excel_graph_mcp.flows import FlowDetector
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(_resolve(file_path))
    detector = FlowDetector(store)
    flows = detector.detect_flows()
    store.close()
    for flow in flows:
        if flow["sheet"] == flow_name or flow_id in str(flow):
            return flow
    return {"error": "Flow not found"}


@server.tool()
def get_architecture_overview_tool(file_path: str) -> dict:
    """High-level workbook architecture from sheet communities."""
    from excel_graph_mcp.communities import CommunityDetector
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(_resolve(file_path))
    detector = CommunityDetector(store)
    result = detector.get_architecture_overview()
    store.close()
    return result


@server.tool()
def list_communities_tool(file_path: str, sort_by: str = "size") -> dict:
    """List detected sheet communities (grouped by cross-refs)."""
    from excel_graph_mcp.communities import CommunityDetector
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(_resolve(file_path))
    detector = CommunityDetector(store)
    communities = detector.detect_communities()
    store.close()
    communities.sort(key=lambda c: c[sort_by], reverse=True)
    return {"communities": communities}


@server.tool()
def get_hub_cells_tool(file_path: str, top_n: int = 10) -> dict:
    """Most-referenced cells (architectural hotspots)."""
    from excel_graph_mcp.analysis import get_hub_cells
    return {"hub_cells": get_hub_cells(file_path, top_n)}


@server.tool()
def get_bridge_cells_tool(file_path: str, top_n: int = 10) -> dict:
    """Cells connecting different sheet communities."""
    from excel_graph_mcp.analysis import get_bridge_cells
    return {"bridge_cells": get_bridge_cells(file_path, top_n)}


@server.tool()
def get_knowledge_gaps_tool(file_path: str) -> dict:
    """Structural weaknesses: orphan sheets, untested calculations."""
    from excel_graph_mcp.analysis import get_knowledge_gaps
    return get_knowledge_gaps(file_path)


@server.tool()
def get_suggested_questions_tool(file_path: str) -> dict:
    """Auto-generated audit questions from graph analysis."""
    from excel_graph_mcp.analysis import get_suggested_questions
    return {"questions": get_suggested_questions(file_path)}


@server.tool()
def visualize_graph_tool(file_path: str, format: str = "html", depth: int = 2) -> dict:
    """Generate interactive HTML graph visualization."""
    from excel_graph_mcp.exports import visualize_graph
    return visualize_graph(file_path, depth)


@server.tool()
def generate_workbook_tool(prompt: str, output_path: str, template: str = "", data_source: str = "", style: str = "professional") -> dict:
    """Create Excel workbook from natural language prompt."""
    from excel_graph_mcp.generation import generate_workbook
    return generate_workbook(prompt, output_path, template, data_source, style)


@server.tool()
def generate_workbook_from_data_tool(data: str, output_path: str, sheet_name: str = "Sheet1", include_charts: bool = False, style: str = "professional") -> dict:
    """Create workbook from provided data (JSON or CSV)."""
    from excel_graph_mcp.generation import generate_workbook_from_data
    return generate_workbook_from_data(data, output_path, sheet_name, include_charts, style)


@server.tool()
def generate_workbook_from_template_tool(template_name: str, output_path: str, customizations: str = "") -> dict:
    """Create workbook from built-in template (expense_tracker, budget_planner, invoice)."""
    from excel_graph_mcp.generation import generate_workbook_from_template
    return generate_workbook_from_template(template_name, output_path, customizations)


@server.tool()
def add_sheet_tool(file_path: str, sheet_name: str, columns: list[dict], formulas: list[dict] = None, formatting: dict = None) -> dict:
    """Add a new sheet to existing workbook."""
    from excel_graph_mcp.generation import add_sheet
    return add_sheet(file_path, sheet_name, columns, formulas, formatting)


@server.tool()
def add_chart_tool(file_path: str, sheet_name: str, chart_type: str, data_range: str, title: str = "") -> dict:
    """Add chart to existing workbook."""
    from excel_graph_mcp.generation import add_chart
    return add_chart(file_path, sheet_name, chart_type, data_range, title)


@server.tool()
def apply_formatting_tool(file_path: str, style: str = "professional", auto_size: bool = True, freeze_panes: dict = None) -> dict:
    """Apply professional formatting to workbook."""
    from excel_graph_mcp.generation import apply_formatting
    return apply_formatting(file_path, style, auto_size, freeze_panes)


@server.tool()
def generate_formulas_tool(file_path: str, sheet_name: str, column: str, formula_type: str = "sum") -> dict:
    """Auto-generate formulas for a column based on context (sum, average, count, max, min)."""
    from excel_graph_mcp.generation import generate_formulas
    return generate_formulas(file_path, sheet_name, column, formula_type)


@server.tool()
def validate_workbook_tool(file_path: str, check_circular: bool = True, check_broken_refs: bool = True) -> dict:
    """Validate generated workbook for errors."""
    from excel_graph_mcp.generation import validate_workbook
    return validate_workbook(file_path, check_circular, check_broken_refs)


@server.tool()
def export_graph_tool(file_path: str, format: str = "json") -> dict:
    """Export graph as JSON, CSV, or GraphML."""
    from excel_graph_mcp.exports import export_as_csv, export_as_graphml, export_as_json
    exporters = {"json": export_as_json, "csv": export_as_csv, "graphml": export_as_graphml}
    exporter = exporters.get(format, export_as_json)
    return exporter(file_path)


def _resolve(file_path: str) -> Path:
    from excel_graph_mcp.tools._common import validate_file_path
    return validate_file_path(file_path)


@server.prompt()
def audit_workbook(file_path: str) -> str:
    """Comprehensive workbook audit."""
    return f"Audit the workbook at {file_path} using Ohh-my-excel tools. Build the graph, check for issues, provide risk assessment."


@server.prompt()
def debug_formula(file_path: str, cell_ref: str) -> str:
    """Guided formula debugging."""
    return f"Debug formula at {cell_ref} in {file_path}. Trace dependencies and check for errors."


@server.prompt()
def data_flow_map(file_path: str) -> str:
    """Document data flows with diagrams."""
    return f"Map data flows in {file_path}. Identify input, calculation, and output sheets."


@server.prompt()
def onboard_analyst(file_path: str) -> str:
    """New analyst orientation."""
    return f"Orient me to {file_path}. Explain structure, key sheets, and important formulas."


@server.prompt()
def pre_merge_check(file_path: str) -> str:
    """Pre-merge change readiness check."""
    return f"Pre-merge check on {file_path}. Check broken refs, circular deps, and risk areas."


def serve(host: str = "localhost", port: int = 8080, transport: str = "stdio"):
    if transport == "stdio":
        server.run(transport="stdio")
    else:
        server.run(transport="sse", host=host, port=port)
