from pathlib import Path
from typing import Optional

from fastmcp import FastMCP

from excel_graph_mcp.constants import VERSION, APP_NAME
from excel_graph_mcp.tools.build import build_or_update_graph, run_postprocess, list_graph_stats
from excel_graph_mcp.tools.query import query_graph, traverse_graph, semantic_search
from excel_graph_mcp.tools.context import get_minimal_context

server = FastMCP(
    APP_NAME,
    version=VERSION,
)


@server.tool()
def build_or_update_graph_tool(
    file_path: str,
    full_rebuild: bool = False,
    detail_level: str = "standard",
) -> dict:
    """Build or incrementally update the Excel knowledge graph for a workbook.
    Returns node/edge counts and per-type breakdown."""
    return build_or_update_graph(file_path, full_rebuild, detail_level)


@server.tool()
def run_postprocess_tool(
    file_path: str,
    flows: bool = True,
    communities: bool = True,
    fts: bool = True,
) -> dict:
    """Run post-processing on the graph: flow detection, community detection, FTS indexing."""
    from excel_graph_mcp.tools.build import run_postprocess as _rp
    return _rp(file_path, flows, communities, fts)


@server.tool()
def list_graph_stats_tool(file_path: str) -> dict:
    """Get aggregate statistics about the knowledge graph for a workbook."""
    return list_graph_stats(file_path)


@server.tool()
def get_minimal_context_tool(file_path: str, task: str = "") -> dict:
    """Ultra-compact ~100 token summary of a workbook with next-tool suggestions."""
    return get_minimal_context(file_path, task)


@server.tool()
def semantic_search_nodes_tool(
    file_path: str,
    query: str,
    kind: str = "",
    limit: int = 20,
) -> dict:
    """Search for sheets, cells, tables, or formulas by name or keyword."""
    return semantic_search(file_path, query, kind, limit)


@server.tool()
def query_graph_tool(
    file_path: str,
    pattern: str,
    target: str,
    detail_level: str = "standard",
) -> dict:
    """Query the graph with predefined patterns: precedents_of, dependents_of, cross_sheet_refs, contains."""
    return query_graph(file_path, pattern, target, detail_level)


@server.tool()
def traverse_graph_tool(
    file_path: str,
    query: str,
    mode: str = "bfs",
    depth: int = 2,
    token_budget: int = 2000,
) -> dict:
    """BFS/DFS traversal from any cell or sheet with a token budget."""
    return traverse_graph(file_path, query, mode, depth, token_budget)


@server.tool()
def get_impact_radius_tool(file_path: str, cell_ref: str, max_depth: int = 2) -> dict:
    """Blast radius analysis — which cells/formulas are affected when a cell changes."""
    p = validate_path(file_path)
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(p)
    results = store.bfs(f"cell:{cell_ref}", max_depth=max_depth, direction="incoming")
    store.close()
    return {"cell_ref": cell_ref, "max_depth": max_depth, "impacted": results}


@server.tool()
def get_formula_dependencies_tool(file_path: str, cell_ref: str, include_ranges: bool = True) -> dict:
    """Get all precedent cells a formula depends on."""
    p = validate_path(file_path)
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(p)
    result = store.bfs(f"formula:{cell_ref}", max_depth=1, direction="outgoing")
    store.close()
    return {"cell_ref": cell_ref, "dependencies": result}


@server.tool()
def find_circular_references_tool(file_path: str) -> dict:
    """Detect circular formula dependencies in the workbook."""
    p = validate_path(file_path)
    from excel_graph_mcp.dependency import DependencyBuilder
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
def list_sheets_tool(file_path: str) -> dict:
    """List all sheets in a workbook with metadata."""
    p = validate_path(file_path)
    from excel_graph_mcp.graph import GraphStore
    store = GraphStore(p)
    from excel_graph_mcp.parser import ParsedWorkbook
    wb = ParsedWorkbook(p)
    store.close()
    return {"sheets": [s.to_dict() for s in wb.sheets]}


@server.tool()
def get_sheet_info_tool(file_path: str, sheet_name: str) -> dict:
    """Get detailed info about a specific sheet."""
    p = validate_path(file_path)
    from excel_graph_mcp.parser import ParsedWorkbook
    wb = ParsedWorkbook(p)
    for s in wb.sheets:
        if s.name == sheet_name:
            return s.to_dict()
    return {"error": f"Sheet '{sheet_name}' not found"}


@server.tool()
def list_tables_tool(file_path: str) -> dict:
    """List all Excel tables with their ranges and columns."""
    p = validate_path(file_path)
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
    p = validate_path(file_path)
    from excel_graph_mcp.parser import ParsedWorkbook
    wb = ParsedWorkbook(p)
    return {"named_ranges": wb.named_ranges}


def validate_path(file_path: str) -> Path:
    from excel_graph_mcp.tools._common import validate_file_path
    return validate_file_path(file_path)


@server.prompt()
def audit_workbook(file_path: str) -> str:
    """Comprehensive workbook audit prompt."""
    return f"I need you to perform a comprehensive audit of the workbook at {file_path}. Use the Ohh-my-excel tools to build the graph, check for issues, and provide a full risk assessment."


@server.prompt()
def debug_formula(file_path: str, cell_ref: str) -> str:
    """Guided formula debugging workflow."""
    return f"I need help debugging the formula in {cell_ref} in the workbook at {file_path}. Please trace its dependencies and check for errors."


@server.prompt()
def data_flow_map(file_path: str) -> str:
    """Document data flows with diagrams."""
    return f"Map out the data flows in the workbook at {file_path}. Identify input sheets, calculation sheets, and output sheets."


@server.prompt()
def onboard_analyst(file_path: str) -> str:
    """New analyst orientation to workbook structure."""
    return f"I'm a new analyst. Please orient me to the workbook at {file_path}. Explain its structure, key sheets, and important formulas."


@server.prompt()
def pre_merge_check(file_path: str) -> str:
    """Workbook change readiness check."""
    return f"Perform a pre-merge check on the workbook at {file_path}. Check for broken refs, circular dependencies, and risk areas."


def serve(host: str = "localhost", port: int = 8080, transport: str = "stdio"):
    if transport == "stdio":
        server.run(transport="stdio")
    else:
        server.run(transport="sse", host=host, port=port)
