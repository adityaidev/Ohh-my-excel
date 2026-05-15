from excel_graph_mcp.graph import GraphStore
from excel_graph_mcp.tools._common import format_result, validate_file_path


def query_graph(file_path: str, pattern: str, target: str, detail_level: str = "standard") -> dict:
    p = validate_file_path(file_path)
    store = GraphStore(p)
    results = []
    if pattern == "precedents_of":
        target_id = f"formula:{target}"
        results = store.bfs(target_id, max_depth=1, direction="outgoing")
    elif pattern == "dependents_of":
        target_id = f"cell:{target}"
        results = store.bfs(target_id, max_depth=1, direction="incoming")
    elif pattern == "cross_sheet_refs":
        edges = store.get_edges(f"sheet:{target}", "outgoing")
        results = [e for e in edges if e["edge_type"] == "CROSS_SHEET_REF"]
    elif pattern == "contains":
        parent_id = f"sheet:{target}" if "!" not in target else target
        results = store.get_edges(parent_id, "outgoing")
    store.close()
    return format_result({"pattern": pattern, "target": target, "results": results}, detail_level)


def traverse_graph(file_path: str, query: str, mode: str = "bfs", depth: int = 2, token_budget: int = 2000) -> dict:
    p = validate_file_path(file_path)
    store = GraphStore(p)
    results = store.search(query, limit=5)
    traversed = []
    for r in results:
        traversed.extend(store.bfs(r["id"], max_depth=depth))
    store.close()
    return {"results": traversed, "mode": mode, "depth": depth}


def semantic_search(file_path: str, query: str, kind: str = "", limit: int = 20) -> dict:
    p = validate_file_path(file_path)
    store = GraphStore(p)
    results = store.search(query, limit=limit)
    if kind:
        results = [r for r in results if r.get("type", "").lower() == kind.lower()]
    store.close()
    return {"results": results, "total": len(results)}
