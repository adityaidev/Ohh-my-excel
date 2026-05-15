from pathlib import Path
from typing import Optional

from excel_graph_mcp.dependency import build_dependency_graph
from excel_graph_mcp.graph import GraphStore
from excel_graph_mcp.tools._common import validate_file_path, format_result


def build_or_update_graph(file_path: str, full_rebuild: bool = False, detail_level: str = "standard") -> dict:
    p = validate_file_path(file_path)
    store, builder = build_dependency_graph(p)
    stats = store.stats()
    result = {
        "file": str(p),
        "status": "built",
        "nodes": stats["nodes"],
        "edges": stats["edges"],
        "by_type": stats["by_type"],
    }
    store.close()
    return format_result(result, detail_level)


def run_postprocess(file_path: str, flows: bool = True, communities: bool = True, fts: bool = True) -> dict:
    p = validate_file_path(file_path)
    store = GraphStore(p)
    result = {"file": str(p), "flows": False, "communities": False, "fts": False}
    if fts:
        store._conn().execute("INSERT INTO nodes_fts SELECT rowid, id, type, sheet, data FROM nodes")
        store._conn().commit()
        result["fts"] = True
    store.close()
    return result


def list_graph_stats(file_path: str) -> dict:
    p = validate_file_path(file_path)
    store = GraphStore(p)
    stats = store.stats()
    store.close()
    return stats
