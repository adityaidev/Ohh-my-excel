from excel_graph_mcp.graph import GraphStore
from excel_graph_mcp.tools._common import validate_file_path, next_tool_suggestions


def get_minimal_context(file_path: str, task: str = "") -> dict:
    p = validate_file_path(file_path)
    store = GraphStore(p)
    stats = store.stats()
    store.close()
    context = {
        "file": str(p),
        "task": task,
        "nodes": stats["nodes"],
        "edges": stats["edges"],
        "by_type": stats.get("by_type", {}),
        "cells": stats.get("by_type", {}).get("Cell", 0),
        "formulas": stats.get("by_type", {}).get("Formula", 0),
        "sheets": stats.get("by_type", {}).get("Sheet", 0),
        "tables": stats.get("by_type", {}).get("Table", 0),
        "has_circular_refs": False,
        "has_broken_refs": False,
        "next_tool_suggestions": next_tool_suggestions(stats),
    }
    return context
