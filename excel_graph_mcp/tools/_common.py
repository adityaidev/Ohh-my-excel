from pathlib import Path
from typing import Any


def validate_file_path(path_str: str) -> Path:
    p = Path(path_str).resolve()
    if not p.exists():
        raise FileNotFoundError(f"File not found: {path_str}")
    if p.suffix not in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xls"):
        raise ValueError(f"Unsupported file format: {p.suffix}")
    return p


def sanitize_name(name: str, max_length: int = 256) -> str:
    import re
    sanitized = re.sub(r"[\x00-\x1f\x7f]", "", name)
    return sanitized[:max_length]


def next_tool_suggestions(context: dict) -> list[str]:
    suggestions = []
    if context.get("nodes", 0) == 0:
        suggestions.append("build_or_update_graph_tool")
    if context.get("has_circular_refs"):
        suggestions.append("find_circular_references_tool")
    if context.get("has_broken_refs"):
        suggestions.append("find_broken_references_tool")
    if not suggestions:
        suggestions.append("query_graph_tool")
        suggestions.append("get_minimal_context_tool")
    return suggestions


def format_result(data: Any, detail_level: str = "standard") -> dict:
    if detail_level == "minimal":
        if isinstance(data, dict):
            return {k: v for k, v in data.items() if not isinstance(v, (list, dict)) or k in ("nodes", "edges")}
    return {"data": data, "detail_level": detail_level}
