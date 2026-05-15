from pathlib import Path
from excel_graph_mcp.tools.build import build_or_update_graph, list_graph_stats
from excel_graph_mcp.tools.context import get_minimal_context
from excel_graph_mcp.tools._common import validate_file_path, sanitize_name

FIXTURES = Path(__file__).parent / "fixtures"


def test_build_tool():
    result = build_or_update_graph(str(FIXTURES / "simple.xlsx"))
    data = result.get("data", result)
    assert data["status"] == "built"
    assert data["nodes"] > 0


def test_graph_stats():
    build_or_update_graph(str(FIXTURES / "simple.xlsx"))
    stats = list_graph_stats(str(FIXTURES / "simple.xlsx"))
    assert "nodes" in stats
    assert "edges" in stats


def test_minimal_context():
    build_or_update_graph(str(FIXTURES / "simple.xlsx"))
    ctx = get_minimal_context(str(FIXTURES / "simple.xlsx"), task="test")
    assert ctx["nodes"] > 0
    assert "next_tool_suggestions" in ctx


def test_validate_valid_path():
    p = validate_file_path(str(FIXTURES / "simple.xlsx"))
    assert p.exists()


def test_validate_invalid_path():
    import pytest
    with pytest.raises(FileNotFoundError):
        validate_file_path(str(FIXTURES / "nonexistent.xlsx"))


def test_validate_invalid_format():
    import pytest
    with pytest.raises(ValueError):
        validate_file_path(str(FIXTURES / "generate_fixtures.py"))


def test_sanitize_name():
    assert sanitize_name("hello\x00world") == "helloworld"
    assert len(sanitize_name("a" * 500)) == 256
