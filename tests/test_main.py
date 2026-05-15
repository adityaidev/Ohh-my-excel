from pathlib import Path
from excel_graph_mcp.main import list_sheets_tool, list_tables_tool, list_named_ranges_tool

FIXTURES = Path(__file__).parent / "fixtures"


def test_list_sheets():
    result = list_sheets_tool(str(FIXTURES / "simple.xlsx"))
    assert "sheets" in result
    assert len(result["sheets"]) > 0
    assert result["sheets"][0]["cell_count"] > 0


def test_list_tables():
    result = list_tables_tool(str(FIXTURES / "simple.xlsx"))
    assert "tables" in result
    assert isinstance(result["tables"], list)


def test_list_named_ranges():
    result = list_named_ranges_tool(str(FIXTURES / "simple.xlsx"))
    assert "named_ranges" in result


def test_build_and_query():
    from excel_graph_mcp.tools.build import build_or_update_graph
    from excel_graph_mcp.tools.query import query_graph
    build_or_update_graph(str(FIXTURES / "simple.xlsx"))
    result = query_graph(str(FIXTURES / "simple.xlsx"), "contains", "Sheet1")
    data = result.get("data", result)
    assert "results" in data
    assert data["pattern"] == "contains"
