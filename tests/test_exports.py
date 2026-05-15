from pathlib import Path
from excel_graph_mcp.exports import export_as_json
from excel_graph_mcp.dependency import build_dependency_graph

FIXTURES = Path(__file__).parent / "fixtures"


def test_export_as_json():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    result = export_as_json(str(FIXTURES / "simple.xlsx"))
    assert "nodes" in result
    assert "edges" in result
    assert result["node_count"] > 0
    store.close()
