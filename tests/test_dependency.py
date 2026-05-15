from pathlib import Path
from excel_graph_mcp.dependency import build_dependency_graph

FIXTURES = Path(__file__).parent / "fixtures"


def test_build_simple():
    store, builder = build_dependency_graph(FIXTURES / "simple.xlsx")
    stats = store.stats()
    assert stats["nodes"] > 0
    assert stats["edges"] > 0
    store.close()


def test_build_cross_sheet():
    store, builder = build_dependency_graph(FIXTURES / "cross_sheet.xlsx")
    stats = store.stats()
    assert stats["nodes"] > 0
    G = store.to_networkx()
    cross_sheet_edges = [
        (s, t) for s, t, d in G.edges(data=True)
        if d.get("edge_type") == "REFERENCES"
    ]
    assert len(cross_sheet_edges) > 0
    store.close()


def test_find_circular_references():
    store, builder = build_dependency_graph(FIXTURES / "simple.xlsx")
    cycles = builder.find_circular_references()
    assert isinstance(cycles, list)
    store.close()


def test_node_types_present():
    store, builder = build_dependency_graph(FIXTURES / "cross_sheet.xlsx")
    stats = store.stats()
    by_type = stats.get("by_type", {})
    assert "Workbook" in by_type
    assert "Sheet" in by_type
    assert "Cell" in by_type
    store.close()
