from pathlib import Path
from excel_graph_mcp.communities import CommunityDetector
from excel_graph_mcp.dependency import build_dependency_graph

FIXTURES = Path(__file__).parent / "fixtures"


def test_detect_communities():
    store, _ = build_dependency_graph(FIXTURES / "cross_sheet.xlsx")
    detector = CommunityDetector(store)
    communities = detector.detect_communities()
    assert len(communities) > 0
    for comm in communities:
        assert "id" in comm
        assert "members" in comm
        assert "size" in comm
    store.close()


def test_architecture_overview():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    detector = CommunityDetector(store)
    overview = detector.get_architecture_overview()
    assert "total_sheets" in overview
    assert "total_cells" in overview
    assert "total_formulas" in overview
    assert "communities" in overview
    store.close()
