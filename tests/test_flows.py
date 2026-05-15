from pathlib import Path
from excel_graph_mcp.flows import FlowDetector
from excel_graph_mcp.graph import GraphStore
from excel_graph_mcp.dependency import build_dependency_graph

FIXTURES = Path(__file__).parent / "fixtures"


def test_detect_flows():
    store, _ = build_dependency_graph(FIXTURES / "cross_sheet.xlsx")
    detector = FlowDetector(store)
    flows = detector.detect_flows()
    assert len(flows) > 0
    for flow in flows:
        assert "sheet" in flow
        assert "inputs" in flow
        assert "calculations" in flow
    store.close()


def test_affected_flows():
    store, _ = build_dependency_graph(FIXTURES / "cross_sheet.xlsx")
    detector = FlowDetector(store)
    impacted = detector.get_affected_flows("Input!A1")
    assert isinstance(impacted, list)
    store.close()


def test_simple_workbook_flows():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    detector = FlowDetector(store)
    flows = detector.detect_flows()
    assert len(flows) > 0
    store.close()
