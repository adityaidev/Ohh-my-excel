from pathlib import Path
from excel_graph_mcp.hyperformula_bridge import HyperFormulaBridge, evaluate_formulas

FIXTURES = Path(__file__).parent / "fixtures"


def test_bridge_availability():
    bridge = HyperFormulaBridge()
    assert isinstance(bridge.is_available(), bool)


def test_evaluate_fallback():
    result = evaluate_formulas({"A1": "=SUM(1,2,3)"}, ["A1"])
    assert isinstance(result, dict)
