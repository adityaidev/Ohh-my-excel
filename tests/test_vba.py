from pathlib import Path
from excel_graph_mcp.vba_analysis import analyze_vba, explain_vba, VBAAnalyzer

FIXTURES = Path(__file__).parent / "fixtures"


def test_analyze_no_vba():
    result = analyze_vba(str(FIXTURES / "simple.xlsx"))
    assert "file" in result
    assert "macro_count" in result


def test_explain_no_vba():
    result = explain_vba(str(FIXTURES / "simple.xlsx"))
    assert "explanation" in result


def test_analyzer_initialization():
    analyzer = VBAAnalyzer(FIXTURES / "simple.xlsx")
    assert isinstance(analyzer.macros, list)
