from pathlib import Path
from excel_graph_mcp.formula_parser import parse_formula, FormulaAnalysis

FIXTURES = Path(__file__).parent / "fixtures"


def test_simple_formula():
    analysis = parse_formula("=B2*C2")
    assert analysis is not None
    assert "B2" in analysis.all_references
    assert "C2" in analysis.all_references
    assert analysis.cell_count >= 2


def test_formula_with_range():
    analysis = parse_formula("=SUM(D2:D10)")
    assert analysis is not None
    assert any("D2:D10" in r for r in analysis.all_references)
    assert "SUM" in analysis.functions_used
    assert analysis.nesting_depth == 1


def test_nested_formula():
    analysis = parse_formula("=IF(A1>0, VLOOKUP(B1, C1:D10, 2, FALSE), 0)")
    assert analysis is not None
    assert "IF" in analysis.functions_used
    assert "VLOOKUP" in analysis.functions_used
    assert analysis.nesting_depth >= 2


def test_cross_sheet_reference():
    analysis = parse_formula("=Input!A1*2")
    assert analysis is not None
    refs = analysis.all_references
    assert any("A1" in r for r in refs)


def test_ambiguous_formula():
    analysis = parse_formula("=INDIRECT(A1)")
    assert analysis is not None
    assert analysis.is_ambiguous
    assert "INDIRECT" in analysis.functions_used


def test_no_formula():
    analysis = parse_formula("hello")
    assert analysis is None


def test_empty_string():
    analysis = parse_formula("")
    assert analysis is None


def test_function_extraction():
    analysis = parse_formula("=SUM(A1:A10)+AVERAGE(B1:B5)+MAX(C1:C3)")
    assert analysis is not None
    for func in ["SUM", "AVERAGE", "MAX"]:
        assert func in analysis.functions_used


def test_deep_nesting():
    analysis = parse_formula("=IF(AND(A1>0, OR(B1<10, C1=5)), SUM(D1:D10), 0)")
    assert analysis is not None
    assert analysis.nesting_depth >= 3
