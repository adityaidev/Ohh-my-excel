from pathlib import Path
from excel_graph_mcp.parser import ParsedWorkbook

FIXTURES = Path(__file__).parent / "fixtures"


def test_parse_simple_workbook():
    wb = ParsedWorkbook(FIXTURES / "simple.xlsx")
    assert wb.name == "simple"
    assert wb.sheet_count >= 1
    sheet = wb.sheets[0]
    assert sheet.name == "Sheet1"
    assert sheet.cell_count > 0
    assert sheet.formula_count > 0


def test_parse_cross_sheet():
    wb = ParsedWorkbook(FIXTURES / "cross_sheet.xlsx")
    assert wb.sheet_count >= 2
    sheets = {s.name: s for s in wb.sheets}
    assert "Input" in sheets
    assert "Calculation" in sheets
    formulas = [c for c in sheets["Calculation"].cells if c.is_formula]
    assert len(formulas) > 0


def test_cell_values():
    wb = ParsedWorkbook(FIXTURES / "simple.xlsx")
    sheet = wb.sheets[0]
    cells = {c.address: c for c in sheet.cells}
    assert "A1" in cells
    assert cells["A1"].value is not None
    assert cells["A1"].is_input


def test_formula_detection():
    wb = ParsedWorkbook(FIXTURES / "simple.xlsx")
    sheet = wb.sheets[0]
    formula_cells = [c for c in sheet.cells if c.is_formula]
    assert len(formula_cells) > 0
    for c in formula_cells:
        assert c.formula is not None
        assert c.formula.startswith("=")
        assert c.analysis is not None


def test_named_ranges():
    wb = ParsedWorkbook(FIXTURES / "simple.xlsx")
    assert hasattr(wb, "named_ranges")


def test_sheet_metadata():
    wb = ParsedWorkbook(FIXTURES / "simple.xlsx")
    d = wb.to_dict()
    assert "sheets" in d
    assert d["sheet_count"] > 0
    assert d["name"] == "simple"
