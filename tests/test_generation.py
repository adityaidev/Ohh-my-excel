from pathlib import Path
from excel_graph_mcp.generation import (
    generate_workbook,
    generate_workbook_from_data,
    generate_workbook_from_template,
    add_sheet,
    add_chart,
    apply_formatting,
    generate_formulas,
    validate_workbook,
)

FIXTURES = Path(__file__).parent / "fixtures"
TMP = FIXTURES.parent / "tmp"
TMP.mkdir(exist_ok=True)


def test_generate_workbook():
    out = TMP / "test_gen.xlsx"
    if out.exists(): out.unlink()
    result = generate_workbook("Test workbook", str(out))
    assert result["status"] == "created"
    assert out.exists()
    out.unlink()


def test_generate_from_json():
    out = TMP / "test_data.xlsx"
    if out.exists(): out.unlink()
    data = '[{"Name":"Alice","Age":30},{"Name":"Bob","Age":25}]'
    result = generate_workbook_from_data(data, str(out))
    assert result["status"] == "created"
    assert out.exists()
    out.unlink()


def test_generate_from_template():
    out = TMP / "test_template.xlsx"
    if out.exists(): out.unlink()
    result = generate_workbook_from_template("expense_tracker", str(out))
    assert result["status"] == "created"
    assert out.exists()
    out.unlink()


def test_add_sheet():
    src = FIXTURES / "simple.xlsx"
    out = TMP / "test_addsheet.xlsx"
    import shutil
    shutil.copy2(src, out)
    result = add_sheet(str(out), "NewSheet", [{"name": "Col1"}, {"name": "Col2"}])
    assert result["status"] == "sheet_added"
    out.unlink()


def test_generate_formulas():
    src = FIXTURES / "simple.xlsx"
    out = TMP / "test_formula.xlsx"
    import shutil
    shutil.copy2(src, out)
    result = generate_formulas(str(out), "Sheet1", "B", "sum")
    assert result["type"] == "sum"
    assert "formula" in result
    out.unlink()


def test_validate_workbook():
    result = validate_workbook(str(FIXTURES / "simple.xlsx"))
    assert "valid" in result
    assert "issues" in result
