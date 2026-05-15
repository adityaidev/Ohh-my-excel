from pathlib import Path
from excel_graph_mcp.analysis import (
    find_cross_sheet_references,
    find_broken_references,
    get_formula_complexity,
    get_formula_dependents,
    get_hub_cells,
    get_bridge_cells,
    get_knowledge_gaps,
    get_suggested_questions,
)
from excel_graph_mcp.dependency import build_dependency_graph

FIXTURES = Path(__file__).parent / "fixtures"


def test_find_cross_sheet_references():
    refs = find_cross_sheet_references(FIXTURES / "cross_sheet.xlsx")
    assert isinstance(refs, list)


def test_find_broken_references():
    broken = find_broken_references(FIXTURES / "simple.xlsx")
    assert isinstance(broken, list)


def test_get_formula_complexity():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    result = get_formula_complexity(FIXTURES / "simple.xlsx", "Sheet1!D2")
    assert isinstance(result, dict)
    store.close()


def test_get_formula_dependents():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    result = get_formula_dependents(FIXTURES / "simple.xlsx", "Sheet1!B2")
    assert "dependents" in result
    store.close()


def test_get_hub_cells():
    store, _ = build_dependency_graph(FIXTURES / "cross_sheet.xlsx")
    hubs = get_hub_cells(FIXTURES / "cross_sheet.xlsx")
    assert isinstance(hubs, list)
    store.close()


def test_get_bridge_cells():
    store, _ = build_dependency_graph(FIXTURES / "cross_sheet.xlsx")
    bridges = get_bridge_cells(FIXTURES / "cross_sheet.xlsx")
    assert isinstance(bridges, list)
    store.close()


def test_get_knowledge_gaps():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    gaps = get_knowledge_gaps(FIXTURES / "simple.xlsx")
    assert "orphan_sheets" in gaps
    assert "input_cells" in gaps
    assert "formula_cells" in gaps
    store.close()


def test_get_suggested_questions():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    questions = get_suggested_questions(FIXTURES / "simple.xlsx")
    assert len(questions) > 0
    store.close()
