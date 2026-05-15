from pathlib import Path
from excel_graph_mcp.templates import list_templates, get_template, get_template_categories


def test_list_all_templates():
    templates = list_templates()
    assert len(templates) >= 26


def test_list_by_category():
    finance = list_templates("Finance")
    assert len(finance) > 0
    for t in finance:
        assert t["category"] == "Finance"


def test_get_template():
    t = get_template("expense_tracker")
    assert "sheets" in t
    assert t["name"] == "Expense Tracker"


def test_get_template_not_found():
    t = get_template("nonexistent")
    assert "error" in t


def test_get_categories():
    cats = get_template_categories()
    assert len(cats) >= 6
    names = [c["name"] for c in cats]
    for expected in ["Finance", "Sales", "Project", "HR", "Marketing", "Personal"]:
        assert expected in names


def test_template_structure():
    templates = list_templates()
    for t in templates:
        assert "id" in t
        assert "name" in t
        assert "category" in t
        assert "description" in t
        assert t["sheets"] >= 1
