from pathlib import Path
from excel_graph_mcp.incremental import IncrementalUpdater
from excel_graph_mcp.dependency import build_dependency_graph

FIXTURES = Path(__file__).parent / "fixtures"


def test_incremental_first_build():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    updater = IncrementalUpdater(FIXTURES / "simple.xlsx")
    assert updater.needs_update()
    result = updater.update()
    assert result["status"] in ("updated", "unchanged")
    store.close()


def test_incremental_no_change():
    store, _ = build_dependency_graph(FIXTURES / "simple.xlsx")
    updater = IncrementalUpdater(FIXTURES / "simple.xlsx")
    updater.update()
    assert not updater.needs_update()
    store.close()
