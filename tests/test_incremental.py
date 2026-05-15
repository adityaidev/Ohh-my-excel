from pathlib import Path
from excel_graph_mcp.incremental import IncrementalUpdater
from excel_graph_mcp.constants import get_graph_dir
import shutil

FIXTURES = Path(__file__).parent / "fixtures"


def setup_module():
    graph_dir = get_graph_dir(FIXTURES / "simple.xlsx")
    if graph_dir.exists():
        shutil.rmtree(graph_dir)


def test_incremental_first_build():
    updater = IncrementalUpdater(FIXTURES / "simple.xlsx")
    assert updater.needs_update()
    result = updater.update()
    assert result["status"] in ("updated", "unchanged")
    assert result["nodes"] >= 0


def test_incremental_no_change():
    updater = IncrementalUpdater(FIXTURES / "simple.xlsx")
    updater.update()
    updater2 = IncrementalUpdater(FIXTURES / "simple.xlsx")
    assert not updater2.needs_update()
