from pathlib import Path
from excel_graph_mcp.graph import GraphStore

FIXTURES = Path(__file__).parent / "fixtures"


import tempfile


def _fresh_store():
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    p = Path(tmp.name)
    p.write_bytes(b"")
    store = GraphStore(p)
    return store, p


def test_graph_initialization():
    store, _ = _fresh_store()
    stats = store.stats()
    assert stats["nodes"] == 0
    assert stats["edges"] == 0
    store.close()


def test_add_and_get_node():
    store, _ = _fresh_store()
    store.add_node("test:node1", "Cell", data={"value": 42})
    node = store.get_node("test:node1")
    assert node is not None
    assert node["type"] == "Cell"
    store.close()


def test_add_and_get_edge():
    store, _ = _fresh_store()
    store.add_node("test:a", "Cell")
    store.add_node("test:b", "Cell")
    store.add_edge("test:a", "test:b", "DEPENDS_ON")
    edges = store.get_edges("test:a", "outgoing")
    assert len(edges) == 1
    assert edges[0]["target_id"] == "test:b"
    store.close()


def test_bfs_traversal():
    store, _ = _fresh_store()
    store.add_node("test:root", "Cell")
    store.add_node("test:child1", "Cell")
    store.add_node("test:child2", "Cell")
    store.add_node("test:grandchild", "Cell")
    store.add_edge("test:root", "test:child1", "DEPENDS_ON")
    store.add_edge("test:root", "test:child2", "DEPENDS_ON")
    store.add_edge("test:child1", "test:grandchild", "DEPENDS_ON")
    results = store.bfs("test:root", max_depth=2, direction="outgoing")
    depth_1 = [r for r in results if r["depth"] == 1]
    depth_2 = [r for r in results if r["depth"] == 2]
    assert len(depth_1) == 2
    assert len(depth_2) == 1
    store.close()


def test_networkx_conversion():
    store, _ = _fresh_store()
    store.add_node("nx:a", "Cell")
    store.add_node("nx:b", "Cell")
    store.add_edge("nx:a", "nx:b", "DEPENDS_ON")
    G = store.to_networkx()
    assert G.has_node("nx:a")
    assert G.has_node("nx:b")
    assert G.has_edge("nx:a", "nx:b")
    store.close()


def test_stats():
    store, _ = _fresh_store()
    before = store.stats()
    store.add_node("stat:1", "Cell")
    store.add_node("stat:2", "Formula")
    stats = store.stats()
    assert stats["nodes"] == before["nodes"] + 2
    assert stats["edges"] == before["edges"]
    store.close()


def test_clear():
    store, p = _fresh_store()
    store.add_node("clear:1", "Cell")
    store.clear()
    assert store.stats()["nodes"] == 0
    store.close()
    p.unlink()
