from pathlib import Path
import json

from excel_graph_mcp.graph import GraphStore
from excel_graph_mcp.constants import get_graph_dir


def export_as_json(file_path: str) -> dict:
    store = GraphStore(Path(file_path))
    nodes = [dict(r) for r in store._conn().execute("SELECT * FROM nodes").fetchall()]
    edges = [dict(r) for r in store._conn().execute("SELECT * FROM edges").fetchall()]
    store.close()
    return {"nodes": nodes, "edges": edges, "node_count": len(nodes), "edge_count": len(edges)}


def export_as_csv(file_path: str) -> dict:
    store = GraphStore(Path(file_path))
    nodes = [dict(r) for r in store._conn().execute("SELECT id, type, sheet FROM nodes").fetchall()]
    edges = [dict(r) for r in store._conn().execute("SELECT source_id, target_id, edge_type, confidence FROM edges").fetchall()]
    store.close()
    export_dir = get_graph_dir(Path(file_path))
    export_dir.mkdir(parents=True, exist_ok=True)
    nodes_path = export_dir / "nodes.csv"
    edges_path = export_dir / "edges.csv"
    with open(nodes_path, "w") as f:
        f.write("id,type,sheet\n")
        for n in nodes:
            f.write(f"{n['id']},{n['type']},{n.get('sheet', '')}\n")
    with open(edges_path, "w") as f:
        f.write("source_id,target_id,edge_type,confidence\n")
        for e in edges:
            f.write(f"{e['source_id']},{e['target_id']},{e['edge_type']},{e['confidence']}\n")
    return {"nodes_file": str(nodes_path), "edges_file": str(edges_path), "node_count": len(nodes), "edge_count": len(edges)}


def export_as_graphml(file_path: str) -> dict:
    store = GraphStore(Path(file_path))
    G = store.to_networkx()
    export_dir = get_graph_dir(Path(file_path))
    export_dir.mkdir(parents=True, exist_ok=True)
    output = export_dir / "graph.graphml"
    import networkx as nx
    nx.write_graphml(G, str(output))
    store.close()
    return {"file": str(output)}


def visualize_graph(file_path: str, depth: int = 2) -> dict:
    store = GraphStore(Path(file_path))
    G = store.to_networkx()
    export_dir = get_graph_dir(Path(file_path))
    export_dir.mkdir(parents=True, exist_ok=True)
    output = export_dir / "graph.html"
    try:
        from pyvis.network import Network
        net = Network(height="800px", width="100%", bgcolor="#1a1a2e", font_color="white")
        net.barnes_hut(gravity=-8000, central_gravity=0.3, spring_length=200)
        for node, data in G.nodes(data=True):
            ntype = data.get("type", "cell")
            color_map = {"Sheet": "#00d2ff", "Cell": "#ffffff", "Formula": "#ff6b6b", "Table": "#4ecdc4", "Range": "#ffe66d", "Chart": "#a8e6cf", "Workbook": "#ffd700", "NamedRange": "#ff9ff3"}
            net.add_node(node, label=str(node)[:30], color=color_map.get(ntype, "#888"), title=ntype)
        for src, tgt, data in G.edges(data=True):
            net.add_edge(src, tgt, title=data.get("edge_type", ""))
        net.show(str(output), notebook=False)
    except ImportError:
        import matplotlib.pyplot as plt
        pos = nx.spring_layout(G)
        plt.figure(figsize=(20, 20))
        nx.draw(G, pos, with_labels=False, node_size=50, node_color="blue", edge_color="gray")
        plt.savefig(str(export_dir / "graph.png"))
        output = export_dir / "graph.png"
    store.close()
    return {"file": str(output)}
