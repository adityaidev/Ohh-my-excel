from pathlib import Path

from excel_graph_mcp.graph import GraphStore


class EmbeddingSearch:
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.store = GraphStore(file_path)
        self._encoder = None

    def _get_encoder(self):
        if self._encoder is not None:
            return self._encoder
        try:
            from sentence_transformers import SentenceTransformer
            self._encoder = SentenceTransformer("all-MiniLM-L6-v2")
        except ImportError:
            self._encoder = None
        return self._encoder

    def is_available(self) -> bool:
        return self._get_encoder() is not None

    def embed_graph(self) -> dict:
        encoder = self._get_encoder()
        if not encoder:
            return {"error": "sentence-transformers not installed. Run: pip install ohh-my-excel[embeddings]"}
        conn = self.store._conn()
        nodes = [dict(r) for r in conn.execute("SELECT id, type, sheet, data FROM nodes").fetchall()]
        texts = []
        node_ids = []
        for node in nodes:
            label = f"{node['type']}: {node['id']}"
            if node.get("sheet"):
                label += f" sheet:{node['sheet']}"
            texts.append(label)
            node_ids.append(node["id"])
        if not texts:
            return {"error": "No nodes to embed"}
        embeddings = encoder.encode(texts, show_progress_bar=False)
        conn.execute("DELETE FROM IF EXISTS node_embeddings")
        conn.execute("CREATE TABLE IF NOT EXISTS node_embeddings (node_id TEXT, embedding BLOB)")
        import numpy as np
        for nid, emb in zip(node_ids, embeddings):
            conn.execute("INSERT INTO node_embeddings VALUES (?, ?)", (nid, emb.astype(np.float32).tobytes()))
        conn.commit()
        return {"nodes_embedded": len(texts), "dimension": embeddings.shape[1]}

    def semantic_search(self, query: str, limit: int = 20) -> list[dict]:
        encoder = self._get_encoder()
        if not encoder:
            return []
        conn = self.store._conn()
        has_table = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='node_embeddings'"
        ).fetchone()
        if not has_table:
            self.embed_graph()
        q_vec = encoder.encode([query])[0]
        rows = conn.execute("SELECT node_id, embedding FROM node_embeddings").fetchall()
        import numpy as np
        from numpy.linalg import norm
        results = []
        for node_id, emb_bytes in rows:
            emb = np.frombuffer(emb_bytes, dtype=np.float32)
            sim = float(np.dot(q_vec, emb) / (norm(q_vec) * norm(emb) + 1e-10))
            results.append((sim, node_id))
        results.sort(key=lambda x: x[0], reverse=True)
        top = []
        for sim, nid in results[:limit]:
            node = self.store.get_node(nid)
            if node:
                top.append({"node": node, "similarity": round(sim, 4)})
        return top


def build_embeddings(file_path: str) -> dict:
    es = EmbeddingSearch(Path(file_path))
    return es.embed_graph()


def semantic_search_vector(file_path: str, query: str, limit: int = 20) -> dict:
    es = EmbeddingSearch(Path(file_path))
    if not es.is_available():
        return {"error": "Embeddings unavailable. Install: pip install ohh-my-excel[embeddings]"}
    results = es.semantic_search(query, limit)
    return {"query": query, "results": results, "count": len(results)}
