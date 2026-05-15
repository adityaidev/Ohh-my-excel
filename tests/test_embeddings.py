from pathlib import Path
from excel_graph_mcp.embeddings import EmbeddingSearch, build_embeddings

FIXTURES = Path(__file__).parent / "fixtures"


def test_embedding_availability():
    es = EmbeddingSearch(FIXTURES / "simple.xlsx")
    assert isinstance(es.is_available(), bool)


def test_build_embeddings():
    result = build_embeddings(str(FIXTURES / "simple.xlsx"))
    assert isinstance(result, dict)
