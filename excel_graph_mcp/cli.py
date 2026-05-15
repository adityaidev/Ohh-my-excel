import argparse
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(
        prog="ohh-my-excel",
        description="Persistent knowledge graph for Excel workbooks — safe, fast, token-efficient",
    )
    sub = parser.add_subparsers(dest="command", help="Available commands")

    build_p = sub.add_parser("build", help="Build knowledge graph for a workbook")
    build_p.add_argument("file_path", help="Path to .xlsx file")
    build_p.add_argument("--full-rebuild", action="store_true", help="Force full rebuild")

    update_p = sub.add_parser("update", help="Incremental update of an existing graph")
    update_p.add_argument("file_path", help="Path to .xlsx file")

    status_p = sub.add_parser("status", help="Show graph status for a workbook")
    status_p.add_argument("file_path", help="Path to .xlsx file")

    serve_p = sub.add_parser("serve", help="Start the MCP server")
    serve_p.add_argument("--host", default="localhost", help="Host to bind (default: localhost)")
    serve_p.add_argument("--port", type=int, default=8080, help="Port (default: 8080)")
    serve_p.add_argument("--transport", choices=["stdio", "sse"], default="stdio", help="Transport (default: stdio)")

    sub.add_parser("version", help="Show version")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    if args.command == "build":
        _cmd_build(args.file_path, args.full_rebuild)
    elif args.command == "update":
        _cmd_update(args.file_path)
    elif args.command == "status":
        _cmd_status(args.file_path)
    elif args.command == "serve":
        _cmd_serve(args.host, args.port, args.transport)
    elif args.command == "version":
        from excel_graph_mcp.constants import VERSION
        print(VERSION)


def _cmd_build(file_path: str, full_rebuild: bool):
    from excel_graph_mcp.tools.build import build_or_update_graph
    result = build_or_update_graph(file_path, full_rebuild=full_rebuild)
    print(f"Built graph: {result.get('nodes', 0)} nodes, {result.get('edges', 0)} edges")


def _cmd_update(file_path: str):
    from excel_graph_mcp.tools.build import build_or_update_graph
    result = build_or_update_graph(file_path, full_rebuild=False)
    print(f"Updated graph: {result.get('nodes', 0)} nodes, {result.get('edges', 0)} edges")


def _cmd_status(file_path: str):
    p = Path(file_path).resolve()
    from excel_graph_mcp.constants import get_db_path
    db = get_db_path(p)
    if db.exists():
        from excel_graph_mcp.graph import GraphStore
        store = GraphStore(p)
        stats = store.stats()
        store.close()
        print(f"Graph for: {p.name}")
        print(f"  Nodes: {stats['nodes']}")
        print(f"  Edges: {stats['edges']}")
        print(f"  By type: {stats.get('by_type', {})}")
    else:
        print(f"No graph found for: {p.name}. Run 'build' first.")


def _cmd_serve(host: str, port: int, transport: str):
    from excel_graph_mcp.main import serve
    serve(host=host, port=port, transport=transport)


if __name__ == "__main__":
    main()
