# CLAUDE.md - Project Context

## Project Overview

**Ohh-my-excel** is a persistent, incrementally-updated knowledge graph for Excel workbooks.
It parses Excel files into a structural graph of sheets, tables, cells, formulas, and their
dependencies — then exposes this graph via MCP tools so AI assistants get precise,
token-efficient context instead of dumping entire spreadsheets.

Three pillars: **safe** (backups, validation, no corruption), **fast** (sub-second graph builds),
**token-efficient** (minimal context per query).

## Graph Tool Usage (Token-Efficient)

1. **First call**: `get_minimal_context(task="<description>")` — costs ~100 tokens
2. **All subsequent calls**: use `detail_level="minimal"` unless you need more
3. **Prefer** `query_graph` with a specific target over broad `list_*` calls
4. The `next_tool_suggestions` field tells you the optimal next step
5. **Target**: ~5 tool calls per task, ~800 total tokens

## Architecture

- **Core Package**: `excel_graph_mcp/` (Python 3.10+)
  - `parser.py` - Excel multi-sheet AST parser (cells, formulas, tables, charts, ranges)
  - `graph.py` - SQLite-backed graph store (nodes, edges, BFS impact analysis)
  - `main.py` - FastMCP server entry point (stdio transport), registers tools + prompts
  - `formula_parser.py` - Excel formula AST parser (cell reference extraction)
  - `dependency.py` - Dependency graph builder (precedents, dependents, range optimization)
  - `incremental.py` - File-hash change detection, file watching
  - `visualization.py` - D3.js / Pyvis interactive HTML graph generator
  - `exports.py` - Export as JSON, CSV, GraphML, Cypher
  - `cli.py` - CLI entry point (install, build, update, watch, status, visualize, serve)

- **Database**: `.excel-graph/graph.db` (SQLite, WAL mode)

## Key Commands

```bash
# Development
uv run pytest tests/ --tb=short -q
uv run ruff check excel_graph_mcp/
```
