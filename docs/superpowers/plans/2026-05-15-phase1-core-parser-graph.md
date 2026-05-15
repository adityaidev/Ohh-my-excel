# Phase 1: Core Parser + Graph Implementation Plan

**Goal:** Excel parser, formula AST extraction, SQLite graph store, basic MCP tools (build, query, stats, context)

**Architecture:** FastMCP server with SQLite WAL graph store. openpyxl reads .xlsx -> formula_parser extracts refs -> graph.py builds nodes/edges -> tools/ exposes via MCP. HyperFormula sidecar optional, `formulas` Python lib primary.

**Tech Stack:** Python 3.10+, FastMCP, openpyxl, NetworkX, SQLite (WAL), `formulas` lib

**Build Order:** constants -> formula_parser -> parser -> graph -> dependency -> tools -> main -> cli -> tests

---

### Task 1: constants.py - Config and constants

**File:** `excel_graph_mcp/constants.py`

- [ ] Write constants: DB path, schema version, node/edge types, default configs
- [ ] Test: `tests/test_constants.py`
- [ ] Commit

### Task 2: formula_parser.py - Formula AST parser

**File:** `excel_graph_mcp/formula_parser.py`

- [ ] Use `formulas` lib to parse Excel formulas into AST
- [ ] Extract cell references (A1, $A$1, Sheet!A1)
- [ ] Extract range references (A1:B10)
- [ ] Extract function names, nesting depth
- [ ] Handle edge cases: INDIRECT (flag AMBIGUOUS), named ranges, structured refs
- [ ] Test: `tests/test_formula_parser.py`
- [ ] Commit

### Task 3: parser.py - Excel workbook parser

**File:** `excel_graph_mcp/parser.py`

- [ ] openpyxl reads .xlsx: sheets, cells, formulas, tables, charts, named ranges
- [ ] Extract all cell values + formulas per sheet
- [ ] Identify tables (ListObjects) and their structured references
- [ ] Extract charts and their data sources
- [ ] Extract named ranges
- [ ] Output: dict of parsed workbook data
- [ ] Test: `tests/test_parser.py`
- [ ] Commit

### Task 4: graph.py - SQLite-backed graph store

**File:** `excel_graph_mcp/graph.py`

- [ ] SQLite schema: nodes table + edges table + FTS5 index
- [ ] WAL mode, concurrent reads
- [ ] Node CRUD: add/update/get/delete
- [ ] Edge CRUD: add/update/get/delete
- [ ] BFS traversal for impact radius
- [ ] Build NetworkX cache in memory
- [ ] Incremental update support (sheet-level hashing)
- [ ] Test: `tests/test_graph.py`
- [ ] Commit

### Task 5: dependency.py - Dependency graph builder

**File:** `excel_graph_mcp/dependency.py`

- [ ] Build full dependency graph from parser output + formula refs
- [ ] Range node optimization (SUM(A1:A100) = 1 edge, not 100)
- [ ] Detect circular references (NetworkX find_cycle)
- [ ] Detect broken references
- [ ] Compute edge confidence (EXTRACTED/INFERRED/AMBIGUOUS)
- [ ] Test: `tests/test_dependency.py`
- [ ] Commit

### Task 6: tools/ - MCP tool implementations

**Files:** `excel_graph_mcp/tools/_common.py`, `build.py`, `query.py`, `context.py`

- [ ] `_common.py`: shared validation, path sanitization, error wrappers
- [ ] `build.py`: build_or_update_graph_tool, run_postprocess_tool, list_graph_stats_tool
- [ ] `query.py`: query_graph_tool, traverse_graph_tool, semantic_search_nodes_tool
- [ ] `context.py`: get_minimal_context_tool
- [ ] Test: `tests/test_tools.py`
- [ ] Commit

### Task 7: main.py - FastMCP server entry point

**File:** `excel_graph_mcp/main.py`

- [ ] FastMCP server with stdio transport
- [ ] Register all Phase 1 tools
- [ ] Register 5 prompt templates
- [ ] Server lifecycle hooks
- [ ] Test: `tests/test_main.py`
- [ ] Commit

### Task 8: cli.py - CLI entry point

**File:** `excel_graph_mcp/cli.py`

- [ ] Commands: build, update, watch, status, visualize, serve
- [ ] Argument parsing with argparse
- [ ] Test: `tests/test_cli.py`
- [ ] Commit
