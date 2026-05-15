# For Reference and architecture look this :: ""C:\Users\adity\Desktop\ohh-my-Excel\code-review-graph""
and create a new repo via using my Github CLI named : Ohh-my-excel

# Excel Graph MCP — Product Requirements Document

Setup : Project name :-- Ohh-my-excel

**Version:** 1.1.0
**Status:** DRAFT
**Created:** 2026-05-15
**Last Updated:** 2026-05-15

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Competitive Landscape](#2-competitive-landscape)
3. [Architecture Overview](#3-architecture-overview)
4. [MCP Tools (38 Tools)](#4-mcp-tools-38-tools)
5. [Workbook Generation — Create from Scratch](#5-workbook-generation--create-from-scratch)
6. [Dependencies & Libraries](#6-dependencies--libraries)
7. [Latency & Accuracy Strategy](#7-latency--accuracy-strategy)
8. [Visualization Strategy](#8-visualization-strategy)
9. [Graph Schema](#9-graph-schema)
10. [Token Efficiency Strategy](#10-token-efficiency-strategy)
11. [Incremental Update Strategy](#11-incremental-update-strategy)
12. [Supported Formats](#12-supported-formats)
13. [Security](#13-security)
14. [Development Phases & Milestones](#14-development-phases--milestones)
15. [Success Metrics](#15-success-metrics)
16. [Risk Assessment](#16-risk-assessment)
17. [Open Questions & Decisions](#17-open-questions--decisions)
18. [References](#18-references)

---

## 1. Executive Summary

**Excel Graph MCP** is a persistent, incrementally-updated knowledge graph for Excel workbooks. It parses Excel files into a structural graph of sheets, tables, cells, formulas, and their dependencies — then exposes this graph via MCP tools so AI assistants (Claude Code, OpenCode, Codex, Cursor, etc.) get precise, token-efficient context instead of dumping entire spreadsheets.

**The Problem:** Existing Excel MCP servers are CRUD-only — they read/write cells and dump raw data. None provide structural understanding of formula dependencies, blast radius analysis, or token-efficient context. When an AI assistant opens a 50-sheet financial model, it reads everything, burning tokens and missing critical relationships.

**The Solution:** Build a Tree-sitter-equivalent for Excel — an AST parser for formulas, a dependency graph engine, and an MCP server that gives AI assistants exactly what they need to know.

**Scope:** 38 MCP tools across graph management, analysis, workbook generation, visualization, and export — plus 5 built-in prompt templates.

---

## 2. Competitive Landscape

> Star counts captured 2026-05-15 from GitHub.

| Server | Stars | Approach | Gap |
|--------|-------|----------|-----|
| haris-musa/excel-mcp-server | 3.8k | OpenPyXL CRUD | No graph, no dependencies, no token efficiency |
| negokaz/excel-mcp-server | 951 | Go + file-based | Read/write only, no structural analysis |
| sbroenne/mcp-server-excel | — | .NET COM API | Controls real Excel, no graph |
| sbraind/excel-mcp-server | 9 | ExcelJS | 34 CRUD tools, no dependency tracking |
| yzfly/mcp-excel-server | 90 | pandas | Analysis-focused, no graph |
| **Ohh-my-Excel (us)** | — | **Knowledge Graph** | **First structural graph + dependency engine** |

---

## 3. Architecture Overview

### 3.1 System Design

```
┌─────────────────────────────────────────────────────────┐
│                    MCP Client (AI Agent)                  │
│         OpenCode / Claude Code / Codex / Cursor          │
└────────────────────┬────────────────────────────────────┘
                     │  stdio / Streamable HTTP
                     ▼
┌─────────────────────────────────────────────────────────┐
│              FastMCP Server (main.py)                     │
│         38 MCP Tools + 5 Prompt Templates                │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│              Graph Store (graph.py)                       │
│         SQLite WAL mode — nodes, edges, BFS              │
│         .excel-graph/graph.db                            │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│              Excel Parser (parser.py)                     │
│         openpyxl + formulas AST + cell ref extraction    │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│              Excel Workbook (.xlsx / .xlsm)              │
│         Sheets → Tables → Cells → Formulas → Ranges      │
└─────────────────────────────────────────────────────────┘
```

### 3.2 Directory Structure

```
excel_graph_mcp/
├── __init__.py
├── __main__.py
├── cli.py                  # CLI entry point (install, build, update, watch, status, visualize, serve)
├── parser.py               # Excel multi-sheet AST parser (cells, formulas, tables, charts, ranges)
├── graph.py                # SQLite-backed graph store (nodes, edges, BFS impact analysis)
├── main.py                 # FastMCP server entry point (stdio transport), registers tools + prompts
├── incremental.py          # File-hash change detection, file watching
├── formula_parser.py       # Excel formula AST parser (cell reference extraction)
├── dependency.py           # Dependency graph builder (precedents, dependents, range optimization)
├── communities.py          # Sheet community detection (grouped by cross-sheet references)
├── flows.py                # Data flow detection (input → calculation → output chains)
├── changes.py              # Risk-scored change impact analysis
├── search.py               # FTS5 hybrid search (keyword + vector)
├── embeddings.py           # Optional vector embeddings
├── visualization.py        # D3.js / Pyvis interactive HTML graph generator
├── exports.py              # Export as JSON, CSV, GraphML, Cypher
├── prompts.py              # MCP prompt templates
├── skills.py               # Skill definitions for AI coding platforms
├── constants.py            # Constants and configuration
├── migrations.py           # Database schema migrations
├── backup.py               # Automatic backup/restore before destructive writes
├── tools/                  # MCP tool implementations
│   ├── __init__.py
│   ├── _common.py
│   ├── build.py
│   ├── query.py
│   ├── context.py
│   ├── analysis_tools.py
│   ├── community_tools.py
│   ├── flows_tools.py
│   ├── docs.py
│   └── registry_tools.py
├── tests/                  # Test suite
├── docs/                   # Documentation
├── skills/                 # Built-in skills
├── pyproject.toml
└── README.md
```

---

## 4. MCP Tools (38 Tools)

### 4.1 Graph Setup & Maintenance

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 1 | `build_or_update_graph_tool` | Build or incrementally update the Excel knowledge graph | `file_path`, `full_rebuild`, `detail_level` |
| 2 | `run_postprocess_tool` | Run post-processing (flows, communities, FTS) | `flows`, `communities`, `fts`, `file_path` |
| 3 | `list_graph_stats_tool` | Get aggregate statistics about the graph | `file_path` |

### 4.2 Excel Exploration

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 4 | `get_minimal_context_tool` | Ultra-compact context for any Excel task (~100 tokens) | `task`, `file_path` |
| 5 | `semantic_search_nodes_tool` | Search for sheets, cells, tables, formulas by name/keyword | `query`, `kind`, `limit`, `file_path` |
| 6 | `query_graph_tool` | Predefined graph queries (precedents_of, dependents_of, references_of, contains, tested_by) | `pattern`, `target`, `file_path` |
| 7 | `traverse_graph_tool` | BFS/DFS traversal from any cell/sheet with token budget | `query`, `mode`, `depth`, `token_budget`, `file_path` |

### 4.3 Dependency & Impact Analysis

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 8 | `get_impact_radius_tool` | Blast radius when a cell/formula changes | `cell_ref`, `max_depth`, `file_path` |
| 9 | `detect_changes_tool` | Risk-scored change analysis between workbook versions | `old_file`, `new_file`, `detail_level` |
| 10 | `get_affected_flows_tool` | Which data flows are impacted by a cell change | `cell_ref`, `file_path` |

### 4.4 Formula Analysis

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 11 | `get_formula_dependencies_tool` | Get all precedent cells a formula depends on | `cell_ref`, `file_path`, `include_ranges` |
| 12 | `get_formula_dependents_tool` | Get all cells that depend on a given cell | `cell_ref`, `file_path` |
| 13 | `find_circular_references_tool` | Detect circular formula dependencies | `file_path` |
| 14 | `find_broken_references_tool` | Find formulas referencing deleted/missing cells | `file_path` |
| 15 | `get_formula_complexity_tool` | Analyze formula complexity (nesting depth, function count) | `cell_ref`, `file_path` |

### 4.5 Sheet & Workbook Analysis

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 16 | `list_sheets_tool` | List all sheets with metadata (cell count, formula count, cross-refs) | `file_path` |
| 17 | `get_sheet_info_tool` | Detailed info about a specific sheet | `sheet_name`, `file_path` |
| 18 | `find_cross_sheet_references_tool` | Find all cross-sheet formula references | `file_path` |
| 19 | `list_tables_tool` | List all Excel tables with their ranges and columns | `file_path` |
| 20 | `list_named_ranges_tool` | List all named ranges and their definitions | `file_path` |

### 4.6 Data Flow Analysis

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 21 | `list_flows_tool` | List data flows (input → calculation → output chains) | `sort_by`, `limit`, `file_path` |
| 22 | `get_flow_tool` | Details of a specific data flow | `flow_id`, `flow_name`, `file_path` |

### 4.7 Architecture & Quality

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 23 | `get_architecture_overview_tool` | High-level workbook architecture from sheet communities | `file_path` |
| 24 | `list_communities_tool` | List detected sheet communities (grouped by cross-refs) | `sort_by`, `file_path` |
| 25 | `get_hub_cells_tool` | Most-referenced cells (architectural hotspots) | `top_n`, `file_path` |
| 26 | `get_bridge_cells_tool` | Cells that connect different sheet communities | `top_n`, `file_path` |
| 27 | `get_knowledge_gaps_tool` | Structural weaknesses (orphan sheets, untested calculations) | `file_path` |
| 28 | `get_suggested_questions_tool` | Auto-generated audit questions from graph analysis | `file_path` |

### 4.8 Visualization & Export

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 29 | `visualize_graph_tool` | Generate interactive HTML graph visualization | `file_path`, `format`, `depth` |
| 30 | `export_graph_tool` | Export the graph as JSON, CSV, GraphML, or Cypher | `file_path`, `format` |

### 4.9 MCP Prompts (5 Built-in)

| # | Prompt | Description |
|---|--------|-------------|
| 1 | `audit_workbook` | Comprehensive workbook audit with risk levels |
| 2 | `data_flow_map` | Data flow documentation with Mermaid diagrams |
| 3 | `debug_formula` | Guided formula debugging workflow |
| 4 | `onboard_analyst` | New analyst orientation to workbook structure |
| 5 | `pre_merge_check` | Workbook change readiness with risk scoring |

---

## 5. Workbook Generation — Create from Scratch

### 5.1 The Missing Feature

Every existing Excel MCP server can READ and EDIT cells. None can CREATE a professional workbook from a natural language prompt. This is the #1 user request:

> *"Create a monthly expense tracker with categories, automatic totals, and a summary dashboard"*
> *"Build a sales pipeline tracker with deal stages, probability weighting, and forecast"*
> *"Make a project timeline with Gantt-style conditional formatting"*

### 5.2 Generation Architecture

```
User Prompt
    │
    ▼
┌─────────────────────────────────────────────────────┐
│  Step 1: LLM Planning (structured output)            │
│  Model: claude-sonnet-4-20250514                     │
│  Prompt → JSON Schema → WorkbookPlan                 │
│  - Sheet definitions, column schemas                 │
│  - Formula templates, chart specs                    │
│  - Formatting rules, table structures                │
│  - On malformed JSON: strip fences, retry once,      │
│    then return structured error to caller            │
└────────────────────┬────────────────────────────────┘
                     │  Validated JSON (Pydantic)
                     ▼
┌─────────────────────────────────────────────────────┐
│  Step 2: Deterministic Execution                     │
│  WorkbookPlan → XlsxWriter/OpenPyXL → .xlsx file    │
│  - No LLM involved — pure Python execution           │
│  - Guaranteed valid output                           │
│  - Retry-safe (deterministic)                        │
└────────────────────┬────────────────────────────────┘
                     │  Generated .xlsx
                     ▼
┌─────────────────────────────────────────────────────┐
│  Step 3: Graph Build (auto-index)                    │
│  .xlsx → Parser → Graph Store                        │
│  - Immediately available for analysis                │
│  - Dependency graph built                            │
└─────────────────────────────────────────────────────┘
```

### 5.3 Workbook Generation MCP Tools (Tools 31–38)

| # | Tool | Description | Parameters |
|---|------|-------------|------------|
| 31 | `generate_workbook_tool` | **PRIMARY** — Create Excel workbook from natural language prompt | `prompt`, `output_path`, `template`, `data_source`, `style` |
| 32 | `generate_workbook_from_data_tool` | Create workbook from provided data (JSON, CSV, pasted table) | `data`, `output_path`, `sheet_name`, `include_charts`, `style` |
| 33 | `generate_workbook_from_template_tool` | Create workbook from built-in template | `template_name`, `output_path`, `customizations` |
| 34 | `add_sheet_tool` | Add a new sheet to existing workbook with structure | `file_path`, `sheet_name`, `columns`, `formulas`, `formatting` |
| 35 | `add_chart_tool` | Add chart to existing workbook | `file_path`, `sheet_name`, `chart_type`, `data_range`, `title` |
| 36 | `apply_formatting_tool` | Apply professional formatting to workbook | `file_path`, `style`, `auto_size`, `freeze_panes` |
| 37 | `generate_formulas_tool` | Auto-generate formulas for columns based on context | `file_path`, `sheet_name`, `column`, `formula_type` |
| 38 | `validate_workbook_tool` | Validate generated workbook (formulas, refs, structure) | `file_path`, `check_circular`, `check_broken_refs` |

> **Note on write tools (34–37):** All tools that modify existing workbooks automatically create a timestamped backup under `.excel-graph/backups/` before writing. Use `restore_backup_tool` (CLI) or the `--restore` flag to roll back. See §5.10 for details.

### 5.4 WorkbookPlan JSON Schema

The LLM generates a structured plan that Python executes deterministically:

```json
{
  "workbook_name": "Monthly Expense Tracker",
  "sheets": [
    {
      "name": "Expenses",
      "columns": [
        {"name": "Date", "type": "date", "format": "YYYY-MM-DD", "width": 12},
        {"name": "Category", "type": "text", "width": 18, "validation": {"type": "list", "values": ["Food", "Transport", "Utilities", "Entertainment", "Other"]}},
        {"name": "Description", "type": "text", "width": 30},
        {"name": "Amount", "type": "currency", "format": "$#,##0.00", "width": 14},
        {"name": "Payment Method", "type": "text", "width": 16, "validation": {"type": "list", "values": ["Cash", "Card", "UPI", "Bank Transfer"]}}
      ],
      "table": {"name": "ExpenseTable", "style": "TableStyleMedium9"},
      "formulas": [
        {"cell": "F2", "formula": "=SUBTOTAL(9,ExpenseTable[Amount])", "label": "Total"}
      ],
      "conditional_formatting": [
        {"range": "D2:D1000", "type": "greater_than", "value": 1000, "format": {"fill": "#FFC7CE", "font_color": "#9C0006"}}
      ]
    },
    {
      "name": "Summary",
      "columns": [
        {"name": "Category", "type": "text", "width": 18},
        {"name": "Total Spent", "type": "currency", "format": "$#,##0.00", "width": 14},
        {"name": "% of Budget", "type": "percentage", "format": "0.0%", "width": 12}
      ],
      "formulas": [
        {"cell": "A2", "formula": "=UNIQUE(Expenses!B:B)"},
        {"cell": "B2", "formula": "=SUMIF(Expenses!B:B,A2,Expenses!D:D)"},
        {"cell": "C2", "formula": "=B2/$B$10"}
      ],
      "chart": {
        "type": "pie",
        "data_range": "A2:B8",
        "title": "Expenses by Category",
        "position": "D2"
      }
    }
  ],
  "styling": {
    "theme": "professional",
    "header_format": {"bold": true, "fill": "#4472C4", "font_color": "#FFFFFF"},
    "auto_size_columns": true,
    "freeze_panes": {"sheet": "Expenses", "cell": "A2"}
  }
}
```

### 5.5 Built-in Templates (20+)

| Category | Templates |
|----------|-----------|
| **Finance** | Expense Tracker, Budget Planner, Invoice Generator, P&L Statement, Cash Flow, Balance Sheet |
| **Sales** | Sales Pipeline, CRM Tracker, Lead Scoring, Revenue Forecast, Commission Calculator |
| **Project** | Project Timeline, Gantt Chart, Sprint Planner, Resource Allocation, Risk Register |
| **HR** | Employee Directory, Attendance Tracker, Performance Review, Leave Calendar |
| **Marketing** | Campaign Tracker, Content Calendar, Social Media Planner, ROI Calculator |
| **Personal** | Habit Tracker, Meal Planner, Workout Log, Travel Budget, Grade Calculator |

### 5.6 Data Import Sources

| Source | Tool | Description |
|--------|------|-------------|
| **JSON data** | `generate_workbook_from_data_tool` | Paste JSON → auto-detect schema → create workbook |
| **CSV data** | `generate_workbook_from_data_tool` | Paste CSV → parse → create formatted workbook |
| **Pasted table** | `generate_workbook_from_data_tool` | Markdown/TSV table → parse → create workbook |
| **API endpoint** | `generate_workbook_from_data_tool` | Fetch JSON from API → transform → create workbook |
| **Database query** | `generate_workbook_from_data_tool` | SQL query results → create workbook |
| **Manual entry** | `generate_workbook_tool` | Natural language → LLM generates sample data → create workbook |

### 5.7 Generation Workflow Patterns

#### Pattern 1: Natural Language → Workbook

```
User: "Create a monthly expense tracker with categories, automatic totals, and a summary dashboard"

1. LLM generates WorkbookPlan JSON (validated against schema)
2. Python executes: XlsxWriter creates .xlsx file
3. Auto-build graph: parser indexes the new workbook
4. Return: file path + summary ("Created 2 sheets, 5 columns, 3 formulas, 1 chart")
```

#### Pattern 2: Data → Workbook

```
User: "Here's my sales data in JSON, make it a professional report with charts"

1. Parse JSON data → detect schema
2. LLM generates WorkbookPlan with charts and formatting
3. Python executes: XlsxWriter creates .xlsx with data + charts
4. Auto-build graph
5. Return: file path + summary
```

#### Pattern 3: Template → Customized Workbook

```
User: "Create a budget planner from template but add a crypto investment section"

1. Load built-in template plan
2. LLM modifies plan with customizations
3. Python executes
4. Auto-build graph
5. Return: file path + summary
```

### 5.8 Library Strategy for Generation

| Task | Library | Why |
|------|---------|-----|
| **Create new workbook** | **XlsxWriter** | 2x faster than openpyxl for writes, best chart support, clean API |
| **Modify existing workbook** | **OpenPyXL** | Read+write support, preserves existing formatting |
| **Bulk data write** | **pandas + XlsxWriter** | Fastest for large datasets (10K+ rows) |
| **Formula evaluation** | **formulas** | Validate formulas before writing |
| **Template loading** | **OpenPyXL** | Read template, then modify with XlsxWriter patterns |

### 5.9 Quality Assurance

| Check | Implementation |
|-------|---------------|
| **Schema validation** | Pydantic validates WorkbookPlan before execution |
| **Formula validation** | `formulas` library parses each formula before writing |
| **Circular reference check** | NetworkX detects cycles in formula dependency graph |
| **Broken reference check** | Verify all sheet/cell references exist |
| **Retry on failure** | If LLM returns malformed JSON: strip fences → retry once → return structured error |
| **File size limit** | Configurable max file size (default: 50MB) |
| **Deterministic output** | Same plan → same file (no randomness in execution) |

### 5.10 Rollback & Backup Strategy

Write tools (`add_sheet_tool`, `add_chart_tool`, `apply_formatting_tool`, `generate_formulas_tool`) can modify existing workbooks. To protect against data loss:

- Before any write, a timestamped backup is created at `.excel-graph/backups/<filename>_<timestamp>.xlsx`
- Backups are retained for 7 days by default (configurable via `EXCEL_GRAPH_BACKUP_RETENTION_DAYS`)
- Restore via CLI: `excel-graph restore <file_path> [--timestamp <ts>]`
- Max 10 backups per workbook (oldest pruned automatically)
- Generation tools that create *new* files do not back up (no pre-existing file to protect)

### 5.11 Token Efficiency for Generation

| Step | Tokens | Notes |
|------|--------|-------|
| LLM planning | 2-5K | One-time cost to generate WorkbookPlan |
| Execution | 0 | Pure Python, no LLM involved |
| Graph build | 0 | Local parsing, no tokens |
| **Total** | **2-5K** | vs 500K+ for reading entire workbook |

### 5.12 Prompt Engineering for Generation

The system prompt for WorkbookPlan generation (model: `claude-sonnet-4-20250514`):

```
You are an Excel workbook architect. Given a user's description, generate a complete
WorkbookPlan JSON that specifies every sheet, column, formula, chart, and formatting rule.

Rules:
1. Every column must have: name, type, width
2. Formulas must use Excel syntax (A1 notation)
3. Cross-sheet references must use 'SheetName'!A1 format
4. Include data validation for categorical columns
5. Add conditional formatting for outliers/errors
6. Create summary sheets with charts for data sheets
7. Use professional color schemes
8. Include auto-totals and subtotals where appropriate

Output: Valid WorkbookPlan JSON only. No preamble, no explanations, no markdown fences.
```

---

## 6. Dependencies & Libraries

### 6.1 Core Dependencies

| Library | Version | Purpose | Why |
|---------|---------|---------|-----|
| **openpyxl** | ≥3.1.2 | Read/write .xlsx files, extract formulas, formatting, charts | Industry standard, read+write support, formula access |
| **formulas** | ≥1.3.4 | Excel formula AST parser, dependency graph builder, workbook evaluation | Parses formulas into AST, extracts cell references, builds dispatch pipes |
| **networkx** | ≥3.2 | Graph data structures, BFS/DFS, betweenness centrality, topological sort | Proven for graph algorithms |
| **fastmcp** | ≥3.2.4 | MCP server framework | Async support, stdio transport |
| **mcp** | ≥1.0.0 | MCP SDK | Protocol implementation |
| **watchdog** | ≥4.0.0 | File system watching for incremental updates | Cross-platform file events |

### 6.2 Visualization Dependencies

| Library | Version | Purpose | Why |
|---------|---------|---------|-----|
| **pyvis** | ≥0.3.2 | Interactive HTML network graphs (vis.js wrapper) | Python-native, generates standalone HTML, drag-and-drop, physics layout |
| **plotly** | ≥5.18.0 | Interactive charts and Sankey diagrams for data flows | Best-in-class interactivity, Sankey for data flow visualization |
| **matplotlib** | ≥3.8.0 | Static graph visualization (fallback) | Publication-quality static graphs |

### 6.3 Optional Dependencies

| Library | Version | Purpose | Why |
|---------|---------|---------|-----|
| **pandas** | ≥2.1.0 | Bulk data reading/writing, chunked processing | Fast bulk operations for large datasets |
| **xlsxwriter** | ≥3.1.9 | High-performance .xlsx creation (write-only) | 2x faster than openpyxl for writes |
| **pyxlsb** | ≥1.0.10 | Read .xlsb (binary Excel) files | Binary format support for large files |
| **oletools** | ≥0.60.1 | Extract VBA macros from .xlsm files | Security audit, macro analysis |
| **sentence-transformers** | ≥2.2.0 | Vector embeddings for semantic search | Optional semantic search |
| **xlcalculator** | — | Formula recalculation engine | Validate formula results |

### 6.4 pyproject.toml Dependencies

```toml
[project]
dependencies = [
    "fastmcp>=3.2.4",
    "mcp>=1.0.0",
    "networkx>=3.2",
    "openpyxl>=3.1.2",
    "formulas>=1.3.4",
    "watchdog>=4.0.0",
    "pyvis>=0.3.2",
    "plotly>=5.18.0",
]

[project.optional-dependencies]
embeddings = ["sentence-transformers>=2.2.0"]
xlsb = ["pyxlsb>=1.0.10"]
macros = ["oletools>=0.60.1"]
performance = ["pandas>=2.1.0", "xlsxwriter>=3.1.9"]
all = [
    "sentence-transformers>=2.2.0",
    "pyxlsb>=1.0.10",
    "oletools>=0.60.1",
    "pandas>=2.1.0",
    "xlsxwriter>=3.1.9",
]
```

---

## 7. Latency & Accuracy Strategy

### 7.1 Lowest Latency

| Technique | Implementation | Expected Impact |
|-----------|---------------|-----------------|
| **Read-only mode** | `openpyxl.load_workbook(data_only=False, read_only=True)` for parsing | 3-5x faster reads, 10x less memory |
| **SHA-256 hash skip** | Hash each sheet's XML content; skip unchanged sheets on incremental update | Sub-second updates for small changes |
| **SQLite WAL mode** | `PRAGMA journal_mode=WAL` for concurrent reads | Zero read locks, parallel queries |
| **Cached directed graph** | NetworkX graph cached in memory, rebuilds only on writes | O(1) query response after initial build |
| **Range node optimization** | Represent `SUM(A1:A1000)` as range node (not 1000 edges) — inspired by HyperFormula | 100x fewer edges for large ranges |
| **Lazy cell loading** | Only load cell values when queried, not during graph build | 50% faster initial build |
| **Async thread offload** | `asyncio.to_thread()` for blocking operations | No event loop blocking |
| **FTS5 index** | SQLite FTS5 for keyword search across sheet/cell names | Sub-10ms search |
| **Batch cell reads** | Read entire sheet ranges in one openpyxl call, not cell-by-cell | 10-50x faster than individual reads |
| **Memory-mapped parsing** | For files >50MB, use streaming parser instead of full load | Constant memory regardless of file size |

### 7.2 Perfect Accuracy

| Technique | Implementation | Why |
|-----------|---------------|-----|
| **AST-based formula parsing** | Use `formulas` library to parse formulas into AST, then extract references via `extract_references_from_ast()` | Regex misses edge cases (quoted sheet names, INDIRECT, etc.) |
| **Named range resolution** | Resolve named ranges to actual cell references before building graph | Named ranges are common in financial models |
| **Cross-sheet reference parsing** | Parse `Sheet1!A1`, `'Sheet Name'!B2:C10` patterns via AST | Cross-sheet refs are the hardest to get right |
| **Range expansion** | Expand range references (`A1:B10`) into individual cell edges with range node optimization | Accurate blast radius requires full expansion |
| **INDIRECT/HYPERLINK handling** | Flag dynamic references as "AMBIGUOUS" edges. Statically-resolvable cases (e.g. `INDIRECT("Sheet1!A1")`) will be resolved in a future version | These can't be fully statically resolved |
| **Array formula support** | Parse multi-cell array formulas as single nodes with multi-cell output | Array formulas affect multiple cells |
| **Table reference support** | Parse structured references (`Table1[Column]`) into cell ranges | Modern Excel uses tables extensively |
| **External workbook refs** | Detect and flag references to external workbooks | Cannot resolve without the external file |
| **Circular reference detection** | Use NetworkX `find_cycle()` to detect circular dependencies | Critical for financial model auditing |
| **Edge confidence scoring** | Three-tier system: EXTRACTED (AST), INFERRED (range expansion), AMBIGUOUS (INDIRECT) | Transparency about certainty |

### 7.3 Performance Targets

| Metric | Target | Measurement |
|--------|--------|-------------|
| Initial build (100-sheet workbook) | <30 seconds | `time build_or_update_graph(full_rebuild=True)` |
| Incremental update (1 cell change) | <500ms | `time build_or_update_graph()` after edit |
| Query response (any tool) | <100ms | P95 latency |
| Search response | <10ms | FTS5 keyword search |
| Memory usage (100-sheet workbook) | <200MB | RSS during build |
| Graph DB size (100-sheet workbook) | <50MB | `.excel-graph/graph.db` file size |

> **Watch mode note:** The watch-mode debounce is set to 5 seconds (not 300ms) to avoid triggering incremental updates mid-build on large workbooks. For workbooks that build in <1s, this can be reduced via `EXCEL_GRAPH_WATCH_DEBOUNCE_MS`.

---

## 8. Visualization Strategy

### 8.1 When User Asks for Visual Graphs

```
User: "Show me the dependency graph for this workbook"
```

**Flow:**
1. `visualize_graph_tool` generates interactive HTML
2. Returns file path: `.excel-graph/graph.html`
3. AI tells user to open in browser

### 8.2 Visualization Types

| Type | Library | Use Case | Output |
|------|---------|----------|--------|
| **Dependency Network** | Pyvis (vis.js) | Cell-to-cell formula dependencies | Interactive HTML with drag, zoom, physics |
| **Data Flow Sankey** | Plotly | Input → Calculation → Output chains | Interactive Sankey diagram |
| **Sheet Architecture** | Plotly | Cross-sheet reference matrix | Heatmap + network overlay |
| **Formula Tree** | Pyvis | Single formula's precedent tree | Hierarchical tree view |
| **Community Map** | Plotly | Sheet communities and coupling | Force-directed graph with clusters |
| **Static Fallback** | Matplotlib | Environments without browser | PNG/SVG export |

### 8.3 Pyvis Implementation (Primary)

```python
from pyvis.network import Network

def generate_dependency_html(graph: nx.DiGraph, output_path: str):
    net = Network(height="800px", width="100%", bgcolor="#1a1a2e", font_color="white")
    net.barnes_hut(gravity=-8000, central_gravity=0.3, spring_length=200)

    for node, data in graph.nodes(data=True):
        node_type = data.get("type", "cell")
        color = {
            "sheet": "#00d2ff",
            "cell": "#ffffff",
            "formula": "#ff6b6b",
            "table": "#4ecdc4",
            "range": "#ffe66d",
            "chart": "#a8e6cf"
        }.get(node_type, "#ffffff")
        net.add_node(node, label=data.get("label", node), color=color, title=data.get("description", ""))

    for src, tgt, data in graph.edges(data=True):
        confidence = data.get("confidence", "EXTRACTED")
        width = {"EXTRACTED": 2, "INFERRED": 1, "AMBIGUOUS": 0.5}.get(confidence, 1)
        color = {"EXTRACTED": "#4ecdc4", "INFERRED": "#ffe66d", "AMBIGUOUS": "#ff6b6b"}.get(confidence, "#888")
        net.add_edge(src, tgt, width=width, color=color, title=data.get("edge_type", ""))

    net.show(output_path, notebook=False)
```

### 8.4 Plotly Sankey for Data Flows

```python
import plotly.graph_objects as go

def generate_sankey_diagram(flows: list[dict], output_path: str):
    fig = go.Figure(data=[go.Sankey(
        node=dict(
            pad=15, thickness=20, line=dict(color="black", width=0.5),
            label=[f["source"] for f in flows] + [f["target"] for f in flows],
            color="blue"
        ),
        link=dict(
            source=[i for i in range(len(flows))],
            target=[i + len(flows) for i in range(len(flows))],
            value=[f["value"] for f in flows]
        )
    )])
    fig.write_html(output_path)
```

---

## 9. Graph Schema

### 9.1 Node Types

| Type | Description | Properties |
|------|-------------|------------|
| `Workbook` | The Excel file itself | `path`, `name`, `sheet_count`, `size_bytes` |
| `Sheet` | A worksheet | `name`, `index`, `cell_count`, `formula_count`, `has_table` |
| `Cell` | A single cell | `address` (A1), `sheet`, `value_type`, `has_formula`, `is_input` |
| `Formula` | A cell containing a formula | `formula_text`, `functions_used`, `nesting_depth`, `cell_count` |
| `Range` | A cell range (optimization node) | `start`, `end`, `sheet`, `cell_count` |
| `Table` | An Excel table | `name`, `range`, `columns`, `sheet` |
| `NamedRange` | A named range | `name`, `range`, `sheet`, `scope` |
| `Chart` | A chart object | `name`, `type`, `data_source`, `sheet` |

### 9.2 Edge Types

| Type | Description | Direction |
|------|-------------|-----------|
| `CONTAINS` | Workbook → Sheet, Sheet → Cell, Sheet → Table | Parent → Child |
| `REFERENCES` | Formula → Cell (precedent) | Dependent → Precedent |
| `RANGE_REF` | Formula → Range | Dependent → Range |
| `DEPENDS_ON` | Cell → Cell (computed) | Dependent → Precedent |
| `CROSS_SHEET_REF` | Cell → Cell (different sheets) | Dependent → Precedent |
| `TABLE_REF` | Formula → Table (structured ref) | Dependent → Table |
| `CHART_SOURCE` | Chart → Data range | Chart → Data |
| `NAMED_REF` | Formula → Named range | Dependent → NamedRange |

---

## 10. Token Efficiency Strategy

### 10.1 The Problem

Existing Excel MCP servers dump entire sheets as text. A 100-sheet workbook with 10,000 cells each = millions of tokens.

### 10.2 Our Solution

> **Methodology:** Estimates below are based on token counts for representative workbooks at each scenario size. Raw-dump counts use openpyxl `sheet.values` serialized as tab-separated text. Graph-query counts measure actual tool response payloads. Numbers will be validated against the benchmark suite in Phase 4.

| Scenario | Without Graph | With Graph | Reduction |
|----------|--------------|------------|-----------|
| Audit 50-sheet workbook | ~500k tokens | ~50k tokens | ~10x |
| Review cell change | ~500k tokens | ~5k tokens | ~100x |
| Find formula dependencies | ~500k tokens | ~2k tokens | ~250x |
| Understand data flow | ~500k tokens | ~10k tokens | ~50x |

### 10.3 Implementation

1. **`get_minimal_context_tool`** — ~100 tokens, returns graph stats + next_tool_suggestions
2. **`detail_level="minimal"`** — All tools support compact output
3. **Structural context over raw data** — Return "Cell C5 depends on A1:A100 via SUM" instead of dumping 100 cell values
4. **`next_tool_suggestions`** — Every response tells AI the optimal next tool
5. **Target: ≤5 tool calls per task, ≤800 total tokens**

---

## 11. Incremental Update Strategy

### 11.1 How It Works

```
File changed? → SHA-256 hash comparison → Only re-parse changed sheets → Update affected graph nodes → Rebuild cached NetworkX graph
```

### 11.2 Change Detection

| Trigger | Detection Method | Action |
|---------|-----------------|--------|
| File save | watchdog file event | Hash comparison, incremental parse |
| Git commit | `git diff --name-only` | Parse only changed files |
| Manual update | CLI `update` command | Hash comparison |
| Watch mode | `excel-graph watch` | 5s debounced auto-update (configurable) |

### 11.3 Hash Strategy

- Hash each sheet's XML content (not entire file)
- Skip sheets where hash unchanged
- Only rebuild graph nodes for changed sheets
- Update affected edges (precedents/dependents)

---

## 12. Supported Formats

| Format | Read | Write | Formula Parsing | Notes |
|--------|------|-------|-----------------|-------|
| `.xlsx` | ✅ | ✅ | ✅ | Primary format |
| `.xlsm` | ✅ | ✅ | ✅ | + VBA macro extraction |
| `.xltx` | ✅ | ✅ | ✅ | Templates |
| `.xltm` | ✅ | ✅ | ✅ | Macro-enabled templates |
| `.xlsb` | ✅ | ❌ | ⚠️ | Binary format, via pyxlsb |
| `.xls` | ⚠️ | ❌ | ❌ | Legacy, via xlrd (limited) |
| `.csv` | ✅ | ✅ | ❌ | No formulas |

---

## 13. Security

| Concern | Mitigation |
|---------|-----------|
| Path traversal | `_validate_file_path()` prevents escaping allowed directories |
| Prompt injection | `_sanitize_name()` strips control characters, caps at 256 chars |
| Macro execution | VBA macros are extracted as text, never executed |
| File size limits | Configurable max file size (default: 100MB read, 50MB generate) |
| External references | Flagged as AMBIGUOUS, never followed automatically |
| No eval/exec | No `eval()`, `exec()`, `pickle`, or `yaml.unsafe_load()` |
| Generation abuse | Rate limiting on `generate_workbook_tool` (default: 10 requests/min per client); configurable via `EXCEL_GRAPH_RATE_LIMIT`. Large prompt inputs (>4K chars) are truncated before LLM call. |
| Backup integrity | Backups are write-once; backup directory is outside the workbook's directory to prevent accidental deletion |

---

## 14. Development Phases & Milestones

### Phase 1: Core Parser + Graph (Weeks 1–2)
- [ ] Excel parser (openpyxl + formulas AST)
- [ ] Formula dependency extraction
- [ ] SQLite graph store
- [ ] Basic MCP tools (build, query, stats)
- [ ] **Milestone v0.1:** Internal alpha — graph builds correctly on 5 reference workbooks

### Phase 2: Analysis Tools (Weeks 3–4)
- [ ] Impact radius / blast radius
- [ ] Circular reference detection
- [ ] Cross-sheet reference mapping
- [ ] Data flow detection
- [ ] **Milestone v0.2:** Core analysis tools passing against real-world financial model test suite

### Phase 3: Advanced Features (Weeks 5–6)
- [ ] Sheet community detection
- [ ] Risk-scored change analysis
- [ ] Visualization (Pyvis + Plotly)
- [ ] Semantic search (embeddings)
- [ ] Workbook generation (tools 31–38)
- [ ] **Milestone v0.3:** Beta — generation tools functional, visualization working

### Phase 4: Polish + Integration (Weeks 7–8)
- [ ] Incremental updates + watch mode
- [ ] Multi-file support
- [ ] Skills for AI platforms
- [ ] Tests (500+), benchmarks, docs
- [ ] Token efficiency benchmark validation (§10.2)
- [ ] **Milestone v1.0:** Public release — all 38 tools, >80% test coverage, README + docs complete

### Test Plan

| Layer | Count | Tooling | Coverage Target |
|-------|-------|---------|----------------|
| Unit (parser, formula_parser, graph) | ~200 | pytest | >90% |
| Integration (MCP tool round-trips) | ~150 | pytest + FastMCP test client | >80% |
| E2E (real workbook scenarios) | ~100 | pytest fixtures with 10+ reference workbooks | Key workflows |
| Performance benchmarks | ~50 | pytest-benchmark | Build/query targets in §7.3 |

Reference workbooks will include: simple budget tracker, 50-sheet financial model, workbook with circular references, workbook with named ranges and structured tables, `.xlsm` with VBA, `.xlsb` binary format.

---

## 15. Success Metrics

| Metric | Target | Measurement |
|--------|--------|-------------|
| Tools implemented | 38 | `mcp.list_tools()` |
| Test coverage | >80% | `pytest --cov` |
| Build time (100 sheets) | <30s | Benchmark |
| Token reduction | 10-100x | Eval suite (validated in Phase 4) |
| Accuracy (formula refs) | >99% | Test against known workbooks |
| GitHub stars (6 months) | 1000+ | GitHub API |

---

## 16. Risk Assessment

| Risk | Probability | Impact | Mitigation |
|------|------------|--------|------------|
| Formula parsing misses edge cases | Medium | High | Extensive test suite with real-world workbooks |
| Large workbook performance | Medium | Medium | Range node optimization, lazy loading, streaming |
| Cross-sheet reference accuracy | Low | High | AST-based parsing (not regex), named range resolution |
| MCP server compatibility | Low | Medium | Use FastMCP, test across Claude Code / OpenCode / Cursor |
| Windows file path issues | Medium | Low | Use `pathlib`, test on Windows early |
| LLM generation produces invalid JSON | Medium | Low | Pydantic validation + one retry + structured error fallback |
| Watch mode triggers during large build | Medium | Medium | 5s debounce + build lock prevents concurrent runs |

---

## 17. Open Questions & Decisions

| # | Question | Decision | Owner | Due |
|---|----------|----------|-------|-----|
| 1 | Should we support Google Sheets? | **Out of scope for v1.** Google Sheets API exists and can be a v2 feature. | PM | — |
| 2 | Should we evaluate formulas? | **Optional.** `formulas` library can evaluate; expose as opt-in `evaluate=True` parameter on relevant tools. Default off. | Eng lead | Phase 3 |
| 3 | Should we support LibreOffice .ods? | **Lower priority.** Possible via `odfpy`. Schedule for v1.1 if there is demand. | PM | Post-v1 |
| 4 | Multi-workbook dependency tracking? | **v2 feature.** Flag external refs as AMBIGUOUS in v1 with a `source_file` property for future resolution. | Eng lead | v2 |
| 5 | Real-time Excel COM integration? | **Out of scope.** Windows-only, complex, breaks cross-platform story. Revisit if enterprise demand emerges. | PM | Post-v1 |

---

## 18. References

- **formulas (AST parser)**: https://github.com/vinci1it2000/formulas
- **HyperFormula dependency graph**: https://hyperformula.handsontable.com/docs/guide/dependency-graph.html
- **Excel Dependency Graph (existing)**: https://github.com/jiteshgurav/formula-dependency-excel
- **ExcelFormulaParser**: https://github.com/Voltaic314/ExcelFormulaParser
- **Pyvis visualization**: https://pyvis.readthedocs.io/
- **Plotly Sankey**: https://plotly.com/python/sankey-diagram/
- **OpenPyXL performance**: https://openpyxl.readthedocs.io/en/3.1/performance.html

---

*This PRD is a living document. Update as the project evolves. Increment the version number and add a changelog entry for every meaningful change.*