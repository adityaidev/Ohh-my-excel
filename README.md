<h1 align="center">Ohh-my-excel</h1>

<p align="center">
  <strong>Safe · Fast · Token-Efficient — Knowledge Graph for Excel Workbooks</strong>
</p>

<p align="center">
  <a href="https://pypi.org/project/ohh-my-excel/"><img src="https://img.shields.io/pypi/v/ohh-my-excel?style=flat-square&color=blue" alt="PyPI"></a>
  <a href="https://github.com/adityaidev/Ohh-my-excel/stargazers"><img src="https://img.shields.io/github/stars/adityaidev/Ohh-my-excel?style=flat-square" alt="Stars"></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square" alt="MIT Licence"></a>
  <a href="https://github.com/adityaidev/Ohh-my-excel/actions/workflows/ci.yml"><img src="https://github.com/adityaidev/Ohh-my-excel/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/python-3.10%2B-blue.svg?style=flat-square" alt="Python 3.10+"></a>
  <a href="https://modelcontextprotocol.io/"><img src="https://img.shields.io/badge/MCP-compatible-green.svg?style=flat-square" alt="MCP"></a>
</p>

<br>

Existing Excel MCP servers are CRUD-only — they read/write cells and dump raw data.
None provide structural understanding of formula dependencies, blast radius analysis,
or token-efficient context. When an AI assistant opens a 50-sheet financial model,
it reads everything, burning tokens and missing critical relationships.

**Ohh-my-excel** fixes that. It parses Excel workbooks into a structural knowledge
graph of sheets, tables, cells, formulas, and dependencies — then exposes it via
[MCP](https://modelcontextprotocol.io/) tools so AI assistants get precise,
token-efficient context instead of dumping entire spreadsheets.

---

## Quick Start

```bash
pip install ohh-my-excel
ohh-my-excel build path/to/workbook.xlsx
```

Then open your AI assistant and ask:
```
Build the Excel knowledge graph for this workbook
```

The initial build takes ~10 seconds for a 100-sheet workbook. After that,
incremental updates complete in under 500ms.

```bash
ohh-my-excel build path/to/workbook.xlsx    # Build or rebuild graph
ohh-my-excel update path/to/workbook.xlsx   # Incremental update
ohh-my-excel status path/to/workbook.xlsx   # Show graph stats
ohh-my-excel serve                          # Start MCP server
```

Requires Python 3.10+.

---

## How It Works

```
┌─────────────────────────────────────────────────────────┐
│                    MCP Client (AI Agent)                  │
│         OpenCode / Claude Code / Codex / Cursor          │
└────────────────────┬────────────────────────────────────┘
                     │  stdio / Streamable HTTP
                     ▼
┌─────────────────────────────────────────────────────────┐
│              FastMCP Server (38 tools + 5 prompts)       │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│              SQLite Graph Store (WAL mode)               │
│         nodes + edges + FTS5 + NetworkX cache           │
│         .excel-graph/graph.db                           │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│              Excel Parser + Formula AST                  │
│         openpyxl + formulas + cell ref extraction       │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│              Excel Workbook (.xlsx / .xlsm)              │
└─────────────────────────────────────────────────────────┘
```

### Blast-radius analysis

When a cell changes, the graph traces every dependent formula, cross-sheet
reference, and data flow that could be affected. Your AI reads only these
instead of scanning the whole workbook.

### Incremental updates in < 500ms

On every file save, a SHA-256 hash comparison detects changes. Only changed
sheets are re-parsed. A 100-sheet workbook re-indexes relevant cells in
under 500ms.

### Token efficiency: 10-250x reduction

| Scenario | Without Graph | With Graph | Reduction |
|----------|--------------|------------|-----------|
| Audit 50-sheet workbook | ~500k tokens | ~50k tokens | ~10x |
| Review cell change | ~500k tokens | ~5k tokens | ~100x |
| Find formula dependencies | ~500k tokens | ~2k tokens | ~250x |
| Understand data flow | ~500k tokens | ~10k tokens | ~50x |

---

## 38 MCP Tools

| Category | Tools | Description |
|----------|-------|-------------|
| **Graph Setup** | 3 | Build, update, post-process, stats |
| **Exploration** | 4 | Minimal context, search, query, traverse |
| **Dependency & Impact** | 3 | Blast radius, change detection, affected flows |
| **Formula Analysis** | 5 | Dependencies, dependents, circular refs, broken refs, complexity |
| **Sheet & Workbook** | 5 | List sheets, sheet info, cross-sheet refs, tables, named ranges |
| **Data Flow** | 2 | List flows, get flow details |
| **Architecture & Quality** | 6 | Communities, hub cells, bridge cells, knowledge gaps, questions, architecture overview |
| **Visualization & Export** | 2 | Interactive HTML graph, JSON/CSV/GraphML export |
| **Workbook Generation** | 8 | Create from prompt, from data, from template, add sheet, add chart, format, formulas, validate |

### Key Tools

| Tool | What it does |
|------|-------------|
| `build_or_update_graph_tool` | Build or incrementally update the Excel knowledge graph |
| `get_minimal_context_tool` | Ultra-compact ~100 token workbook summary — call this first |
| `get_impact_radius_tool` | Blast radius when a cell/formula changes |
| `query_graph_tool` | Precedents, dependents, cross-sheet refs, contains |
| `traverse_graph_tool` | BFS/DFS from any cell/sheet with token budget |
| `find_circular_references_tool` | Detect circular formula dependencies |
| `detect_changes_tool` | Risk-scored change analysis between versions |
| `get_architecture_overview_tool` | High-level workbook structure from sheet communities |
| `generate_workbook_tool` | Create workbook from natural language prompt |
| `visualize_graph_tool` | Generate interactive HTML dependency graph |

### MCP Prompts (5)

`audit_workbook`, `debug_formula`, `data_flow_map`, `onboard_analyst`, `pre_merge_check`

---

## Three Pillars

### Safe

- **Automatic backups** before every destructive write (`.excel-graph/backups/`)
- **Formula validation** via AST parsing before writing
- **Circular reference detection** via NetworkX before saving
- **No eval/exec** — zero arbitrary code execution
- **Path traversal prevention** — `_validate_file_path()` sanitizes all input
- **Max file size limits** — configurable (default: 100MB read, 50MB generate)
- **Backup retention** — 7 days by default, max 10 per workbook

### Fast

- **SQLite WAL mode** — concurrent reads, zero lock contention
- **SHA-256 hash skip** — only re-parse changed sheets
- **Range node optimization** — `SUM(A1:A100)` = 1 edge, not 100
- **Lazy cell loading** — values loaded on query, not during build
- **Batch cell reads** — entire sheet ranges in one call
- **Cached NetworkX graph** — O(1) query response after build
- **Async thread offload** — no event loop blocking

### Token-Efficient

- **`get_minimal_context_tool`** — ~100 tokens: stats + next_tool_suggestions
- **`detail_level="minimal"`** — all tools support compact output
- **Structural context over raw data** — "Cell C5 depends on A1:A100 via SUM" instead of 100 cell values
- **`next_tool_suggestions`** — every response tells AI the optimal next tool
- **Target: ≤5 tool calls per task, ≤800 total tokens**

---

## Architecture

```
excel_graph_mcp/
├── __init__.py
├── constants.py       # Config, schema, performance targets, env vars
├── formula_parser.py  # Formula AST → cell/range refs, functions, nesting depth
├── parser.py          # openpyxl multi-sheet parser (cells, formulas, tables, charts)
├── graph.py           # SQLite WAL graph store (nodes, edges, FTS5, BFS, NetworkX cache)
├── dependency.py      # Dependency graph builder, circular/broken ref detection
├── flows.py           # Data flow detection (input → calculation → output chains)
├── communities.py     # Sheet community detection + architecture overview
├── changes.py         # Risk-scored version-to-version change analysis
├── analysis.py        # Cross-sheet refs, hub/bridge cells, knowledge gaps, suggested questions
├── exports.py         # JSON/CSV/GraphML export + Pyvis/Matplotlib visualization
├── incremental.py     # SHA-256 hash change detection, auto-update
├── generation.py      # 8 workbook creation/modification/validation tools
├── main.py            # FastMCP server — 38 tools + 5 prompts
├── cli.py             # CLI: build, update, status, serve, version
└── tools/             # Tool implementations (build, query, context)
```

### Graph Schema

**Node types:** Workbook, Sheet, Cell, Formula, Range, Table, NamedRange, Chart

**Edge types:** CONTAINS, REFERENCES, RANGE_REF, DEPENDS_ON, CROSS_SHEET_REF, TABLE_REF, CHART_SOURCE, NAMED_REF

**Confidence levels:** EXTRACTED (AST), INFERRED (range expansion), AMBIGUOUS (INDIRECT/OFFSET)

---

## Configuration

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `EXCEL_GRAPH_BACKUP_RETENTION_DAYS` | Days to keep backups | `7` |
| `EXCEL_GRAPH_MAX_BACKUPS` | Max backups per workbook | `10` |
| `EXCEL_GRAPH_WATCH_DEBOUNCE_MS` | Watch mode debounce | `5000` |
| `EXCEL_GRAPH_RATE_LIMIT` | Generation requests/min | `10` |
| `EXCEL_GRAPH_MAX_FILE_SIZE_MB` | Max read file size | `100` |
| `EXCEL_GRAPH_GENERATE_MAX_SIZE_MB` | Max generate file size | `50` |

### Optional Dependency Groups

```bash
pip install ohh-my-excel[embeddings]     # Vector embeddings (sentence-transformers)
pip install ohh-my-excel[visualization]  # Pyvis + Plotly visualization
pip install ohh-my-excel[dev]            # Development (pytest, ruff, bandit)
```

---

## Supported Formats

| Format | Read | Write | Formula Parsing | Notes |
|--------|------|-------|-----------------|-------|
| `.xlsx` | ✅ | ✅ | ✅ | Primary format |
| `.xlsm` | ✅ | ✅ | ✅ | + VBA macro extraction |
| `.xltx` | ✅ | ✅ | ✅ | Templates |
| `.xltm` | ✅ | ✅ | ✅ | Macro-enabled templates |
| `.xlsb` | ✅ | ❌ | ⚠️ | Binary format (via pyxlsb) |
| `.xls` | ⚠️ | ❌ | ❌ | Legacy (via xlrd) |
| `.csv` | ✅ | ✅ | ❌ | No formulas |

---

## Performance Targets

| Metric | Target |
|--------|--------|
| Initial build (100-sheet workbook) | <30s |
| Incremental update (1 cell change) | <500ms |
| Query response (any tool) | <100ms P95 |
| Search response | <10ms |
| Memory usage (100-sheet workbook) | <200MB |
| Graph DB size (100-sheet workbook) | <50MB |

---

## Development

```bash
git clone https://github.com/adityaidev/Ohh-my-excel.git
cd Ohh-my-excel
pip install -e ".[dev]"
pytest
```

### Running CI Checks Locally

```bash
pytest tests/ --tb=short -q                    # Run tests
ruff check excel_graph_mcp/                     # Lint
bandit -r excel_graph_mcp/ -c pyproject.toml    # Security scan
```

---

## What's Next

| Feature | Status |
|---------|--------|
| **HyperFormula sidecar** — dual formula engine (JS + Python) | Planned |
| **VBA analysis tools** — extract, summarize, explain VBA macros | Planned |
| **Template expansion** — 20+ built-in templates across 6 categories | Planned |
| **Embedding-based semantic search** — vector search via sentence-transformers | Planned |
| **Power Query / DAX roadmap** — financial analyst integration | Future |
| **Excel Online / Graph API** — read from SharePoint/OneDrive | Future |
| **Claude in Excel differentiation** — MCP-native, any-AI-client support | Ongoing |

---

## Licence

MIT. See [LICENSE](LICENSE).

<p align="center">
<br>
<a href="https://github.com/adityaidev/Ohh-my-excel">github.com/adityaidev/Ohh-my-excel</a><br><br>
<code>pip install ohh-my-excel && ohh-my-excel build workbook.xlsx</code><br>
<sub>38 MCP Tools · 5 Prompts · 62 Tests · Safe · Fast · Token-Efficient</sub>
</p>
