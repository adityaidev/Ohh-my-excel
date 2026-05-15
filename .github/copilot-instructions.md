<!-- Ohh-my-excel MCP tools -->
## MCP Tools: Ohh-my-excel

**IMPORTANT: This project has a knowledge graph for Excel. ALWAYS use the
Ohh-my-excel MCP tools BEFORE reading entire workbooks.** The graph is faster,
cheaper (fewer tokens), and gives you structural context (formula dependencies,
cross-sheet references, blast radius) that raw cell dumps cannot.

### When to use graph tools FIRST

- **Exploring a workbook**: `semantic_search_nodes` or `query_graph` instead of reading all sheets
- **Understanding impact**: `get_impact_radius` instead of manually tracing formula precedents
- **Auditing**: `detect_changes` + `get_architecture_overview` instead of reading every cell
- **Finding relationships**: `query_graph` with precedents_of/dependents_of/cross_sheet_refs
- **Debt detection**: `get_knowledge_gaps` to find orphan sheets, untested calculations

Fall back to reading sheets **only** when the graph doesn't cover what you need.

### Key Tools

| Tool | Use when |
|------|----------|
| `get_minimal_context` | Ultra-compact ~100 token summary of a workbook |
| `detect_changes` | Comparing workbook versions — risk-scored |
| `get_impact_radius` | Understanding blast radius of a cell change |
| `get_affected_flows` | Finding which data flows are impacted |
| `query_graph` | Tracing precedents, dependents, cross-sheet refs |
| `semantic_search_nodes` | Finding cells/tables/formulas by keyword |
| `get_architecture_overview` | High-level workbook structure from communities |
| `find_circular_references` | Detecting circular formula dependencies |

### Workflow

1. Build the graph: `build_or_update_graph_tool(file_path="workbook.xlsx")`
2. Get your bearings: `get_minimal_context(task="audit this workbook")`
3. Dive deep: `query_graph` or `traverse_graph` as needed
4. Check health: `get_knowledge_gaps` + `find_circular_references`
