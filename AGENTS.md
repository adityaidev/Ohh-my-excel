# Agent Instructions

## Quick Reference

```bash
# Development
pytest tests/ --tb=short -q     # Run tests
ruff check excel_graph_mcp/     # Lint
bandit -r excel_graph_mcp/      # Security scan
```

## Session Completion

When ending a work session:
1. Run quality gates (tests, lint, security)
2. Commit and push all changes
3. Verify push succeeded

## MCP Tools

This project has a knowledge graph for Excel. Use Ohh-my-excel MCP tools
BEFORE reading entire workbooks. The graph is faster, cheaper, and gives
structural context that raw cell dumps cannot.
