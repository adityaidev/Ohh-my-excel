# Ohh-my-excel

**Persistent, incrementally-updated knowledge graph for Excel workbooks.**

Safe · Fast · Token-Efficient

A Model Context Protocol (MCP) server that parses Excel workbooks into a structural
knowledge graph of sheets, tables, cells, formulas, and dependencies — then exposes
it via MCP tools so AI assistants get precise, token-efficient context instead of
dumping entire spreadsheets.

## Why

Existing Excel MCP servers are CRUD-only — they read/write cells and dump raw data.
None provide structural understanding of formula dependencies, blast radius analysis,
or token-efficient context. When an AI assistant opens a 50-sheet financial model,
it reads everything, burning tokens and missing critical relationships.

## Three Pillars

- **Safe** — Automatic backups before writes, formula validation, circular ref detection
- **Fast** — Sub-second graph builds, streaming responses, incremental updates
- **Token-Efficient** — ~100 token workbook summaries, ~800 token full analysis sessions

## Status

Pre-alpha. PRD complete. Phase 1 (Core Parser + Graph) in development.
