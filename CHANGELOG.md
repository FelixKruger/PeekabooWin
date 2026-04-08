# Changelog

## 0.1.0 — 2026-04-08

Initial public release.

### Added

- Window and screen capture to PNG
- Structured `see` snapshots with indexed UI Automation controls
- Built-in Windows OCR on captures and snapshots
- Click and scroll by visible label with OCR fallback
- Snapshot-based click, scroll, and drag
- Menu listing and menu clicks
- Standard dialog listing and button clicks
- Window focus, move, resize, maximize, minimize, restore, and wait
- App launch, list, switch, and quit
- Mouse move, click, drag, and scroll at coordinates
- Keyboard type, press, and hotkey commands
- Interaction profiles with human-paced mode
- Scroll-and-harvest OCR pipeline for long documents and threads
- Three-tier window capture with black frame detection and automatic fallback
- Harvest auto-fallback from window to screen capture when OCR returns empty
- Fuzzy overlap matching for harvest deduplication across OCR-variable passes
- Higher-level recipes and plain-language goal planning
- Workflow runner for JSON-defined automation sequences
- AI-ready snapshot summaries
- CLI with full command surface
- MCP server over local stdio for Claude Code and other AI tools
- Single-file desktop UI (`PeekabooWin.hta`)
- Project `.mcp.json` for Claude Code auto-discovery
