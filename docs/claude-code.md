# Using PeekabooWin With Claude Code

PeekabooWin is built to work well with Claude Code through the local MCP server in the repo root `.mcp.json`.

## What Claude Should Use PeekabooWin For

Use Claude for reasoning.

Use PeekabooWin for:

- seeing what is on the Windows desktop
- clicking, scrolling, typing, and focusing apps
- OCRing visible text
- harvesting long local threads into files Claude can read afterward

The clean split is:

- PeekabooWin manipulates and reads the desktop
- Claude interprets the results and decides what to do next

## Setup

1. Open this repository in Claude Code.
2. Approve the `peekaboo-win` MCP server when Claude asks.
3. Confirm Claude can see the server through `/mcp`.

The repo already contains this config:

```json
{
  "mcpServers": {
    "peekaboo-win": {
      "type": "stdio",
      "command": "node",
      "args": ["./bin/peekaboo-win-mcp.js"]
    }
  }
}
```

## Best First Test

Open a local text file in Notepad and ask Claude:

```text
Use PeekabooWin to harvest the visible Notepad window.
Scroll down up to 10 times, save the transcript to .\exports\claude-harvest.txt,
then read that file and summarize what it contains.
```

Expected flow:

1. Claude calls `harvest_scroll_text`
2. PeekabooWin saves:
   - `.\\exports\\claude-harvest.txt`
   - `.\\exports\\claude-harvest.json`
3. Claude reads the `.txt` file
4. Claude summarizes or extracts what you asked for

By default, harvest output will not overwrite an existing file. Reuse the same path only if you also pass `overwrite`.

## Recommended Patterns

### Pattern 1: Inspect Then Act

Good for buttons, menus, and forms.

Ask Claude to:

1. call `ui_snapshot`
2. inspect controls and OCR text
3. call `element_click`, `snapshot_click`, `type_text`, or `menu_click`

### Pattern 2: Harvest Then Reason

Good for chats, logs, transcripts, and long local documents.

Ask Claude to:

1. call `harvest_scroll_text`
2. save output inside the repo
3. read the saved file
4. summarize, classify, search, or transform it

### Pattern 3: Goal Layer

Good when you want a shorter prompt.

Ask Claude to use:

- `goal_plan` to see what PeekabooWin thinks the task is
- `goal_run` to execute it

Example:

```text
Use PeekabooWin goal_run for: "scrape this thread"
Target the Notepad window and save the output to .\exports\thread.txt.
```

## Good Prompt Examples

### Click A Visible Button

```text
Use PeekabooWin to inspect the Settings window, then click the visible label "Bluetooth".
```

### Read A Screen

```text
Use PeekabooWin to capture the current Notepad window and tell me what text is visible.
```

### Harvest A Long Local Thread

```text
Use PeekabooWin to harvest the visible Discord thread from the target window.
Scroll down up to 40 times, save the full transcript to .\exports\discord-thread.txt,
then read that file and give me a concise summary plus action items.
```

## Tips

- Save harvested output inside the repo when possible so Claude can read it directly afterward.
- Harvest output does not overwrite existing files unless you explicitly allow it.
- If scrolling at the center of the window is not precise enough, pass `scrollLabel` for the pane or list you want.
- Start with lower `maxSteps` while testing, then increase it.
- Use disposable apps first, not your browser or work-critical apps.

## Troubleshooting

### Claude cannot see the MCP server

- Reopen the repo in Claude Code
- Run `/mcp`
- Make sure the project `.mcp.json` is present

### Harvest output is too small

- Increase `maxSteps`
- Make sure the target window is correct
- Use a more specific window title

### Scroll is hitting the wrong area

- Provide a better target window
- Use `scrollLabel`
- Use a smaller, more controlled test app first
