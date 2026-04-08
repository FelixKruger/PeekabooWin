# Contributing to PeekabooWin

## Getting Started

```powershell
git clone https://github.com/FelixKruger/PeekabooWin.git
cd PeekabooWin
npm install
npm test
```

## Development

```powershell
node .\bin\peekaboo-win.js --help       # CLI
node .\bin\peekaboo-win-mcp.js          # MCP server
npm run ui                               # Desktop UI
```

## Testing

Run the test suite before submitting changes:

```powershell
npm test
```

Tips:

- Use Notepad or Paint for live automation tests, not your browser or work apps.
- Tests use dependency injection so they run without a live desktop session.
- Keep harvest overwrite protection in place.

## Submitting Changes

1. Fork the repo and create a branch.
2. Make your changes in small, focused commits.
3. Run `npm test` and confirm all tests pass.
4. Open a pull request with a clear description of what changed and why.

## Code Style

- ESM modules, Node.js 22+
- PowerShell backend for native Windows automation
- Clear names over comments
- No external runtime dependencies beyond Node.js and PowerShell
