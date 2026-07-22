# Contributing

Contributions are welcome, especially reproducible samples that expose a general reconstruction problem.

## Before opening a pull request

1. Keep every source presentation unchanged; tests must write to a separate output directory.
2. Do not commit private or proprietary decks, reference images, fonts, or rendered slides.
3. Keep Windows PowerShell scripts ASCII-only unless a BOM-safe encoding strategy is added and tested under Windows PowerShell 5.1.
4. Add deterministic stage logging and structured JSON output to new automation scripts.
5. Preserve protected-picture behavior.
6. Keep MCP tool arguments explicit and pass child-process arguments as arrays; never interpolate user paths into a shell command.
7. Run:

```bash
python tools/validate_repository.py
node tests/test_mcp_server.mjs
```

When working in a Codex development environment, also run the plugin and skill validators documented in the root README.

## Useful issue evidence

- Windows and PowerPoint versions and architecture
- The failing command without copied shell prompts
- Structured JSON output
- A redacted reference/render pair when licensing permits
- Whether the source deck contains groups, freeforms, SVG conversions, or protected pictures

## Pull request scope

Prefer one focused capability or bug fix per pull request. Explain how it affects visual fidelity, editability, and source-file safety separately.
