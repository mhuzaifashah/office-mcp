# Design principles

Rough guidelines for how tools are shaped. They follow Anthropic's
[Writing effective tools for agents](https://www.anthropic.com/engineering/writing-tools-for-agents),
the MCP specification, and well-built servers such as chrome-devtools-mcp.

## Few, task-shaped tools

Each app exposes four tools: `<app>_read`, `<app>_exec`, `<app>_snapshot`,
`<app>_write`. The operations behind them (about 120 for Word) are values of a
`command` or an `op`, not separate tools. Every tool definition costs context
before the agent does anything, and many similar tools make it pick the wrong
one more often.

## Names say which app

Tool names start with the app (`word_`, `excel_`, `ppt_`). Parameter names are
unambiguous (`anchor_id`, `table`, `start_anchor`), and closed sets are enums
in the schema, so a typo fails validation instead of doing something else.

## Small outputs

- Results are compact JSON; prose (Markdown, the guide) is plain text.
- Output over `--max-output-chars` is truncated and followed by a
  `[truncated]` note saying what was dropped and how to ask for less. Lists
  keep whole items, so the JSON stays valid.
- Images are capped (`--snapshot-max-dim`, `--snapshot-max-pages`).

## Errors the agent can act on

Failures are tool results with `isError: true`, not protocol errors, so the
model reads them. Each has a stable `code`, a `retryable` flag, and a `hint`
saying what to do next (re-read the outline, add `occurrence`, ask the user
for an allowed folder). A failed batch also reports the failing op's index and
whether it was rolled back.

## Safe by default

- Edits are one undo step; a failed batch rolls back.
- No file is written unless the user allows a folder (`--save-dir`).
- Network paths and URLs are refused.
- `--read-only` registers only tools annotated `readOnlyHint`.
- The server never saves on exit or closes the user's Office apps.

## One server, an app per toolset

`office_mcp/server/` knows nothing about Word. Each app package provides a
toolset (`office_mcp/<app>/tools.py`) with its tool definitions, its paragraph
of the server instructions, and its resources. `--apps` chooses which toolsets
load; an app's COM thread starts on its first tool call.

## Adding an app

1. Write `office_mcp/<app>/tools.py` with a toolset class following
   `office_mcp.server.tool_definition.Toolset`.
2. Register it in `_TOOLSETS` in `office_mcp/server/build.py` and add the app
   to `IMPLEMENTED_APPS` in `office_mcp/server/config.py`.
3. Write the app's guide and serve it from `<app>_read(command="guide")` and an
   `office-mcp://<app>/guide` resource.
4. Add tests with a mocked app, then run `scripts/gen_tool_reference.py`.
