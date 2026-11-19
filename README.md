# Office MCP

An MCP server that lets AI agents work in the Microsoft Office apps open on
your Windows PC. It edits the document you have open right now, not a copy of
the file. You watch the changes appear, and each change the agent makes is a
single Ctrl+Z.

Word is supported today. Excel and PowerPoint are next and will join the same
server.

**What an agent can do in Word**

- Read structure: outline, paragraphs, tables, comments, tracked changes, styles.
- Edit with about 120 operations: text, styles, lists, tables, images, charts,
  equations, footnotes, fields, tables of contents, headers and footers, comments.
- Batch many edits into one call. If any edit fails, the whole batch is undone.
- Render pages as images to check the layout the way you see it.
- Save, save as, or export to PDF, but only into folders you allow.

## Requirements

- Windows with the desktop version of Microsoft Word installed.
- [uv](https://docs.astral.sh/uv/getting-started/installation/), which installs
  Python and the server for you.
- Optional: Microsoft Excel, used by Word to insert charts.

The server must run as a Windows process. If your agent runs in WSL, start
the Windows `uvx.exe` from its MCP config.

## Getting started

Add the server to your MCP client. The standard config is:

```json
{
  "mcpServers": {
    "office": {
      "command": "uvx",
      "args": ["office-mcp", "--save-dir", "C:\\AI output"]
    }
  }
}
```

Without `--save-dir` the agent can read and edit open documents, but it cannot
write any file.

<details>
<summary>Claude Code</summary>

```sh
claude mcp add office --scope user -- uvx office-mcp --save-dir "C:\AI output"
```

</details>

<details>
<summary>Claude Desktop</summary>

Open Settings, then Developer, then Edit Config, and add the standard config
above to `claude_desktop_config.json`. Restart Claude Desktop.

</details>

<details>
<summary>Cursor</summary>

Add the standard config to `%USERPROFILE%\.cursor\mcp.json`, or to
`.cursor\mcp.json` in a project.

</details>

<details>
<summary>VS Code (GitHub Copilot)</summary>

Add to `.vscode/mcp.json`:

```json
{
  "servers": {
    "office": {
      "command": "uvx",
      "args": ["office-mcp"]
    }
  }
}
```

</details>

<details>
<summary>Codex CLI</summary>

Add to `%USERPROFILE%\.codex\config.toml`:

```toml
[mcp_servers.office]
command = "uvx"
args = ["office-mcp"]
startup_timeout_ms = 20_000
```

</details>

<details>
<summary>From a source checkout</summary>

```json
{
  "mcpServers": {
    "office": {
      "command": "uv",
      "args": ["run", "--directory", "C:\\path\\to\\office-mcp", "office-mcp"]
    }
  }
}
```

</details>

Then open a document in Word and ask your agent, for example: *"Read the
outline of my open document and tighten the introduction."*

## Tools

Each app has four tools. The many operations sit behind them, so the agent's
tool list stays short.

| Tool | Does | Changes things |
| --- | --- | --- |
| `word_read` | Every read: status, outline, text, tables, find, comments, the guide, ... | no |
| `word_exec` | Every edit, as a batch of ops applied as one undo step | yes |
| `word_snapshot` | Renders pages as PNG images | no |
| `word_write` | Open, save, save as, export PDF, Track Changes on or off | yes |

Agents learn the ops from the built-in guide: `word_read(command="guide")`,
also served as the `office-mcp://word/guide` resource. All parameters are in
the [tool reference](docs/tool-reference.md).

## Configuration

Every flag has an environment variable. A flag wins over its variable.
Directory lists in variables are separated by `;`.

| Flag | Variable | Default | Meaning |
| --- | --- | --- | --- |
| `--apps LIST` | `OFFICE_MCP_APPS` | `word` | Apps to expose, comma-separated |
| `--read-only` | `OFFICE_MCP_READ_ONLY=1` | off | Only the tools that never change a document or file |
| `--save-dir DIR` | `OFFICE_MCP_SAVE_DIRS` | none | Folders saves and exports may write into; repeatable. None means saving is off |
| `--image-dir DIR` | `OFFICE_MCP_IMAGE_DIRS` | any local path | Folders images may be inserted from; repeatable |
| `--max-output-chars N` | | 60000 | Longest text one tool call returns before it is truncated |
| `--snapshot-max-dim PX` | | 1200 | Default cap on a rendered page's long edge |
| `--snapshot-max-pages N` | | 10 | Most pages one snapshot call returns |
| `--log-level LEVEL` | | `WARNING` | Log level for messages on stderr |
| | `WORDCOM_CALL_TIMEOUT` | 300 | Seconds to wait for Word before a call fails as busy; `0` waits forever |

## Safety

- **Your documents stay yours.** The server edits the open document in place
  and never saves on its own. Every call is one undo step, so you can step back
  through what the agent did with Ctrl+Z.
- **No file writes unless you allow a folder.** Saves and PDF exports outside
  `--save-dir` folders are refused.
- **Local files only.** Network paths (`\\server\share`) and URLs are refused for
  documents and images, so a document cannot make the server contact another
  machine.
- **Read-only mode.** `--read-only` removes every tool that changes something.
- Tools carry MCP annotations (`readOnlyHint`, `destructiveHint`) so clients
  can ask you before running edits.

## Troubleshooting

- **`word_not_running`**: start Word and open the document, or let the agent
  open a file with `word_write(command="open", path=...)`.
- **`word_busy`**: a dialog is open in Word (Save As, a warning). Close it and
  the agent can retry.
- **`path_not_allowed`** when saving: start the server with `--save-dir`.
- **Nothing happens in WSL**: Office automation needs the Windows build of the
  server; see Requirements.

## Development

```sh
uv sync
uv run pytest                     # mocked Word, runs anywhere on Windows
uv run pytest -m smoke            # against a real running Word
uv run python scripts/gen_tool_reference.py   # after changing a tool
```

See [docs/design-principles.md](docs/design-principles.md) for how tools are
shaped and how to add an app.
