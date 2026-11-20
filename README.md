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

The server is a Windows program. Agents that run in WSL start it through
Windows; see [Agents running in WSL](#agents-running-in-wsl).

## Getting started

Most MCP clients take this standard config:

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

The first start downloads the server and its dependencies, which can take
15 to 30 seconds and trip a client's startup timeout. Run
`uvx office-mcp --version` once in a terminal first, and later starts take
about 2 seconds.

Then open a document in Word and ask your agent, for example: *"Read the
outline of my open document and tighten the introduction."*

### Add it to your agent

Some Word calls (exporting a long PDF, rendering many pages) take more than a
minute. Where a client cuts tool calls off sooner, the entries below raise its
limit to 300 seconds, the same time the server waits for a busy Word.

<details>
<summary>Antigravity</summary>

In the Antigravity app open Customizations, then Installed, then MCP Servers,
and choose Add MCP Server. In the IDE open the "..." menu in the agent panel,
then MCP Servers, Manage MCP Servers, View raw config. Add the standard config
to `mcp_config.json` (global: `~/.gemini/config/mcp_config.json`; a server's
menu has Copy Config File Path), with a longer tool timeout:

```json
{
  "mcpServers": {
    "office": {
      "command": "uvx",
      "args": ["office-mcp", "--save-dir", "C:\\AI output"],
      "timeoutSeconds": 300
    }
  }
}
```

Antigravity saves large images from tools to a file and gives the agent the
file, so `word_snapshot` pages reach the model that way.

</details>

<details>
<summary>Claude Code</summary>

```sh
claude mcp add office --scope user -- uvx office-mcp --save-dir "C:\AI output"
```

No timeout changes are needed. Calls that run longer than two minutes move to
the background on their own.

</details>

<details>
<summary>Claude Desktop</summary>

Open Settings, then Developer, then Edit Config, and add the standard config
to `claude_desktop_config.json`. Restart Claude Desktop.

</details>

<details>
<summary>Cline</summary>

From the Cline CLI:

```sh
cline mcp add office --yes -- uvx office-mcp --save-dir "C:\AI output"
```

Or in the extension open the MCP Servers view, choose Configure MCP Servers,
and add the standard config with a `timeout`:

```json
{
  "mcpServers": {
    "office": {
      "command": "uvx",
      "args": ["office-mcp", "--save-dir", "C:\\AI output"],
      "timeout": 300
    }
  }
}
```

`timeout` is in seconds and covers startup and every tool call. Without it,
Cline stops tool calls after 60 seconds and startup after 10.

</details>

<details>
<summary>Codex (CLI, IDE extension, app)</summary>

```sh
codex mcp add office -- uvx office-mcp --save-dir "C:\AI output"
```

Then raise the timeouts in `~/.codex/config.toml`
(`%USERPROFILE%\.codex\config.toml` on Windows). The CLI, the IDE extension
and the ChatGPT desktop app share this file.

```toml
[mcp_servers.office]
command = "uvx"
args = ["office-mcp", "--save-dir", 'C:\AI output']
startup_timeout_sec = 60
tool_timeout_sec = 300
```

Codex passes only a few environment variables to servers, so use flags rather
than `OFFICE_MCP_*` variables, or list them under `env`.

</details>

<details>
<summary>Cursor</summary>

Add the standard config to `%USERPROFILE%\.cursor\mcp.json`, or to
`.cursor\mcp.json` in a project.

</details>

<details>
<summary>Hermes Agent</summary>

```sh
hermes mcp add office --command uvx --args office-mcp --save-dir "C:\AI output"
```

`--args` takes everything after it, so keep it last. Or edit the config
(`~/.hermes/config.yaml`; on native Windows it is under `%LOCALAPPDATA%\hermes`):

```yaml
mcp_servers:
  office:
    command: uvx
    args: [office-mcp, --save-dir, 'C:\AI output']
```

The default tool timeout, 300 seconds, is enough. In WSL, see
[Agents running in WSL](#agents-running-in-wsl).

</details>

<details>
<summary>OpenClaw</summary>

```sh
openclaw mcp add office --command uvx --arg office-mcp --arg --save-dir --arg "C:\AI output"
```

Then set a longer request timeout (the default is 60 seconds) in
`~/.openclaw/openclaw.json`:

```json5
{
  mcp: {
    servers: {
      office: {
        command: "uvx",
        args: ["office-mcp", "--save-dir", "C:\\AI output"],
        requestTimeoutMs: 300000,
      },
    },
  },
}
```

Office automation only reaches apps in your signed-in desktop session. If the
OpenClaw Gateway runs as a background service, the server cannot see Word, so
run the Gateway in your own session.

</details>

<details>
<summary>OpenCode</summary>

```sh
opencode mcp add office -- uvx office-mcp --save-dir "C:\AI output"
```

Or add to `opencode.json` (global: `~/.config/opencode/opencode.json`):

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "office": {
      "type": "local",
      "command": ["uvx", "office-mcp", "--save-dir", "C:\\AI output"],
      "enabled": true,
      "timeout": 300000
    }
  }
}
```

`timeout` is in milliseconds. Leave it out and tool calls may be stopped after
30 seconds.

</details>

<details>
<summary>VS Code (GitHub Copilot)</summary>

Add to `.vscode/mcp.json`:

```json
{
  "servers": {
    "office": {
      "command": "uvx",
      "args": ["office-mcp", "--save-dir", "C:\\AI output"]
    }
  }
}
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

Other clients: use the standard config wherever the client adds a local
(stdio) MCP server.

### Agents running in WSL

The server must run on Windows, but an agent inside WSL can start it there.
Use the Windows `uvx.exe` as the command, or go through `cmd.exe`:

```sh
claude mcp add office --scope user -- uvx.exe office-mcp --save-dir "C:\AI output"
hermes mcp add office --command cmd.exe --args /c uvx office-mcp --save-dir "C:\AI output"
```

- `uvx.exe` is found when WSL adds the Windows PATH, which it does by default.
  Otherwise give its full path as WSL sees it (by default uv installs it in
  `%USERPROFILE%\.local\bin`).
- `cmd.exe` prints a "UNC paths are not supported" warning when the agent
  starts from a Linux folder. The warning is harmless; starting the agent from a
  folder on a Windows drive avoids it.
- Linux environment variables do not reach the Windows process, so pass
  settings as flags. Paths are Windows paths (`C:\...`).

### Client support

The server speaks MCP over stdio and works with every protocol version from
2024-11-05 on. Everything an agent needs is in the tool descriptions and the
guide, so a client that skips the server instructions or resources still
works.

| Client | Sees `word_snapshot` images | Server instructions | Timeout to set |
| --- | --- | --- | --- |
| Antigravity | yes, saved to a file first | yes | `timeoutSeconds: 300` |
| Claude Code | yes | yes | none |
| Cline | yes | not documented | `timeout: 300` |
| Codex | yes, with image-capable models | yes | `tool_timeout_sec = 300` |
| Hermes Agent | yes | not documented | none |
| OpenClaw | yes | yes | `requestTimeoutMs: 300000` |
| OpenCode | yes | yes | `timeout: 300000` |

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
the [tool reference](https://github.com/mhuzaifashah/office-mcp/blob/main/docs/tool-reference.md).

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
- **The server fails to start the first time**: the first start downloads the
  server and can outlast the client's startup timeout. Run
  `uvx office-mcp --version` once, then restart the client.
- **A long call fails with a timeout**: raise the client's tool timeout; see
  [Add it to your agent](#add-it-to-your-agent).
- **Nothing happens in WSL**: Office automation needs the Windows build of the
  server; see [Agents running in WSL](#agents-running-in-wsl).

## Development

```sh
uv sync
uv run pytest                     # mocked Word, runs anywhere on Windows
uv run pytest -m smoke            # against a real running Word
uv run python scripts/gen_tool_reference.py   # after changing a tool
```

See [docs/design-principles.md](https://github.com/mhuzaifashah/office-mcp/blob/main/docs/design-principles.md) for how tools are
shaped and how to add an app.

## License

Office MCP is created by Muhammad Huzaifa and licensed under the
[Apache License 2.0](https://github.com/mhuzaifashah/office-mcp/blob/main/LICENSE).
You may use, change and redistribute it, including commercially. If you
redistribute it or a modified version, keep the
[NOTICE](https://github.com/mhuzaifashah/office-mcp/blob/main/NOTICE) file,
which credits the author.
