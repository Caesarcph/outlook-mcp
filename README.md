# outlook-mcp

An [MCP](https://modelcontextprotocol.io) server that lets Claude (or any MCP client) read, write, reply to, forward and organize mail in **classic Outlook on Windows**, through Outlook's local COM interface.

**Why:** many school and work Microsoft 365 tenants block third-party apps unless an admin approves them, so the official connectors and the Microsoft Graph API are off the table. You can still sign in to the Outlook desktop app. This server talks to that already-signed-in Outlook, so **it needs no admin consent, no app registration and no tokens**.

> 中文简介：通过 Windows 上**经典版 Outlook** 的本地 COM 接口，让 Claude 读写、回复、转发、整理邮件。适用于学校或公司邮箱需要管理员授权、官方连接器连不上，但能在 Outlook 桌面版登录的情况。

## Features

| Tool | What it does |
|---|---|
| `list_accounts` | Email addresses configured in Outlook |
| `list_folders` | Mail folders, as paths usable by other tools |
| `search_emails` | Recent mail in a folder, newest first, with keyword and unread filters |
| `read_email` | Full body, To / CC / BCC and attachment names |
| `send_email` | New mail with CC, BCC and attachments, from any configured account |
| `reply_email` | Reply or reply-all **inside the original thread**, keeping the quoted history |
| `forward_email` | Forward with the original attachments, plus optional extra recipients and files |
| `send_draft` | Send a draft created earlier |
| `mark_emails` | Read / unread, flag / unflag |
| `create_folder` / `move_emails` | Organize mail |
| `delete_emails` | Move to Deleted Items (always recoverable) |
| `save_attachments` | Save attachments to disk so Claude can read them |

### Safety defaults

- **Drafts first.** Every sending tool saves a draft unless it is called with `send=true`.
- **No permanent deletion.** `delete_emails` only moves mail to Deleted Items, and it refuses mail that is already there.
- **Attach, never launch.** The server connects only to an Outlook you already have open. Launching a hidden Outlook can hang forever on dialogs you can't see.
- Each tool carries MCP annotations (read-only / destructive), so clients can ask before running the risky ones.

## Requirements

- Windows with **classic Outlook** (Microsoft 365 / Office desktop). The web-based "new Outlook" has no COM interface and is **not** supported.
- [uv](https://docs.astral.sh/uv/) (Python 3.10+ is installed for you).
- Outlook must be **open** while you use the server.

## Install

```bash
git clone https://github.com/Caesarcph/outlook-mcp.git
cd outlook-mcp
uv sync
cp config.example.json config.json   # then edit it, see below
```

### Register with Claude Code (all projects)

```bash
claude mcp add -s user outlook -- uv run --directory "C:/path/to/outlook-mcp" server.py
```

### Register with Claude Desktop (Chat)

Add this to `%APPDATA%\Claude\claude_desktop_config.json`, then restart Claude Desktop:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\path\\to\\uv.exe",
      "args": ["run", "--directory", "C:\\path\\to\\outlook-mcp", "server.py"]
    }
  }
}
```

Use the absolute path to `uv.exe` (find it with `where uv`). GUI apps don't always inherit your terminal's `PATH`.

## Configure

`config.json` is gitignored, so your addresses stay on your machine. Every field is optional.

| Key | Purpose |
|---|---|
| `mailboxes` | Labels and folder paths, told to Claude so "check my school email" maps to the right mailbox. The default account's inbox is `inbox`; for other accounts use `address@domain/inbox`. |
| `notes` | Free text passed to Claude, for example "my personal Gmail is handled by the Gmail connector". |
| `keep_domains` | Senders from these domains (and their subdomains) are never filtered out. |
| `block_domains` | Senders from these domains are hidden from `search_emails`. |
| `promo_keywords` | Mail from other senders whose subject contains any of these words is hidden. |

Set `OUTLOOK_MCP_CONFIG` to load the config from another path.

## Usage

Once it's registered, just ask in plain language:

- "Summarize this week's unread mail in my school inbox."
- "Reply to the email from the registrar, CC my advisor, and save it as a draft."
- "Forward the syllabus email to my personal address with the PDF attached."
- "Move all newsletters from last month into a folder called Newsletters."

## Troubleshooting

| Symptom | Fix |
|---|---|
| `Outlook is not running` | Open classic Outlook and retry. |
| Outlook won't start: `0xc0000142` | A hidden Outlook started by an earlier script is stuck. End `OUTLOOK.EXE` in Task Manager, then open Outlook normally. |
| A change to `server.py` or `config.json` has no effect | Each conversation starts its own server process and keeps it until the conversation ends. Start a new conversation. |
| `Folder not found` for "Inbox" | Folder names are localized (for example "收件箱"). Use the lowercase keywords `inbox`, `sent`, `drafts`, `deleted` and `junk`, which work in any language. |
| A sent mail never arrives | Run `list_accounts`. An account with no mailbox loaded is refused by the server instead of silently "sending". |
| Attachment "path does not exist" | The path has to be visible to Outlook itself. Use a normal folder, not a sandboxed or temporary one. |

## License

[MIT](LICENSE)
