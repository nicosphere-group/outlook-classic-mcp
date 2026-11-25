# outlook-classic-mcp

[日本語](README.md) | English

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets Claude read, search, draft and send email through **classic Outlook for Windows**. It talks to the Outlook already installed on your PC, so any account set up there (Exchange, IMAP, POP3) works as it is.

Sending always opens a confirmation window on your screen. Nothing is sent unless you press **Send** yourself.

[![Download](https://img.shields.io/badge/Download-outlook--classic--mcp.mcpb-2563eb?style=for-the-badge)](https://github.com/nicosphere-group/outlook-classic-mcp/releases/latest/download/outlook-classic-mcp.mcpb)

1. Download `outlook-classic-mcp.mcpb` with the button above. Do not unzip it.
2. In Claude Desktop, open **Settings** → **Extensions** → **Advanced settings** → **Install Extension** and choose the file.
3. Click **Install**. Start classic Outlook, then ask Claude "What came in today?"

## Features

- **Works with your existing Outlook.** No mail passwords, no server settings, no cloud sign-in.
- **Nothing extra to install.** A Claude Desktop extension with no executable of its own: it runs on the Node.js built into Claude and on Windows PowerShell, which is part of Windows.
- **Sending needs your click.** Claude cannot press the buttons in the confirmation window.
- **Stays on your PC.** The server makes no network connections and collects no usage data.

## Tools

| Tool | Description |
|---|---|
| `list_folders` | List mail folders with total and unread counts |
| `list_accounts` | List accounts and data files |
| `list_emails` | List emails in a folder, filtered by date, unread state and count |
| `search_emails` | Search by subject, sender, body and date |
| `read_email` | Read the body, recipients and attachment names of an email |
| `list_attachments` | List the attachments of an email |
| `save_attachment` | Save an attachment into the allowed folder |
| `create_draft` | Create a draft of a new email |
| `create_reply_draft` | Create a draft reply or reply-all |
| `create_forward_draft` | Create a draft forward |
| `update_draft` | Change an existing draft |
| `send_draft` | Send an existing draft after you confirm it on screen |

Not planned for the first release: deleting or moving email, and the calendar.

## How sending works

Only an existing draft can be sent. There is no tool that sends a new email directly.

1. You ask Claude to send a draft.
2. A confirmation window opens on your screen. It shows the sender, To, CC, BCC, subject, attachments and the full body.
3. You choose one of three buttons.

| Button | What happens |
|---|---|
| **Send** | The email is sent. |
| **Open in Outlook** | The draft opens in Outlook. Nothing is sent. |
| **Don't send** | Nothing is sent. The draft stays where it is. |

- Closing the window, or leaving it unanswered for 3 minutes, sends nothing.
- A warning is shown at the top when there are many recipients, when it is a reply-all, when a recipient is outside your organization, or when there are attachments.
- The window appears every time. There is no "don't ask again" option and no setting that skips it.
- If the draft changes while the window is open, it is not sent.
- Each send, and each time you decline, is recorded in `%LOCALAPPDATA%\outlook-classic-mcp\audit.log` with the date, sender account, recipients and subject. The body is not recorded.

## Requirements

- Windows 10 or 11
- Classic Outlook 2016 or later, with at least one mail account set up
- Claude Desktop for Windows, or Claude Code

The new Outlook app and Outlook on the web are not supported. Claude in a browser and the Claude mobile apps cannot use this server, because it has to run on the same PC as Outlook.

On Team and Enterprise plans, an administrator may need to allow the extension first.

## Installation

### Claude Desktop

1. Download [`outlook-classic-mcp.mcpb`](https://github.com/nicosphere-group/outlook-classic-mcp/releases/latest/download/outlook-classic-mcp.mcpb). This is always the latest version; older versions are on the [releases page](https://github.com/nicosphere-group/outlook-classic-mcp/releases).
2. In Claude Desktop, open **Settings** → **Extensions** → **Advanced settings** → **Install Extension** and choose the file.

### Claude Code

Build the program from source, then register it:

```
claude mcp add outlook-classic -- node C:\path\to\extension\server\index.js
```

## Usage

Start classic Outlook, then ask Claude in plain words:

- "What came in today?"
- "Find the email from Tanaka about the estimate."
- "Draft a reply saying next Wednesday works."
- "Make it a little more polite."
- "Send it." A confirmation window opens.

## Configuration

No configuration is needed. The confirmation window is always shown, whatever the settings are.

In Claude Desktop, the extension's settings screen lets you switch sending on or off and choose the folder for saved attachments. Elsewhere (for example Claude Code), use these environment variables.

| Environment variable | Default | Description |
|---|---|---|
| `OUTLOOK_MCP_ALLOW_SEND` | `1` | Set to `0` to remove the sending tool. Reading and drafting still work. |
| `OUTLOOK_MCP_SAVE_DIR` | `Downloads\outlook-classic-mcp` | Folder where attachments are saved. |
| `OUTLOOK_MCP_ATTACH_DIRS` | Documents, Desktop, Downloads | Folders that files can be attached from, separated by `;`. |
| `OUTLOOK_MCP_CONFIRM_TIMEOUT` | `180` | Seconds before the confirmation window closes by itself (15 to 600). |
| `OUTLOOK_MCP_LANG` | Automatic | Language of the confirmation window: `ja` or `en`. |

## Security and privacy

- The server makes no network connections. It only talks to Outlook on the same PC.
- Email content reaches Claude only when a tool returns it in a conversation.
- Email content is treated as untrusted. The server tells Claude not to follow instructions found inside an email.
- Administrator rights are not needed. Running Outlook or Claude as administrator prevents the connection.
- Do not turn off Outlook's security prompts in the registry to use this server.
- In Claude Code, Claude can also run shell commands, so the confirmation window is not the only route to Outlook there. Keep permission prompts for shell commands turned on.

## Troubleshooting

**"Outlook did not respond"**
Start Outlook normally. The connection fails when Outlook or Claude is running as administrator.

**"Classic Outlook was not found"**
The PC has only the new Outlook app. Classic Outlook is required.

**Outlook warns that a program is trying to access your email**
Outlook shows this when antivirus software is turned off or out of date. Turn it on and update it.

**A sent email does not arrive (POP3 or IMAP)**
The email is placed in Outlook's Outbox and goes out at the next send/receive. Keep Outlook open.

## Building

Requirements: Node.js 18 or later.

```
npm ci
npm test
npm run pack
```

This produces `dist/outlook-classic-mcp.mcpb`. The extension contains no executable of its own: it runs on the Node.js that Claude Desktop ships with, and reaches Outlook through Windows PowerShell, which is part of Windows.

On Windows, `powershell -NoProfile -ExecutionPolicy Bypass -File tests\helper.tests.ps1` checks the PowerShell helper against a stand-in for Outlook.

- [skills/outlook-classic/SKILL.md](skills/outlook-classic/SKILL.md): a skill that teaches Claude how to use the tools safely

## License

[MIT](LICENSE)
