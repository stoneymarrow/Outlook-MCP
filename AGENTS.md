# Outlook MCP Server

MCP server for Microsoft 365 email + calendar + contacts. Uses Graph API with client credentials flow so auth never expires.

## Architecture
- Pure Graph API transport - no LLM calls inside the server
- Client credentials flow (application permissions) via MSAL - auth never expires
- Single httpx.AsyncClient with 30s timeout, reused across all tool calls
- Attachment bytes stream from Graph's raw endpoint to a temporary local file; only delivery metadata crosses MCP
- Credentials in `.env` (gitignored), registered in `~/.claude/.mcp.json` and Claude Desktop
- Shutdown: `main._run_stdio` exits via `os._exit` after transport + lifespan cleanup (not a plain `mcp.run`). This skips interpreter finalization, which otherwise races the anyio stdio worker thread and self-aborts (`_enter_buffered_busy` SIGABRT). Do not revert it; regression-guarded in `tests/test_shutdown_regression.py`

## Azure AD App
- Application registration in your tenant
- Declared permissions: `Mail.Read`, `Calendars.Read`, `Contacts.Read` (application, admin-consented); see the regression-tested constants in `main.py`
- Client secret rotation: track expiry in your Azure portal

## Tools (21)

### Email (9)
- `read_inbox` - list emails, filter by sender/subject/date
- `read_email` - full content by ID
- `list_folder_messages` - list messages in any folder
- `search_emails` - keyword search
- `send_email` - send with to/cc/body
- `reply_email` - reply or reply-all to a thread
- `forward_email` - forward to new recipients
- `move_email` - move to folder by ID
- `suggest_folders` - group inbox by sender domain for folder planning

### Filing (1)
- `file_emails` - bulk move emails by folder name, auto-creates missing folders

### Attachments (2)
- `list_attachments` - list every paginated attachment with its kind
- `download_attachment` - stream file bytes into a caller-selected directory and return metadata only

### Folders (3)
- `list_folders` - all folders with children
- `create_folder` - top-level or nested
- `delete_folder` - delete an empty folder, or force deletion explicitly

### Calendar (4)
- `list_events` - events in a date range (expands recurring)
- `get_event` - full event details with body and Teams join URL
- `create_event` - create event/meeting with attendees, location, Teams link
- `delete_event` - cancel/delete an event

### Contacts (2)
- `list_contacts` - list Outlook contacts
- `search_contacts` - search by name, email, or company

## Usage patterns
- **Daily email summary:** `read_inbox(since=...)` → `read_email(id)` per email → Claude summarizes
- **Interactive:** any tool available directly in Claude Code and Claude Desktop sessions
- **Meeting scheduling:** `search_contacts("name")` → `create_event(attendees=[...], is_online_meeting=True)`
- **AI email filing:** `read_inbox(since=...)` → Claude decides folder per email → `file_emails(moves=[...])`

## Dependencies
httpx, mcp[cli], msal, python-dotenv - that's it.

## Maintaining this file

Keep this file for knowledge useful to almost every future agent session in this project.
Do not repeat what the codebase already shows; point to the authoritative file or command instead.
Prefer rewriting or pruning existing entries over appending new ones.
When updating this file, preserve this bar for all agents and keep entries concise.
