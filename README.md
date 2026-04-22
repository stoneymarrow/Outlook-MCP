# Outlook MCP Server

MCP server for Microsoft 365 email, calendar, and contacts using client credentials flow. Auth never expires — no browser login, no token refresh needed.

## Why this exists

Off-the-shelf Microsoft 365 MCP servers use delegated auth (device code flow), which expires and requires re-login. They also load 130+ tool schemas into every conversation even when you only need email. This server uses application permissions via MSAL — authenticate once at the Azure level, and it works forever.

## Azure AD Setup

### 1. Register an app

Go to **Azure Portal > App Registrations > New Registration**.

### 2. Add application permissions

Under **API Permissions > Add a permission > Microsoft Graph > Application permissions**, add:

| Permission | Purpose |
|-----------|---------|
| `Mail.ReadWrite` | Read emails, move between folders, manage folders, attachments |
| `Mail.Send` | Send, reply, forward emails |
| `Calendars.ReadWrite` | List, create, delete calendar events |
| `Contacts.Read` | List and search contacts |

### 3. Grant admin consent

Click **Grant admin consent for [your tenant]**. All four permissions must show a green checkmark.

### 4. Create a client secret

Under **Certificates & secrets > New client secret**. Copy the value immediately — you won't see it again.

## Install

```bash
git clone <this-repo>
cd Outlook-MCP
cp .env.example .env   # fill in your credentials
uv sync
```

### `.env` file

```
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-app-client-id
AZURE_CLIENT_SECRET=your-client-secret
OUTLOOK_USER_EMAIL=you@yourdomain.com
```

## Configuration

### Claude Code

Add to `~/.claude/.mcp.json`:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "uv",
      "args": ["run", "main.py"],
      "cwd": "/path/to/Outlook-MCP"
    }
  }
}
```

### Claude Desktop

Add to `%APPDATA%/Claude/claude_desktop_config.json` (Windows) or `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS):

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "uv",
      "args": ["run", "main.py"],
      "cwd": "/path/to/Outlook-MCP"
    }
  }
}
```

## Tools (19)

### Email

#### `read_inbox`
List recent inbox emails, newest first. Supports filtering.

```
read_inbox(top=10, sender="john@example.com")
read_inbox(since="2026-04-01", subject="invoice")
```

| Param | Type | Default | Description |
|-------|------|---------|-------------|
| `top` | int | 20 | Max emails (max 50) |
| `sender` | str | None | Filter by exact sender email |
| `subject` | str | None | Filter by subject (contains) |
| `since` | str | None | Only after this date (YYYY-MM-DD) |

#### `read_email`
Get full content of a specific email.

```
read_email(message_id="AAMkAG...")
```

Returns: subject, from, to, cc, body (HTML/text), attachments flag.

#### `search_emails`
Keyword search across subject, body, and sender.

```
search_emails(query="quarterly report", top=10)
```

#### `send_email`
Send a new email.

```
send_email(
    to=["john@example.com"],
    subject="Meeting follow-up",
    body="Thanks for the discussion today.",
    cc=["boss@example.com"]
)
```

#### `reply_email`
Reply to an existing email thread.

```
reply_email(message_id="AAMkAG...", body="Noted, thanks.")
reply_email(message_id="AAMkAG...", body="Sharing with the team.", reply_all=True)
```

#### `forward_email`
Forward an email to new recipients.

```
forward_email(
    message_id="AAMkAG...",
    to=["colleague@example.com"],
    body="FYI — see below."
)
```

#### `move_email`
Move an email to a different folder.

```
move_email(message_id="AAMkAG...", destination_folder_id="AAMkAG...")
```

Use `list_folders` to find folder IDs.

#### `suggest_folders`
Analyze recent inbox emails and group by sender domain. Useful for planning folder structure.

```
suggest_folders(top=100)
```

Returns domains ranked by email count with sample subjects and email IDs, so you can then call `create_folder` and `move_email` to organize.

### Filing

#### `file_emails`
Bulk move emails to folders by name. Resolves names to IDs automatically, creates folders that don't exist, supports nested paths with slash notation.

```
file_emails(moves=[
    {"email_id": "AAMkAG...", "folder": "Clients/Acme"},
    {"email_id": "AAMkAG...", "folder": "Invoices"},
    {"email_id": "AAMkAG...", "folder": "Newsletters"},
])
```

Creates `Clients/Acme` as a nested folder if it doesn't exist. Returns count of filed emails and per-email status.

### Attachments

#### `list_attachments`
List attachments on an email.

```
list_attachments(message_id="AAMkAG...")
```

Returns: id, name, contentType, size (bytes).

#### `get_attachment`
Download an attachment (base64-encoded content).

```
get_attachment(message_id="AAMkAG...", attachment_id="AAMkAG...")
```

### Folders

#### `list_folders`
List all mail folders and their child folders.

```
list_folders()
```

Returns: id, name, total count, unread count, children.

#### `create_folder`
Create a new mail folder.

```
create_folder(name="Invoices")
create_folder(name="2026", parent_folder_id="AAMkAG...")  # nested
```

### Calendar

#### `list_events`
List calendar events in a date range. Expands recurring events into individual occurrences.

```
list_events(start="2026-04-01T00:00:00", end="2026-04-07T23:59:59")
```

Returns: subject, start/end times, location, attendees with response status, online meeting flag.

#### `get_event`
Get full event details including body and Teams join URL.

```
get_event(event_id="AAMkAG...")
```

#### `create_event`
Create a calendar event or meeting. Sends invites automatically.

```
# Simple event
create_event(
    subject="Lunch with Alice",
    start="2026-04-05T12:30:00",
    end="2026-04-05T13:30:00",
    location="Taj Lands End"
)

# Teams meeting with attendees
create_event(
    subject="Project Kickoff",
    start="2026-04-07T10:00:00",
    end="2026-04-07T11:00:00",
    attendees=["alice@example.com", "bob@example.com"],
    body="Agenda:\n1. Timeline\n2. Scope\n3. Next steps",
    is_online_meeting=True
)
```

| Param | Type | Default | Description |
|-------|------|---------|-------------|
| `subject` | str | required | Event title |
| `start` | str | required | Start datetime (ISO 8601) |
| `end` | str | required | End datetime (ISO 8601) |
| `attendees` | list[str] | None | Email addresses — sends invites |
| `location` | str | None | Location name |
| `body` | str | None | Description/agenda |
| `body_type` | str | "Text" | "Text" or "HTML" |
| `is_online_meeting` | bool | False | Generate Teams meeting link |
| `time_zone` | str | "India Standard Time" | Time zone for start/end |

#### `delete_event`
Delete or cancel a calendar event.

```
delete_event(event_id="AAMkAG...")
```

### Contacts

#### `list_contacts`
List Outlook contacts sorted by name.

```
list_contacts(top=50)
```

Returns: name, emails, company, job title, phone numbers.

#### `search_contacts`
Search contacts by name, email, or company.

```
search_contacts(query="Alice")
search_contacts(query="example.com")
```

## Common workflows

### Daily email summary
```
read_inbox(since="2026-04-02") → get list of emails
read_email(id) for each → get full content
Claude summarizes → writes to Obsidian Emails/2026-04-02.md
```

### Schedule a meeting
```
search_contacts("Alice") → get email address
list_events(start=..., end=...) → check availability
create_event(subject=..., attendees=[...], is_online_meeting=True)
```

### AI-powered email filing
```
read_inbox(since="2026-04-01") → Claude sees subjects + senders
list_folders() → Claude sees existing folder structure
file_emails(moves=[...]) → bulk move with folder auto-creation
```

### Organize inbox by domain
```
suggest_folders(top=200) → see domains ranked by volume
file_emails(moves=[...]) → bulk move by folder name
```

### Reply with attachment check
```
read_inbox(top=5) → pick an email
list_attachments(message_id) → check what's attached
reply_email(message_id, body="Received the document, reviewing now.")
```
