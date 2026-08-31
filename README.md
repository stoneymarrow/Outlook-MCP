# Outlook MCP Server

MCP server for Microsoft 365 email, calendar, and contacts using client credentials flow.
Authentication does not require browser login or delegated-token refresh.

## Why this exists

Off-the-shelf Microsoft 365 MCP servers use delegated auth (device code flow), which expires and requires re-login.
They also load 130+ tool schemas into every conversation even when you only need email.
This server uses application permissions via MSAL.

## Azure AD Setup

### 1. Register an app

Go to **Azure Portal > App Registrations > New Registration**.

### 2. Choose an application-permission profile

Under **API Permissions > Add a permission > Microsoft Graph > Application permissions**, choose one profile.

#### Recommended: read-only intake

| Permission | Purpose |
|-----------|---------|
| `Mail.Read` | Read emails and pull attachments without mailbox mutation |
| `Calendars.Read` | List calendar events |
| `Contacts.Read` | List and search contacts |

Use this profile for intake deployments.
Graph rejects any tool operation that needs a permission outside this set.
Do not add write-capable application permissions to an intake deployment; use a separate app registration and deployment when writes are intentional.

#### Optional: write-enabled email

To use all paths in `send_email`, `reply_email`, and `forward_email`, replace `Mail.Read` with `Mail.ReadWrite` and add `Mail.Send`.
Keep `Calendars.Read` and `Contacts.Read` if the deployment also uses the read tools.
This profile covers the email operations below, not unrelated calendar mutations.

Microsoft's current permission tables list these least-privileged **application** permissions:

| Operation used by this server | Least application permission |
|-----------|---------|
| New message ([`sendMail`](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)) | `Mail.Send` |
| Direct reply or reply-all ([`reply`](https://learn.microsoft.com/en-us/graph/api/message-reply?view=graph-rest-1.0), [`replyAll`](https://learn.microsoft.com/en-us/graph/api/message-replyall?view=graph-rest-1.0)) | `Mail.Send` |
| Direct forward ([`forward`](https://learn.microsoft.com/en-us/graph/api/message-forward?view=graph-rest-1.0)) | `Mail.Send` |
| Create a reply or reply-all draft ([`createReply`](https://learn.microsoft.com/en-us/graph/api/message-createreply?view=graph-rest-1.0), [`createReplyAll`](https://learn.microsoft.com/en-us/graph/api/message-createreplyall?view=graph-rest-1.0)) | `Mail.ReadWrite` |
| Create a forward draft ([`createForward`](https://learn.microsoft.com/en-us/graph/api/message-createforward?view=graph-rest-1.0)) | `Mail.ReadWrite` |
| [Add a file attachment](https://learn.microsoft.com/en-us/graph/api/message-post-attachments?view=graph-rest-1.0) to a draft | `Mail.ReadWrite` |
| [Send an existing draft](https://learn.microsoft.com/en-us/graph/api/message-send?view=graph-rest-1.0) | `Mail.Send` |

Without local attachments, replies and forwards use the direct actions and need `Mail.Send`.
With local attachments, they create a draft, add files, and send the draft, so the complete flow needs both `Mail.ReadWrite` and `Mail.Send`.
`Mail.ReadWrite` does not include permission to send mail.

### 3. Grant admin consent

Click **Grant admin consent for [your tenant]**.
Every permission in the selected profile must show a green checkmark.

At runtime, client-credential authentication requests `https://graph.microsoft.com/.default`.
Azure app registration and admin consent therefore determine the application permissions in the token.
The read-only constants in `main.py` and their regression tests document and guard the recommended repository profile; they do not request individual OAuth scopes, remove permissions already granted in Azure, or force a write-enabled registration to be read-only.
A deployment whose Azure registration includes and has admin consent for the write permissions above can therefore execute the corresponding retained mutation tools.

### 4. Create a client secret

Under **Certificates & secrets > New client secret**.
Copy the value immediately because it is shown only once.

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

## Tools (21)

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

#### `list_folder_messages`
List messages in any mail folder, with an optional lower date bound.

```
list_folder_messages(folder_id="AAMkAG...", top=50, since="2026-04-01")
```

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
    cc=["boss@example.com"],
    attachments=["/path/to/agenda.pdf"]
)
```

#### `reply_email`
Reply to an existing email thread.

```
reply_email(message_id="AAMkAG...", body="Noted, thanks.")
reply_email(message_id="AAMkAG...", body="Sharing with the team.", reply_all=True)
reply_email(
    message_id="AAMkAG...",
    body="Attaching the backup.",
    attachments=["/path/to/backup.xlsx"]
)
```

#### `forward_email`
Forward an email to new recipients.

```
forward_email(
    message_id="AAMkAG...",
    to=["colleague@example.com"],
    body="FYI - see below.",
    attachments=["/path/to/context.pdf"]
)
```

For `send_email`, `reply_email`, and `forward_email`, `attachments` is an optional list of local file paths. The server validates each path before any Graph write request, enforces the same 100 MiB per-file local guard used by attachment downloads, and sends each file as a Microsoft Graph `fileAttachment` with filename, detected content type, and base64 content. That local guard is not a promise that Graph accepts files of that size; in particular, Microsoft documents the draft [file-attachment operation](https://learn.microsoft.com/en-us/graph/api/message-post-attachments?view=graph-rest-1.0) used by attached replies and forwards as limited to attachments under 3 MB.

A routed live check on 2026-08-31 verified the new-message path in a write-enabled deployment: Graph returned `202 Accepted` for a self-addressed message with a small local text attachment, and inbox readback found the same subject with `hasAttachments=true`. This demonstrates the permission-dependent capability without changing the recommended read-only intake profile. Replies, forwards, and multiple-file request sequences are covered by the repository's mocked test suite rather than that live check.

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
List every attachment on an email.
The tool follows every Graph `@odata.nextLink`, including pages after the first.

```
list_attachments(message_id="AAMkAG...")
```

Returns: id, name, contentType, size (bytes), kind, and inline status.
The kind is `fileAttachment`, `itemAttachment`, or `referenceAttachment`.

#### `download_attachment`
Deliver a file attachment to a caller-selected local directory.
The file bytes are streamed directly from Graph's raw attachment endpoint and never enter the MCP response.

```
download_attachment(
    message_id="AAMkAG...",
    attachment_id="AAMkAG...",
    destination_directory="/path/to/quarantine",
)
```

Returns only delivery metadata: original filename, saved filename and path, content type, delivered byte size, SHA-256, attachment kind, and inline status.
Ordinary filenames are preserved.
Unsafe path characters and traversal sequences are sanitized, collisions receive numeric suffixes, existing files are never overwritten, and incomplete temporary files are removed.

The default decoded-size limit is 100 MiB and the default total download timeout is 30 seconds.
Callers can lower either limit with `max_bytes` and `timeout_seconds`.
Inline `fileAttachment` values are supported.
Downloads of `itemAttachment` and `referenceAttachment` fail with the typed `unsupported_attachment_kind` error.

The server writes bytes only and never opens, parses, or executes them.
The caller owns quarantine policy and its processed-message or processed-attachment ID ledger.

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

#### `delete_folder`
Delete an empty mail folder, or pass `force=True` to delete a non-empty folder.

```
delete_folder(folder_id="AAMkAG...")
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
| `attendees` | list[str] | None | Email addresses - sends invites |
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
