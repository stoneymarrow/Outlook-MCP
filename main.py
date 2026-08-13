"""Outlook MCP Server - client credentials flow, no token expiry."""

import asyncio
import base64
import hashlib
import json
import mimetypes
import os
import signal
import stat
import threading
import time
import traceback
import uuid
from contextlib import asynccontextmanager
from pathlib import Path
from urllib.parse import quote, urlsplit

import httpx
import msal
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

load_dotenv()

TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USER_EMAIL = os.environ["OUTLOOK_USER_EMAIL"]

GRAPH_BASE = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}"
GRAPH_TIMEOUT = 30.0

# Client-credential tokens use Graph's mandatory ``.default`` scope. These
# constants are the authoritative application-permission contract for the Azure
# registration and are regression-tested to remain read-only.
READ_ONLY_GRAPH_PERMISSIONS = frozenset(
    {"Mail.Read", "Mail.ReadBasic", "Calendars.Read", "Contacts.Read"}
)
REQUESTED_GRAPH_PERMISSIONS = frozenset(
    {"Mail.Read", "Calendars.Read", "Contacts.Read"}
)
READ_ONLY_MAIL_PERMISSIONS = frozenset(
    permission
    for permission in READ_ONLY_GRAPH_PERMISSIONS
    if permission.startswith("Mail.")
)
REQUESTED_MAIL_PERMISSIONS = frozenset(
    permission
    for permission in REQUESTED_GRAPH_PERMISSIONS
    if permission.startswith("Mail.")
)

DEFAULT_ATTACHMENT_MAX_BYTES = 100 * 1024 * 1024
DEFAULT_ATTACHMENT_TIMEOUT_SECONDS = 30.0

# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
_msal_app: msal.ConfidentialClientApplication | None = None


def _get_token() -> str:
    global _msal_app
    if _msal_app is None:
        _msal_app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET,
        )
    result = _msal_app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" in result:
        return result["access_token"]
    raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")


# ---------------------------------------------------------------------------
# HTTP client lifecycle - single client reused across all tool calls
# ---------------------------------------------------------------------------
_http: httpx.AsyncClient | None = None


@asynccontextmanager
async def _lifespan(server):
    global _http
    _http = httpx.AsyncClient(
        base_url=GRAPH_BASE,
        timeout=GRAPH_TIMEOUT,
    )
    try:
        yield
    finally:
        await _http.aclose()
        _http = None


def _auth_headers() -> dict[str, str]:
    return {"Authorization": f"Bearer {_get_token()}"}


async def _graph_get(
    path: str,
    params: dict | None = None,
    timeout: float | None = None,
) -> dict:
    request_options: dict = {
        "headers": _auth_headers(),
        "params": params,
    }
    if timeout is not None:
        request_options["timeout"] = timeout
    r = await _http.get(path, **request_options)
    if r.status_code >= 400:
        _raise_graph_error(r)
    return r.json()


def _graph_path_segment(value: str) -> str:
    """Encode an opaque Graph identifier as one URL path segment."""
    return quote(value, safe="")


async def _graph_post(path: str, body: dict) -> dict:
    r = await _http.post(
        path,
        headers={**_auth_headers(), "Content-Type": "application/json"},
        json=body,
    )
    if r.status_code >= 400:
        _raise_graph_error(r)
    return r.json()


async def _graph_post_no_response(path: str, body: dict | None = None) -> None:
    request_options: dict = {
        "headers": {**_auth_headers(), "Content-Type": "application/json"},
    }
    if body is not None:
        request_options["json"] = body
    r = await _http.post(path, **request_options)
    if r.status_code >= 400:
        _raise_graph_error(r)


def _raise_graph_error(r: httpx.Response):
    try:
        err = r.json().get("error", {})
        msg = err.get("message", r.text)
    except Exception:
        msg = r.text
    raise RuntimeError(f"Graph API {r.status_code}: {msg}")


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------
_MSG_LIST_FIELDS = (
    "id,internetMessageId,subject,from,receivedDateTime,isRead,hasAttachments"
)
_MSG_FULL_FIELDS = f"{_MSG_LIST_FIELDS},body,toRecipients,ccRecipients"


def _fmt_message(m: dict, full: bool = False) -> dict:
    sender = m.get("from", {}).get("emailAddress", {}) or {}
    out = {
        "id": m["id"],
        "internetMessageId": m.get("internetMessageId", ""),
        "subject": m.get("subject", "(no subject)"),
        "from": sender.get("address", "unknown"),
        "fromName": sender.get("name", ""),
        "received": m.get("receivedDateTime", ""),
        "isRead": m.get("isRead", False),
        "hasAttachments": m.get("hasAttachments", False),
    }
    if full:
        body = m.get("body", {})
        out["bodyType"] = body.get("contentType", "text")
        out["body"] = body.get("content", "")
        out["to"] = [r["emailAddress"]["address"] for r in m.get("toRecipients", [])]
        out["cc"] = [r["emailAddress"]["address"] for r in m.get("ccRecipients", [])]
    return out


def _fmt_folder(f: dict) -> dict:
    return {
        "id": f["id"],
        "name": f["displayName"],
        "total": f.get("totalItemCount", 0),
        "unread": f.get("unreadItemCount", 0),
    }


def _escape_odata(s: str) -> str:
    """Escape single quotes for OData filter strings."""
    return s.replace("'", "''")


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------
mcp = FastMCP("outlook-mcp", lifespan=_lifespan)


@mcp.tool()
async def read_inbox(
    top: int = 20,
    sender: str | None = None,
    subject: str | None = None,
    since: str | None = None,
) -> str:
    """List recent inbox emails, newest first.

    Args:
        top: Number of emails to return (max 50).
        sender: Filter by exact sender email address.
        subject: Filter by subject (contains match).
        since: Only emails after this date (YYYY-MM-DD).
    """
    top = min(top, 50)
    filters = []
    if sender:
        filters.append(f"from/emailAddress/address eq '{_escape_odata(sender)}'")
    if subject:
        filters.append(f"contains(subject, '{_escape_odata(subject)}')")
    if since:
        filters.append(f"receivedDateTime ge {since}T00:00:00Z")

    params: dict[str, str] = {
        "$top": str(top),
        "$orderby": "receivedDateTime desc",
        "$select": _MSG_LIST_FIELDS,
    }
    if filters:
        params["$filter"] = " and ".join(filters)

    data = await _graph_get("/mailFolders/inbox/messages", params)
    return json.dumps([_fmt_message(m) for m in data.get("value", [])], indent=1)


@mcp.tool()
async def read_email(message_id: str) -> str:
    """Get full content of a specific email by its ID."""
    data = await _graph_get(
        f"/messages/{message_id}",
        params={"$select": _MSG_FULL_FIELDS},
    )
    return json.dumps(_fmt_message(data, full=True), indent=1)


@mcp.tool()
async def list_folder_messages(
    folder_id: str,
    top: int = 50,
    since: str | None = None,
) -> str:
    """List messages in a specific mail folder, newest first.

    Works for any folder (use list_folders to find IDs). Unlike read_inbox,
    which is hardcoded to the Inbox folder.

    Args:
        folder_id: ID of the folder to list (from list_folders).
        top: Number of messages to return (max 100).
        since: Only messages after this date (YYYY-MM-DD).
    """
    top = min(top, 100)
    params: dict[str, str] = {
        "$top": str(top),
        "$orderby": "receivedDateTime desc",
        "$select": _MSG_LIST_FIELDS,
    }
    if since:
        params["$filter"] = f"receivedDateTime ge {since}T00:00:00Z"

    data = await _graph_get(f"/mailFolders/{folder_id}/messages", params)
    return json.dumps([_fmt_message(m) for m in data.get("value", [])], indent=1)


@mcp.tool()
async def search_emails(query: str, top: int = 20) -> str:
    """Search emails by keyword across subject, body, and sender.

    Args:
        query: Search keyword or phrase.
        top: Max results (max 50).
    """
    top = min(top, 50)
    data = await _graph_get(
        "/messages",
        params={
            "$search": f'"{query}"',
            "$top": str(top),
            "$select": _MSG_LIST_FIELDS,
        },
    )
    return json.dumps([_fmt_message(m) for m in data.get("value", [])], indent=1)


@mcp.tool()
async def file_emails(moves: list[dict]) -> str:
    """Move multiple emails to folders by folder name in a single call.

    Resolves folder names to IDs automatically. Supports nested folders
    with slash notation (e.g. "Clients/Acme"). Creates folders that
    don't exist yet.

    Args:
        moves: List of {"email_id": "...", "folder": "FolderName"} dicts.
               Use slash for nested folders: "Parent/Child".

    Example:
        file_emails(moves=[
            {"email_id": "AAMk...", "folder": "Clients/Acme"},
            {"email_id": "AAMk...", "folder": "Invoices"},
            {"email_id": "AAMk...", "folder": "Newsletters"},
        ])
    """
    # Build folder name→ID lookup (including children)
    folder_data = await _graph_get("/mailFolders", params={"$top": "100"})
    name_to_id: dict[str, str] = {}
    parent_ids: dict[str, str] = {}  # name → id for top-level folders
    for f in folder_data.get("value", []):
        name = f["displayName"]
        fid = f["id"]
        name_to_id[name.lower()] = fid
        parent_ids[name.lower()] = fid
        if f.get("childFolderCount", 0) > 0:
            children = await _graph_get(
                f"/mailFolders/{fid}/childFolders",
                params={"$top": "100"},
            )
            for c in children.get("value", []):
                child_name = c["displayName"]
                name_to_id[f"{name}/{child_name}".lower()] = c["id"]

    results = []
    for move in moves:
        email_id = move["email_id"]
        folder_path = move["folder"]
        folder_key = folder_path.lower()

        # Resolve or create the folder
        if folder_key in name_to_id:
            dest_id = name_to_id[folder_key]
        else:
            # Create the folder (handle nested paths)
            parts = folder_path.split("/")
            current_parent = None
            for i, part in enumerate(parts):
                partial_key = "/".join(parts[: i + 1]).lower()
                if partial_key in name_to_id:
                    current_parent = name_to_id[partial_key]
                else:
                    path = (
                        f"/mailFolders/{current_parent}/childFolders"
                        if current_parent
                        else "/mailFolders"
                    )
                    created = await _graph_post(path, {"displayName": part})
                    current_parent = created["id"]
                    name_to_id[partial_key] = current_parent
            dest_id = current_parent

        # Move the email. Graph's /move returns the new message resource,
        # which has a NEW folder-scoped id (the old id becomes stale) but
        # the SAME stable internetMessageId across moves.
        try:
            data = await _graph_post(
                f"/messages/{email_id}/move",
                {"destinationId": dest_id},
            )
            results.append(
                {
                    "email_id": email_id,
                    "new_id": data.get("id", ""),
                    "internet_message_id": data.get("internetMessageId", ""),
                    "folder": folder_path,
                    "status": "filed",
                }
            )
        except RuntimeError as e:
            results.append(
                {"email_id": email_id, "folder": folder_path, "status": f"error: {e}"}
            )

    filed = sum(1 for r in results if r["status"] == "filed")
    return json.dumps(
        {"filed": filed, "total": len(results), "results": results}, indent=1
    )


@mcp.tool()
async def move_email(message_id: str, destination_folder_id: str) -> str:
    """Move an email to a different folder.

    Args:
        message_id: ID of the email to move.
        destination_folder_id: ID of the target folder (use list_folders to find IDs).
    """
    data = await _graph_post(
        f"/messages/{message_id}/move",
        {"destinationId": destination_folder_id},
    )
    return json.dumps(
        {
            "status": "moved",
            "id": data["id"],
            "to_folder": destination_folder_id,
        }
    )


@mcp.tool()
async def list_folders() -> str:
    """List all mail folders and their child folders."""
    data = await _graph_get("/mailFolders", params={"$top": "100"})
    folders = []
    for f in data.get("value", []):
        folder = _fmt_folder(f)
        if f.get("childFolderCount", 0) > 0:
            children = await _graph_get(
                f"/mailFolders/{f['id']}/childFolders",
                params={"$top": "100"},
            )
            folder["children"] = [_fmt_folder(c) for c in children.get("value", [])]
        folders.append(folder)
    return json.dumps(folders, indent=1)


@mcp.tool()
async def create_folder(name: str, parent_folder_id: str | None = None) -> str:
    """Create a new mail folder.

    Args:
        name: Name for the new folder.
        parent_folder_id: Create as subfolder of this folder. Top-level if omitted.
    """
    path = (
        f"/mailFolders/{parent_folder_id}/childFolders"
        if parent_folder_id
        else "/mailFolders"
    )
    data = await _graph_post(path, {"displayName": name})
    return json.dumps(_fmt_folder(data), indent=1)


@mcp.tool()
async def delete_folder(folder_id: str, force: bool = False) -> str:
    """Delete a mail folder.

    By default, refuses to delete folders that contain messages or child
    folders - move them out first with move_email/file_emails. Pass
    force=True to delete anyway (contents go to Deleted Items).

    Args:
        folder_id: ID of the folder to delete (use list_folders to find IDs).
        force: If True, delete even if the folder is non-empty.
    """
    if not force:
        info = await _graph_get(
            f"/mailFolders/{folder_id}",
            params={"$select": "displayName,totalItemCount,childFolderCount"},
        )
        total = info.get("totalItemCount", 0)
        children = info.get("childFolderCount", 0)
        if total > 0 or children > 0:
            return json.dumps(
                {
                    "status": "refused",
                    "reason": "folder not empty",
                    "name": info.get("displayName", ""),
                    "totalItemCount": total,
                    "childFolderCount": children,
                    "hint": "move messages out first, or pass force=True",
                },
                indent=1,
            )

    r = await _http.delete(
        f"/mailFolders/{folder_id}",
        headers=_auth_headers(),
    )
    if r.status_code >= 400:
        _raise_graph_error(r)
    return json.dumps({"status": "deleted", "id": folder_id})


def _local_file_attachments(attachment_paths: list[str] | None) -> list[dict]:
    if not attachment_paths:
        return []

    attachments: list[dict] = []
    for raw_path in attachment_paths:
        path = Path(raw_path).expanduser()
        try:
            stat_result = path.stat()
        except OSError as exc:
            raise ValueError(f"Invalid attachment path {raw_path!r}: {exc.strerror}") from exc

        if not stat.S_ISREG(stat_result.st_mode):
            raise ValueError(f"Invalid attachment path {raw_path!r}: not a regular file")
        if stat_result.st_size > DEFAULT_ATTACHMENT_MAX_BYTES:
            raise ValueError(
                f"Invalid attachment path {raw_path!r}: file is {stat_result.st_size} bytes, "
                f"exceeding the {DEFAULT_ATTACHMENT_MAX_BYTES} byte limit"
            )

        try:
            content = path.read_bytes()
        except OSError as exc:
            raise ValueError(f"Invalid attachment path {raw_path!r}: {exc.strerror}") from exc

        attachments.append(
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": path.name,
                "contentType": mimetypes.guess_type(path.name)[0]
                or "application/octet-stream",
                "contentBytes": base64.b64encode(content).decode("ascii"),
            }
        )
    return attachments


async def _add_attachments_to_message(message_id: str, attachments: list[dict]) -> None:
    encoded_message_id = _graph_path_segment(message_id)
    for attachment in attachments:
        await _graph_post(
            f"/messages/{encoded_message_id}/attachments",
            attachment,
        )


@mcp.tool()
async def send_email(
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    body_type: str = "Text",
    attachments: list[str] | None = None,
) -> str:
    """Send an email.

    Args:
        to: List of recipient email addresses.
        subject: Email subject line.
        body: Email body content.
        cc: Optional list of CC email addresses.
        body_type: "Text" for plain text or "HTML" for rich content.
        attachments: Optional local file paths to attach. Files are validated and
            encoded as Graph fileAttachment values before sending.
    """
    file_attachments = _local_file_attachments(attachments)
    message: dict = {
        "subject": subject,
        "body": {"contentType": body_type, "content": body},
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
    }
    if cc:
        message["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc]
    if file_attachments:
        message["attachments"] = file_attachments

    await _graph_post_no_response("/sendMail", {"message": message})
    return json.dumps({"status": "sent", "to": to, "subject": subject})


@mcp.tool()
async def suggest_folders(top: int = 50) -> str:
    """Return recent inbox emails grouped by sender domain for folder planning.

    Returns sender domains ranked by email count with sample subjects and
    email IDs. Use this data to decide folder structure, then call
    create_folder and move_email to execute.

    Args:
        top: Number of recent inbox emails to analyze (max 200).
    """
    top = min(top, 200)
    data = await _graph_get(
        "/mailFolders/inbox/messages",
        params={
            "$top": str(top),
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime",
        },
    )

    groups: dict[str, list[dict]] = {}
    for m in data.get("value", []):
        addr = m.get("from", {}).get("emailAddress", {}).get("address", "unknown")
        domain = addr.rsplit("@", 1)[-1] if "@" in addr else "unknown"
        groups.setdefault(domain, []).append(
            {
                "id": m["id"],
                "subject": m.get("subject", ""),
                "from": addr,
                "received": m.get("receivedDateTime", ""),
            }
        )

    ranked = [
        {
            "domain": domain,
            "count": len(msgs),
            "sample_subjects": [m["subject"] for m in msgs[:5]],
            "email_ids": [m["id"] for m in msgs],
        }
        for domain, msgs in sorted(groups.items(), key=lambda x: -len(x[1]))
    ]

    return json.dumps(
        {"analyzed": len(data.get("value", [])), "by_domain": ranked}, indent=1
    )


# ---------------------------------------------------------------------------
# Reply / Forward
# ---------------------------------------------------------------------------
@mcp.tool()
async def reply_email(
    message_id: str,
    body: str,
    reply_all: bool = False,
    body_type: str = "Text",
    attachments: list[str] | None = None,
) -> str:
    """Reply to an email thread.

    Args:
        message_id: ID of the email to reply to.
        body: Reply body content.
        reply_all: True to reply to all recipients, False for sender only.
        body_type: "Text" or "HTML".
        attachments: Optional local file paths to attach. Files are validated and
            encoded as Graph fileAttachment values before sending.
    """
    file_attachments = _local_file_attachments(attachments)
    if not file_attachments:
        action = "replyAll" if reply_all else "reply"
        await _graph_post_no_response(
            f"/messages/{message_id}/{action}",
            {"comment": body},
        )
        return json.dumps({"status": "replied", "id": message_id, "replyAll": reply_all})

    encoded_message_id = _graph_path_segment(message_id)
    draft_action = "createReplyAll" if reply_all else "createReply"
    draft = await _graph_post(
        f"/messages/{encoded_message_id}/{draft_action}",
        {"message": {"body": {"contentType": body_type, "content": body}}},
    )
    draft_id = draft["id"]
    await _add_attachments_to_message(draft_id, file_attachments)
    await _graph_post_no_response(f"/messages/{_graph_path_segment(draft_id)}/send")
    return json.dumps({"status": "replied", "id": message_id, "replyAll": reply_all})


@mcp.tool()
async def forward_email(
    message_id: str,
    to: list[str],
    body: str | None = None,
    body_type: str = "Text",
    attachments: list[str] | None = None,
) -> str:
    """Forward an email to new recipients.

    Args:
        message_id: ID of the email to forward.
        to: List of recipient email addresses.
        body: Optional comment to include above the forwarded message.
        body_type: "Text" or "HTML".
        attachments: Optional local file paths to attach. Files are validated and
            encoded as Graph fileAttachment values before sending.
    """
    file_attachments = _local_file_attachments(attachments)
    payload: dict = {
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
    }
    if body:
        payload["comment"] = body
    if not file_attachments:
        await _graph_post_no_response(
            f"/messages/{message_id}/forward",
            payload,
        )
        return json.dumps({"status": "forwarded", "id": message_id, "to": to})

    encoded_message_id = _graph_path_segment(message_id)
    message: dict = {
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
    }
    if body is not None:
        message["body"] = {"contentType": body_type, "content": body}
    draft = await _graph_post(
        f"/messages/{encoded_message_id}/createForward",
        {"message": message},
    )
    draft_id = draft["id"]
    await _add_attachments_to_message(draft_id, file_attachments)
    await _graph_post_no_response(f"/messages/{_graph_path_segment(draft_id)}/send")
    return json.dumps({"status": "forwarded", "id": message_id, "to": to})


# ---------------------------------------------------------------------------
# Attachments
# ---------------------------------------------------------------------------
_ATTACH_LIST_FIELDS = "id,name,contentType,size,isInline"
_ATTACHMENT_KINDS = frozenset(
    {"fileAttachment", "itemAttachment", "referenceAttachment"}
)


class AttachmentDeliveryError(RuntimeError):
    """Base class for typed attachment-delivery failures."""

    code = "attachment_delivery_error"


class UnsupportedAttachmentKindError(AttachmentDeliveryError):
    """Raised when Graph cannot provide a supported file payload."""

    code = "unsupported_attachment_kind"

    def __init__(self, attachment_id: str, kind: str) -> None:
        self.attachment_id = attachment_id
        self.kind = kind
        super().__init__(
            f"{self.code}: attachment {attachment_id!r} has kind {kind!r}; "
            "only fileAttachment downloads are supported"
        )


class AttachmentTooLargeError(AttachmentDeliveryError):
    """Raised before a download can exceed the configured decoded-size cap."""

    code = "attachment_too_large"


class AttachmentSizeMismatchError(AttachmentDeliveryError):
    """Raised when Graph attachment metadata is internally invalid.

    Note: this never compares the delivered decoded byte count to Graph's
    declared ``size`` field. Graph reports the MIME/base64-encoded size while
    the raw ``$value`` endpoint returns decoded bytes, so the two legitimately
    differ. It only guards against nonsensical metadata (e.g. a negative size).
    """

    code = "attachment_size_mismatch"


class AttachmentIntegrityError(AttachmentDeliveryError):
    """Raised when delivered bytes fail structural integrity validation.

    The delivered file is preserved on disk (never deleted) so a false-positive
    gate can never destroy valid data; ``path`` names where the bytes were kept.
    """

    code = "attachment_integrity_failed"

    def __init__(self, filename: str, path: str, reason: str):
        self.filename = filename
        self.path = path
        self.reason = reason
        super().__init__(
            f"{self.code}: delivered attachment {filename!r} failed integrity "
            f"validation ({reason}); the bytes were preserved for inspection "
            f"at {path}"
        )


# Structural integrity is validated from magic-number signatures, not by fully
# parsing content. We key off the delivered bytes themselves rather than Graph's
# declared content type, since a content type can misdescribe the payload.
_INTEGRITY_HEADER_BYTES = 8
_INTEGRITY_TAIL_BYTES = 2048
_PDF_MAGIC = b"%PDF-"
_PDF_EOF = b"%%EOF"
_ZIP_MAGICS = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
_ZIP_EOCD = b"PK\x05\x06"


def _verify_delivered_integrity(
    *,
    filename: str,
    path: str,
    header: bytes,
    tail: bytes,
    delivered_size: int,
    has_declared_size: bool,
    declared_size: int,
) -> None:
    """Validate delivered bytes on their own terms, never against the MIME size.

    Confirms structural integrity for formats the payload's own magic bytes
    identify (PDF trailer, ZIP end-of-central-directory), which catches genuine
    truncation. Unknown formats carry no structural claim and pass on the hash
    alone. Raises :class:`AttachmentIntegrityError` (leaving the file in place)
    when a recognized container is truncated.
    """
    if has_declared_size and declared_size > 0 and delivered_size == 0:
        # Graph expected content but the raw endpoint returned nothing.
        raise AttachmentIntegrityError(
            filename, path, "delivered zero bytes for a non-empty attachment"
        )
    if header.startswith(_PDF_MAGIC):
        if _PDF_EOF not in tail:
            raise AttachmentIntegrityError(
                filename, path, "PDF is missing its %%EOF trailer (truncated)"
            )
    elif any(header.startswith(magic) for magic in _ZIP_MAGICS):
        if _ZIP_EOCD not in tail:
            raise AttachmentIntegrityError(
                filename,
                path,
                "ZIP container is missing its end-of-central-directory "
                "record (truncated)",
            )


class UnsafeAttachmentDestinationError(AttachmentDeliveryError):
    """Raised when a destination cannot be used without following a symlink."""

    code = "unsafe_attachment_destination"


def _attachment_kind(attachment: dict) -> str:
    odata_type = str(attachment.get("@odata.type", ""))
    kind = odata_type.rsplit(".", 1)[-1].lstrip("#")
    return kind if kind in _ATTACHMENT_KINDS else "unknownAttachment"


def _fmt_attachment(a: dict) -> dict:
    return {
        "id": a["id"],
        "name": a.get("name", "unnamed"),
        "contentType": a.get("contentType", ""),
        "size": a.get("size", 0),
        "kind": _attachment_kind(a),
        "isInline": bool(a.get("isInline", False)),
    }


def _validate_graph_next_link(next_link: str) -> str:
    """Keep bearer credentials on Microsoft Graph while following pagination."""
    parsed = urlsplit(next_link)
    if parsed.scheme or parsed.netloc:
        if parsed.scheme != "https" or parsed.hostname != "graph.microsoft.com":
            raise RuntimeError("Graph pagination returned an untrusted nextLink")
    elif not next_link.startswith("/"):
        raise RuntimeError("Graph pagination returned an invalid nextLink")
    return next_link


async def _list_all_attachments(message_id: str) -> list[dict]:
    path = f"/messages/{_graph_path_segment(message_id)}/attachments"
    params: dict | None = {"$select": _ATTACH_LIST_FIELDS}
    attachments: list[dict] = []
    seen_links: set[str] = set()

    while path:
        data = await _graph_get(path, params=params)
        attachments.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        path = _validate_graph_next_link(str(next_link))
        if path in seen_links:
            raise RuntimeError("Graph pagination returned a repeated nextLink")
        seen_links.add(path)
        params = None

    return attachments


@mcp.tool()
async def list_attachments(message_id: str) -> str:
    """List every attachment on an email, following Graph pagination.

    Args:
        message_id: ID of the email.

    Returns file, inline file, item, and reference attachment metadata. Item
    and reference attachments are visible here but cannot be downloaded.
    """
    attachments = await _list_all_attachments(message_id)
    return json.dumps([_fmt_attachment(a) for a in attachments], indent=1)


def _sanitize_attachment_filename(filename: str) -> str:
    """Return one safe path segment while preserving ordinary filenames."""
    sanitized = filename.replace("/", "_").replace("\\", "_")
    sanitized = "".join(
        "_" if ord(character) < 32 or ord(character) == 127 else character
        for character in sanitized
    )
    while ".." in sanitized:
        sanitized = sanitized.replace("..", "_")
    sanitized = sanitized.replace("\x00", "_")
    if sanitized in {"", ".", ".."}:
        return "unnamed-attachment"
    return sanitized


def _safe_destination_directory(destination_directory: str) -> Path:
    if not destination_directory or "\x00" in destination_directory:
        raise UnsafeAttachmentDestinationError("Invalid destination directory")

    requested_destination = Path(destination_directory).expanduser().absolute()
    try:
        requested_info = requested_destination.lstat()
    except FileNotFoundError:
        pass
    else:
        if stat.S_ISLNK(requested_info.st_mode):
            raise UnsafeAttachmentDestinationError(
                "Destination directory cannot be a symlink"
            )

    destination = requested_destination.resolve(strict=False)
    destination.mkdir(mode=0o700, parents=True, exist_ok=True)
    try:
        info = destination.lstat()
    except OSError as error:
        raise UnsafeAttachmentDestinationError(
            f"Cannot inspect destination directory: {error}"
        ) from error
    if stat.S_ISLNK(info.st_mode) or not stat.S_ISDIR(info.st_mode):
        raise UnsafeAttachmentDestinationError(
            "Destination must be a real directory, not a symlink or file"
        )
    return destination


def _deduplicated_filename(directory_fd: int, filename: str) -> tuple[str, int]:
    suffix = Path(filename).suffix
    stem = filename[: -len(suffix)] if suffix else filename
    collision_number = 0
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW

    while True:
        candidate = (
            filename
            if collision_number == 0
            else f"{stem} ({collision_number}){suffix}"
        )
        try:
            reservation_fd = os.open(
                candidate,
                flags,
                0o600,
                dir_fd=directory_fd,
            )
            return candidate, reservation_fd
        except FileExistsError:
            collision_number += 1


def _unlink_if_present(filename: str | None, directory_fd: int) -> None:
    if not filename:
        return
    try:
        os.unlink(filename, dir_fd=directory_fd)
    except FileNotFoundError:
        pass


async def _download_file_attachment(
    message_id: str,
    attachment: dict,
    destination_directory: str,
    max_bytes: int,
    timeout_seconds: float,
) -> dict:
    if max_bytes <= 0:
        raise ValueError("max_bytes must be greater than zero")
    if timeout_seconds <= 0:
        raise ValueError("timeout_seconds must be greater than zero")

    has_declared_size = attachment.get("size") is not None
    declared_size = int(attachment.get("size", 0) or 0)
    if declared_size < 0:
        raise AttachmentSizeMismatchError(
            f"Graph declared an invalid attachment size of {declared_size} bytes"
        )
    if declared_size > max_bytes:
        raise AttachmentTooLargeError(
            f"Attachment declares {declared_size} bytes, above the "
            f"{max_bytes}-byte limit"
        )

    destination = _safe_destination_directory(destination_directory)
    directory_flags = os.O_RDONLY
    if hasattr(os, "O_DIRECTORY"):
        directory_flags |= os.O_DIRECTORY
    if hasattr(os, "O_NOFOLLOW"):
        directory_flags |= os.O_NOFOLLOW
    directory_fd = os.open(destination, directory_flags)
    temporary_name = f".attachment-{uuid.uuid4().hex}.partial"
    temporary_fd: int | None = None
    final_name: str | None = None
    reservation_fd: int | None = None
    delivered = False

    try:
        temporary_flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
        if hasattr(os, "O_NOFOLLOW"):
            temporary_flags |= os.O_NOFOLLOW
        temporary_fd = os.open(
            temporary_name,
            temporary_flags,
            0o600,
            dir_fd=directory_fd,
        )

        attachment_id = str(attachment["id"])
        raw_path = (
            f"/messages/{_graph_path_segment(message_id)}/attachments/"
            f"{_graph_path_segment(attachment_id)}/$value"
        )
        delivered_size = 0
        digest = hashlib.sha256()
        header = b""
        tail = b""

        with os.fdopen(temporary_fd, "wb", closefd=True) as output:
            temporary_fd = None
            async with asyncio.timeout(timeout_seconds):
                async with _http.stream(
                    "GET",
                    raw_path,
                    headers=_auth_headers(),
                    timeout=httpx.Timeout(timeout_seconds),
                ) as response:
                    if response.status_code >= 400:
                        await response.aread()
                        _raise_graph_error(response)
                    content_length = response.headers.get("Content-Length")
                    if content_length and int(content_length) > max_bytes:
                        raise AttachmentTooLargeError(
                            f"Attachment response exceeds the {max_bytes}-byte limit"
                        )
                    async for chunk in response.aiter_bytes():
                        delivered_size += len(chunk)
                        if delivered_size > max_bytes:
                            raise AttachmentTooLargeError(
                                f"Attachment exceeded the {max_bytes}-byte limit"
                            )
                        output.write(chunk)
                        digest.update(chunk)
                        if len(header) < _INTEGRITY_HEADER_BYTES:
                            header += chunk[: _INTEGRITY_HEADER_BYTES - len(header)]
                        tail = (tail + chunk)[-_INTEGRITY_TAIL_BYTES:]
            output.flush()
            os.fsync(output.fileno())

        safe_name = _sanitize_attachment_filename(
            str(attachment.get("name") or "unnamed-attachment")
        )
        final_name, reservation_fd = _deduplicated_filename(
            directory_fd,
            safe_name,
        )
        os.close(reservation_fd)
        reservation_fd = None
        os.replace(
            temporary_name,
            final_name,
            src_dir_fd=directory_fd,
            dst_dir_fd=directory_fd,
        )
        temporary_name = ""
        # The bytes are now on disk under their final name. Any validation from
        # here on must never delete them, so a false-positive gate can never
        # destroy valid data.
        delivered = True

        _verify_delivered_integrity(
            filename=str(attachment.get("name") or "unnamed-attachment"),
            path=str(destination / final_name),
            header=header,
            tail=tail,
            delivered_size=delivered_size,
            has_declared_size=has_declared_size,
            declared_size=declared_size,
        )

        return {
            "id": attachment["id"],
            "originalFilename": str(attachment.get("name") or "unnamed-attachment"),
            "filename": final_name,
            "path": str(destination / final_name),
            "contentType": attachment.get("contentType", ""),
            "size": delivered_size,
            "sha256": digest.hexdigest(),
            "kind": "fileAttachment",
            "isInline": bool(attachment.get("isInline", False)),
        }
    except BaseException:
        if temporary_fd is not None:
            os.close(temporary_fd)
        if reservation_fd is not None:
            os.close(reservation_fd)
        _unlink_if_present(temporary_name, directory_fd)
        # Once the bytes are delivered under their final name they are never
        # removed, even on a validation failure: the caller must be able to
        # inspect what actually arrived.
        if not delivered:
            _unlink_if_present(final_name, directory_fd)
        raise
    finally:
        os.close(directory_fd)


@mcp.tool()
async def download_attachment(
    message_id: str,
    attachment_id: str,
    destination_directory: str,
    max_bytes: int = DEFAULT_ATTACHMENT_MAX_BYTES,
    timeout_seconds: float = DEFAULT_ATTACHMENT_TIMEOUT_SECONDS,
) -> str:
    """Deliver a file attachment byte-faithfully to a local directory.

    Args:
        message_id: ID of the email.
        attachment_id: ID of the attachment (use list_attachments to find IDs).
        destination_directory: Caller-selected local quarantine or intake directory.
        max_bytes: Maximum delivered byte count (default 100 MiB).
        timeout_seconds: Total request timeout in seconds (default 30).

    The response contains metadata only (including a sha256 of the delivered
    bytes). Delivered bytes are validated on their own terms - hashed, and
    structurally integrity-checked from their own magic-number signatures to
    catch truncation - never by comparing the decoded byte count to Graph's
    MIME/base64 ``size`` field, which legitimately differs. The server never
    fully parses or executes the bytes, and never deletes a delivered file: a
    payload that fails integrity is preserved on disk and the error names it.
    Consumers own quarantine policy and processed-ID state. Only fileAttachment,
    including inline files, is downloadable.
    """
    async with asyncio.timeout(timeout_seconds):
        attachment = await _graph_get(
            (
                f"/messages/{_graph_path_segment(message_id)}/attachments/"
                f"{_graph_path_segment(attachment_id)}"
            ),
            params={"$select": _ATTACH_LIST_FIELDS},
            timeout=timeout_seconds,
        )
    kind = _attachment_kind(attachment)
    if kind != "fileAttachment":
        raise UnsupportedAttachmentKindError(attachment_id, kind)

    response = await _download_file_attachment(
        message_id=message_id,
        attachment=attachment,
        destination_directory=destination_directory,
        max_bytes=max_bytes,
        timeout_seconds=timeout_seconds,
    )
    return json.dumps(response, indent=1)


# ---------------------------------------------------------------------------
# Contacts
# ---------------------------------------------------------------------------
_CONTACT_FIELDS = (
    "id,displayName,emailAddresses,companyName,jobTitle,mobilePhone,businessPhones"
)


def _fmt_contact(c: dict) -> dict:
    emails = c.get("emailAddresses", [])
    return {
        "id": c["id"],
        "name": c.get("displayName", ""),
        "emails": [e.get("address", "") for e in emails],
        "company": c.get("companyName", ""),
        "jobTitle": c.get("jobTitle", ""),
        "mobile": c.get("mobilePhone", ""),
        "phones": c.get("businessPhones", []),
    }


@mcp.tool()
async def list_contacts(top: int = 50) -> str:
    """List Outlook contacts.

    Args:
        top: Max contacts to return (max 100).
    """
    top = min(top, 100)
    data = await _graph_get(
        "/contacts",
        params={
            "$top": str(top),
            "$orderby": "displayName",
            "$select": _CONTACT_FIELDS,
        },
    )
    return json.dumps([_fmt_contact(c) for c in data.get("value", [])], indent=1)


@mcp.tool()
async def search_contacts(query: str, top: int = 20) -> str:
    """Search contacts by name or email.

    Args:
        query: Search keyword (matches name, email, company).
        top: Max results (max 50).
    """
    top = min(top, 50)
    data = await _graph_get(
        "/contacts",
        params={
            "$search": f'"{query}"',
            "$top": str(top),
            "$select": _CONTACT_FIELDS,
        },
    )
    return json.dumps([_fmt_contact(c) for c in data.get("value", [])], indent=1)


# ---------------------------------------------------------------------------
# Calendar helpers
# ---------------------------------------------------------------------------
_EVENT_LIST_FIELDS = (
    "id,subject,start,end,location,organizer,attendees,isOnlineMeeting,webLink"
)
_EVENT_FULL_FIELDS = f"{_EVENT_LIST_FIELDS},body,onlineMeeting"


def _fmt_event(e: dict, full: bool = False) -> dict:
    out = {
        "id": e["id"],
        "subject": e.get("subject", "(no subject)"),
        "start": e.get("start", {}).get("dateTime", ""),
        "end": e.get("end", {}).get("dateTime", ""),
        "timeZone": e.get("start", {}).get("timeZone", ""),
        "location": e.get("location", {}).get("displayName", ""),
        "organizer": e.get("organizer", {}).get("emailAddress", {}).get("address", ""),
        "isOnlineMeeting": e.get("isOnlineMeeting", False),
        "attendees": [
            {
                "email": a["emailAddress"]["address"],
                "type": a.get("type", "required"),
                "response": a.get("status", {}).get("response", "none"),
            }
            for a in e.get("attendees", [])
        ],
    }
    if full:
        body = e.get("body", {})
        out["bodyType"] = body.get("contentType", "text")
        out["body"] = body.get("content", "")
        meeting = e.get("onlineMeeting") or {}
        out["joinUrl"] = meeting.get("joinUrl", "")
    return out


# ---------------------------------------------------------------------------
# Calendar tools
# ---------------------------------------------------------------------------
@mcp.tool()
async def list_events(
    start: str,
    end: str,
    top: int = 25,
) -> str:
    """List calendar events in a date range.

    Args:
        start: Start date/time in ISO 8601 (e.g. 2026-04-01T00:00:00).
        end: End date/time in ISO 8601 (e.g. 2026-04-07T23:59:59).
        top: Max events to return (max 50).
    """
    top = min(top, 50)
    data = await _graph_get(
        "/calendarView",
        params={
            "startDateTime": start,
            "endDateTime": end,
            "$top": str(top),
            "$orderby": "start/dateTime",
            "$select": _EVENT_LIST_FIELDS,
        },
    )
    return json.dumps([_fmt_event(e) for e in data.get("value", [])], indent=1)


@mcp.tool()
async def get_event(event_id: str) -> str:
    """Get full details of a calendar event by ID."""
    data = await _graph_get(
        f"/events/{event_id}",
        params={"$select": _EVENT_FULL_FIELDS},
    )
    return json.dumps(_fmt_event(data, full=True), indent=1)


@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    attendees: list[str] | None = None,
    location: str | None = None,
    body: str | None = None,
    body_type: str = "Text",
    is_online_meeting: bool = False,
    time_zone: str = "India Standard Time",
) -> str:
    """Create a calendar event or meeting.

    Args:
        subject: Event title.
        start: Start date/time (e.g. 2026-04-05T10:00:00).
        end: End date/time (e.g. 2026-04-05T11:00:00).
        attendees: List of attendee email addresses. Sends invite automatically.
        location: Location name (e.g. "Conference Room A").
        body: Event description/agenda.
        body_type: "Text" or "HTML".
        is_online_meeting: Set true to generate a Teams meeting link.
        time_zone: IANA or Windows time zone (default: India Standard Time).
    """
    event: dict = {
        "subject": subject,
        "start": {"dateTime": start, "timeZone": time_zone},
        "end": {"dateTime": end, "timeZone": time_zone},
        "isOnlineMeeting": is_online_meeting,
    }
    if attendees:
        event["attendees"] = [
            {"emailAddress": {"address": a}, "type": "required"} for a in attendees
        ]
    if location:
        event["location"] = {"displayName": location}
    if body:
        event["body"] = {"contentType": body_type, "content": body}

    data = await _graph_post("/events", event)
    return json.dumps(_fmt_event(data), indent=1)


@mcp.tool()
async def delete_event(event_id: str) -> str:
    """Delete/cancel a calendar event by ID.

    Args:
        event_id: ID of the event to delete (use list_events to find IDs).
    """
    r = await _http.delete(
        f"/events/{event_id}",
        headers=_auth_headers(),
    )
    if r.status_code >= 400:
        _raise_graph_error(r)
    return json.dumps({"status": "deleted", "id": event_id})


# How long a signalled shutdown may spend unwinding before we stop waiting on
# it. The clean unwind finishes in milliseconds; anything still running after
# this is wedged on the uncancellable stdin read and never will finish.
_SHUTDOWN_GRACE_SECONDS = 2.0

# Set by the signal handler to arm the watchdog. Module scope so the watchdog
# thread and the handler share one object without a closure cell.
_shutdown_requested = threading.Event()


def _run_stdio() -> None:
    """Serve over stdio, then leave without interpreter finalization.

    Two separate hazards live at shutdown, both caused by the same thing: the
    anyio stdio transport reads stdin on a worker thread, that thread blocks in
    ``readline`` holding the ``BufferedReader`` lock, and it cannot be
    cancelled (``to_thread.run_sync`` is not cancellable).

    1. Aborting. If CPython finalizes while the worker still holds the buffer
       lock, ``finalize_modules`` deallocates a stdio ``TextIOWrapper``, its
       close reaches ``_enter_buffered_busy``, that fails to acquire the lock
       and calls ``Py_FatalError`` -> ``abort()`` (SIGABRT), which macOS turns
       into a "Python quit unexpectedly" crash report. Skipping finalization
       with ``os._exit`` removes the hazard entirely.

    2. Wedging. On stdin EOF and on SIGTERM the transport and the
       httpx-closing lifespan (``_lifespan``) unwind cleanly, ``mcp.run``
       returns, and ``os._exit(0)`` below runs. On SIGINT it does not:
       anyio cancels the task group, then waits forever for the stdin worker
       that cannot be cancelled. ``mcp.run`` never returns, so the ``os._exit``
       below is unreachable and the server hangs until something kills it.

    So a signalled shutdown arms a watchdog before raising KeyboardInterrupt.
    The clean unwind wins the race in the normal case and exits below with
    httpx closed; if it wedges, the watchdog leaves without finalization -
    still no abort, and no hang either.

    This remedy is derived locally from the diagnosed mechanism: there is no
    upstream python-sdk issue tracking either symptom as of mcp 1.26.0 (issue
    #575 is an unrelated Windows cleanup bug; #1933 is a different,
    ValueError-on-closed-stdio symptom).
    """

    def _watchdog() -> None:
        _shutdown_requested.wait()
        time.sleep(_SHUTDOWN_GRACE_SECONDS)
        # Reached only if the unwind is wedged: exit code stays 0 because the
        # server did its work and the client has already disconnected.
        os._exit(0)

    # Started up front, not from the handler: creating a thread inside a signal
    # handler can deadlock against a thread-start already in progress.
    threading.Thread(target=_watchdog, name="shutdown-watchdog", daemon=True).start()

    def _terminate(signum: int, frame: object) -> None:
        # Re-raise as KeyboardInterrupt so anyio unwinds the transport and runs
        # the lifespan httpx cleanup exactly as it does for Ctrl+C.
        _shutdown_requested.set()
        raise KeyboardInterrupt

    signal.signal(signal.SIGTERM, _terminate)
    signal.signal(signal.SIGINT, _terminate)

    try:
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        # SIGINT/SIGTERM: cleanup already ran during the anyio unwind.
        pass
    except BaseException:
        # A genuine failure must not be masked by the clean-exit path. Surface
        # it, then exit nonzero - still without finalization, which could also
        # race the I/O worker.
        traceback.print_exc()
        os._exit(1)
    os._exit(0)


if __name__ == "__main__":
    _run_stdio()
