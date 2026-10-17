"""MCP server for classic Outlook (Windows, COM automation).

Reads mail that classic Outlook has already synced locally, so no
Microsoft 365 admin consent is needed. Can write new mail, reply / reply-all
and forward inside the original thread, with CC / BCC / attachments; saves a
draft unless told to send. Can mark read/unread, flag, move, create folders,
save attachments, and delete -- delete only moves mail to Deleted Items, so
it is always recoverable; nothing here deletes permanently.

stdout is the MCP protocol channel: log to stderr only, never print().
"""

import html
import json
import logging
import os
import re
import sys
from datetime import datetime, timedelta

import pythoncom
import win32com.client
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

logging.basicConfig(stream=sys.stderr, level=logging.INFO)
log = logging.getLogger("outlook-mcp")

# ---------------------------------------------------------------------------
# Personal settings live in config.json (gitignored); see config.example.json.
# ---------------------------------------------------------------------------

CONFIG_PATH = os.environ.get("OUTLOOK_MCP_CONFIG") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def _load_config() -> dict:
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        log.info("No config at %s; using defaults", CONFIG_PATH)
        return {}


CONFIG = _load_config()


def _instructions() -> str:
    parts = ["The user's mailboxes in classic Outlook (Outlook must be open)."]
    for box in CONFIG.get("mailboxes", []):
        parts.append(f"{box['label']}: folder='{box['folder']}'.")
    if CONFIG.get("notes"):
        parts.append(CONFIG["notes"])
    parts.append(
        "To answer or pass on an existing email use reply_email / forward_email with its "
        "entry_id so it stays in the original thread; send_email is only for new mail. "
        "Prefer send=False (draft) unless the user explicitly asked to send; "
        "send_draft sends a draft later. delete_emails is recoverable (moves to Deleted Items). "
        "Saved attachments can then be read from disk."
    )
    return " ".join(parts)


mcp = MCPServer("outlook", instructions=_instructions())
READ_ONLY = ToolAnnotations(read_only_hint=True, open_world_hint=False)
SENDS_MAIL = ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=True)
ORGANIZES = ToolAnnotations(read_only_hint=False, destructive_hint=False, idempotent_hint=True, open_world_hint=False)
DELETES = ToolAnnotations(read_only_hint=False, destructive_hint=True, open_world_hint=False)

# Outlook constants (OlDefaultFolders / OlObjectClass / OlItemType)
DEFAULT_FOLDERS = {"inbox": 6, "sent": 5, "drafts": 16, "deleted": 3, "junk": 23}
OL_MAIL = 43
OL_MAIL_ITEM = 0
OL_TO, OL_CC, OL_BCC = 1, 2, 3  # OlMailRecipientType
OL_FORMAT_PLAIN = 1  # OlBodyFormat
OL_NO_FLAG, OL_FLAG_MARKED = 0, 2  # OlFlagStatus
PREVIEW_CHARS = 200


# ---------------------------------------------------------------------------
# Filter rules (configured in config.json)
# ---------------------------------------------------------------------------

# Mail from these domains (and their subdomains) is always kept.
KEEP_DOMAINS = {d.lower() for d in CONFIG.get("keep_domains", [])}
# Mail from these domains is hidden (marketing, automated notices...).
BLOCK_DOMAINS = {d.lower() for d in CONFIG.get("block_domains", [])}
# Other mail whose subject contains any of these is treated as promo.
PROMO_KEYWORDS = tuple(k.lower() for k in CONFIG.get(
    "promo_keywords", ["unsubscribe", "newsletter", "% off", "sale", "deal", "webinar", "promo"]))


def _domain_matches(domain: str, domains: set[str]) -> bool:
    return any(domain == d or domain.endswith("." + d) for d in domains)


def is_relevant(email: dict) -> bool:
    """Decide whether an email should be returned by search_emails.

    Blocklist approach: keep everything unless it is clearly noise, so an
    unexpected but important sender is never silently hidden.
    """
    domain = email["sender_email"].lower().rpartition("@")[2]
    if not domain:
        return True  # unknown sender address: don't risk hiding it
    if _domain_matches(domain, KEEP_DOMAINS):
        return True
    if _domain_matches(domain, BLOCK_DOMAINS):
        return False
    subject = email["subject"].lower()
    return not any(k in subject for k in PROMO_KEYWORDS)


# ---------------------------------------------------------------------------
# COM helpers
# ---------------------------------------------------------------------------

def _namespace():
    # COM is per-thread; MCP may call tools from different threads.
    pythoncom.CoInitialize()
    # Attach to the Outlook the user already has open. Dispatch() would silently
    # launch a headless instance that can hang on invisible dialogs.
    try:
        app = win32com.client.GetActiveObject("Outlook.Application")
    except pythoncom.com_error:
        raise ToolError("Outlook is not running. Open classic Outlook first, then retry.")
    return app.GetNamespace("MAPI")


def _resolve_folder(ns, folder: str):
    """'inbox' / 'sent' / ... or a path like 'Inbox/Courses' or 'account@x.ca/Inbox/X'."""
    parts = [p for p in folder.replace("\\", "/").split("/") if p]
    # Optional leading store (account) name; otherwise the default store.
    store = ns.DefaultStore
    for st in ns.Stores:
        if parts and st.DisplayName.lower() == parts[0].lower():
            store, parts = st, parts[1:]
            break
    if not parts:
        raise ToolError(f"Give a folder after the account name, e.g. '{folder.rstrip('/')}/inbox'.")
    # Well-known folders by ID, not name: names are localized ('收件箱' is the inbox).
    if parts[0].lower() in DEFAULT_FOLDERS:
        current, parts = store.GetDefaultFolder(DEFAULT_FOLDERS[parts[0].lower()]), parts[1:]
    else:
        current = store.GetRootFolder()
    for part in parts:
        match = next((f for f in current.Folders if f.Name.lower() == part.lower()), None)
        if match is None:
            raise ToolError(f"Folder not found: {part!r} under {current.Name!r}")
        current = match
    return current


def _sender_email(item) -> str:
    try:
        if item.SenderEmailType == "EX":  # Exchange users have an X500 address
            user = item.Sender.GetExchangeUser()
            if user is not None:
                return user.PrimarySmtpAddress
        return item.SenderEmailAddress or ""
    except Exception:
        return ""


def _local_dt(com_time) -> datetime:
    # pywin32 tags Outlook's local times as UTC; strip tzinfo to compare with now().
    return datetime(com_time.year, com_time.month, com_time.day,
                    com_time.hour, com_time.minute, com_time.second)


def _summarize(item) -> dict:
    body = (item.Body or "").strip()
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "",
        "sender_name": item.SenderName or "",
        "sender_email": _sender_email(item),
        "received": _local_dt(item.ReceivedTime),
        "preview": " ".join(body[:PREVIEW_CHARS].split()),
        "unread": bool(item.UnRead),
    }


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

@mcp.tool(annotations=READ_ONLY)
def list_folders(max_depth: int = 2) -> list[str]:
    """List mail folders (as paths usable by search_emails' `folder` argument)."""
    ns = _namespace()
    paths: list[str] = []

    def walk(folder, prefix: str, depth: int):
        for sub in folder.Folders:
            path = f"{prefix}/{sub.Name}"
            paths.append(path)
            if depth < max_depth:
                walk(sub, path, depth + 1)

    for store in ns.Folders:
        walk(store, store.Name, 1)
    return paths


@mcp.tool(annotations=READ_ONLY)
def search_emails(
    query: str = "",
    folder: str = "inbox",
    days: int = 7,
    limit: int = 20,
    unread_only: bool = False,
) -> list[dict]:
    """Search recent emails in a folder, newest first.

    query: case-insensitive keyword matched against subject, sender and preview
           (empty = everything). folder: 'inbox' (default account), 'sent',
           'you@other-account.com/inbox' for another account, or a path from
           list_folders. days: how far back to look.
           Returns summaries; use read_email(entry_id) for the full body.
    """
    ns = _namespace()
    items = _resolve_folder(ns, folder).Items
    items.Sort("[ReceivedTime]", True)  # newest first

    cutoff = datetime.now() - timedelta(days=days)
    needle = query.lower().strip()
    results: list[dict] = []

    item = items.GetFirst()
    while item is not None and len(results) < limit:
        if item.Class == OL_MAIL:
            email = _summarize(item)
            if email["received"] < cutoff:
                break  # sorted, so everything after is older
            haystack = f'{email["subject"]} {email["sender_name"]} {email["sender_email"]} {email["preview"]}'.lower()
            if ((not needle or needle in haystack)
                    and (not unread_only or email["unread"])
                    and is_relevant(email)):
                email["received"] = email["received"].isoformat(timespec="minutes")
                results.append(email)
        item = items.GetNext()

    log.info("search_emails folder=%s query=%r -> %d results", folder, query, len(results))
    return results


@mcp.tool(annotations=READ_ONLY)
def read_email(entry_id: str, max_chars: int = 8000) -> dict:
    """Read one email's full content by entry_id (from search_emails)."""
    ns = _namespace()
    item = ns.GetItemFromID(entry_id)
    email = _summarize(item)
    email["received"] = email["received"].isoformat(timespec="minutes")
    body = (item.Body or "").strip()
    email.pop("preview")
    email.update({
        "to": item.To or "",
        "cc": item.CC or "",
        "bcc": item.BCC or "",  # only populated on mail you sent
        "body": body[:max_chars],
        "truncated": len(body) > max_chars,
        "attachments": [item.Attachments.Item(i + 1).FileName
                        for i in range(item.Attachments.Count)],
    })
    return email


@mcp.tool(annotations=READ_ONLY)
def list_accounts() -> list[str]:
    """List the email addresses of the accounts configured in Outlook."""
    return [acc.SmtpAddress for acc in _namespace().Accounts]


# ---------------------------------------------------------------------------
# Compose helpers (shared by send / reply / forward)
# ---------------------------------------------------------------------------

def _clean(addrs: list[str] | None) -> list[str]:
    return [a.strip() for a in (addrs or []) if a and a.strip()]


def _check_attachments(paths: list[str] | None) -> list[str]:
    resolved = [os.path.abspath(os.path.expanduser(p)) for p in _clean(paths)]
    missing = [p for p in resolved if not os.path.isfile(p)]
    if missing:
        raise ToolError(f"Attachment file(s) not found: {missing}")
    return resolved


def _set_account(ns, mail, from_account: str):
    account = next((a for a in ns.Accounts if a.SmtpAddress.lower() == from_account.lower()), None)
    if account is None:
        raise ToolError(f"No Outlook account {from_account!r}. Use list_accounts().")
    try:
        account.DeliveryStore.DisplayName
    except pythoncom.com_error:
        # Without a mailbox, Send() "succeeds" but the mail goes nowhere.
        raise ToolError(f"Account {from_account!r} has no mailbox loaded in Outlook; nothing was sent.")
    # pywin32 can't assign object-valued properties with '='; 64209 is the
    # SendUsingAccount DISPID, 8 = DISPATCH_PROPERTYPUTREF.
    mail._oleobj_.Invoke(64209, 0, 8, 0, account)
    if mail.SendUsingAccount.SmtpAddress.lower() != from_account.lower():
        raise ToolError("Could not set the sending account; nothing was sent.")


def _prepend_text(mail, text: str):
    """Put new text above the quoted original instead of replacing the body."""
    if not text:
        return
    if mail.BodyFormat == OL_FORMAT_PLAIN:
        mail.Body = text + "\r\n\r\n" + (mail.Body or "")
        return
    block = ('<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">'
             + html.escape(text).replace("\n", "<br>") + "</div><br>")
    body = mail.HTMLBody or ""
    tag = re.search(r"<body[^>]*>", body, re.IGNORECASE)
    mail.HTMLBody = body[:tag.end()] + block + body[tag.end():] if tag else block + body


def _compose(ns, mail, body, to, cc, bcc, attachments, from_account) -> None:
    files = _check_attachments(attachments)  # validate before touching the item
    if from_account:
        _set_account(ns, mail, from_account)
    for addrs, kind in ((to, OL_TO), (cc, OL_CC), (bcc, OL_BCC)):
        for addr in _clean(addrs):
            mail.Recipients.Add(addr).Type = kind
    for path in files:
        try:
            mail.Attachments.Add(path)
        except pythoncom.com_error as e:
            raise ToolError(f"Outlook could not attach {path!r}: {e.excepinfo[2] if e.excepinfo else e}")
    _prepend_text(mail, body)


def _addresses_of(mail, kind: int) -> list[str]:
    """SMTP addresses of one recipient type (mail.To etc. only hold display names)."""
    out = []
    for i in range(mail.Recipients.Count):
        r = mail.Recipients.Item(i + 1)
        if r.Type != kind:
            continue
        addr = r.Address or r.Name
        try:
            if r.AddressEntry is not None and r.AddressEntry.Type == "EX":
                user = r.AddressEntry.GetExchangeUser()
                addr = user.PrimarySmtpAddress if user else addr
        except pythoncom.com_error:
            pass
        out.append(addr)
    return out


def _finish(mail, action: str, send: bool) -> dict:
    if mail.Recipients.Count == 0:
        raise ToolError("No recipients given; nothing was sent.")
    if not mail.Recipients.ResolveAll():
        bad = [mail.Recipients.Item(i + 1).Name for i in range(mail.Recipients.Count)
               if not mail.Recipients.Item(i + 1).Resolved]
        raise ToolError(f"Could not resolve recipient(s) {bad}; nothing was sent.")
    # Capture details first: after Send() the item is gone and can't be read.
    info = {
        "action": action,
        "from": mail.SendUsingAccount.SmtpAddress if mail.SendUsingAccount else "(account of original / default)",
        "to": _addresses_of(mail, OL_TO),
        "cc": _addresses_of(mail, OL_CC),
        "bcc": _addresses_of(mail, OL_BCC),
        "subject": mail.Subject or "",
        "attachments": [mail.Attachments.Item(i + 1).FileName for i in range(mail.Attachments.Count)],
    }
    if send:
        mail.Send()
        info["status"] = "sent"
    else:
        mail.Save()
        info.update(status="draft", entry_id=mail.EntryID)
    log.info("%s %s to=%s cc=%s bcc=%s", action, info["status"], info["to"], info["cc"], info["bcc"])
    return info


def _original(ns, entry_id: str):
    try:
        return ns.GetItemFromID(entry_id)
    except pythoncom.com_error:
        raise ToolError(f"No email with entry_id {entry_id!r}. Get one from search_emails.")


# ---------------------------------------------------------------------------
# Sending tools
# ---------------------------------------------------------------------------

@mcp.tool(annotations=SENDS_MAIL)
def send_email(
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    attachments: list[str] | None = None,
    from_account: str = "",
    send: bool = False,
) -> dict:
    """Write a NEW email (not a reply). Saves a draft by default.

    to / cc / bcc: recipient email addresses.
    attachments: absolute file paths on this computer.
    from_account: one of list_accounts(); empty = Outlook's default account.
    send: False saves to Drafts for the user to review; True sends immediately.
    To answer an existing email use reply_email; to pass one on use forward_email.
    """
    ns = _namespace()
    mail = ns.Application.CreateItem(OL_MAIL_ITEM)
    mail.Subject = subject
    _compose(ns, mail, body, to, cc, bcc, attachments, from_account)
    return _finish(mail, "new", send)


@mcp.tool(annotations=SENDS_MAIL)
def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    attachments: list[str] | None = None,
    send: bool = False,
) -> dict:
    """Reply inside the original email's thread (keeps quoted history). Draft by default.

    entry_id: the email to reply to (from search_emails / read_email).
    body: your reply text; it is placed above the quoted original.
    reply_all: True replies to the sender and everyone on To/CC.
    cc / bcc: extra recipients added on top of the reply's own.
    attachments: absolute file paths. The reply is sent from the account that
    received the original. send: False = save draft, True = send now.
    """
    ns = _namespace()
    original = _original(ns, entry_id)
    mail = original.ReplyAll() if reply_all else original.Reply()
    _compose(ns, mail, body, None, cc, bcc, attachments, "")
    return _finish(mail, "reply_all" if reply_all else "reply", send)


@mcp.tool(annotations=SENDS_MAIL)
def forward_email(
    entry_id: str,
    to: list[str],
    body: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    attachments: list[str] | None = None,
    send: bool = False,
) -> dict:
    """Forward an email (original content and attachments included). Draft by default.

    entry_id: the email to forward (from search_emails / read_email).
    to / cc / bcc: who to forward it to. body: optional note above the original.
    attachments: extra absolute file paths to add. send: False = draft, True = send now.
    """
    ns = _namespace()
    mail = _original(ns, entry_id).Forward()
    _compose(ns, mail, body, to, cc, bcc, attachments, "")
    return _finish(mail, "forward", send)


@mcp.tool(annotations=SENDS_MAIL)
def send_draft(entry_id: str) -> dict:
    """Send a draft created earlier by send_email / reply_email / forward_email."""
    ns = _namespace()
    mail = _original(ns, entry_id)
    if mail.Sent:
        raise ToolError("That item has already been sent; it is not a draft.")
    return _finish(mail, "send_draft", True)


# ---------------------------------------------------------------------------
# Mailbox management tools
# ---------------------------------------------------------------------------

def _items(ns, entry_ids: list[str]):
    ids = _clean(entry_ids)
    if not ids:
        raise ToolError("No entry_ids given.")
    return [(eid, _original(ns, eid)) for eid in ids]


@mcp.tool(annotations=ORGANIZES)
def mark_emails(entry_ids: list[str], read: bool | None = None, flagged: bool | None = None) -> list[dict]:
    """Mark emails read/unread and/or flag/unflag them for follow-up.

    entry_ids: emails from search_emails. read: True = read, False = unread,
    omit to leave as is. flagged: True = flag, False = clear flag, omit to leave.
    """
    if read is None and flagged is None:
        raise ToolError("Nothing to change: pass read and/or flagged.")
    ns = _namespace()
    results = []
    for eid, item in _items(ns, entry_ids):
        if read is not None:
            item.UnRead = not read
        if flagged is not None:
            item.FlagStatus = OL_FLAG_MARKED if flagged else OL_NO_FLAG
        item.Save()
        results.append({"entry_id": eid, "subject": item.Subject or "",
                        "unread": bool(item.UnRead), "flagged": item.FlagStatus == OL_FLAG_MARKED})
    return results


@mcp.tool(annotations=ORGANIZES)
def create_folder(parent_folder: str, name: str) -> dict:
    """Create a mail folder, e.g. parent_folder='you@school.edu/inbox', name='Courses'.
    Returns its path for use with move_emails / search_emails. Existing folder is reused."""
    ns = _namespace()
    parent = _resolve_folder(ns, parent_folder)
    existing = next((f for f in parent.Folders if f.Name.lower() == name.lower()), None)
    folder = existing or parent.Folders.Add(name)
    return {"path": f"{parent_folder.rstrip('/')}/{folder.Name}", "created": existing is None}


@mcp.tool(annotations=ORGANIZES)
def move_emails(entry_ids: list[str], folder: str) -> list[dict]:
    """Move emails to another folder (a path like those from list_folders).
    Moving changes an email's entry_id; the new one is returned."""
    ns = _namespace()
    target = _resolve_folder(ns, folder)
    results = []
    for eid, item in _items(ns, entry_ids):
        moved = item.Move(target)
        results.append({"old_entry_id": eid, "entry_id": moved.EntryID,
                        "subject": moved.Subject or "", "folder": target.FolderPath})
    return results


@mcp.tool(annotations=DELETES)
def delete_emails(entry_ids: list[str]) -> list[dict]:
    """Delete emails by moving them to that account's Deleted Items (recoverable).

    This tool never deletes permanently; the user can restore from Deleted Items.
    """
    ns = _namespace()
    results = []
    for eid, item in _items(ns, entry_ids):
        trash = item.Parent.Store.GetDefaultFolder(DEFAULT_FOLDERS["deleted"])
        if item.Parent.EntryID == trash.EntryID:
            raise ToolError(f"{item.Subject!r} is already in Deleted Items; this tool won't delete permanently.")
        moved = item.Move(trash)  # explicit move: Delete() can purge on some store types
        results.append({"old_entry_id": eid, "entry_id": moved.EntryID,
                        "subject": moved.Subject or "", "folder": trash.FolderPath})
    return results


def _unique_path(directory: str, filename: str) -> str:
    stem, ext = os.path.splitext(re.sub(r'[<>:"/\\|?*]', "_", filename) or "attachment")
    path, n = os.path.join(directory, stem + ext), 1
    while os.path.exists(path):
        path, n = os.path.join(directory, f"{stem} ({n}){ext}"), n + 1
    return path


@mcp.tool(annotations=ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=False))
def save_attachments(entry_id: str, directory: str = "", names: list[str] | None = None) -> list[str]:
    """Save an email's attachments to disk and return the saved file paths.

    directory: where to save (default: Downloads/outlook-attachments).
    names: only these attachment file names (default: all). Existing files are
    never overwritten; a ' (1)' suffix is added instead.
    """
    ns = _namespace()
    item = _original(ns, entry_id)
    directory = os.path.abspath(os.path.expanduser(directory or "~/Downloads/outlook-attachments"))
    os.makedirs(directory, exist_ok=True)
    wanted = {n.lower() for n in _clean(names)}
    saved = []
    for i in range(item.Attachments.Count):
        att = item.Attachments.Item(i + 1)
        if wanted and att.FileName.lower() not in wanted:
            continue
        path = _unique_path(directory, att.FileName)
        att.SaveAsFile(path)
        saved.append(path)
    if not saved:
        raise ToolError("No matching attachments on that email.")
    return saved


if __name__ == "__main__":
    mcp.run()  # stdio transport
