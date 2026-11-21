"""Sources de mails : Outlook COM (Windows), Microsoft Graph, dossier JSON (démo/tests).

Chaque source expose :
  fetch(since)        -> list[Mail]            mails reçus depuis `since`
  sent_after(since)   -> dict[conv_id, list[datetime]]   mails ENVOYÉS (réponses)
"""
from __future__ import annotations

import json
import logging
import os
import re
from datetime import datetime, timezone
from pathlib import Path

from .models import Mail

log = logging.getLogger(__name__)
MAX_BODY = 20000


def matches_target(mail: Mail, target: str) -> bool:
    t = target.lower().strip()
    return t in [a.lower() for a in mail.to + mail.cc]


# --------------------------------------------------------------------------- COM
class OutlookComSource:
    """Outlook de bureau classique (win32com). NE fonctionne PAS avec le « nouvel Outlook »."""
    PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

    def __init__(self, cfg: dict):
        import win32com.client  # noqa: WPS433 (import tardif : Windows uniquement)
        self.cfg = cfg["source"]
        self.target = cfg["target_address"]
        self.ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _folder(self, kind: int):
        mailbox = self.cfg.get("mailbox")
        if mailbox:  # boîte partagée / alias géré comme boîte
            rcpt = self.ns.CreateRecipient(mailbox)
            rcpt.Resolve()
            base = self.ns.GetSharedDefaultFolder(rcpt, kind)
        else:
            base = self.ns.GetDefaultFolder(kind)
        if kind == 6 and self.cfg.get("folder", "Inbox") not in ("Inbox", "Boîte de réception"):
            base = base.Folders[self.cfg["folder"]]
        return base

    def _smtp(self, recipient) -> str:
        try:
            return recipient.PropertyAccessor.GetProperty(self.PR_SMTP).lower()
        except Exception:
            return (getattr(recipient, "Address", "") or "").lower()

    def _sender(self, item) -> str:
        if getattr(item, "SenderEmailType", "") == "EX":
            try:
                return item.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
            except Exception:
                pass
        return (item.SenderEmailAddress or "").lower()

    def fetch(self, since: datetime) -> list[Mail]:
        items = self._folder(6).Items
        items.Sort("[ReceivedTime]", True)
        flt = "[ReceivedTime] >= '%s'" % since.strftime("%m/%d/%Y %I:%M %p")
        out: list[Mail] = []
        for it in items.Restrict(flt):
            if getattr(it, "Class", 0) != 43:  # 43 = MailItem
                continue
            to, cc = [], []
            for r in it.Recipients:
                (to if r.Type == 1 else cc if r.Type == 2 else []).append(self._smtp(r))
            out.append(Mail(
                id=it.EntryID, conversation_id=it.ConversationID,
                received=it.ReceivedTime.replace(tzinfo=None),
                sender=self._sender(it), sender_name=it.SenderName or "",
                to=to, cc=cc, subject=it.Subject or "", body=(it.Body or "")[:MAX_BODY],
                has_attachments=it.Attachments.Count > 0,
                importance={0: "low", 1: "normal", 2: "high"}.get(it.Importance, "normal")))
        return out

    def sent_after(self, since: datetime) -> dict[str, list[datetime]]:
        items = self._folder(5).Items
        flt = "[SentOn] >= '%s'" % since.strftime("%m/%d/%Y %I:%M %p")
        res: dict[str, list[datetime]] = {}
        for it in items.Restrict(flt):
            try:
                res.setdefault(it.ConversationID, []).append(it.SentOn.replace(tzinfo=None))
            except Exception:
                continue
        return res


# ------------------------------------------------------------------------- Graph
class GraphSource:
    """Microsoft Graph (permissions d'application Mail.Read + Mail.Send, flux client_credentials)."""
    API = "https://graph.microsoft.com/v1.0"

    def __init__(self, cfg: dict):
        import msal
        import requests
        g = cfg["graph"]
        secret = os.environ.get(g["client_secret_env"], "")
        if not secret:
            raise RuntimeError(f"Variable d'environnement {g['client_secret_env']} absente.")
        self.app = msal.ConfidentialClientApplication(
            g["client_id"], authority=f"https://login.microsoftonline.com/{g['tenant_id']}",
            client_credential=secret)
        self.requests = requests
        self.cfg = cfg["source"]
        self.target = cfg["target_address"]
        self.mailbox = cfg["source"].get("mailbox") or cfg["target_address"]

    def _token(self) -> str:
        r = self.app.acquire_token_for_client(["https://graph.microsoft.com/.default"])
        if "access_token" not in r:
            raise RuntimeError(f"Auth Graph échouée : {r.get('error_description', r)}")
        return r["access_token"]

    def _get_all(self, url: str, params: dict | None = None):
        headers = {"Authorization": f"Bearer {self._token()}",
                   "Prefer": 'outlook.body-content-type="text"'}
        while url:
            r = self.requests.get(url, headers=headers, params=params, timeout=60)
            r.raise_for_status()
            data = r.json()
            yield from data.get("value", [])
            url, params = data.get("@odata.nextLink"), None

    @staticmethod
    def _local(iso: str) -> datetime:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00").split(".")[0] + "+00:00")
        return dt.astimezone().replace(tzinfo=None)

    @staticmethod
    def _utc(dt: datetime) -> str:
        return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def fetch(self, since: datetime) -> list[Mail]:
        folder = self.cfg.get("folder", "Inbox")
        folder = "inbox" if folder in ("Inbox", "Boîte de réception") else folder
        url = f"{self.API}/users/{self.mailbox}/mailFolders/{folder}/messages"
        params = {"$filter": f"receivedDateTime ge {self._utc(since)}", "$top": 50,
                  "$select": "id,conversationId,receivedDateTime,from,toRecipients,ccRecipients,"
                             "subject,body,hasAttachments,importance"}
        out = []
        for m in self._get_all(url, params):
            f = (m.get("from") or {}).get("emailAddress", {})
            rc = lambda k: [x["emailAddress"]["address"].lower() for x in m.get(k, [])]  # noqa: E731
            out.append(Mail(
                id=m["id"], conversation_id=m["conversationId"],
                received=self._local(m["receivedDateTime"]),
                sender=f.get("address", "").lower(), sender_name=f.get("name", ""),
                to=rc("toRecipients"), cc=rc("ccRecipients"), subject=m.get("subject") or "",
                body=(m.get("body", {}).get("content") or "")[:MAX_BODY],
                has_attachments=bool(m.get("hasAttachments")), importance=m.get("importance", "normal")))
        return out

    def sent_after(self, since: datetime) -> dict[str, list[datetime]]:
        url = f"{self.API}/users/{self.mailbox}/mailFolders/sentitems/messages"
        params = {"$filter": f"sentDateTime ge {self._utc(since)}", "$top": 50,
                  "$select": "conversationId,sentDateTime"}
        res: dict[str, list[datetime]] = {}
        for m in self._get_all(url, params):
            res.setdefault(m["conversationId"], []).append(self._local(m["sentDateTime"]))
        return res


# ------------------------------------------------------------- Dossier JSON (démo)
class JsonFolderSource:
    """Lit des fichiers .json (un mail par fichier) : démo, tests, import depuis autre outil.

    Format : {id, conversation_id, received, sender, sender_name, to, cc, subject, body,
              kind: "received" | "sent"}
    """
    def __init__(self, cfg: dict):
        d = Path(cfg["source"].get("path", "sample_mails"))
        self.dir = d if d.is_absolute() else cfg["_base"] / d

    def _load(self) -> list[dict]:
        return [json.loads(p.read_text(encoding="utf-8")) for p in sorted(self.dir.glob("*.json"))]

    def fetch(self, since: datetime) -> list[Mail]:
        out = []
        for d in self._load():
            if d.get("kind", "received") != "received":
                continue
            rec = datetime.fromisoformat(d["received"])
            if rec >= since:
                out.append(Mail(d["id"], d.get("conversation_id", d["id"]), rec, d["sender"],
                                d.get("sender_name", d["sender"]), d["to"], d.get("cc", []),
                                d["subject"], d.get("body", ""), d.get("has_attachments", False),
                                d.get("importance", "normal")))
        return out

    def sent_after(self, since: datetime) -> dict[str, list[datetime]]:
        res: dict[str, list[datetime]] = {}
        for d in self._load():
            if d.get("kind") == "sent":
                res.setdefault(d["conversation_id"], []).append(datetime.fromisoformat(d["received"]))
        return res


def make_source(cfg: dict):
    t = cfg["source"]["type"]
    return {"outlook_com": OutlookComSource, "graph": GraphSource,
            "json_folder": JsonFolderSource}[t](cfg)


_PREFIX = re.compile(r"^\s*((re|tr|fw|fwd|rv|aw)\s*:\s*)+", re.I)


def clean_subject(s: str) -> str:
    return _PREFIX.sub("", s or "").strip()
