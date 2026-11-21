"""Envoi des alertes / rapports : Outlook COM, Graph, SMTP, ou fichier (dry-run)."""
from __future__ import annotations

import base64
import logging
import os
import smtplib
import time
from email.message import EmailMessage
from pathlib import Path

from .models import Outgoing

log = logging.getLogger(__name__)


class FileNotifier:
    """Écrit les mails dans un dossier au lieu de les envoyer (test / validation à blanc)."""
    def __init__(self, cfg: dict):
        self.dir = cfg["_outbox"]
        self.dir.mkdir(parents=True, exist_ok=True)

    def send(self, m: Outgoing) -> None:
        safe = "".join(c if c.isalnum() else "_" for c in m.subject)[:60]
        p = self.dir / f"{int(time.time() * 1000)}_{safe}.html"
        p.write_text(f"<!-- À: {', '.join(m.to)} | PJ: {m.attachments} -->\n<h3>{m.subject}</h3>\n{m.html}",
                     encoding="utf-8")
        log.info("[dry-run] mail écrit : %s", p)


class OutlookComNotifier:
    def __init__(self, cfg: dict):
        import win32com.client
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.sender = cfg["source"].get("mailbox", "")

    def send(self, m: Outgoing) -> None:
        item = self.app.CreateItem(0)
        item.To = ";".join(m.to)
        item.Subject = m.subject
        item.HTMLBody = m.html
        if self.sender:  # envoyer au nom d'une boîte partagée (droit « Envoyer en tant que »)
            item.SentOnBehalfOfName = self.sender
        for a in m.attachments:
            item.Attachments.Add(str(Path(a).resolve()))
        item.Send()


class SmtpNotifier:
    def __init__(self, cfg: dict):
        self.c = cfg["notifier"]["smtp"]

    def send(self, m: Outgoing) -> None:
        msg = EmailMessage()
        msg["From"], msg["To"], msg["Subject"] = self.c["from"], ", ".join(m.to), m.subject
        msg.set_content("Ce message nécessite un client compatible HTML.")
        msg.add_alternative(m.html, subtype="html")
        for a in m.attachments:
            p = Path(a)
            msg.add_attachment(p.read_bytes(), maintype="application", subtype="octet-stream",
                               filename=p.name)
        with smtplib.SMTP(self.c["host"], self.c["port"], timeout=60) as s:
            s.starttls()
            if self.c["user"]:
                s.login(self.c["user"], os.environ[self.c["password_env"]])
            s.send_message(msg)


class GraphNotifier:
    def __init__(self, cfg: dict):
        from .sources import GraphSource
        self.g = GraphSource(cfg)

    def send(self, m: Outgoing) -> None:
        body = {"message": {
            "subject": m.subject, "body": {"contentType": "HTML", "content": m.html},
            "toRecipients": [{"emailAddress": {"address": a}} for a in m.to],
            "attachments": [{"@odata.type": "#microsoft.graph.fileAttachment", "name": Path(a).name,
                             "contentBytes": base64.b64encode(Path(a).read_bytes()).decode()}
                            for a in m.attachments]}}
        r = self.g.requests.post(f"{self.g.API}/users/{self.g.mailbox}/sendMail",
                                 headers={"Authorization": f"Bearer {self.g._token()}"},
                                 json=body, timeout=60)
        r.raise_for_status()


def make_notifier(cfg: dict):
    return {"outlook_com": OutlookComNotifier, "graph": GraphNotifier, "smtp": SmtpNotifier,
            "file": FileNotifier}[cfg["notifier"]["type"]](cfg)
