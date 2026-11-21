from __future__ import annotations

import hashlib
import json
import logging
from datetime import date, datetime, timedelta
from html import escape

from . import tracker
from .classify import categorize, is_ignored
from .models import Outgoing
from .notifier import make_notifier
from .sources import clean_subject, make_source, matches_target
from .summarize import Summarizer

log = logging.getLogger(__name__)


def elapsed_days(start: datetime, end: datetime, business: bool) -> int:
    d0, d1 = start.date(), end.date()
    if not business:
        return max((d1 - d0).days, 0)
    n, d = 0, d0
    while d < d1:
        d += timedelta(days=1)
        n += d.weekday() < 5
    return n


def alert_level(rec: dict, now: datetime, cfg: dict) -> str:
    if rec["status"] in tracker.CLOSED:
        return ""
    if rec["priority"] == "Critique":
        return "Critique"
    if rec["age"] >= cfg["sla"]["days"]:
        return "SLA dépassé"
    dl = rec.get("deadline")
    if dl:
        if dl.date() < now.date():
            return "Échéance dépassée"
        if (dl.date() - now.date()).days <= cfg["sla"]["deadline_warning_days"]:
            return "Échéance proche"
    return ""


def _table(rows: list[dict]) -> str:
    th = "".join(f"<th style='border:1px solid #ccc;padding:4px 8px;background:#1F3864;color:#fff'>{h}</th>"
                 for h in ("Reçu le", "Expéditeur", "Objet", "Résumé", "Âge (j)"))
    body = ""
    for r in rows:
        cells = (r["received"].strftime("%d/%m %H:%M"), r["sender"], r["subject"], r["summary"], r["age"])
        body += "<tr>" + "".join(f"<td style='border:1px solid #ccc;padding:4px 8px'>{escape(str(c))}</td>"
                                 for c in cells) + "</tr>"
    return f"<table style='border-collapse:collapse;font-family:Calibri,sans-serif;font-size:13px'><tr>{th}</tr>{body}</table>"


class Agent:
    def __init__(self, cfg: dict, source=None, notifier=None, now: datetime | None = None):
        self.cfg = cfg
        self.now = now or datetime.now()
        self.source = source or make_source(cfg)
        self.notifier = notifier
        self._notifier_ready = notifier is not None
        self.summarizer = Summarizer(cfg)
        self.state_path = cfg["_out"] / "state.json"
        self.state = json.loads(self.state_path.read_text()) if self.state_path.exists() else {"alerted": {}}

    # -- utilitaires
    def get_notifier(self):
        if not self._notifier_ready:
            self.notifier, self._notifier_ready = make_notifier(self.cfg), True
        return self.notifier

    def _mark(self, key: str) -> None:
        self.state["alerted"][key] = self.now.strftime(tracker.FMT)
        self.cfg["_out"].mkdir(parents=True, exist_ok=True)
        self.state_path.write_text(json.dumps(self.state, indent=1))

    def _send(self, to: list[str], subject: str, html: str) -> bool:
        if not to:
            log.warning("Aucun destinataire d'alerte configuré (alerts.recipients).")
            return False
        try:
            self.get_notifier().send(Outgoing(to, subject, html))
            return True
        except Exception:
            log.exception("Échec d'envoi : %s", subject)
            return False

    # -- cycle principal
    def run(self) -> dict:
        cfg, now = self.cfg, self.now
        records = tracker.load_records(cfg["_db"])
        for rid, ts in self.state["alerted"].items():
            if rid in records and not rid.startswith("crit:"):
                records[rid]["alert_sent"] = records[rid].get("alert_sent") or tracker._to_dt(ts)
        known_conv = {r["conv"] for r in records.values()}
        since = now - timedelta(days=cfg["source"]["lookback_days"])

        # 1) nouveaux mails -> une demande par conversation
        mails = sorted(self.source.fetch(since), key=lambda m: m.received)
        new = 0
        for m in mails:
            if m.conversation_id in known_conv or is_ignored(m, cfg):
                continue
            if m.sender == cfg["target_address"].lower() or not matches_target(m, cfg["target_address"]):
                continue
            a = self.summarizer.analyze(m, now)
            rid = hashlib.sha1(m.conversation_id.encode()).hexdigest()[:10]
            records[rid] = {
                "id": rid, "received": m.received, "sender": m.sender, "subject": m.subject,
                "category": categorize(m.subject, cfg), "summary": a.summary, "action": a.action,
                "priority": "Critique" if a.critical else a.priority, "deadline": a.deadline,
                "status": "Nouveau", "replied": None, "age": 0, "alert": "", "alert_sent": None,
                "notes": None, "conv": m.conversation_id}
            known_conv.add(m.conversation_id)
            new += 1
            if a.critical and cfg["alerts"]["immediate_on_critical"] and f"crit:{rid}" not in self.state["alerted"]:
                html = (f"<p><b>Demande critique</b> reçue sur {escape(cfg['target_address'])} :</p>"
                        f"<p><b>{escape(m.subject)}</b><br>De : {escape(m.sender)}<br>{escape(a.summary)}</p>")
                if self._send(cfg["alerts"]["recipients"], f"[CRITIQUE] {m.subject}", html):
                    self._mark(f"crit:{rid}")

        # 2) détection des réponses (courriers envoyés dans la même conversation)
        open_recs = [r for r in records.values() if r["status"] in ("Nouveau", "En cours")]
        if open_recs:
            sent = self.source.sent_after(min(r["received"] for r in open_recs))
            for r in open_recs:
                after = sorted(t for t in sent.get(r["conv"], []) if t >= r["received"])
                if after:
                    r["status"], r["replied"] = "Répondu", after[0]

        # 3) âge, alertes
        for r in records.values():
            end = r["replied"] if r["status"] in tracker.CLOSED and r["replied"] else now
            r["age"] = elapsed_days(r["received"], end, cfg["sla"]["business_days"])
            r["alert"] = alert_level(r, now, cfg)

        # 4) e-mails « sans retour après N jours »
        due = [r for r in records.values()
               if r["status"] in ("Nouveau", "En cours") and r["age"] >= cfg["sla"]["days"]
               and not r["alert_sent"]]
        sent_alerts = self._sla_alerts(due)
        for r in due:
            if r["id"] in sent_alerts:
                r["alert_sent"] = now

        # 5) base SQLite + export CSV pour Power BI
        tracker.save_records(cfg["_db"], records)
        tracker.export_powerbi(cfg["_powerbi"], records, cfg, clean_subject, now)
        stats = {"nouveaux": new, "total": len(records), "alertes_sla_envoyees": len(sent_alerts)}
        log.info("Passage terminé : %s", stats)
        return stats

    def _sla_alerts(self, due: list[dict]) -> set[str]:
        if not due:
            return set()
        cfg, rcpt = self.cfg, self.cfg["alerts"]["recipients"]
        n = cfg["sla"]["days"]
        unit = "jours ouvrés" if cfg["sla"]["business_days"] else "jours"
        done: set[str] = set()
        if cfg["alerts"]["mode"] == "per_mail":
            for r in due:
                html = (f"<p>Aucune réponse après {n} {unit} :</p><p><b>{escape(r['subject'])}</b><br>"
                        f"De : {escape(r['sender'])} - reçu le {r['received']:%d/%m/%Y %H:%M}<br>"
                        f"{escape(r['summary'])}</p>")
                if self._send(rcpt, f"[Sans réponse {n} j] {r['subject']}", html):
                    done.add(r["id"])
        else:
            html = f"<p>{len(due)} demande(s) sans réponse après {n} {unit} :</p>" + _table(due)
            if self._send(rcpt, f"[Sans réponse] {len(due)} demande(s) en attente depuis {n} {unit}", html):
                done = {r["id"] for r in due}
        for rid in done:
            self._mark(rid)
        return done
