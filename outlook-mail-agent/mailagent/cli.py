from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler

from .agent import Agent
from .sources import clean_subject
from .config import load_config
from .report import build_stats, period_for, report_email
from . import tracker


def weekly_report(agent: Agent, ref: datetime | None, send: bool = True) -> None:
    cfg = agent.cfg
    start, end = period_for(agent.now, ref)
    st = build_stats(tracker.load_records(cfg["_db"]), start, end)
    rcpt = cfg["weekly_report"]["recipients"]
    mail = report_email(st, start, end, cfg["target_address"], rcpt)
    path = cfg["_reports"] / f"reporting_{start:%G-S%V}.html"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(f"<h3>{mail.subject}</h3>{mail.html}", encoding="utf-8")
    logging.info("Rapport écrit : %s", path)
    if send and rcpt:
        agent.get_notifier().send(mail)
        logging.info("Rapport envoyé à %s", rcpt)
    elif send:
        logging.warning("weekly_report.recipients vide : rapport généré mais non envoyé.")


def set_status(cfg: dict, args: list[str]) -> int:
    if len(args) < 2 or args[1] not in tracker.STATUSES:
        print(f"Usage : set-status ID STATUT [NOTES]   STATUT parmi {tracker.STATUSES}")
        return 2
    recs = tracker.load_records(cfg["_db"])
    if args[0] not in recs:
        print(f"ID inconnu : {args[0]}")
        return 2
    r = recs[args[0]]
    r["status"] = args[1]
    if args[1] in tracker.CLOSED and not r["replied"]:
        r["replied"] = datetime.now()
    if len(args) > 2:
        r["notes"] = " ".join(args[2:])
    ag = Agent(cfg, source=object())  # pas de relève : on réécrit juste la base et les CSV
    tracker.save_records(cfg["_db"], recs)
    tracker.export_powerbi(cfg["_powerbi"], recs, cfg, clean_subject, ag.now)
    print(f"{r['id']} -> {args[1]}")
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="mailagent", description="Agent de suivi des demandes Outlook")
    p.add_argument("command", choices=["run", "weekly-report", "set-status"],
                   help="run = relève + base + CSV Power BI + alertes ; weekly-report = run puis reporting ; "
                        "set-status ID STATUT [NOTES] = changer le statut d'une demande")
    p.add_argument("args", nargs="*", help="(set-status) ID, STATUT (Nouveau|En cours|Répondu|Clôturé), NOTES")
    p.add_argument("-c", "--config", default="config.yaml")
    p.add_argument("--date", help="(weekly-report) rapporter les 7 jours précédant cette date AAAA-MM-JJ")
    p.add_argument("--no-send", action="store_true", help="ne pas envoyer le rapport par mail")
    a = p.parse_args(argv)

    cfg = load_config(a.config)
    cfg["_out"].mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=[
        logging.StreamHandler(sys.stdout),
        RotatingFileHandler(cfg["_out"] / "agent.log", maxBytes=1_000_000, backupCount=5, encoding="utf-8")])
    try:
        if a.command == "set-status":
            return set_status(cfg, a.args)
        agent = Agent(cfg)
        agent.run()
        if a.command == "weekly-report":
            weekly_report(agent, datetime.fromisoformat(a.date) if a.date else None, not a.no_send)
        return 0
    except Exception:
        logging.exception("Échec de l'agent")
        return 1
