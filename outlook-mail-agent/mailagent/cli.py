from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler

from .agent import Agent
from .config import load_config
from .report import build_stats, period_for, report_email, write_report_xlsx
from . import tracker


def weekly_report(agent: Agent, ref: datetime | None, send: bool = True) -> None:
    cfg = agent.cfg
    start, end = period_for(agent.now, ref)
    st = build_stats(tracker.load_records(cfg["_tracking"]), start, end)
    path = cfg["_reports"] / f"reporting_{start:%G-S%V}.xlsx"
    write_report_xlsx(path, st, start, end, cfg["target_address"])
    logging.info("Rapport écrit : %s", path)
    rcpt = cfg["weekly_report"]["recipients"]
    if send and rcpt:
        agent.get_notifier().send(report_email(st, start, end, cfg["target_address"], path, rcpt))
        logging.info("Rapport envoyé à %s", rcpt)
    elif send:
        logging.warning("weekly_report.recipients vide : rapport généré mais non envoyé.")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="mailagent", description="Agent de suivi des demandes Outlook")
    p.add_argument("command", choices=["run", "weekly-report"],
                   help="run = relève + Excel + alertes ; weekly-report = run puis reporting")
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
        agent = Agent(cfg)
        agent.run()
        if a.command == "weekly-report":
            weekly_report(agent, datetime.fromisoformat(a.date) if a.date else None, not a.no_send)
        return 0
    except Exception:
        logging.exception("Échec de l'agent")
        return 1
