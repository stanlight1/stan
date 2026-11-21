"""Stockage (SQLite = source de vérité) et export CSV pour Power BI.

Une ligne = une conversation. Power BI ne lit que les CSV exportés (lecture seule).
"""
from __future__ import annotations

import csv
import sqlite3
import time
from datetime import datetime
from pathlib import Path

STATUSES = ["Nouveau", "En cours", "Répondu", "Clôturé"]
CLOSED = {"Répondu", "Clôturé"}
PRIORITIES = ["Critique", "Haute", "Normale", "Basse"]
FMT = "%Y-%m-%d %H:%M:%S"

KEYS = ["id", "received", "sender", "subject", "category", "summary", "action", "priority",
        "deadline", "status", "replied", "age", "alert", "alert_sent", "notes", "conv"]
DATE_KEYS = {"received", "deadline", "replied", "alert_sent"}


def _to_dt(v):
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v
    try:
        return datetime.fromisoformat(str(v))
    except ValueError:
        return None


def _connect(db: Path) -> sqlite3.Connection:
    db.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(db, timeout=30)
    cols = ", ".join(f"{k} {'INTEGER' if k == 'age' else 'TEXT'}" + (" PRIMARY KEY" if k == "id" else "")
                     for k in KEYS)
    con.execute(f"CREATE TABLE IF NOT EXISTS demandes ({cols})")
    return con


def load_records(db: Path) -> dict[str, dict]:
    if not db.exists():
        return {}
    con = _connect(db)
    try:
        recs = {}
        for row in con.execute(f"SELECT {', '.join(KEYS)} FROM demandes"):
            r = dict(zip(KEYS, row))
            for k in DATE_KEYS:
                r[k] = _to_dt(r[k])
            r["age"] = r["age"] or 0
            recs[r["id"]] = r
        return recs
    finally:
        con.close()


def save_records(db: Path, records: dict[str, dict]) -> None:
    con = _connect(db)
    try:
        with con:  # transaction atomique
            con.execute("DELETE FROM demandes")
            con.executemany(
                f"INSERT INTO demandes VALUES ({','.join('?' * len(KEYS))})",
                [[r[k].isoformat(sep=' ', timespec='seconds') if k in DATE_KEYS and r.get(k) else r.get(k)
                  for k in KEYS] for r in records.values()])
    finally:
        con.close()


def _atomic_write(path: Path, write) -> None:
    """Écrit un fichier via un temporaire ; réessaie si un autre programme le verrouille."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", newline="", encoding="utf-8-sig") as f:  # BOM : accents corrects dans Power BI/Excel
        write(f)
    for attempt in range(5):
        try:
            tmp.replace(path)
            return
        except PermissionError:
            time.sleep(2 * (attempt + 1))
    raise PermissionError(f"{path.name} est verrouillé par un autre programme.")


def export_powerbi(folder: Path, records: dict[str, dict], cfg: dict, clean_subject, now: datetime) -> None:
    """Demandes.csv (table de faits) + Parametres.csv (SLA, date de mise à jour)."""
    head = ["ID", "Recu_le", "Expediteur", "Domaine_expediteur", "Objet", "Objet_nettoye", "Categorie",
            "Resume", "Action_attendue", "Priorite", "Echeance", "Statut", "Ouverte", "Repondu_le",
            "Age_jours", "Delai_reponse_jours", "Alerte", "Alerte_envoyee_le", "Notes"]
    d = lambda v, f=FMT: v.strftime(f) if v else ""  # noqa: E731
    rows = sorted(records.values(), key=lambda r: r["received"], reverse=True)

    def write(f):
        w = csv.writer(f)
        w.writerow(head)
        for r in rows:
            w.writerow([r["id"], d(r["received"]), r["sender"], r["sender"].split("@")[-1],
                        r["subject"], clean_subject(r["subject"]), r["category"],
                        r["summary"], r["action"] or "", r["priority"], d(r["deadline"], "%Y-%m-%d"),
                        r["status"], 0 if r["status"] in CLOSED else 1, d(r["replied"]),
                        r["age"], r["age"] if r["status"] == "Répondu" else "",
                        r["alert"] or "", d(r["alert_sent"]), r["notes"] or ""])
    _atomic_write(folder / "Demandes.csv", write)

    def write_p(f):
        w = csv.writer(f)
        w.writerow(["Parametre", "Valeur"])
        w.writerows([["SLA_jours", cfg["sla"]["days"]],
                     ["Jours_ouvres", int(cfg["sla"]["business_days"])],
                     ["Adresse_suivie", cfg["target_address"]],
                     ["Derniere_maj", d(now)]])
    _atomic_write(folder / "Parametres.csv", write_p)
