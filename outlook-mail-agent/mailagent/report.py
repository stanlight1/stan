from __future__ import annotations

from collections import Counter
from datetime import datetime, timedelta
from html import escape
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font

from . import tracker
from .models import Outgoing


def period_for(now: datetime, ref: datetime | None = None) -> tuple[datetime, datetime]:
    """Semaine calendaire précédente (lundi 00:00 -> lundi suivant 00:00), ou 7 jours avant `ref`."""
    end = (ref or now).replace(hour=0, minute=0, second=0, microsecond=0)
    if ref is None:
        end -= timedelta(days=end.weekday())  # lundi de la semaine en cours
    return end - timedelta(days=7), end


def build_stats(records: dict[str, dict], start: datetime, end: datetime) -> dict:
    inp = [r for r in records.values() if start <= r["received"] < end]
    replied = [r for r in inp if r["status"] == "Répondu"]
    openr = [r for r in records.values() if r["status"] in ("Nouveau", "En cours")]
    return {
        "in_period": sorted(inp, key=lambda r: (r["category"], r["received"])),
        "by_cat": Counter(r["category"] for r in inp),
        "by_status": Counter(r["status"] for r in inp),
        "by_prio": Counter(r["priority"] for r in inp),
        "senders": Counter(r["sender"] for r in inp).most_common(5),
        "avg_resp": (sum(r["age"] for r in replied) / len(replied)) if replied else None,
        "open_late": sorted([r for r in openr if r["alert"]], key=lambda r: -r["age"]),
        "open_total": len(openr),
    }


def write_report_xlsx(path: Path, st: dict, start: datetime, end: datetime, target: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    ws["A1"] = f"Reporting hebdomadaire - {target}"
    ws["A1"].font = Font(bold=True, size=15)
    ws["A2"] = f"Période : {start:%d/%m/%Y} au {(end - timedelta(seconds=1)):%d/%m/%Y}"
    rows = [("Demandes reçues", len(st["in_period"])),
            ("Répondues (sur la période)", st["by_status"].get("Répondu", 0)),
            ("Clôturées", st["by_status"].get("Clôturé", 0)),
            ("Délai moyen de réponse (j)", round(st["avg_resp"], 1) if st["avg_resp"] is not None else "-"),
            ("Demandes ouvertes (total)", st["open_total"]),
            ("Dont en alerte", len(st["open_late"]))]
    for i, (k, v) in enumerate(rows, 4):
        ws.cell(i, 1, k).font = Font(bold=True)
        ws.cell(i, 2, v)
    r0 = 12
    for j, h in enumerate(("Catégorie", "Demandes"), 1):
        ws.cell(r0, j, h).font = Font(bold=True)
    for i, (k, v) in enumerate(sorted(st["by_cat"].items(), key=lambda x: -x[1]), r0 + 1):
        ws.cell(i, 1, k), ws.cell(i, 2, v)
    n = len(st["by_cat"])
    if n:
        ch = BarChart()
        ch.title, ch.height, ch.width = "Demandes par catégorie", 7, 14
        ch.add_data(Reference(ws, min_col=2, min_row=r0, max_row=r0 + n), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=r0 + 1, max_row=r0 + n))
        ws.add_chart(ch, "D4")
    ws.column_dimensions["A"].width = 36

    def sheet(title: str, data: list[dict]):
        s = wb.create_sheet(title)
        keys = ["received", "sender", "subject", "category", "summary", "priority", "status", "age", "alert"]
        t = {k: h for k, h, _ in tracker.COLUMNS}
        s.append([t[k] for k in keys])
        for r in data:
            s.append([r.get(k) for k in keys])
        for c, w in zip("ABCDEFGHI", (17, 28, 40, 18, 60, 11, 11, 8, 18)):
            s.column_dimensions[c].width = w
        for c in s[1]:
            c.font = Font(bold=True)
        for row in s.iter_rows(min_row=2, max_col=1):
            row[0].number_format = tracker.FMT

    sheet("Demandes de la période", st["in_period"])
    sheet("Ouvertes en alerte", st["open_late"])
    tracker._write_atomic(wb, path)


def report_email(st: dict, start: datetime, end: datetime, target: str, attachment: Path, to: list[str]) -> Outgoing:
    def tbl(counter) -> str:
        return "<table>" + "".join(f"<tr><td>{escape(str(k))}</td><td><b>{v}</b></td></tr>"
                                   for k, v in counter) + "</table>"
    late = "".join(f"<li><b>{escape(r['subject'])}</b> ({escape(r['sender'])}, {r['age']} j) - {escape(r['alert'])}</li>"
                   for r in st["open_late"][:15]) or "<li>Aucune</li>"
    avg = f"{st['avg_resp']:.1f} j" if st["avg_resp"] is not None else "n/a"
    html = (f"<p>Semaine du <b>{start:%d/%m/%Y}</b> au <b>{(end - timedelta(seconds=1)):%d/%m/%Y}</b> "
            f"- adresse <b>{escape(target)}</b></p>"
            f"<p><b>{len(st['in_period'])}</b> demande(s) reçue(s) - délai moyen de réponse : {avg} - "
            f"{st['open_total']} ouverte(s) au total dont {len(st['open_late'])} en alerte.</p>"
            f"<h4>Par catégorie</h4>{tbl(sorted(st['by_cat'].items(), key=lambda x: -x[1]))}"
            f"<h4>Par statut</h4>{tbl(st['by_status'].items())}"
            f"<h4>Principaux expéditeurs</h4>{tbl(st['senders'])}"
            f"<h4>Demandes ouvertes en alerte</h4><ul>{late}</ul>"
            f"<p>Détail dans le fichier joint.</p>")
    return Outgoing(to, f"Reporting hebdo {target} - semaine du {start:%d/%m/%Y}", html, [str(attachment)])
