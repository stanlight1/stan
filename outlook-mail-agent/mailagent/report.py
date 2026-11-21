from __future__ import annotations

from collections import Counter
from datetime import datetime, timedelta
from html import escape

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


def report_email(st: dict, start: datetime, end: datetime, target: str, to: list[str]) -> Outgoing:
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
            f"<p>Détail interactif : rapport Power BI (filtre sur la période).</p>")
    return Outgoing(to, f"Reporting hebdo {target} - semaine du {start:%d/%m/%Y}", html)
