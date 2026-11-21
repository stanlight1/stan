"""Tableau de bord HTML autonome (aucun script externe, aucune licence) régénéré à chaque passage."""
from __future__ import annotations

from collections import Counter
from datetime import datetime, timedelta
from html import escape as e
from pathlib import Path

from . import tracker

COLORS = {"Critique": "#F8696B", "SLA dépassé": "#F8696B", "Échéance dépassée": "#FFB366",
          "Échéance proche": "#FFEB84"}
CSS = """body{font-family:Segoe UI,Calibri,sans-serif;margin:0;background:#f3f5f9;color:#222}
header{background:#1F3864;color:#fff;padding:14px 24px}header small{opacity:.8}
main{padding:18px 24px;max-width:1400px;margin:auto}.kpis{display:flex;gap:12px;flex-wrap:wrap}
.k{background:#fff;border-radius:8px;padding:12px 18px;min-width:150px;box-shadow:0 1px 3px #0002}
.k b{display:block;font-size:28px}.k.al b{color:#c0392b}.row{display:flex;gap:16px;flex-wrap:wrap;margin-top:16px}
.card{background:#fff;border-radius:8px;padding:14px 18px;box-shadow:0 1px 3px #0002;flex:1;min-width:320px}
h2{font-size:15px;margin:0 0 10px}.bar{display:flex;align-items:center;margin:4px 0;font-size:13px}
.bar span:first-child{width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bar i{display:block;height:14px;background:#2E75B6;border-radius:2px}.bar i.o{background:#ED7D31}
.bar em{margin-left:6px;font-style:normal;color:#555}table{border-collapse:collapse;width:100%;font-size:13px}
th{background:#1F3864;color:#fff;text-align:left;padding:5px 8px;cursor:pointer}td{padding:5px 8px;border-bottom:1px solid #e3e6ec;vertical-align:top}
input{padding:6px 10px;width:320px;margin-bottom:8px;border:1px solid #bbb;border-radius:4px}
details{margin:4px 0}summary{cursor:pointer;font-weight:600}"""
JS = """const q=document.getElementById('q');q.oninput=()=>{const v=q.value.toLowerCase();
document.querySelectorAll('#todo tr.d').forEach(r=>r.style.display=r.textContent.toLowerCase().includes(v)?'':'none')};
document.querySelectorAll('#todo th').forEach((h,i)=>h.onclick=()=>{const b=document.querySelector('#todo tbody'),
rows=[...b.rows];const n=rows.every(r=>!isNaN(parseFloat(r.cells[i].textContent)));
rows.sort((a,c)=>n?parseFloat(c.cells[i].textContent)-parseFloat(a.cells[i].textContent):a.cells[i].textContent.localeCompare(c.cells[i].textContent));
rows.forEach(r=>b.appendChild(r))});"""


def _bars(items: list[tuple[str, int, int | None]]) -> str:
    mx = max([t for _, t, _ in items] + [1])
    out = ""
    for label, tot, opn in items:
        o = f"<div class='bar'><span></span><i class='o' style='width:{opn / mx * 220:.0f}px'></i><em>{opn} ouv.</em></div>" if opn else ""
        out += (f"<div class='bar'><span title='{e(label)}'>{e(label)}</span>"
                f"<i style='width:{tot / mx * 220:.0f}px'></i><em>{tot}</em></div>{o}")
    return out or "<p>Aucune donnée.</p>"


def write_dashboard(path: Path, records: dict[str, dict], cfg: dict, clean_subject, now: datetime) -> None:
    recs = list(records.values())
    opn = [r for r in recs if r["status"] in ("Nouveau", "En cours")]
    alerts = [r for r in opn if r["alert"]]
    done = [r for r in recs if r["status"] == "Répondu"]
    avg = f"{sum(r['age'] for r in done) / len(done):.1f} j" if done else "n/a"
    sla = [r for r in done if r["age"] <= cfg["sla"]["days"]]
    pct = f"{len(sla) / len(done):.0%}" if done else "n/a"
    kp = [("Demandes reçues", len(recs), ""), ("Ouvertes", len(opn), ""), ("En alerte", len(alerts), "al"),
          ("Critiques ouvertes", sum(r["priority"] == "Critique" for r in opn), "al"),
          ("Délai moyen de réponse", avg, ""), (f"Réponse ≤ {cfg['sla']['days']} j", pct, "")]
    cats = [c["name"] for c in cfg["categories"]] + [cfg["default_category"]]
    by_cat = Counter(r["category"] for r in recs)
    by_cat_open = Counter(r["category"] for r in opn)
    monday = lambda d: d.date() - timedelta(days=d.weekday())  # noqa: E731
    weeks = [monday(now) - timedelta(weeks=i) for i in range(7, -1, -1)]
    wk = Counter(monday(r["received"]) for r in recs)
    todo = sorted(opn, key=lambda r: (not r["alert"], -r["age"]))
    rows = ""
    for r in todo:
        dl = f"{r['deadline']:%d/%m/%Y}" if r["deadline"] else ""
        bg = COLORS.get(r["alert"], "")
        rows += (f"<tr class='d'><td>{r['received']:%d/%m %H:%M}</td><td>{e(r['sender'])}</td>"
                 f"<td>{e(r['subject'])}</td><td>{e(r['category'])}</td>"
                 f"<td>{e(r['summary'])}<br><i>{e(r['action'] or '')}</i></td><td>{e(r['priority'])}</td>"
                 f"<td>{dl}</td><td>{r['age']}</td>"
                 f"<td style='background:{bg}'>{e(r['alert'])}</td><td>{e(r['id'])}</td></tr>")
    by_subject = ""
    for c in cats:
        sub = sorted((r for r in recs if r["category"] == c), key=lambda r: (clean_subject(r["subject"]).lower(), r["received"]))
        if sub:
            li = "".join(f"<li>{e(clean_subject(r['subject']))} - {e(r['sender'])} - {r['received']:%d/%m} "
                         f"<b>[{e(r['status'])}]</b></li>" for r in sub)
            by_subject += f"<details><summary>{e(c)} ({len(sub)})</summary><ul>{li}</ul></details>"
    html = f"""<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><title>Suivi des demandes</title>
<style>{CSS}</style></head><body><header><h1 style="margin:0">Suivi des demandes - {e(cfg['target_address'])}</h1>
<small>Mis à jour le {now:%d/%m/%Y à %H:%M} (régénéré par l'agent à chaque passage - actualisez la page)</small></header><main>
<div class="kpis">{''.join(f"<div class='k {c}'><b>{v}</b>{e(l)}</div>" for l, v, c in kp)}</div>
<div class="row"><div class="card"><h2>Par catégorie (bleu : total, orange : ouvertes)</h2>{_bars([(c, by_cat[c], by_cat_open[c]) for c in cats])}</div>
<div class="card"><h2>Par statut</h2>{_bars([(s, sum(r['status'] == s for r in recs), None) for s in tracker.STATUSES])}</div>
<div class="card"><h2>Reçues par semaine</h2>{_bars([(f"sem. du {w:%d/%m}", wk[w], None) for w in weeks])}</div></div>
<div class="row"><div class="card"><h2>À traiter ({len(opn)}) - alertes en premier</h2>
<input id="q" placeholder="Filtrer (objet, expéditeur, catégorie…)">
<table id="todo"><thead><tr><th>Reçu</th><th>Expéditeur</th><th>Objet</th><th>Catégorie</th><th>Résumé / action</th><th>Priorité</th><th>Échéance</th><th>Âge (j)</th><th>Alerte</th><th>ID</th></tr></thead><tbody>{rows}</tbody></table></div></div>
<div class="row"><div class="card"><h2>Tri par objet</h2>{by_subject or '<p>Aucune donnée.</p>'}</div></div>
</main><script>{JS}</script></body></html>"""
    tracker._atomic_write(path, lambda f: f.write(html))
