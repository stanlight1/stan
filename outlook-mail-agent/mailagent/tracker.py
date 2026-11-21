"""Fichiers Excel : suivi (+ tableau de bord) et tri par objet.

Le fichier de suivi est aussi la base de l'agent : une ligne = une conversation.
L'utilisateur ne modifie que « Statut » et « Notes » ; le reste est régénéré à chaque passage.
"""
from __future__ import annotations

import re
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

STATUSES = ["Nouveau", "En cours", "Répondu", "Clôturé"]
CLOSED = {"Répondu", "Clôturé"}
PRIORITIES = ["Critique", "Haute", "Normale", "Basse"]

COLUMNS = [  # (clé, titre, largeur)
    ("id", "ID", 12), ("received", "Reçu le", 17), ("sender", "Expéditeur", 28),
    ("subject", "Objet", 40), ("category", "Catégorie", 18), ("summary", "Résumé", 60),
    ("action", "Action attendue", 30), ("priority", "Priorité", 11), ("deadline", "Échéance", 12),
    ("status", "Statut", 11), ("replied", "Répondu le", 17), ("age", "Âge (j)", 8),
    ("alert", "Alerte", 18), ("alert_sent", "Alerte envoyée le", 17), ("notes", "Notes", 30),
    ("conv", "Conversation", 12),
]
KEYS = [c[0] for c in COLUMNS]
DATE_KEYS = {"received", "replied", "alert_sent", "deadline"}
FMT = "%Y-%m-%d %H:%M"
HEAD = PatternFill("solid", fgColor="1F3864")
RED, ORANGE, YELLOW, GREEN = (PatternFill("solid", fgColor=c) for c in ("F8696B", "FFB366", "FFEB84", "A9D08E"))


def _write_atomic(wb: Workbook, path: Path) -> None:
    """Sauvegarde avec nouvel essai : Excel verrouille le fichier tant qu'il est ouvert."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp.xlsx")
    wb.save(tmp)
    for attempt in range(5):
        try:
            tmp.replace(path)
            return
        except PermissionError:
            time.sleep(2 * (attempt + 1))
    raise PermissionError(f"{path.name} est ouvert dans Excel : fermez-le (version prête : {tmp.name}).")


def _to_dt(v):
    if isinstance(v, datetime) or v is None or v == "":
        return v or None
    try:
        return datetime.strptime(str(v)[:16], FMT) if len(str(v)) > 10 else datetime.strptime(str(v)[:10], "%Y-%m-%d")
    except ValueError:
        return None


def load_records(path: Path) -> dict[str, dict]:
    if not path.exists():
        return {}
    ws = load_workbook(path, data_only=True)["Suivi"]
    head = [c.value for c in ws[1]]
    title2key = {t: k for k, t, _ in COLUMNS}
    recs = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        r = {title2key[h]: v for h, v in zip(head, row) if h in title2key}
        if not r.get("id"):
            continue
        for k in DATE_KEYS:
            r[k] = _to_dt(r.get(k))
        r["status"] = r.get("status") if r.get("status") in STATUSES else "Nouveau"
        recs[r["id"]] = r
    return recs


def _style_header(ws, ncols: int) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), HEAD
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def _put_row(ws, r: int, rec: dict, keys: list[str]) -> None:
    for c, k in enumerate(keys, 1):
        v = rec.get(k)
        cell = ws.cell(r, c, v if v != "" else None)
        if k in ("received", "replied", "alert_sent"):
            cell.number_format = FMT
        elif k == "deadline":
            cell.number_format = "dd/mm/yyyy"
        if k in ("summary", "subject", "action", "notes"):
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def save_tracking(path: Path, records: dict[str, dict], categories: list[str]) -> None:
    wb = Workbook()
    dash = wb.active
    dash.title = "Tableau de bord"
    ws = wb.create_sheet("Suivi")
    ws.append([t for _, t, _ in COLUMNS])
    _style_header(ws, len(COLUMNS))
    for i, (_, _, w) in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    rows = sorted(records.values(), key=lambda r: r["received"], reverse=True)
    for r, rec in enumerate(rows, 2):
        _put_row(ws, r, rec, KEYS)
    last = max(len(rows) + 1, 2)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last}"

    col = {k: get_column_letter(i) for i, k in enumerate(KEYS, 1)}
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"{col['status']}2:{col['status']}5000")
    a = f"{col['alert']}2:{col['alert']}5000"
    ws.conditional_formatting.add(a, CellIsRule(operator="equal", formula=['"Critique"'], fill=RED))
    ws.conditional_formatting.add(a, CellIsRule(operator="equal", formula=['"SLA dépassé"'], fill=RED))
    ws.conditional_formatting.add(a, CellIsRule(operator="equal", formula=['"Échéance dépassée"'], fill=ORANGE))
    ws.conditional_formatting.add(a, CellIsRule(operator="equal", formula=['"Échéance proche"'], fill=YELLOW))
    s = f"{col['status']}2:{col['status']}5000"
    ws.conditional_formatting.add(s, FormulaRule(formula=[f'OR({col["status"]}2="Répondu",{col["status"]}2="Clôturé")'], fill=GREEN))

    _dashboard(dash, categories, col)
    _write_atomic(wb, path)


def _dashboard(ws, categories: list[str], col: dict[str, str]) -> None:
    S = lambda k: f"Suivi!${col[k]}:${col[k]}"  # noqa: E731
    ws["A1"], ws["A1"].font = "Tableau de bord - suivi des demandes", Font(bold=True, size=16)
    ws["A2"] = "Mis à jour par l'agent à chaque passage ; les chiffres se recalculent à l'ouverture."
    ws["A2"].font = Font(italic=True, color="666666")
    kpis = [
        ("Demandes reçues", f"=COUNTA({S('id')})-1"),
        ("Ouvertes (Nouveau / En cours)", f'=COUNTIF({S("status")},"Nouveau")+COUNTIF({S("status")},"En cours")'),
        ("Répondues / Clôturées", f'=COUNTIF({S("status")},"Répondu")+COUNTIF({S("status")},"Clôturé")'),
        ("En alerte", f'=COUNTIF({S("alert")},"?*")'),
        ("Dont SLA dépassé", f'=COUNTIF({S("alert")},"SLA dépassé")'),
        ("Délai moyen de réponse (j)", f'=IFERROR(AVERAGEIFS({S("age")},{S("status")},"Répondu"),"-")'),
    ]
    for i, (label, f) in enumerate(kpis):
        r = 4 + i
        ws.cell(r, 1, label).font = Font(bold=True)
        c = ws.cell(r, 2, f)
        c.font, c.alignment = Font(bold=True, size=13), Alignment(horizontal="left")
    ws.cell(4 + len(kpis) - 1, 2).number_format = "0.0"

    def table(top: int, title: str, items: list[str], key: str):
        ws.cell(top, 1, title)
        for j, h in enumerate([title, "Total", "Ouvertes", "En alerte"], 1):
            c = ws.cell(top, j, h)
            c.font, c.fill = Font(bold=True, color="FFFFFF"), HEAD
        for i, it in enumerate(items, 1):
            r = top + i
            ws.cell(r, 1, it)
            ws.cell(r, 2, f'=COUNTIF({S(key)},A{r})')
            ws.cell(r, 3, f'=COUNTIFS({S(key)},A{r},{S("status")},"Nouveau")+COUNTIFS({S(key)},A{r},{S("status")},"En cours")')
            ws.cell(r, 4, f'=COUNTIFS({S(key)},A{r},{S("alert")},"?*")')
        return top, top + len(items)

    t1, e1 = table(12, "Catégorie", categories, "category")
    t2, e2 = table(e1 + 3, "Statut", STATUSES, "status")
    t3, e3 = table(e2 + 3, "Priorité", PRIORITIES, "priority")

    bar = BarChart()
    bar.title, bar.type, bar.height, bar.width = "Demandes par catégorie", "bar", 8, 16
    bar.add_data(Reference(ws, min_col=2, max_col=4, min_row=t1, max_row=e1), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=t1 + 1, max_row=e1))
    ws.add_chart(bar, "G3")
    pie = PieChart()
    pie.title, pie.height, pie.width = "Répartition par statut", 8, 10
    pie.add_data(Reference(ws, min_col=2, min_row=t2, max_row=e2), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=t2 + 1, max_row=e2))
    ws.add_chart(pie, "G20")
    ws.column_dimensions["A"].width = 34
    for c in "BCD":
        ws.column_dimensions[c].width = 12


def _sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "-", name)[:28] or "Feuille"
    n, i = base, 2
    while n.lower() in used:
        n, i = f"{base[:25]}_{i}", i + 1
    used.add(n.lower())
    return n


def save_sorted(path: Path, records: dict[str, dict], categories: list[str], clean_subject) -> None:
    """Un onglet « Toutes » + un onglet par catégorie, triés par objet (puis par date)."""
    keys = ["received", "sender", "subject", "category", "summary", "priority", "status", "alert"]
    titles = {k: t for k, t, _ in COLUMNS}
    widths = {k: w for k, _, w in COLUMNS}
    recs = sorted(records.values(), key=lambda r: (clean_subject(r["subject"]).lower(), r["received"]))
    wb = Workbook()
    used: set[str] = set()
    sheets = [("Toutes", recs)] + [(c, [r for r in recs if r["category"] == c]) for c in categories]
    first = True
    for name, rows in sheets:
        if not first and not rows:
            continue
        ws = wb.active if first else wb.create_sheet()
        ws.title = _sheet_name(name, used)
        first = False
        ws.append([titles[k] for k in keys])
        _style_header(ws, len(keys))
        for i, k in enumerate(keys, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths[k]
        for r, rec in enumerate(rows, 2):
            _put_row(ws, r, rec, keys)
        ws.auto_filter.ref = f"A1:{get_column_letter(len(keys))}{max(len(rows) + 1, 2)}"
    _write_atomic(wb, path)
