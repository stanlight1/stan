from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

from mailagent.agent import Agent, elapsed_days
from mailagent.cli import weekly_report
from mailagent.config import load_config
from mailagent.demo import TARGET, make_demo

ROOT = Path(__file__).resolve().parent.parent
NOW = datetime(2026, 11, 20, 10, 0)  # vendredi


@pytest.fixture
def cfg(tmp_path):
    make_demo(tmp_path / "sample_mails", now=NOW)
    text = (ROOT / "config.example.yaml").read_text(encoding="utf-8")
    text = (text.replace('"support@votre-societe.fr"', f'"{TARGET}"')
            .replace("type: outlook_com        #", "type: json_folder        #")
            .replace("  # path: sample_mails", "  path: sample_mails")
            .replace("notifier:\n  type: outlook_com", "notifier:\n  type: file"))
    (tmp_path / "config.yaml").write_text(text, encoding="utf-8")
    return load_config(tmp_path / "config.yaml")


def test_business_days():
    assert elapsed_days(datetime(2026, 11, 20, 9), datetime(2026, 11, 23, 9), True) == 1   # ven -> lun
    assert elapsed_days(datetime(2026, 11, 20, 9), datetime(2026, 11, 23, 9), False) == 3


def rows(path, sheet="Suivi"):
    ws = load_workbook(path)[sheet]
    head = [c.value for c in ws[1]]
    return [dict(zip(head, r)) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]


def test_full_cycle(cfg):
    stats = Agent(cfg, now=NOW).run()
    assert stats["total"] == 5                      # newsletter ignorée
    by_subj = {r["Objet"]: r for r in rows(cfg["_tracking"])}
    assert by_subj["Demande d'accès VPN - nouveau collaborateur"]["Catégorie"] == "Accès & comptes"
    inv = by_subj["RE: Facture n°2024-118 incohérente"]
    assert inv["Catégorie"] == "Facturation" and inv["Statut"] == "Répondu"
    assert by_subj["URGENT - Application de paie bloquante"]["Priorité"] in ("Haute", "Critique")
    assert inv["Échéance"].day == 30
    # 6 jours -> 4 jours ouvrés (VPN) : SLA dépassé ; alerte envoyée une seule fois
    vpn = by_subj["Demande d'accès VPN - nouveau collaborateur"]
    assert vpn["Alerte"] == "SLA dépassé" and vpn["Alerte envoyée le"]
    sent1 = list(cfg["_outbox"].glob("*.html"))
    assert sent1
    Agent(cfg, now=NOW).run()
    assert len(list(cfg["_outbox"].glob("*.html"))) == len(sent1)   # pas de doublon d'alertes


def test_statut_manuel_conserve(cfg):
    Agent(cfg, now=NOW).run()
    from openpyxl import load_workbook as lw
    wb = lw(cfg["_tracking"])
    ws = wb["Suivi"]
    head = [c.value for c in ws[1]]
    for r in ws.iter_rows(min_row=2):
        if "contrat" in r[head.index("Objet")].value.lower():
            r[head.index("Statut")].value = "Clôturé"
            r[head.index("Notes")].value = "traité par tel"
    wb.save(cfg["_tracking"])
    Agent(cfg, now=NOW).run()
    c = [r for r in rows(cfg["_tracking"]) if "contrat" in r["Objet"].lower()][0]
    assert c["Statut"] == "Clôturé" and c["Notes"] == "traité par tel" and c["Alerte"] in (None, "")


def test_tri_et_dashboard_et_rapport(cfg):
    ag = Agent(cfg, now=NOW)
    ag.run()
    wb = load_workbook(cfg["_sorted"])
    assert {"Toutes", "Facturation", "Accès & comptes"} <= set(wb.sheetnames)
    assert "Tableau de bord" in load_workbook(cfg["_tracking"]).sheetnames
    weekly_report(ag, ref=datetime(2026, 11, 21))
    assert list(cfg["_reports"].glob("reporting_*.xlsx"))
    assert any("Reporting hebdo" in p.read_text(encoding="utf-8") for p in cfg["_outbox"].glob("*.html"))
