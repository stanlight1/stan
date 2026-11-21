import csv
from datetime import datetime
from pathlib import Path

import pytest

from mailagent.agent import Agent, elapsed_days
from mailagent.cli import main, weekly_report
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


def rows(cfg):
    with open(cfg["_powerbi"] / "Demandes.csv", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def test_full_cycle(cfg):
    stats = Agent(cfg, now=NOW).run()
    assert stats["total"] == 5                      # newsletter ignorée
    by_subj = {r["Objet"]: r for r in rows(cfg)}
    vpn = by_subj["Demande d'accès VPN - nouveau collaborateur"]
    assert vpn["Categorie"] == "Accès & comptes"
    inv = by_subj["RE: Facture n°2024-118 incohérente"]
    assert inv["Categorie"] == "Facturation" and inv["Statut"] == "Répondu" and inv["Ouverte"] == "0"
    assert inv["Objet_nettoye"].startswith("Facture") and inv["Echeance"].endswith("-11-30")
    assert by_subj["URGENT - Application de paie bloquante"]["Priorite"] in ("Haute", "Critique")
    # 6 jours -> 4 jours ouvrés : SLA dépassé ; alerte envoyée une seule fois
    assert vpn["Alerte"] == "SLA dépassé" and vpn["Alerte_envoyee_le"]
    sent1 = list(cfg["_outbox"].glob("*.html"))
    assert sent1
    Agent(cfg, now=NOW).run()
    assert len(list(cfg["_outbox"].glob("*.html"))) == len(sent1)   # pas de doublon d'alertes
    with open(cfg["_powerbi"] / "Parametres.csv", encoding="utf-8-sig") as f:
        assert "SLA_jours,3" in f.read()


def test_set_status_conserve(cfg):
    Agent(cfg, now=NOW).run()
    rid = [r["ID"] for r in rows(cfg) if "contrat" in r["Objet"].lower()][0]
    assert main(["set-status", rid, "Clôturé", "traité", "par", "tel", "-c", str(cfg["_base"] / "config.yaml")]) == 0
    Agent(cfg, now=NOW).run()
    c = [r for r in rows(cfg) if r["ID"] == rid][0]
    assert c["Statut"] == "Clôturé" and c["Notes"] == "traité par tel" and c["Alerte"] == ""


def test_rapport(cfg):
    ag = Agent(cfg, now=NOW)
    ag.run()
    weekly_report(ag, ref=datetime(2026, 11, 21))
    assert list(cfg["_reports"].glob("reporting_*.html"))
    assert any("Reporting hebdo" in p.read_text(encoding="utf-8") for p in cfg["_outbox"].glob("*.html"))


def test_dashboard_html_echappe(cfg):
    import json
    f = cfg["_base"] / "sample_mails" / "x.json"
    f.write_text(json.dumps({"id": "x", "conversation_id": "cx", "received": NOW.isoformat(),
                             "sender": "a@b.fr", "to": [TARGET], "subject": "Test <script>alert(1)</script>",
                             "body": "<img src=x onerror=alert(1)>"}), encoding="utf-8")
    Agent(cfg, now=NOW).run()
    h = (cfg["_out"] / "tableau_de_bord.html").read_text(encoding="utf-8")
    assert "Suivi des demandes" in h and "À traiter" in h
    assert "<script>alert(1)" not in h and "<img src=x" not in h
