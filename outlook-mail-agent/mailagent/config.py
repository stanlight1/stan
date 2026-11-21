from __future__ import annotations

from pathlib import Path

import yaml

DEFAULTS = {
    "source": {"type": "outlook_com", "mailbox": "", "folder": "Inbox", "lookback_days": 14},
    "target_address": "",
    "paths": {"output_dir": "output", "tracking": "suivi_demandes.xlsx",
              "sorted": "tri_par_objet.xlsx", "reports_dir": "reports"},
    "sla": {"days": 3, "business_days": True, "deadline_warning_days": 2},
    "alerts": {"recipients": [], "mode": "digest", "immediate_on_critical": True},
    "weekly_report": {"recipients": []},
    "notifier": {"type": "outlook_com", "dry_run_dir": "output/outbox",
                 "smtp": {"host": "", "port": 587, "user": "", "password_env": "SMTP_PASSWORD",
                          "from": ""}},
    "graph": {"tenant_id": "", "client_id": "", "client_secret_env": "GRAPH_CLIENT_SECRET"},
    "ai": {"enabled": True, "model": "claude-haiku-5-5", "api_key_env": "ANTHROPIC_API_KEY",
           "max_body_chars": 4000},
    "categories": [],
    "default_category": "Autres",
    "ignore_subjects": [],
    "ignore_senders": [],
    "urgent_keywords": ["urgent", "asap", "bloquant", "critique", "immédiat", "dès que possible"],
}


def _merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in (over or {}).items():
        out[k] = _merge(base[k], v) if isinstance(v, dict) and isinstance(base.get(k), dict) else v
    return out


def load_config(path: str | Path) -> dict:
    path = Path(path)
    with path.open(encoding="utf-8") as f:
        cfg = _merge(DEFAULTS, yaml.safe_load(f) or {})
    # Les chemins relatifs sont résolus par rapport au fichier de config.
    base = path.resolve().parent
    out = Path(cfg["paths"]["output_dir"])
    out = out if out.is_absolute() else base / out
    cfg["_out"] = out
    cfg["_tracking"] = out / cfg["paths"]["tracking"]
    cfg["_sorted"] = out / cfg["paths"]["sorted"]
    cfg["_reports"] = out / cfg["paths"]["reports_dir"]
    d = Path(cfg["notifier"]["dry_run_dir"])
    cfg["_outbox"] = d if d.is_absolute() else base / d
    cfg["_base"] = base
    if not cfg["target_address"]:
        raise ValueError("config: 'target_address' est obligatoire (adresse dont on suit les mails).")
    return cfg
