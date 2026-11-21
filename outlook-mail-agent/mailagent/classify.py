from __future__ import annotations

import re

from .sources import clean_subject


def categorize(subject: str, cfg: dict) -> str:
    """Catégorie déterminée par l'OBJET du mail (règles de config, ordre = priorité)."""
    s = clean_subject(subject).lower()
    for cat in cfg["categories"]:
        for kw in cat.get("subject_contains", []):
            if kw.lower() in s:
                return cat["name"]
        for rx in cat.get("subject_regex", []):
            if re.search(rx, s, re.I):
                return cat["name"]
    return cfg["default_category"]


def is_ignored(mail, cfg: dict) -> bool:
    s = mail.subject.lower()
    if any(x.lower() in s for x in cfg["ignore_subjects"]):
        return True
    return any(x.lower() in mail.sender for x in cfg["ignore_senders"])
