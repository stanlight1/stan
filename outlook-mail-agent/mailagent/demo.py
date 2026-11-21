"""Génère des mails fictifs (dates relatives à aujourd'hui) pour tester l'agent sans Outlook.

    python -m mailagent.demo            # écrit dans ./sample_mails
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

TARGET = "support@exemple.fr"


def make_demo(folder: Path, target: str = TARGET, now: datetime | None = None) -> None:
    now = now or datetime.now()
    folder.mkdir(parents=True, exist_ok=True)
    d = lambda days, h=9: (now - timedelta(days=days)).replace(hour=h, minute=0, second=0, microsecond=0)  # noqa: E731
    mails = [
        ("m1", 6, "Marie Durand", "marie@client-a.fr", "Demande d'accès VPN - nouveau collaborateur",
         "Bonjour, pouvez-vous créer un accès VPN pour Paul Martin qui arrive lundi ? Merci d'avance."),
        ("m2", 5, "Jean Petit", "jean@client-b.fr", "RE: Facture n°2024-118 incohérente",
         "Bonjour, le montant de la facture 2024-118 ne correspond pas au devis. Merci de corriger avant le 30/11."),
        ("m3", 1, "Sophie Leroy", "sophie@client-a.fr", "URGENT - Application de paie bloquante",
         "Impossible de lancer la paie, l'application est bloquante depuis ce matin. Urgent !"),
        ("m4", 4, "Luc Bernard", "luc@client-c.fr", "Question sur le contrat de maintenance",
         "Bonjour, nous souhaitons connaître les conditions de renouvellement du contrat de maintenance."),
        ("m5", 2, "Anne Roux", "anne@client-b.fr", "Newsletter de novembre",
         "Notre newsletter mensuelle est disponible en ligne."),
        ("m6", 0, "Paul Moreau", "paul@client-a.fr", "Incident imprimante étage 2",
         "L'imprimante de l'étage 2 est en panne depuis hier. Pouvez-vous intervenir ?"),
    ]
    for i, (mid, ago, name, addr, subj, body) in enumerate(mails):
        (folder / f"{mid}.json").write_text(json.dumps({
            "id": mid, "conversation_id": f"conv-{mid}", "received": d(ago, 8 + i).isoformat(),
            "sender": addr, "sender_name": name, "to": [target], "cc": [], "subject": subj,
            "body": body, "kind": "received"}, ensure_ascii=False, indent=1), encoding="utf-8")
    # une réponse envoyée à m2 (la demande m2 doit passer « Répondu »)
    (folder / "s2.json").write_text(json.dumps({
        "id": "s2", "conversation_id": "conv-m2", "received": d(3, 14).isoformat(), "sender": target,
        "to": ["jean@client-b.fr"], "subject": "RE: Facture n°2024-118 incohérente",
        "body": "Corrigé.", "kind": "sent"}, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    make_demo(Path(sys.argv[1] if len(sys.argv) > 1 else "sample_mails"))
    print("Mails de démo écrits.")
