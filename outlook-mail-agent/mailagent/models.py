from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class Mail:
    """Mail normalisé, indépendant du backend (heure locale, sans fuseau)."""
    id: str
    conversation_id: str
    received: datetime
    sender: str
    sender_name: str
    to: list[str]
    cc: list[str]
    subject: str
    body: str
    has_attachments: bool = False
    importance: str = "normal"  # low | normal | high


@dataclass
class Analysis:
    summary: str
    priority: str = "Normale"          # Haute | Normale | Basse
    deadline: datetime | None = None   # échéance citée dans le mail
    action: str = ""                   # action attendue
    critical: bool = False             # alerte immédiate demandée


@dataclass
class Outgoing:
    to: list[str]
    subject: str
    html: str
    attachments: list[str] = field(default_factory=list)
