"""Hent Outlook-mails og byg ATS-items uden at ændre mails.

main.py bruger get_relevant_queue_items() og ejer ATS-oprettelsen.
populate_queue() bevares til den asynkrone procesvariant.
Begge indgange bruger samme mailhentning og validering.
"""

from __future__ import annotations

import asyncio
import hashlib
import inspect
import logging
import re
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo

import configuration as config
import outlook_service

logger = logging.getLogger(__name__)


class PopulateQueueFejl(RuntimeError):
    """Kontrolleret fejl under hentning eller oprettelse af kø-items."""


@dataclass(slots=True)
class PopulateQueueResultat:
    """Optælling fra én producer-kørsel."""

    mails_laest: int = 0
    relevante_mails: int = 0
    items_tilfoejet: int = 0
    mails_ignoreret: int = 0
    dubletter_i_mailresultat: int = 0


def _kontroller_konfiguration() -> None:
    """Kontrollér de samme aliases som integrationstesten."""
    adresse = config.OUTLOOK_MAILBOX_ADDRESS
    if not isinstance(adresse, str) or not adresse.strip():
        raise PopulateQueueFejl("OUTLOOK_MAILBOX_ADDRESS mangler.")
    if config.MAILBOX_NAME != adresse:
        raise PopulateQueueFejl(
            "MAILBOX_NAME skal være lig OUTLOOK_MAILBOX_ADDRESS."
        )
    if config.MAIL_FOLDER_NAME != config.OUTLOOK_FOLDER_NAME:
        raise PopulateQueueFejl(
            "MAIL_FOLDER_NAME skal være lig OUTLOOK_FOLDER_NAME."
        )
    mappe = config.OUTLOOK_DEFAULT_FOLDER
    if not isinstance(mappe, str) or not mappe.strip():
        raise PopulateQueueFejl("OUTLOOK_DEFAULT_FOLDER mangler.")


def _kontroller_mail(mail: Any, *, index: int) -> None:
    """Validér mailkontrakten uden pytest eller assert."""
    if not isinstance(mail, Mapping):
        raise PopulateQueueFejl(f"Mail nr. {index} skal være en dictionary.")
    for felt in ("mail_id", "received_at", "subject", "has_attachments"):
        if felt not in mail:
            raise PopulateQueueFejl(f"Mail nr. {index} mangler {felt}.")
    if not isinstance(mail["mail_id"], str) or not mail["mail_id"].strip():
        raise PopulateQueueFejl(f"Mail nr. {index} har ugyldigt mail_id.")
    if not isinstance(mail["subject"], str):
        raise PopulateQueueFejl(f"Mail nr. {index}: subject skal være str.")
    if not isinstance(mail["has_attachments"], bool):
        raise PopulateQueueFejl(
            f"Mail nr. {index}: has_attachments skal være bool."
        )
    parse_received_at(mail["received_at"])
    message_id = mail.get("internet_message_id")
    if message_id is not None and not isinstance(message_id, str):
        raise PopulateQueueFejl(
            f"Mail nr. {index}: internet_message_id skal være str eller None."
        )
    if "attachment_names" in mail and not isinstance(mail["attachment_names"], list):
        raise PopulateQueueFejl(
            f"Mail nr. {index}: attachment_names skal være en liste."
        )
    if "attachment_count" in mail and (
        isinstance(mail["attachment_count"], bool)
        or not isinstance(mail["attachment_count"], int)
        or mail["attachment_count"] < 0
    ):
        raise PopulateQueueFejl(
            f"Mail nr. {index}: attachment_count skal være et ikke-negativt heltal."
        )


async def _hent_queue_items() -> tuple[list[dict[str, Any]], PopulateQueueResultat]:
    """Hent gennem den eksisterende service og byg validerede items."""
    _kontroller_konfiguration()
    logger.info(
        "Henter mails. Postkasse=%s, mappe=%s",
        config.OUTLOOK_MAILBOX_ADDRESS,
        config.OUTLOOK_DEFAULT_FOLDER,
    )
    mails = await outlook_service.hent_mails()
    if not isinstance(mails, list):
        raise PopulateQueueFejl("outlook_service.hent_mails() skal returnere en liste.")
    resultat = PopulateQueueResultat(mails_laest=len(mails))
    queue_items: list[dict[str, Any]] = []
    sete_mail_ids: set[str] = set()
    referencer: dict[str, str] = {}
    for index, mail in enumerate(mails, start=1):
        _kontroller_mail(mail, index=index)
        mail_id = mail["mail_id"].strip()
        if mail_id in sete_mail_ids:
            resultat.dubletter_i_mailresultat += 1
            continue
        sete_mail_ids.add(mail_id)
        subject = mail["subject"].strip()
        mail_type = find_mail_type(subject)
        if mail_type is None:
            resultat.mails_ignoreret += 1
            continue
        received_at = parse_received_at(mail["received_at"])
        reference = opret_queue_reference(
            mail_type=mail_type,
            received_at=received_at,
            mail_id=mail_id,
        )
        if reference in referencer and referencer[reference] != mail_id:
            raise PopulateQueueFejl(
                "To forskellige mails fik samme ATS-reference. "
                f"reference={reference!r}. Ret referenceformatet i "
                "config.py, før der oprettes items."
            )
        referencer[reference] = mail_id
        data = opret_item_data(
            mail=mail,
            mail_id=mail_id,
            subject=subject,
            mail_type=mail_type,
            received_at=received_at,
        )
        config.validate_queue_data(data)
        queue_items.append({"data": data, "reference": reference})
        resultat.relevante_mails += 1
    logger.info(
        "Mailhentning afsluttet. Hentet=%s, relevant=%s, ignoreret=%s, dubletter=%s",
        resultat.mails_laest,
        resultat.relevante_mails,
        resultat.mails_ignoreret,
        resultat.dubletter_i_mailresultat,
    )
    return queue_items, resultat


def get_relevant_queue_items() -> list[dict[str, Any]]:
    """Returnér item-specifikationer; main.py opretter dem i ATS."""
    return asyncio.run(_build_relevant_queue_items())


async def _build_relevant_queue_items() -> list[dict[str, Any]]:
    queue_items, resultat = await _hent_queue_items()
    print(
        f"Mails hentet: {resultat.mails_laest}. "
        f"Relevante queue-items fundet: {len(queue_items)}."
    )
    return queue_items


async def populate_queue(workqueue: Any, debug: bool = False) -> int:
    """Asynkron producer: returnér antal oprettede ATS-items."""
    logger.info("Populate queue startet (debug=%s)", debug)
    queue_items, resultat = await _hent_queue_items()
    for queue_item in queue_items:
        try:
            await _tilfoej_item(
                workqueue=workqueue,
                data=queue_item["data"],
                reference=queue_item["reference"],
            )
        except Exception as error:
            raise PopulateQueueFejl("ATS kunne ikke oprette et kø-item.") from error
        resultat.items_tilfoejet += 1
    _print_resultat(resultat)
    return resultat.items_tilfoejet


def find_mail_type(subject: str) -> str | None:
    """Bevar de eksisterende case-insensitive emneregler."""
    normalised_subject = normaliser_subject(subject)
    if re.match(
        config.DIGITALPOST_BETALING_EMNE_PATTERN,
        normalised_subject,
        flags=re.IGNORECASE,
    ):
        return config.MAIL_TYPE_DIGITALPOST_BETALING
    for configured_subject, mail_type in config.SUBJECT_TO_MAIL_TYPE.items():
        if mail_type == config.MAIL_TYPE_DIGITALPOST_BETALING:
            continue
        if normaliser_subject(configured_subject) in normalised_subject:
            return mail_type
    return None


def normaliser_subject(subject: str) -> str:
    value = str(subject or "").replace("\u00a0", " ").strip()
    prefixes = tuple(
        str(prefix).strip().casefold()
        for prefix in config.SUBJECT_PREFIXES
        if str(prefix).strip()
    )
    while prefixes:
        matching_prefix = next(
            (prefix for prefix in prefixes if value.casefold().startswith(prefix)),
            None,
        )
        if matching_prefix is None:
            break
        value = value[len(matching_prefix):].lstrip()
    return re.sub(r"\s+", " ", value).casefold().strip()


def opret_queue_reference(
    *,
    mail_type: str,
    received_at: datetime,
    mail_id: str,
) -> str:
    """Opret en stabil og unik ATS-reference for én Outlook-mail."""
    normaliseret_mail_type = str(mail_type or "").strip()
    if not normaliseret_mail_type:
        raise PopulateQueueFejl("Mailtypen mangler.")

    normaliseret_mail_id = str(mail_id or "").strip()
    if not normaliseret_mail_id:
        raise PopulateQueueFejl(
            "Mail-id mangler og kan derfor ikke indgå i ATS-referencen."
        )

    if received_at.tzinfo is None or received_at.utcoffset() is None:
        raise PopulateQueueFejl("received_at mangler tidszone.")

    hash_length = config.QUEUE_REFERENCE_MAIL_ID_HASH_LENGTH
    if isinstance(hash_length, bool) or not isinstance(hash_length, int):
        raise PopulateQueueFejl(
            "QUEUE_REFERENCE_MAIL_ID_HASH_LENGTH skal være et heltal."
        )
    if not 8 <= hash_length <= 64:
        raise PopulateQueueFejl(
            "QUEUE_REFERENCE_MAIL_ID_HASH_LENGTH skal være mellem 8 og 64."
        )

    local_time = received_at.astimezone(ZoneInfo(config.TIMEZONE))
    formatted_time = local_time.strftime(config.QUEUE_REFERENCE_DATETIME_FORMAT)
    mail_id_hash = hashlib.sha256(
        normaliseret_mail_id.encode("utf-8")
    ).hexdigest()[:hash_length]

    return (
        f"{normaliseret_mail_type}"
        f"{config.QUEUE_REFERENCE_SEPARATOR}"
        f"{formatted_time}"
        f"{config.QUEUE_REFERENCE_SEPARATOR}"
        f"{mail_id_hash}"
    )

def opret_item_data(
    *,
    mail: Mapping[str, Any],
    mail_id: str,
    subject: str,
    mail_type: str,
    received_at: datetime,
) -> dict[str, Any]:
    """Bevar datafelterne, som den eksisterende behandel.py læser."""
    return {
        "mail_id": mail_id,
        "internet_message_id": mail.get("internet_message_id"),
        "received_at": received_at.isoformat(timespec="seconds"),
        "subject": subject,
        "mail_type": mail_type,
        "box": {"schema_version": config.BOX_SCHEMA_VERSION},
    }


def parse_received_at(value: Any) -> datetime:
    """Returnér et tidspunkt med en gyldig tidszone."""
    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, str) and value.strip():
        try:
            parsed = datetime.fromisoformat(value.strip())
        except ValueError as error:
            raise PopulateQueueFejl("received_at er ikke gyldig ISO 8601.") from error
    else:
        raise PopulateQueueFejl("received_at skal være datetime eller ISO 8601-tekst.")
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        raise PopulateQueueFejl("received_at mangler tidszone.")
    return parsed


async def _tilfoej_item(
    *, workqueue: Any, data: dict[str, Any], reference: str
) -> None:
    result = workqueue.add_item(data=data, reference=reference)
    if inspect.isawaitable(result):
        await result


def _print_resultat(resultat: PopulateQueueResultat) -> None:
    print("\nPOPULATE QUEUE AFSLUTTET")
    print(f"Postkasse: {config.OUTLOOK_MAILBOX_ADDRESS}")
    print(f"Mails hentet: {resultat.mails_laest}")
    print(f"Relevante mails: {resultat.relevante_mails}")
    print(f"Ignorerede mails: {resultat.mails_ignoreret}")
    print(f"Dubletter i mailresultatet: {resultat.dubletter_i_mailresultat}")
    print(f"Items tilføjet: {resultat.items_tilfoejet}")
