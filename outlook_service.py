"""Outlook-service for processen mails-i-forsikring.

Servicen er et kort adapterlag oven på:

    q_outlook_api.functionality.mail_api

q_outlook_api håndterer selv authentication og Microsoft Graph. Denne fil
indeholder derfor ingen credentials, tokens, client secrets eller HTTP-klienter.

Processtrukturen bevares:
- populate_queue.py kalder hent_mails().
- behandel.py behandler ét ATS-item og kalder de øvrige funktioner.
- Denne service opretter eller afslutter aldrig ATS-items.
"""

from __future__ import annotations

import asyncio
import html
from datetime import datetime
from typing import Any, Mapping

from q_outlook_api.functionality import mail_api

import configuration as config


class OutlookServiceError(RuntimeError):
    """Basisklasse for kontrollerede Outlook-servicefejl."""


class OutlookConfigurationError(OutlookServiceError):
    """En nødvendig procesindstilling mangler."""


class OutlookDataError(OutlookServiceError):
    """q_outlook_api returnerede manglende eller ugyldige data."""


class OutlookMailNotFoundError(OutlookServiceError):
    """Mailen findes ikke længere i Outlook."""


class OutlookAttachmentError(OutlookServiceError):
    """En PDF eller anden vedhæftning kunne ikke anvendes."""


class OutlookSendError(OutlookServiceError):
    """En mail kunne ikke sendes eller videresendes."""


# Danske aliases til eksisterende proceskode.
OutlookServiceFejl = OutlookServiceError
OutlookKonfigurationsfejl = OutlookConfigurationError
OutlookDatafejl = OutlookDataError
OutlookMailIkkeFundetFejl = OutlookMailNotFoundError
OutlookAfsendelsesfejl = OutlookSendError


async def hent_mails(
    *,
    folder: str | None = None,
    unread_only: bool | None = None,
    page_size: int | None = None,
    max_mails: int | None = None,
    mailbox_name: str | None = None,
    folder_name: str | None = None,
    select_fields: tuple[str, ...] | None = None,
) -> list[dict[str, Any]]:
    """Hent mailmetadata fra Outlook uden at ændre postkassen.

    Argumenterne ``mailbox_name``, ``folder_name`` og ``select_fields``
    accepteres af hensyn til populate_queue.py. q_outlook_api bestemmer selv
    Graph-felterne, og ``select_fields`` bruges derfor ikke i kaldet.
    """
    del select_fields

    if mailbox_name:
        forventet_navn = str(
            getattr(config, "OUTLOOK_MAILBOX_NAME", "Forsikring")
        )
        if mailbox_name.casefold() not in {
            forventet_navn.casefold(),
            _user_mail().casefold(),
        }:
            raise OutlookConfigurationError(
                "Den angivne postkasse matcher ikke konfigurationen. "
                f"Modtog: {mailbox_name!r}."
            )

    valgt_folder = (
        folder_name
        or folder
        or getattr(config, "OUTLOOK_DEFAULT_FOLDER", None)
        or getattr(config, "OUTLOOK_FOLDER_NAME", None)
        or getattr(config, "MAIL_FOLDER_NAME", None)
        or "inbox"
    )
    graph_folder = _graph_folder(str(valgt_folder))

    limit = max_mails
    if limit is None and page_size is not None:
        limit = int(page_size)

    formatted_mails = await asyncio.to_thread(
        mail_api.get_mails,
        user_mail=_user_mail(),
        folder=graph_folder,
        raw=False,
        limit=limit,
        apply_rules=False,
        include_attachments=False,
        get_inline=False,
        prefer_plain_text=True,
    )

    if not isinstance(formatted_mails, list):
        raise OutlookDataError(
            "q_outlook_api.get_mails() returnerede ikke en liste. "
            f"Modtog: {type(formatted_mails).__name__}."
        )

    resultater = [_normaliser_mail(mail) for mail in formatted_mails]

    if unread_only is None:
        unread_only = bool(
            getattr(config, "OUTLOOK_UNREAD_ONLY", False)
        )

    if unread_only:
        resultater = [mail for mail in resultater if not mail["is_read"]]

    if max_mails is not None:
        resultater = resultater[:max_mails]

    return resultater


async def hent_mail(mail_id: str) -> dict[str, Any]:
    """Hent én konkret mail ved immutable Outlook-ID."""
    mail_id = _paakraevet_tekst(mail_id, "mail_id")

    try:
        mails = await asyncio.to_thread(
            mail_api.get_mails,
            user_mail=_user_mail(),
            message_id=mail_id,
            raw=False,
            include_attachments=False,
            get_inline=False,
            prefer_plain_text=True,
        )
    except mail_api.MailNotFoundError as error:
        raise OutlookMailNotFoundError(
            f"Mailen blev ikke fundet: {mail_id}"
        ) from error

    if not isinstance(mails, list):
        raise OutlookDataError(
            "q_outlook_api.get_mails(message_id=...) returnerede ikke en liste."
        )

    if not mails:
        raise OutlookMailNotFoundError(
            f"Mailen blev ikke fundet: {mail_id}"
        )

    return _normaliser_mail(mails[0])


async def mail_findes(mail_id: str) -> bool:
    """Returnér True, hvis mailen stadig findes."""
    try:
        await hent_mail(mail_id)
        return True
    except OutlookMailNotFoundError:
        return False


async def hent_vedhaeftninger(
    mail: Any | None = None,
    *,
    mail_id: str | None = None,
    get_inline: bool = False,
) -> list[dict[str, Any]]:
    """Hent reelle vedhæftninger som bytes fra q_outlook_api."""
    resolved_mail_id = _mail_id(mail_id if mail_id is not None else mail)

    try:
        attachments = await asyncio.to_thread(
            mail_api.get_attachments,
            user_mail=_user_mail(),
            message_id=resolved_mail_id,
            get_inline=get_inline,
        )
    except mail_api.MailNotFoundError as error:
        raise OutlookMailNotFoundError(
            f"Mailen blev ikke fundet: {resolved_mail_id}"
        ) from error
    except mail_api.AttachmentDownloadError as error:
        raise OutlookAttachmentError(
            "En vedhæftning kunne ikke hentes."
        ) from error

    if not isinstance(attachments, list):
        raise OutlookDataError(
            "q_outlook_api.get_attachments() returnerede ikke en liste."
        )

    return [_normaliser_vedhaeftning(item) for item in attachments]


async def hent_pdf_vedhaeftninger(
    mail: Any | None = None,
    *,
    mail_id: str | None = None,
) -> list[dict[str, Any]]:
    """Hent ikke-inline PDF-vedhæftninger med indhold i ``content``."""
    attachments = await hent_vedhaeftninger(
        mail,
        mail_id=mail_id,
        get_inline=False,
    )

    return [
        attachment
        for attachment in attachments
        if attachment["name"].casefold().endswith(".pdf")
        or attachment["content_type"].split(";", 1)[0].casefold()
        == "application/pdf"
    ]


async def hent_eneste_pdf(
    mail: Any | None = None,
    *,
    mail_id: str | None = None,
) -> dict[str, Any]:
    """Returnér præcis én PDF eller rejs en kontrolleret fejl."""
    pdf_filer = await hent_pdf_vedhaeftninger(mail, mail_id=mail_id)

    if not pdf_filer:
        raise OutlookAttachmentError(
            "Mailen indeholder ingen læsbar PDF-vedhæftning."
        )

    if len(pdf_filer) > 1:
        navne = [pdf["name"] for pdf in pdf_filer]
        raise OutlookAttachmentError(
            "Mailen indeholder flere PDF-filer, og korrekt PDF er tvetydig. "
            f"Filer: {navne!r}."
        )

    return pdf_filer[0]


async def videresend_mail(
    mail: Any | None = None,
    modtager: str | None = None,
    *,
    mail_id: str | None = None,
    recipient: str | None = None,
    comment_html: str | None = None,
) -> bool:
    """Videresend originalmailen med Graph-genereret mailbody og attachments."""
    resolved_mail_id = _mail_id(mail_id if mail_id is not None else mail)
    resolved_recipient = _valider_email(recipient or modtager)
    body = (
        comment_html
        if comment_html is not None
        else str(getattr(config, "OUTLOOK_FORWARD_COMMENT", ""))
    )
    forward_mode = str(
        getattr(config, "OUTLOOK_FORWARD_MODE", "comment")
    )

    try:
        result = await asyncio.to_thread(
            mail_api.forward_mail,
            user_mail=_user_mail(),
            message_id=resolved_mail_id,
            forward={
                "to": [resolved_recipient],
                "cc": [],
                "bcc": [],
                "body": body,
                "forward_mode": forward_mode,
            },
        )
    except mail_api.MailNotFoundError as error:
        raise OutlookMailNotFoundError(
            f"Mailen blev ikke fundet: {resolved_mail_id}"
        ) from error
    except Exception as error:
        raise OutlookSendError(
            f"Mailen kunne ikke videresendes til {resolved_recipient}."
        ) from error

    if isinstance(result, Mapping):
        return bool(result.get("success", False))

    return result is True


async def send_mail(
    *,
    recipient: str,
    subject: str,
    body_html: str,
    attachments: list[dict[str, Any]] | None = None,
    save_to_sent_items: bool = True,
) -> bool:
    """Send en ny mail gennem q_outlook_api."""
    recipient = _valider_email(recipient)
    subject = _paakraevet_tekst(subject, "subject")
    body_html = _paakraevet_tekst(body_html, "body_html")

    q_attachments = []
    for attachment in attachments or []:
        content = attachment.get("content") or attachment.get("content_bytes")
        if not isinstance(content, bytes):
            raise OutlookAttachmentError(
                "Vedhæftningens content skal være bytes."
            )
        q_attachments.append(
            {
                "name": attachment.get("name"),
                "content_bytes": content,
                "content_type": attachment.get("content_type"),
            }
        )

    try:
        return bool(
            await asyncio.to_thread(
                mail_api.send_mail,
                user_mail=_user_mail(),
                mail={
                    "subject": subject,
                    "body": body_html,
                    "content_type": "HTML",
                    "to": [recipient],
                    "cc": [],
                    "bcc": [],
                    "attachments": q_attachments,
                    "save_to_sent_items": save_to_sent_items,
                },
            )
        )
    except Exception as error:
        raise OutlookSendError(
            f"Mailen kunne ikke sendes til {recipient}."
        ) from error


async def send_mail_til_forsikring(
    *,
    emne: str,
    besked: str,
    manuel_id: int,
    nummerplade: str | None = None,
    stelnummer: str | None = None,
) -> bool:
    """Send standardiseret mail med nummerplade og stelnummer til Forsikring."""
    if isinstance(manuel_id, bool) or not isinstance(manuel_id, int):
        raise TypeError("manuel_id skal være et heltal.")
    if manuel_id <= 0:
        raise ValueError("manuel_id skal være større end 0.")

    normalized_subject = _paakraevet_tekst(emne, "emne")
    normalized_message = _paakraevet_tekst(besked, "besked")
    normalized_registration = str(nummerplade or "").strip() or "Mangler"
    normalized_chassis_number = str(stelnummer or "").strip() or "Mangler"

    body_html = (
        f"<p><strong>Manuel ID:</strong> {manuel_id}</p>"
        f"<p>{html.escape(normalized_message)}</p>"
        '<table style="border-collapse: collapse; margin-top: 16px;">'
        "<tr>"
        '<td style="padding: 4px 16px 4px 0;">'
        "<strong>Nummerplade:</strong>"
        "</td>"
        '<td style="padding: 4px 0;">'
        f"{html.escape(normalized_registration)}"
        "</td>"
        "</tr>"
        "<tr>"
        '<td style="padding: 4px 16px 4px 0;">'
        "<strong>Stelnummer:</strong>"
        "</td>"
        '<td style="padding: 4px 0;">'
        f"{html.escape(normalized_chassis_number)}"
        "</td>"
        "</tr>"
        "</table>"
        "<p>Originalmailen er ikke flyttet eller slettet og skal "
        "behandles manuelt.</p>"
    )
    return await send_mail(
        recipient=config.FORSIKRINGSTEAM_EMAIL,
        subject=(
            f"[Manuel {manuel_id}] {normalized_subject} - "
            f"{normalized_registration}"
        ),
        body_html=body_html,
        attachments=None,
        save_to_sent_items=True,
    )

async def send_manglende_afdelingsmail(
    *,
    modtager: str,
    original_mail: Any,
    mail_subject: str | None,
    received_at: str | datetime | None,
    identifikator_type: str,
    identifikator_vaerdi: str,
    nummerplade: str | None = None,
    stelnummer: str | None = None,
) -> bool:
    """Send separat informationsmail. Originalmailen ændres ikke."""
    lines = [
        "Robotten har fundet køretøjet i Insubiz, men Afdelingsmail mangler.",
        "Originalmailen er ikke flyttet eller slettet og skal behandles manuelt.",
        "",
        f"Identifikatortype: {identifikator_type}",
        f"Identifikator: {identifikator_vaerdi}",
        f"Nummerplade: {nummerplade or 'Mangler'}",
        f"Stelnummer: {stelnummer or 'Mangler'}",
        f"Oprindeligt emne: {mail_subject or _felt(original_mail, 'subject') or 'Mangler'}",
        f"Mail modtaget: {_format_tidspunkt(received_at)}",
        f"Original mail-ID: {_mail_id(original_mail)}",
        f"State: {getattr(config, 'STATE_MANGLENDE_AFDELINGSMAIL', 'Manuel behandling')}",
    ]

    body_html = "<br>".join(html.escape(line) for line in lines)

    return await send_mail(
        recipient=modtager,
        subject=str(
            getattr(
                config,
                "OUTLOOK_MANUAL_SUBJECT",
                "Manuel behandling - Afdelingsmail mangler på køretøj",
            )
        ),
        body_html=body_html,
        attachments=None,
        save_to_sent_items=True,
    )


async def flyt_mail(
    mail: Any | None = None,
    destination_folder_id: str | None = None,
    *,
    mail_id: str | None = None,
    destination_folder_name: str | None = None,
) -> dict[str, Any]:
    """Flyt mailen til et Graph-mappe-ID eller et kendt mappenavn."""
    resolved_mail_id = _mail_id(mail_id if mail_id is not None else mail)
    destination = destination_folder_id or destination_folder_name
    destination = _paakraevet_tekst(destination, "destination_folder_id")

    try:
        result = await asyncio.to_thread(
            mail_api.move_mail,
            user_mail=_user_mail(),
            message_id=resolved_mail_id,
            destination_folder_id=_graph_folder(destination),
        )
    except mail_api.MailNotFoundError as error:
        raise OutlookMailNotFoundError(
            "Mailen eller destinationsmappen findes ikke."
        ) from error

    return _normaliser_mail(result)


async def slet_mail(
    mail: Any | None = None,
    *,
    mail_id: str | None = None,
    permanent_delete: bool = False,
    ignore_if_missing: bool = True,
) -> bool:
    """Flyt til Slettet post eller udfør permanent sletning."""
    resolved_mail_id = _mail_id(mail_id if mail_id is not None else mail)

    try:
        result = await asyncio.to_thread(
            mail_api.delete_mail,
            user_mail=_user_mail(),
            message_id=resolved_mail_id,
            permanent_delete=permanent_delete,
        )
    except mail_api.MailNotFoundError as error:
        if ignore_if_missing:
            return False
        raise OutlookMailNotFoundError(
            f"Mailen blev ikke fundet: {resolved_mail_id}"
        ) from error

    return bool(isinstance(result, Mapping) and result.get("success"))


async def set_mail_read_status(
    *,
    mail_id: str,
    is_read: bool,
) -> bool:
    """Markér mailen som læst eller ulæst."""
    if not isinstance(is_read, bool):
        raise TypeError("is_read skal være boolsk.")

    try:
        return bool(
            await asyncio.to_thread(
                mail_api.mark_as_read,
                user_mail=_user_mail(),
                message_id=_mail_id(mail_id),
                is_read=is_read,
            )
        )
    except mail_api.MailNotFoundError as error:
        raise OutlookMailNotFoundError(
            f"Mailen blev ikke fundet: {mail_id}"
        ) from error


async def hent_mapper(
    *,
    include_hidden: bool = True,
    include_children: bool = True,
) -> list[dict[str, Any]]:
    """Hent postkassens mapper gennem q_outlook_api."""
    result = await asyncio.to_thread(
        mail_api.get_folders,
        user_mail=_user_mail(),
        include_hidden=include_hidden,
        include_children=include_children,
    )

    if not isinstance(result, list):
        raise OutlookDataError(
            "q_outlook_api.get_folders() returnerede ikke en liste."
        )

    return [dict(folder) for folder in result]


def _normaliser_mail(mail: Any) -> dict[str, Any]:
    """Map q_outlook_api-formatet til processens stabile mailkontrakt."""
    if not isinstance(mail, Mapping):
        raise OutlookDataError(
            "q_outlook_api returnerede ikke en dictionary for en mail."
        )

    mail_id = mail.get("message_id") or mail.get("id")
    received_at = (
        mail.get("received_datetime_utc")
        or mail.get("received_utc")
    )

    if not mail_id:
        raise OutlookDataError("Mailen mangler message_id/id.")

    if not received_at:
        raise OutlookDataError(
            f"Mailen {mail_id} mangler received_datetime_utc/received_utc."
        )

    if isinstance(received_at, datetime):
        received_at = received_at.isoformat()

    return {
        "mail_id": str(mail_id),
        "id": str(mail_id),
        "internet_message_id": mail.get("internet_message_id"),
        "conversation_id": mail.get("conversation_id"),
        "received_at": str(received_at),
        "received_at_danish": mail.get("received_datetime_danish")
        or mail.get("received_danish_str"),
        "subject": str(mail.get("subject") or ""),
        "has_attachments": bool(mail.get("has_attachments", False)),
        "attachment_count": int(mail.get("attachment_count") or 0),
        "attachment_names": list(mail.get("attachment_names") or []),
        "is_read": bool(mail.get("is_read", False)),
        "categories": list(mail.get("categories") or []),
        "folder": mail.get("folder"),
        "mailbox": mail.get("mailbox"),
    }


def _normaliser_vedhaeftning(attachment: Any) -> dict[str, Any]:
    if not isinstance(attachment, Mapping):
        raise OutlookDataError(
            "q_outlook_api returnerede ikke en dictionary for en vedhæftning."
        )

    content = attachment.get("content_bytes")
    if not isinstance(content, bytes):
        raise OutlookAttachmentError(
            f"Vedhæftningen {attachment.get('name')!r} mangler content_bytes."
        )

    return {
        "id": attachment.get("id"),
        "attachment_id": attachment.get("id"),
        "name": str(attachment.get("name") or "Vedhæftet fil"),
        "original_name": attachment.get("original_name"),
        "content_type": str(
            attachment.get("content_type") or "application/octet-stream"
        ),
        "content": content,
        "content_bytes": content,
        "size": int(attachment.get("size") or len(content)),
        "is_inline": bool(attachment.get("is_inline", False)),
        "attachment_type": attachment.get("attachment_type"),
        "item_type": attachment.get("item_type"),
    }


def _user_mail() -> str:
    value = (
        getattr(config, "OUTLOOK_MAILBOX_ADDRESS", None)
        or getattr(config, "OUTLOOK_MAILBOX_USER_ID", None)
        or getattr(config, "MAILBOX_USER_ID", None)
        or getattr(config, "MAILBOX_NAME", None)
    )

    return _paakraevet_tekst(value, "OUTLOOK_MAILBOX_ADDRESS")


def _graph_folder(value: str) -> str:
    mapping = {
        "indbakke": "inbox",
        "inbox": "inbox",
        "sendt post": "sentitems",
        "sendte elementer": "sentitems",
        "slettet post": "deleteditems",
        "slettede elementer": "deleteditems",
        "kladder": "drafts",
        "uønsket mail": "junkemail",
        "arkiv": "archive",
    }
    cleaned = _paakraevet_tekst(value, "folder").strip()
    return mapping.get(cleaned.casefold(), cleaned)


def _mail_id(mail: Any) -> str:
    if isinstance(mail, str):
        value = mail
    elif isinstance(mail, Mapping):
        value = mail.get("mail_id") or mail.get("message_id") or mail.get("id")
    else:
        value = getattr(mail, "mail_id", None) or getattr(mail, "id", None)

    return _paakraevet_tekst(value, "mail_id")


def _felt(objekt: Any, *navne: str) -> Any:
    if isinstance(objekt, Mapping):
        for navn in navne:
            if navn in objekt:
                return objekt[navn]
        return None

    for navn in navne:
        if hasattr(objekt, navn):
            return getattr(objekt, navn)

    return None


def _paakraevet_tekst(value: Any, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise OutlookConfigurationError(
            f"{field_name} skal være en ikke-tom tekststreng."
        )
    return value.strip()


def _valider_email(value: Any) -> str:
    email = _paakraevet_tekst(value, "email")
    if email.count("@") != 1 or any(character.isspace() for character in email):
        raise ValueError(f"Ugyldig mailadresse: {email!r}.")
    local_part, domain_part = email.rsplit("@", 1)
    if not local_part or not domain_part or "." not in domain_part:
        raise ValueError(f"Ugyldig mailadresse: {email!r}.")
    return email


def _format_tidspunkt(value: str | datetime | None) -> str:
    if isinstance(value, datetime):
        return value.isoformat(timespec="seconds")
    return str(value or "Mangler")


__all__ = [
    "OutlookAttachmentError",
    "OutlookConfigurationError",
    "OutlookDataError",
    "OutlookMailNotFoundError",
    "OutlookSendError",
    "OutlookServiceError",
    "flyt_mail",
    "hent_eneste_pdf",
    "hent_mail",
    "hent_mails",
    "hent_mapper",
    "hent_pdf_vedhaeftninger",
    "hent_vedhaeftninger",
    "mail_findes",
    "send_mail",
    "send_mail_til_forsikring",
    "send_manglende_afdelingsmail",
    "set_mail_read_status",
    "slet_mail",
    "videresend_mail",
]
