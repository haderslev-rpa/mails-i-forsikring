"""Behandling af ét ATS-kø-item for mails-i-forsikring.

Processtrukturen må ikke brydes:

1. populate_queue.py opretter kø-items.
2. main.py henter ét item ad gangen.
3. main.py kalder process_item(item).
4. Denne fil behandler præcis dét ene item.

Checkpoints gemmes i itemets Box data og reducerer risikoen for gentagen
videresendelse ved genkørsel. Outlook-authentication håndteres internt af q_outlook_api via
outlook_service.py.
"""

from __future__ import annotations

import inspect
import io
import re
import logging
from copy import deepcopy
from datetime import datetime, timezone
from typing import Any, Mapping, MutableMapping

from pypdf import PdfReader

import configuration as config
import insubiz_service
import manuel_behandling
import outlook_service
import vehicle_matcher
from pdf_beregner import (
    IdentifikatorIkkeFundetFejl,
    PdfBeregningFejl,
    TvetydigIdentifikatorFejl,
    udtraek_stel_eller_nummerplade,
)

logger = logging.getLogger(__name__)

class Behandlingsfejl(RuntimeError):
    """Basisklasse for kontrollerede fejl i itembehandlingen."""

class Forretningsfejl(Behandlingsfejl):
    """Itemet kan ikke behandles automatisk på grund af data eller regler."""

class Systemfejl(Behandlingsfejl):
    """En integration eller teknisk handling fejlede."""

async def process_item(item: Any) -> dict[str, Any]:
    """Processkabelonens offentlige entrypoint.

    Behandler præcis ét kø-item og returnerer det afsluttende Box data. Main.py
    ejer fortsat kø-loopet og den overordnede Automation Server-livscyklus.
    """
    return await behandel(item)

# START: Erstat hele den eksisterende behandel()-funktion

async def behandel(item: Any) -> dict[str, Any]:
    """Behandl ét mail-item idempotent ved hjælp af Box-checkpoints."""

    data = _hent_item_data(item)
    config.validate_queue_data(data)

    mail_id = _paakraevet_tekst(
        data.get(config.QUEUE_DATA_MAIL_ID),
        config.QUEUE_DATA_MAIL_ID,
    )

    mail_type = _paakraevet_tekst(
        data.get(config.QUEUE_DATA_MAIL_TYPE),
        config.QUEUE_DATA_MAIL_TYPE,
    ).upper()

    box_data = _normaliser_box_data(
        _hent_box_data(item)
    )

    # Gem den normaliserede Box-struktur med det samme.
    await _gem_box_data(item, box_data)

    logger.info(
        "Starter itembehandling. mail_id=%s, mail_type=%s",
        mail_id,
        mail_type,
    )

    try:
        # ------------------------------------------------------------
        # DIGITALPOST-BETALING
        # ------------------------------------------------------------

        if mail_type == config.MAIL_TYPE_DIGITALPOST_BETALING:
            # Funktionen:
            # 1. henter maindoc.pdf
            # 2. læser registreringsnummeret
            # 3. finder køretøjet i Insubiz
            # 4. gemmer identifikator og køretøj i Box
            await _gem_digitalpost_oplysninger(
                item,
                box_data,
                mail_id,
            )

            koeretoej = box_data[config.BOX_KEY_VEHICLE]

            registreringsnummer = _normaliser_valgfri_tekst(
                koeretoej.get(config.BOX_KEY_REGISTRATION)
            )

            stelnummer = _normaliser_valgfri_tekst(
                koeretoej.get(config.BOX_KEY_CHASSIS_NUMBER)
            )

            afdelingsmail = _normaliser_email(
                koeretoej.get(config.BOX_KEY_DEPARTMENT_EMAIL)
            )

            beslutning = manuel_behandling.vurder_manuel_behandling(
                registreringsnummer=registreringsnummer,
                stelnummer=stelnummer,
                antal_fundne=1,
                afdelingsmail=afdelingsmail,
            )

            if beslutning is not None:
                return await _haandter_manuel_behandling(
                    item=item,
                    box_data=box_data,
                    beslutning=beslutning,
                )

            # Typekontrollen ovenfor har sikret, at afdelingsmail
            # er en gyldig mailadresse.
            assert afdelingsmail is not None

            await _videresend_og_slet(
                item=item,
                box_data=box_data,
                mail_id=mail_id,
                afdelingsmail=afdelingsmail,
            )

            await _afslut_item(
                item,
                status=config.STATUS_COMPLETED,
                state=None,
            )

            logger.info(
                "Digitalpost-mail afsluttet. "
                "mail_id=%s, modtager=%s",
                mail_id,
                afdelingsmail,
            )

            return box_data

        # ------------------------------------------------------------
        # ØVRIGE UNDERSTØTTEDE MAILTYPER
        # ------------------------------------------------------------

        mail = await _hent_mail(mail_id)

        identifikator = await _hent_eller_beregn_identifikator(
            item=item,
            box_data=box_data,
            mail=mail,
            mail_id=mail_id,
            mail_type=mail_type,
        )

        koeretoejer = await _hent_koeretoejer()

        try:
            koeretoej = _find_koeretoej(
                koeretoejer=koeretoejer,
                identifikator=identifikator,
            )

        except vehicle_matcher.KoeretoejIkkeFundetFejl:
            beslutning = manuel_behandling.vurder_manuel_behandling(
                registreringsnummer=None,
                stelnummer=None,
                antal_fundne=0,
                afdelingsmail=None,
            )

            return await _haandter_manuel_behandling(
                item=item,
                box_data=box_data,
                beslutning=_paakraevet_manuel_beslutning(
                    beslutning
                ),
            )

        except vehicle_matcher.FlereKoeretoejerFundetFejl:
            beslutning = manuel_behandling.vurder_manuel_behandling(
                registreringsnummer=None,
                stelnummer=None,
                antal_fundne=2,
                afdelingsmail=None,
            )

            return await _haandter_manuel_behandling(
                item=item,
                box_data=box_data,
                beslutning=_paakraevet_manuel_beslutning(
                    beslutning
                ),
            )

        _gem_koeretoej_i_box(
            box_data,
            koeretoej,
        )

        await _gem_box_data(
            item,
            box_data,
        )

        registreringsnummer = _normaliser_valgfri_tekst(
            koeretoej.get("nummerplade")
        )

        stelnummer = _normaliser_valgfri_tekst(
            koeretoej.get("stelnummer")
        )

        afdelingsmail = _normaliser_email(
            koeretoej.get("afdelingsmail")
        )

        beslutning = manuel_behandling.vurder_manuel_behandling(
            registreringsnummer=registreringsnummer,
            stelnummer=stelnummer,
            antal_fundne=1,
            afdelingsmail=afdelingsmail,
        )

        if beslutning is not None:
            return await _haandter_manuel_behandling(
                item=item,
                box_data=box_data,
                beslutning=beslutning,
            )

        # Typekontrollen ovenfor har sikret, at afdelingsmail
        # er en gyldig mailadresse.
        assert afdelingsmail is not None

        await _videresend_og_slet(
            item=item,
            box_data=box_data,
            mail_id=mail_id,
            afdelingsmail=afdelingsmail,
        )

        await _afslut_item(
            item,
            status=config.STATUS_COMPLETED,
            state=None,
        )

        logger.info(
            "Item afsluttet. mail_id=%s, modtager=%s",
            mail_id,
            afdelingsmail,
        )

        return box_data

    except Forretningsfejl:
        raise

    except Systemfejl:
        raise

    except Exception as error:
        logger.exception(
            "Teknisk fejl under behandling af mail_id=%s",
            mail_id,
        )

        raise Systemfejl(
            f"Teknisk fejl under behandling af mail {mail_id}"
        ) from error

# SLUT: behandel()-funktionen

async def _hent_mail(mail_id: str) -> dict[str, Any]:
    try:
        return await outlook_service.hent_mail(mail_id)
    except outlook_service.OutlookMailNotFoundError as error:
        raise Forretningsfejl(config.SYSTEM_ERROR_MAIL_NOT_FOUND) from error

async def _hent_eller_beregn_identifikator(
    *,
    item: Any,
    box_data: dict[str, Any],
    mail: Mapping[str, Any],
    mail_id: str,
    mail_type: str,
) -> dict[str, str]:
    gemt = box_data[config.BOX_KEY_IDENTIFIER]
    gemt_type = gemt.get(config.BOX_KEY_IDENTIFIER_TYPE)
    gemt_vaerdi = gemt.get(config.BOX_KEY_IDENTIFIER_VALUE)

    if gemt_type and gemt_vaerdi:
        return {
            "type": str(gemt_type),
            "vaerdi": str(gemt_vaerdi),
        }

    try:
        pdf = await outlook_service.hent_eneste_pdf(mail_id=mail_id)
    except outlook_service.OutlookAttachmentError as error:
        raise Forretningsfejl(str(error)) from error

    try:
        resultat = udtraek_stel_eller_nummerplade(
            mail_type,
            pdf["content"],
        )
    except IdentifikatorIkkeFundetFejl as error:
        raise Forretningsfejl(
            config.BUSINESS_ERROR_IDENTIFIER_NOT_FOUND
        ) from error
    except TvetydigIdentifikatorFejl as error:
        raise Forretningsfejl(str(error)) from error
    except PdfBeregningFejl as error:
        raise Forretningsfejl(str(error)) from error

    identifikator = {
        "type": resultat.type,
        "vaerdi": resultat.vaerdi,
    }

    box_data[config.BOX_KEY_IDENTIFIER] = {
        config.BOX_KEY_IDENTIFIER_TYPE: identifikator["type"],
        config.BOX_KEY_IDENTIFIER_VALUE: identifikator["vaerdi"],
        config.BOX_KEY_ATTACHMENT_NAME: pdf["name"],
    }
    await _gem_box_data(item, box_data)

    return identifikator

async def _gem_digitalpost_oplysninger(
    item: Any,
    box_data: dict[str, Any],
    mail_id: str,
) -> None:
    """Læs maindoc.pdf og find køretøjet via nummerplade, stelnummer eller begge."""
    identifikator = box_data[config.BOX_KEY_IDENTIFIER]
    nummerplade = _normaliser_valgfri_tekst(
        identifikator.get(config.BOX_KEY_REGISTRATION)
    )
    stelnummer = _normaliser_valgfri_tekst(
        identifikator.get(config.BOX_KEY_CHASSIS_NUMBER)
    )

    if not nummerplade and not stelnummer:
        pdf_filer = await outlook_service.hent_pdf_vedhaeftninger(
            mail_id=mail_id
        )
        maindoc = [
            pdf
            for pdf in pdf_filer
            if str(pdf.get("name", "")).strip().casefold()
            == "maindoc.pdf"
        ]
        if len(maindoc) != 1:
            raise Forretningsfejl("Forventede præcis én maindoc.pdf.")

        fundne = _laes_digitalpost_identifikatorer(maindoc[0]["content"])
        nummerplade = fundne.get("nummerplade")
        stelnummer = fundne.get("stelnummer")

        primartype = "nummerplade" if nummerplade else "stelnummer"
        primaervaerdi = nummerplade or stelnummer
        box_data[config.BOX_KEY_IDENTIFIER] = {
            config.BOX_KEY_IDENTIFIER_TYPE: primartype,
            config.BOX_KEY_IDENTIFIER_VALUE: primaervaerdi,
            config.BOX_KEY_REGISTRATION: nummerplade,
            config.BOX_KEY_CHASSIS_NUMBER: stelnummer,
            config.BOX_KEY_ATTACHMENT_NAME: maindoc[0]["name"],
        }
        await _gem_box_data(item, box_data)

    if box_data[config.BOX_KEY_VEHICLE].get(
        config.BOX_KEY_VEHICLE_FOUND
    ):
        return

    koeretoejer = await _hent_koeretoejer()
    koeretoej = _find_digitalpost_koeretoej(
        koeretoejer=koeretoejer,
        nummerplade=nummerplade,
        stelnummer=stelnummer,
    )
    _gem_koeretoej_i_box(box_data, koeretoej)
    await _gem_box_data(item, box_data)


def _laes_digitalpost_identifikatorer(content: bytes) -> dict[str, str | None]:
    """Udtræk højst ét registreringsnummer og højst ét stelnummer fra PDF'en."""
    if not isinstance(content, bytes) or not content.startswith(b"%PDF-"):
        raise Forretningsfejl(
            "maindoc.pdf indeholder ikke gyldige PDF-bytes."
        )
    try:
        tekst = "\n".join(
            page.extract_text() or ""
            for page in PdfReader(io.BytesIO(content)).pages
        )
    except Exception as error:
        raise Forretningsfejl("maindoc.pdf kunne ikke læses.") from error

    nummerplader = {
        _normaliser_koeretoejsidentifikator(match.group(1))
        for match in re.finditer(
            r"(?:Registreringsnummer|Registreringsnr|Reg\.?\s*nr\.?)"
            r"\s*:\s*([A-Z0-9][A-Z0-9 -]{3,10}[A-Z0-9])",
            tekst,
            flags=re.IGNORECASE,
        )
    }
    nummerplader = {
        value
        for value in nummerplader
        if re.fullmatch(config.REGISTRATION_PATTERN, value)
    }

    stelnumre = {
        _normaliser_koeretoejsidentifikator(match.group(1))
        for match in re.finditer(
            r"(?:Stelnummer|Stelnr\.?|Stel\s*nr\.?|VIN)"
            r"\s*:\s*([A-HJ-NPR-Z0-9]{17})",
            tekst,
            flags=re.IGNORECASE,
        )
    }
    stelnumre = {
        value
        for value in stelnumre
        if re.fullmatch(config.VIN_PATTERN, value)
    }

    if len(nummerplader) > 1:
        raise Forretningsfejl(
            "maindoc.pdf indeholder flere forskellige registreringsnumre: "
            + ", ".join(sorted(nummerplader))
        )
    if len(stelnumre) > 1:
        raise Forretningsfejl(
            "maindoc.pdf indeholder flere forskellige stelnumre: "
            + ", ".join(sorted(stelnumre))
        )
    if not nummerplader and not stelnumre:
        raise Forretningsfejl(
            "maindoc.pdf indeholder hverken et entydigt "
            "registreringsnummer eller stelnummer."
        )

    return {
        "nummerplade": next(iter(nummerplader), None),
        "stelnummer": next(iter(stelnumre), None),
    }


def _normaliser_koeretoejsidentifikator(value: str) -> str:
    """Fjern mellemrum og bindestreger og anvend store bogstaver."""
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def _find_digitalpost_koeretoej(
    *,
    koeretoejer: list[dict[str, Any]],
    nummerplade: str | None,
    stelnummer: str | None,
) -> dict[str, Any]:
    """Find ét køretøj via en eller begge identifikatorer."""
    fundne: list[tuple[str, dict[str, Any]]] = []
    fejl: list[str] = []

    for identifikator_type, identifikator_vaerdi in (
        ("nummerplade", nummerplade),
        ("stelnummer", stelnummer),
    ):
        if not identifikator_vaerdi:
            continue
        try:
            koeretoej = _find_koeretoej(
                koeretoejer=koeretoejer,
                identifikator={
                    "type": identifikator_type,
                    "vaerdi": identifikator_vaerdi,
                },
            )
        except vehicle_matcher.KoeretoejIkkeFundetFejl:
            fejl.append(
                f"{identifikator_type}={identifikator_vaerdi} gav intet match"
            )
            continue
        except vehicle_matcher.FlereKoeretoejerFundetFejl as error:
            raise Forretningsfejl(
                "Flere køretøjer i Insubiz matcher "
                f"{identifikator_type}={identifikator_vaerdi}."
            ) from error
        fundne.append((identifikator_type, koeretoej))

    if not fundne:
        detalje = "; ".join(fejl) or "ingen identifikatorer kunne anvendes"
        raise Forretningsfejl(
            "Intet entydigt køretøj i Insubiz. " + detalje + "."
        )

    if len(fundne) == 2 and not _samme_koeretoej(
        fundne[0][1],
        fundne[1][1],
    ):
        raise Forretningsfejl(
            "Registreringsnummer og stelnummer fra maindoc.pdf matcher "
            "to forskellige køretøjer i Insubiz."
        )

    return fundne[0][1]


def _samme_koeretoej(
    foerste: Mapping[str, Any],
    anden: Mapping[str, Any],
) -> bool:
    """Sammenlign køretøjer på normaliseret nummerplade og stelnummer."""
    for felt in ("nummerplade", "stelnummer"):
        foerste_vaerdi = _normaliser_valgfri_tekst(foerste.get(felt))
        anden_vaerdi = _normaliser_valgfri_tekst(anden.get(felt))
        if foerste_vaerdi and anden_vaerdi:
            if (
                _normaliser_koeretoejsidentifikator(foerste_vaerdi)
                != _normaliser_koeretoejsidentifikator(anden_vaerdi)
            ):
                return False
    return True

async def _hent_koeretoejer() -> list[dict[str, Any]]:
    try:
        koeretoejer = await insubiz_service.hent_koeretoejer_liste()
    except Exception as error:
        raise Systemfejl(
            "Kunne ikke hente køretøjslisten fra Insubiz"
        ) from error

    if not koeretoejer:
        raise Systemfejl("Insubiz returnerede en tom køretøjsliste")

    return koeretoejer

def _find_koeretoej(
    *,
    koeretoejer: list[dict[str, Any]],
    identifikator: Mapping[str, str],
) -> dict[str, Any]:
    """Finder ét entydigt køretøj og bevarer matcherens fejltyper."""
    return vehicle_matcher.find_koeretoej(
        koeretoejer=koeretoejer,
        identifikator_type=identifikator["type"],
        identifikator_vaerdi=identifikator["vaerdi"],
    )

def _gem_koeretoej_i_box(
    box_data: dict[str, Any],
    koeretoej: Mapping[str, Any],
) -> None:
    box_data[config.BOX_KEY_VEHICLE] = {
        config.BOX_KEY_VEHICLE_FOUND: True,
        config.BOX_KEY_REGISTRATION: koeretoej.get("nummerplade"),
        config.BOX_KEY_CHASSIS_NUMBER: koeretoej.get("stelnummer"),
        config.BOX_KEY_DEPARTMENT_EMAIL: koeretoej.get("afdelingsmail"),
    }

async def _haandter_manuel_behandling(
    *,
    item: Any,
    box_data: dict[str, Any],
    beslutning: manuel_behandling.ManuelBehandling,
) -> dict[str, Any]:
    """Send én manuel mail med køretøjsdata og afslut itemet manuelt."""
    levering = box_data[config.BOX_KEY_DELIVERY]
    koeretoej = box_data.get(config.BOX_KEY_VEHICLE, {})
    if not isinstance(koeretoej, Mapping):
        koeretoej = {}
    nummerplade = _normaliser_valgfri_tekst(
        koeretoej.get(config.BOX_KEY_REGISTRATION)
    )
    stelnummer = _normaliser_valgfri_tekst(
        koeretoej.get(config.BOX_KEY_CHASSIS_NUMBER)
    )
    if not levering.get(config.BOX_KEY_MANUAL_EMAIL_SENT, False):
        sendt = await outlook_service.send_mail_til_forsikring(
            emne=beslutning.emne,
            besked=beslutning.besked,
            manuel_id=int(beslutning.manuel_id),
            nummerplade=nummerplade,
            stelnummer=stelnummer,
        )
        if not sendt:
            raise Systemfejl(
                "Outlook bekræftede ikke mailen om manuel behandling. "
                f"Manuel-id: {beslutning.manuel_id}."
            )
        levering[config.BOX_KEY_MANUAL_EMAIL_SENT] = True
        levering[config.BOX_KEY_MANUAL_EMAIL_SENT_AT] = _utc_now()
        levering[config.BOX_KEY_MANUAL_ID] = beslutning.manuel_id
        await _gem_box_data(item, box_data)
    await _afslut_item(
        item,
        status=config.STATUS_MANUAL,
        state=(
            f"{config.STATE_MANUEL_BEHANDLING} - "
            f"{beslutning.manuel_id}: {beslutning.emne}"
        ),
    )
    logger.info(
        "Item sendt til manuel behandling. manuel_id=%s, emne=%s, "
        "nummerplade=%s, stelnummer=%s",
        beslutning.manuel_id,
        beslutning.emne,
        nummerplade or "Mangler",
        stelnummer or "Mangler",
    )
    return box_data

def _paakraevet_manuel_beslutning(
    value: manuel_behandling.ManuelBehandling | None,
) -> manuel_behandling.ManuelBehandling:
    if value is None:
        raise Systemfejl("Forventede en beslutning om manuel behandling.")
    return value

async def _videresend_og_slet(
    *,
    item: Any,
    box_data: dict[str, Any],
    mail_id: str,
    afdelingsmail: str,
) -> None:
    levering = box_data[config.BOX_KEY_DELIVERY]

    # -------------------------------------------------
    # VIDERESEND MAIL
    # -------------------------------------------------

    if not levering.get(
        config.BOX_KEY_SENT_TO_DEPARTMENT,
        False,
    ):
        from q_outlook_api.functionality.mail_api import (
            forward_mail,
        )

        resultat = forward_mail(
            user_mail=config.OUTLOOK_MAILBOX_ADDRESS,
            message_id=mail_id,
            forward={
                "to": [afdelingsmail],
                "subject": "",
                "body": config.FORWARD_BODY_HTML,
                "forward_mode": "formatted",
            },
        )

        if not resultat.get("success"):
            raise Systemfejl(
                f"Outlook bekræftede ikke videresendelse til "
                f"{afdelingsmail}"
            )

        # Gem checkpoint FØR sletning
        levering[
            config.BOX_KEY_SENT_TO_DEPARTMENT
        ] = True

        levering[
            config.BOX_KEY_DEPARTMENT_SENT_AT
        ] = _utc_now()

        await _gem_box_data(
            item,
            box_data,
        )

    # -------------------------------------------------
    # FLYT TIL SLETTET POST
    # -------------------------------------------------

    if not levering.get(
        config.BOX_KEY_MAIL_DELETED,
        False,
    ):
        from q_outlook_api.functionality.mail_api import (
            delete_mail,
        )

        resultat = delete_mail(
            user_mail=config.OUTLOOK_MAILBOX_ADDRESS,
            message_id=mail_id,
            permanent_delete=False,
        )

        if not resultat.get("success"):
            raise Systemfejl(
                f"Mailen kunne ikke flyttes til "
                f"Slettet post. Mail-id: {mail_id}"
            )

        levering[
            config.BOX_KEY_MAIL_DELETED
        ] = True

        await _gem_box_data(
            item,
            box_data,
        )

        logger.info(
            "Mail videresendt og flyttet til "
            "Slettet post. mail_id=%s",
            mail_id,
        )

# ---------------------------------------------------------------------------
# ATS-ADAPTERE
# ---------------------------------------------------------------------------

def _hent_item_data(item: Any) -> dict[str, Any]:
    value = _hent_vaerdi(
        item,
        "data",
        "item_data",
        "payload",
    )

    if callable(value):
        value = value()

    if not isinstance(value, Mapping):
        raise TypeError(
            "Kø-itemet mangler dictionary-feltet data/item_data/payload."
        )

    return dict(value)

def _hent_box_data(item: Any) -> dict[str, Any]:
    data = _hent_item_data(item)
    box = data.get("box")
    if box is None:
        return {}
    if not isinstance(box, Mapping):
        raise TypeError("item.data['box'] skal være en dictionary.")
    return deepcopy(dict(box))

async def _gem_box_data(item: Any, box_data: dict[str, Any]) -> None:
    """Gem hele item.data inklusive box gennem ATS-klientens update()."""
    data = _hent_item_data(item)
    data["box"] = deepcopy(box_data)
    update = getattr(item, "update", None)
    if not callable(update):
        raise Systemfejl("ATS-itemet mangler update(data); Box kan ikke gemmes.")
    await _maybe_await(update(data))
    if isinstance(item, MutableMapping):
        item["data"] = data

def _kald_med_box_data(method: Any, box_data: dict[str, Any]) -> Any:
    signature = inspect.signature(method)
    parameters = signature.parameters

    for name in ("box_data", "data", "value"):
        if name in parameters:
            return method(**{name: deepcopy(box_data)})

    return method(deepcopy(box_data))

async def _afslut_item(
    item: Any,
    *,
    status: str,
    state: str | None,
) -> None:
    """Gem forretningsresultatet; main.py afslutter ATS-itemet."""
    from q_haderslev_vbo.automation_server.ats_update_item_data import (
        update_item_data,
    )

    if status not in {config.STATUS_COMPLETED, config.STATUS_MANUAL}:
        raise Systemfejl(
            f"Ukendt forretningsstatus ved afslutning: {status!r}"
        )

    data = _hent_item_data(item)
    besked = state or status
    update_item_data(
        data,
        item=item,
        status=status,
        status_code=besked,
        state=besked,
    )
    logger.info(
        "Forretningsresultat gemt. status=%s, besked=%s",
        status,
        besked,
    )

async def _maybe_await(value: Any) -> Any:
    if inspect.isawaitable(value):
        return await value
    return value

def _hent_vaerdi(item: Any, *names: str) -> Any:
    if isinstance(item, Mapping):
        for name in names:
            if name in item:
                return item[name]
        return None

    for name in names:
        if hasattr(item, name):
            return getattr(item, name)

    return None

# ---------------------------------------------------------------------------
# GENERELLE HJÆLPERE
# ---------------------------------------------------------------------------

def _normaliser_box_data(value: Mapping[str, Any]) -> dict[str, Any]:
    box_data = deepcopy(dict(value))
    box_data.setdefault(config.BOX_KEY_SCHEMA_VERSION, config.BOX_SCHEMA_VERSION)
    box_data.setdefault(config.BOX_KEY_IDENTIFIER, {})
    box_data.setdefault(
        config.BOX_KEY_VEHICLE,
        {config.BOX_KEY_VEHICLE_FOUND: False},
    )
    box_data.setdefault(
        config.BOX_KEY_DELIVERY,
        {
            config.BOX_KEY_SENT_TO_DEPARTMENT: False,
            config.BOX_KEY_MAIL_DELETED: False,
            config.BOX_KEY_MANUAL_EMAIL_SENT: False,
        },
    )
    return box_data

def _normaliser_valgfri_tekst(value: Any) -> str | None:
    if value in (None, ""):
        return None
    normalized = str(value).strip()
    return normalized or None

def _normaliser_email(value: Any) -> str | None:
    if value in (None, ""):
        return None
    email = str(value).strip()
    if not email:
        return None
    if email.count("@") != 1 or any(character.isspace() for character in email):
        return None
    local_part, domain = email.rsplit("@", 1)
    if not local_part or not domain or "." not in domain:
        return None
    return email

def _paakraevet_tekst(value: Any, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise Forretningsfejl(
            f"Item data mangler det obligatoriske felt {field_name!r}."
        )
    return value.strip()

def _mail_id(mail: Mapping[str, Any]) -> str:
    return str(
        mail.get("mail_id")
        or mail.get("message_id")
        or mail.get("id")
        or ""
    ).strip()

def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")

__all__ = [
    "Behandlingsfejl",
    "Forretningsfejl",
    "Systemfejl",
    "behandel",
    "process_item",
]
