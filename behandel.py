"""Behandling af ét ATS-kø-item for mails-i-forsikring.

Processtrukturen må ikke brydes:

1. populate_queue.py opretter kø-items.
2. main.py henter ét item ad gangen.
3. main.py kalder process_item(item).
4. Denne fil behandler præcis dét ene item.

Checkpoints gemmes i itemets Box data og reducerer risikoen for gentagen
videresendelse ved genkørsel.

Outlook-authentication håndteres internt af q_outlook_api via
outlook_service.py.
"""

from __future__ import annotations

import inspect
import io
import logging
import re
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

    Behandler præcis ét kø-item og returnerer det afsluttende Box data.
    main.py ejer fortsat kø-loopet og den overordnede Automation
    Server-livscyklus.
    """
    return await behandel(item)


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

    box_data = _normaliser_box_data(_hent_box_data(item))

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
            antal_fundne = await _gem_digitalpost_oplysninger(
                item,
                box_data,
                mail_id,
            )

            # Intet køretøjsmatch:
            # Afslut itemet som Completed uden at sende eller flytte mail.
            if antal_fundne == 0:
                return await _afslut_uden_koeretoejsmatch(
                    item=item,
                    box_data=box_data,
                )

            # Flere køretøjsmatch bruger fortsat den eksisterende
            # manuelle behandling, som sender informationsmail.
            if antal_fundne > 1:
                beslutning = manuel_behandling.vurder_manuel_behandling(
                    registreringsnummer=None,
                    stelnummer=None,
                    antal_fundne=antal_fundne,
                    afdelingsmail=None,
                )
                return await _haandter_beslutning(
                    item=item,
                    box_data=box_data,
                    beslutning=_paakraevet_manuel_beslutning(beslutning),
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
                return await _haandter_beslutning(
                    item=item,
                    box_data=box_data,
                    beslutning=beslutning,
                )

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
                "Digitalpost-mail afsluttet. mail_id=%s, modtager=%s",
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
            _gem_identifikator_som_ufundet_koeretoej(
                box_data=box_data,
                identifikator=identifikator,
            )
            await _gem_box_data(item, box_data)

            return await _afslut_uden_koeretoejsmatch(
                item=item,
                box_data=box_data,
            )

        except vehicle_matcher.FlereKoeretoejerFundetFejl:
            _gem_identifikator_som_ufundet_koeretoej(
                box_data=box_data,
                identifikator=identifikator,
            )
            await _gem_box_data(item, box_data)

            beslutning = manuel_behandling.vurder_manuel_behandling(
                registreringsnummer=None,
                stelnummer=None,
                antal_fundne=2,
                afdelingsmail=None,
            )
            return await _haandter_beslutning(
                item=item,
                box_data=box_data,
                beslutning=_paakraevet_manuel_beslutning(beslutning),
            )

        _gem_koeretoej_i_box(box_data, koeretoej)
        await _gem_box_data(item, box_data)

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
            return await _haandter_beslutning(
                item=item,
                box_data=box_data,
                beslutning=beslutning,
            )

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


async def _hent_mail(mail_id: str) -> dict[str, Any]:
    """Hent én mail fra Outlook."""
    try:
        return await outlook_service.hent_mail(mail_id)
    except outlook_service.OutlookMailNotFoundError as error:
        raise Forretningsfejl(
            config.SYSTEM_ERROR_MAIL_NOT_FOUND
        ) from error


async def _hent_eller_beregn_identifikator(
    *,
    item: Any,
    box_data: dict[str, Any],
    mail: Mapping[str, Any],
    mail_id: str,
    mail_type: str,
) -> dict[str, str]:
    """Hent identifikatoren fra Box eller beregn den fra PDF-filen."""
    del mail

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
) -> int:
    """Læs maindoc.pdf og find køretøjet i Insubiz.

    Output:
        0 hvis ingen køretøjer matcher.
        1 hvis der findes ét entydigt køretøj.
        2 hvis flere køretøjer matcher, eller nummerplade og stelnummer
        peger på forskellige køretøjer.
    """
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
            raise Forretningsfejl(
                "Forventede præcis én maindoc.pdf."
            )

        fundne = _laes_digitalpost_identifikatorer(
            maindoc[0]["content"]
        )
        nummerplade = fundne.get("nummerplade")
        stelnummer = fundne.get("stelnummer")
        primaer_type = "nummerplade" if nummerplade else "stelnummer"
        primaer_vaerdi = nummerplade or stelnummer

        box_data[config.BOX_KEY_IDENTIFIER] = {
            config.BOX_KEY_IDENTIFIER_TYPE: primaer_type,
            config.BOX_KEY_IDENTIFIER_VALUE: primaer_vaerdi,
            config.BOX_KEY_REGISTRATION: nummerplade,
            config.BOX_KEY_CHASSIS_NUMBER: stelnummer,
            config.BOX_KEY_ATTACHMENT_NAME: maindoc[0]["name"],
        }
        await _gem_box_data(item, box_data)

    if box_data[config.BOX_KEY_VEHICLE].get(
        config.BOX_KEY_VEHICLE_FOUND
    ):
        return 1

    koeretoejer = await _hent_koeretoejer()

    try:
        koeretoej = _find_digitalpost_koeretoej(
            koeretoejer=koeretoejer,
            nummerplade=nummerplade,
            stelnummer=stelnummer,
        )
    except vehicle_matcher.KoeretoejIkkeFundetFejl:
        _gem_ufundet_digitalpost_koeretoej_i_box(
            box_data=box_data,
            nummerplade=nummerplade,
            stelnummer=stelnummer,
        )
        await _gem_box_data(item, box_data)

        logger.info(
            "Intet køretøjsmatch i Insubiz. "
            "Itemet afsluttes som Completed uden mailhandling. "
            "nummerplade=%s, stelnummer=%s",
            nummerplade or "Mangler",
            stelnummer or "Mangler",
        )
        return 0

    except vehicle_matcher.FlereKoeretoejerFundetFejl:
        _gem_ufundet_digitalpost_koeretoej_i_box(
            box_data=box_data,
            nummerplade=nummerplade,
            stelnummer=stelnummer,
        )
        await _gem_box_data(item, box_data)

        logger.info(
            "Flere køretøjer matcher identifikatoren i Insubiz. "
            "Itemet sendes til manuel behandling. "
            "nummerplade=%s, stelnummer=%s",
            nummerplade or "Mangler",
            stelnummer or "Mangler",
        )
        return 2

    except DigitalpostIdentifikatorKonfliktFejl:
        _gem_ufundet_digitalpost_koeretoej_i_box(
            box_data=box_data,
            nummerplade=nummerplade,
            stelnummer=stelnummer,
        )
        await _gem_box_data(item, box_data)

        logger.info(
            "Nummerplade og stelnummer matcher forskellige køretøjer. "
            "Itemet sendes til manuel behandling. "
            "nummerplade=%s, stelnummer=%s",
            nummerplade or "Mangler",
            stelnummer or "Mangler",
        )
        return 2

    _gem_koeretoej_i_box(box_data, koeretoej)
    await _gem_box_data(item, box_data)
    return 1


def _laes_digitalpost_identifikatorer(
    content: bytes,
) -> dict[str, str | None]:
    """Udtræk højst ét registreringsnummer og ét stelnummer fra PDF'en."""
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
        raise Forretningsfejl(
            "maindoc.pdf kunne ikke læses."
        ) from error

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
            "maindoc.pdf indeholder flere forskellige "
            "registreringsnumre: "
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


class DigitalpostIdentifikatorKonfliktFejl(Forretningsfejl):
    """Nummerplade og stelnummer peger på forskellige køretøjer."""


def _find_digitalpost_koeretoej(
    *,
    koeretoejer: list[dict[str, Any]],
    nummerplade: str | None,
    stelnummer: str | None,
) -> dict[str, Any]:
    """Find ét køretøj via nummerplade, stelnummer eller begge."""
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
                f"{identifikator_type}={identifikator_vaerdi} "
                "gav intet match"
            )
            continue
        except vehicle_matcher.FlereKoeretoejerFundetFejl:
            raise

        fundne.append((identifikator_type, koeretoej))

    if not fundne:
        detalje = "; ".join(fejl) or "ingen identifikatorer kunne anvendes"
        raise vehicle_matcher.KoeretoejIkkeFundetFejl(
            "Intet entydigt køretøj i Insubiz. "
            f"{detalje}."
        )

    if len(fundne) == 2 and not _samme_koeretoej(
        fundne[0][1],
        fundne[1][1],
    ):
        raise DigitalpostIdentifikatorKonfliktFejl(
            "Registreringsnummer og stelnummer fra maindoc.pdf "
            "matcher to forskellige køretøjer i Insubiz."
        )

    return fundne[0][1]


def _samme_koeretoej(
    foerste: Mapping[str, Any],
    anden: Mapping[str, Any],
) -> bool:
    """Sammenlign køretøjer på nummerplade og stelnummer."""
    sammenlignede_felter = 0

    for felt in ("nummerplade", "stelnummer"):
        foerste_vaerdi = _normaliser_valgfri_tekst(foerste.get(felt))
        anden_vaerdi = _normaliser_valgfri_tekst(anden.get(felt))

        if foerste_vaerdi and anden_vaerdi:
            sammenlignede_felter += 1
            if (
                _normaliser_koeretoejsidentifikator(foerste_vaerdi)
                != _normaliser_koeretoejsidentifikator(anden_vaerdi)
            ):
                return False

    return sammenlignede_felter > 0


async def _hent_koeretoejer() -> list[dict[str, Any]]:
    """Hent køretøjslisten fra Insubiz."""
    try:
        koeretoejer = await insubiz_service.hent_koeretoejer_liste()
    except Exception as error:
        raise Systemfejl(
            "Kunne ikke hente køretøjslisten fra Insubiz"
        ) from error

    if not koeretoejer:
        raise Systemfejl(
            "Insubiz returnerede en tom køretøjsliste"
        )

    return koeretoejer


def _find_koeretoej(
    *,
    koeretoejer: list[dict[str, Any]],
    identifikator: Mapping[str, str],
) -> dict[str, Any]:
    """Find ét entydigt køretøj og bevar matcherens fejltyper."""
    return vehicle_matcher.find_koeretoej(
        koeretoejer=koeretoejer,
        identifikator_type=identifikator["type"],
        identifikator_vaerdi=identifikator["vaerdi"],
    )


def _gem_koeretoej_i_box(
    box_data: dict[str, Any],
    koeretoej: Mapping[str, Any],
) -> None:
    """Gem et fundet køretøj i Box."""
    box_data[config.BOX_KEY_VEHICLE] = {
        config.BOX_KEY_VEHICLE_FOUND: True,
        config.BOX_KEY_REGISTRATION: koeretoej.get("nummerplade"),
        config.BOX_KEY_CHASSIS_NUMBER: koeretoej.get("stelnummer"),
        config.BOX_KEY_DEPARTMENT_EMAIL: koeretoej.get("afdelingsmail"),
    }


def _gem_ufundet_digitalpost_koeretoej_i_box(
    *,
    box_data: dict[str, Any],
    nummerplade: str | None,
    stelnummer: str | None,
) -> None:
    """Gem PDF-identifikatorerne, selvom køretøjet ikke blev fundet."""
    box_data[config.BOX_KEY_VEHICLE] = {
        config.BOX_KEY_VEHICLE_FOUND: False,
        config.BOX_KEY_REGISTRATION: nummerplade,
        config.BOX_KEY_CHASSIS_NUMBER: stelnummer,
        config.BOX_KEY_DEPARTMENT_EMAIL: None,
    }


def _gem_identifikator_som_ufundet_koeretoej(
    *,
    box_data: dict[str, Any],
    identifikator: Mapping[str, str],
) -> None:
    """Gem identifikatoren, så den kan ses på det manuelle item."""
    identifikator_type = str(
        identifikator.get("type", "")
    ).strip().casefold()
    identifikator_vaerdi = _normaliser_valgfri_tekst(
        identifikator.get("vaerdi")
    )

    nummerplade = (
        identifikator_vaerdi
        if identifikator_type == "nummerplade"
        else None
    )
    stelnummer = (
        identifikator_vaerdi
        if identifikator_type == "stelnummer"
        else None
    )

    box_data[config.BOX_KEY_VEHICLE] = {
        config.BOX_KEY_VEHICLE_FOUND: False,
        config.BOX_KEY_REGISTRATION: nummerplade,
        config.BOX_KEY_CHASSIS_NUMBER: stelnummer,
        config.BOX_KEY_DEPARTMENT_EMAIL: None,
    }


async def _afslut_uden_koeretoejsmatch(
    *,
    item: Any,
    box_data: dict[str, Any],
) -> dict[str, Any]:
    """Afslut itemet, når køretøjet ikke findes i Insubiz.

    Der sendes ingen mail. Originalmailen bliver hverken videresendt,
    slettet eller ændret. main.py omsætter STATUS_COMPLETED til den
    tekniske ATS-status COMPLETED og fortsætter med næste kø-item.
    """
    koeretoej = box_data.get(config.BOX_KEY_VEHICLE, {})
    if not isinstance(koeretoej, Mapping):
        koeretoej = {}

    nummerplade = _normaliser_valgfri_tekst(
        koeretoej.get(config.BOX_KEY_REGISTRATION)
    )
    stelnummer = _normaliser_valgfri_tekst(
        koeretoej.get(config.BOX_KEY_CHASSIS_NUMBER)
    )

    await _gem_box_data(item, box_data)

    await _afslut_item(
        item,
        status=config.STATUS_COMPLETED,
        state="Intet køretøjsmatch i Insubiz",
    )

    logger.info(
        "Item afsluttet uden køretøjsmatch. "
        "Der er ikke sendt, videresendt eller slettet nogen mail. "
        "nummerplade=%s, stelnummer=%s",
        nummerplade or "Mangler",
        stelnummer or "Mangler",
    )

    return box_data


def _manuel_beslutning_skal_afsluttes_uden_mail(
    beslutning: manuel_behandling.ManuelBehandling,
) -> bool:
    """Returnér True, når beslutningen kun skal afslutte itemet.

    Manglende registreringsnummer eller stelnummer skal ikke udløse en
    informationsmail. Itemet afsluttes som Completed, og originalmailen
    bliver liggende urørt.
    """
    return beslutning.manuel_id in {
        manuel_behandling.MANUEL_ID_MANGLENDE_REGISTRERINGSNUMMER,
        manuel_behandling.MANUEL_ID_MANGLENDE_STELNUMMER,
    }


async def _afslut_beslutning_uden_mail(
    *,
    item: Any,
    box_data: dict[str, Any],
    beslutning: manuel_behandling.ManuelBehandling,
) -> dict[str, Any]:
    """Afslut en tidligere manuel beslutning som Completed uden mail.

    Der sendes ingen informationsmail. Originalmailen bliver hverken
    videresendt, slettet eller ændret.
    """
    levering = box_data[config.BOX_KEY_DELIVERY]
    levering[config.BOX_KEY_MANUAL_ID] = beslutning.manuel_id
    levering[config.BOX_KEY_MANUAL_EMAIL_SENT] = False

    await _gem_box_data(item, box_data)

    await _afslut_item(
        item,
        status=config.STATUS_COMPLETED,
        state=f"Afsluttet - {beslutning.emne}",
    )

    logger.info(
        "Item afsluttet uden mail. manuel_id=%s, emne=%s",
        beslutning.manuel_id,
        beslutning.emne,
    )

    return box_data


async def _haandter_beslutning(
    *,
    item: Any,
    box_data: dict[str, Any],
    beslutning: manuel_behandling.ManuelBehandling,
) -> dict[str, Any]:
    """Afslut en manuel beslutning efter de gældende forretningsregler.

    Manglende nummerplade eller stelnummer afsluttes uden mail.
    Øvrige manuelle beslutninger beholder deres informationsmail, men
    ATS-itemet afsluttes altid som Completed.
    """
    if _manuel_beslutning_skal_afsluttes_uden_mail(beslutning):
        return await _afslut_beslutning_uden_mail(
            item=item,
            box_data=box_data,
            beslutning=beslutning,
        )

    return await _haandter_manuel_behandling(
        item=item,
        box_data=box_data,
        beslutning=beslutning,
    )


async def _haandter_manuel_behandling(
    *,
    item: Any,
    box_data: dict[str, Any],
    beslutning: manuel_behandling.ManuelBehandling,
) -> dict[str, Any]:
    """Send én manuel mail og afslut itemet til manuel behandling."""
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
        status=config.STATUS_COMPLETED,
        state=(
            f"Afsluttet - {beslutning.manuel_id}: "
            f"{beslutning.emne}"
        ),
    )

    logger.info(
        "Item afsluttet efter manuel regel. "
        "manuel_id=%s, emne=%s, nummerplade=%s, stelnummer=%s",
        beslutning.manuel_id,
        beslutning.emne,
        nummerplade or "Mangler",
        stelnummer or "Mangler",
    )

    return box_data


def _paakraevet_manuel_beslutning(
    value: manuel_behandling.ManuelBehandling | None,
) -> manuel_behandling.ManuelBehandling:
    """Returnér beslutningen eller fejl, hvis den mangler."""
    if value is None:
        raise Systemfejl(
            "Forventede en beslutning om manuel behandling."
        )
    return value


async def _videresend_og_slet(
    *,
    item: Any,
    box_data: dict[str, Any],
    mail_id: str,
    afdelingsmail: str,
) -> None:
    """Videresend mailen og flyt originalen til Slettet post."""
    levering = box_data[config.BOX_KEY_DELIVERY]

    if not levering.get(
        config.BOX_KEY_SENT_TO_DEPARTMENT,
        False,
    ):
        from q_outlook_api.functionality.mail_api import forward_mail

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
                "Outlook bekræftede ikke videresendelse "
                f"til {afdelingsmail}"
            )

        levering[config.BOX_KEY_SENT_TO_DEPARTMENT] = True
        levering[config.BOX_KEY_DEPARTMENT_SENT_AT] = _utc_now()
        await _gem_box_data(item, box_data)

    if not levering.get(
        config.BOX_KEY_MAIL_DELETED,
        False,
    ):
        from q_outlook_api.functionality.mail_api import delete_mail

        resultat = delete_mail(
            user_mail=config.OUTLOOK_MAILBOX_ADDRESS,
            message_id=mail_id,
            permanent_delete=False,
        )

        if not resultat.get("success"):
            raise Systemfejl(
                "Mailen kunne ikke flyttes til Slettet post. "
                f"Mail-id: {mail_id}"
            )

        levering[config.BOX_KEY_MAIL_DELETED] = True
        await _gem_box_data(item, box_data)

        logger.info(
            "Mail videresendt og flyttet til Slettet post. mail_id=%s",
            mail_id,
        )


# ---------------------------------------------------------------------------
# ATS-ADAPTERE
# ---------------------------------------------------------------------------


def _hent_item_data(item: Any) -> dict[str, Any]:
    """Hent itemets data-dictionary."""
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
            "Kø-itemet mangler dictionary-feltet "
            "data/item_data/payload."
        )

    return dict(value)


def _hent_box_data(item: Any) -> dict[str, Any]:
    """Hent en kopi af itemets Box data."""
    data = _hent_item_data(item)
    box = data.get("box")

    if box is None:
        return {}

    if not isinstance(box, Mapping):
        raise TypeError(
            "item.data['box'] skal være en dictionary."
        )

    return deepcopy(dict(box))


async def _gem_box_data(
    item: Any,
    box_data: dict[str, Any],
) -> None:
    """Gem hele item.data inklusive Box gennem ATS-klientens update()."""
    data = _hent_item_data(item)
    data["box"] = deepcopy(box_data)

    update = getattr(item, "update", None)
    if not callable(update):
        raise Systemfejl(
            "ATS-itemet mangler update(data); Box kan ikke gemmes."
        )

    await _maybe_await(update(data))

    if isinstance(item, MutableMapping):
        item["data"] = data


def _kald_med_box_data(
    method: Any,
    box_data: dict[str, Any],
) -> Any:
    """Kald en metode med Box data ud fra metodens signatur."""
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

    if status != config.STATUS_COMPLETED:
        raise Systemfejl(
            "Ukendt forretningsstatus ved afslutning: "
            f"{status!r}"
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
    """Await værdien, hvis værdien er awaitable."""
    if inspect.isawaitable(value):
        return await value
    return value


def _hent_vaerdi(item: Any, *names: str) -> Any:
    """Hent første eksisterende værdi fra dictionary eller objekt."""
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


def _normaliser_box_data(
    value: Mapping[str, Any],
) -> dict[str, Any]:
    """Sørg for, at Box indeholder alle nødvendige strukturer."""
    box_data = deepcopy(dict(value))

    box_data.setdefault(
        config.BOX_KEY_SCHEMA_VERSION,
        config.BOX_SCHEMA_VERSION,
    )
    box_data.setdefault(
        config.BOX_KEY_IDENTIFIER,
        {},
    )
    box_data.setdefault(
        config.BOX_KEY_VEHICLE,
        {
            config.BOX_KEY_VEHICLE_FOUND: False,
        },
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
    """Returnér trimmet tekst eller None."""
    if value in (None, ""):
        return None

    normalized = str(value).strip()
    return normalized or None


def _normaliser_email(value: Any) -> str | None:
    """Returnér en simpel valideret mailadresse eller None."""
    if value in (None, ""):
        return None

    email = str(value).strip()
    if not email:
        return None

    if email.count("@") != 1 or any(
        character.isspace()
        for character in email
    ):
        return None

    local_part, domain = email.rsplit("@", 1)
    if not local_part or not domain or "." not in domain:
        return None

    return email


def _paakraevet_tekst(value: Any, field_name: str) -> str:
    """Returnér obligatorisk tekst eller rejs en forretningsfejl."""
    if not isinstance(value, str) or not value.strip():
        raise Forretningsfejl(
            "Item data mangler det obligatoriske felt "
            f"{field_name!r}."
        )

    return value.strip()


def _mail_id(mail: Mapping[str, Any]) -> str:
    """Hent mail-id fra en mail-dictionary."""
    return str(
        mail.get("mail_id")
        or mail.get("message_id")
        or mail.get("id")
        or ""
    ).strip()


def _utc_now() -> str:
    """Returnér aktuelt UTC-tidspunkt som ISO-tekst."""
    return datetime.now(timezone.utc).isoformat(
        timespec="milliseconds"
    )


__all__ = [
    "Behandlingsfejl",
    "Forretningsfejl",
    "Systemfejl",
    "behandel",
    "process_item",
]
