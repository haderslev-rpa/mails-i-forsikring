"""Entydigt opslag af køretøjer for processen mails-i-forsikring.

Modulet har ét ansvar: at matche et stelnummer eller en nummerplade mod den
normaliserede køretøjsliste fra insubiz_service.py.

Modulet:
- udfører ingen HTTP-kald
- læser eller skriver ikke ATS-items
- ændrer ikke Box data
- foretager ingen Outlook-handlinger
- returnerer kun ét entydigt køretøj

Offentlig funktion
------------------
find_koeretoej(
    koeretoejer,
    identifikator_type,
    identifikator_vaerdi,
) -> dict

Forventet input fra insubiz_service.py
--------------------------------------
{
    "stelnummer": "WVWZZZ12345678901",
    "nummerplade": "AB12345",
    "afdelingsmail": "afdeling@haderslev.dk",
    ...
}
"""

from __future__ import annotations

import re
import unicodedata
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Iterable, Mapping, Sequence

import configuration as config

class VehicleMatcherFejl(RuntimeError):
    """Basisklasse for kontrollerede fejl i køretøjsopslaget."""


class UgyldigIdentifikatorFejl(VehicleMatcherFejl):
    """Identifikatortype eller identifikatorværdi er ugyldig."""


class KoeretoejIkkeFundetFejl(VehicleMatcherFejl):
    """Ingen køretøjer matcher identifikatoren."""


class FlereKoeretoejerFundetFejl(VehicleMatcherFejl):
    """Flere forskellige køretøjer matcher identifikatoren."""


class UgyldigKoeretoejslisteFejl(VehicleMatcherFejl):
    """Køretøjslisten har et ugyldigt format."""


@dataclass(frozen=True, slots=True)
class NormaliseretIdentifikator:
    """Valideret identifikator, der kan anvendes til opslag."""

    type: str
    vaerdi: str


STELNUMMER = "stelnummer"
NUMMERPLADE = "nummerplade"
TILLADE_IDENTIFIKATORTYPER = frozenset({STELNUMMER, NUMMERPLADE})

# Et VIN består normalt af 17 tegn og indeholder ikke I, O eller Q.
VIN_PATTERN = re.compile(r"[A-HJ-NPR-Z0-9]{17}")

# Nummerplader fra eksporten kan have flere formater. Matcherens opgave er at
# normalisere og lave et eksakt opslag, ikke at begrænse Insubiz' pladetyper
# unødigt. Standardgrænserne kan ændres i config.py.
REGISTRATION_ALLOWED_PATTERN = re.compile(r"[A-Z0-9]+")


# ---------------------------------------------------------------------------
# Offentlig funktion
# ---------------------------------------------------------------------------

def find_koeretoej(
    koeretoejer: Sequence[Mapping[str, Any]] | Iterable[Mapping[str, Any]],
    identifikator_type: str,
    identifikator_vaerdi: str,
) -> dict[str, Any]:
    """Find præcis ét køretøj ved eksakt, normaliseret identifikator.

    Args:
        koeretoejer: Køretøjslisten fra ``insubiz_service``.
        identifikator_type: ``stelnummer`` eller ``nummerplade``.
        identifikator_vaerdi: Værdien udtrukket fra PDF-filen.

    Returns:
        En kopi af det entydige køretøj med de kanoniske felter:
        ``fundet``, ``nummerplade``, ``stelnummer`` og ``afdelingsmail``.
        Andre kolonner fra Insubiz bevares.

    Raises:
        UgyldigIdentifikatorFejl: Type eller værdi er ugyldig.
        UgyldigKoeretoejslisteFejl: Listen eller en række er ugyldig.
        KoeretoejIkkeFundetFejl: Ingen rækker matcher.
        FlereKoeretoejerFundetFejl: Flere forskellige køretøjer matcher.
    """
    identifikator = normaliser_identifikator(
        identifikator_type=identifikator_type,
        identifikator_vaerdi=identifikator_vaerdi,
    )
    liste = _materialiser_koeretoejer(koeretoejer)

    match = [
        _normaliser_koeretoej(row, index=index)
        for index, row in enumerate(liste, start=1)
        if _matcher_row(row, identifikator)
    ]

    if not match:
        raise KoeretoejIkkeFundetFejl(
            f"Intet køretøj blev fundet med "
            f"{identifikator.type}={identifikator.vaerdi}"
        )

    unikke_match = _dedupliker_identiske_raekker(match)
    if len(unikke_match) > 1:
        detaljer = _format_match_detaljer(unikke_match)
        raise FlereKoeretoejerFundetFejl(
            f"Flere køretøjer matcher "
            f"{identifikator.type}={identifikator.vaerdi}: {detaljer}"
        )

    resultat = unikke_match[0]
    resultat["fundet"] = True
    return resultat


def normaliser_identifikator(
    *,
    identifikator_type: str,
    identifikator_vaerdi: str,
) -> NormaliseretIdentifikator:
    """Normalisér og validér en identifikator fra PDF-beregningen."""
    type_ = _normaliser_type(identifikator_type)

    if type_ == STELNUMMER:
        vaerdi = normaliser_stelnummer(identifikator_vaerdi)
        if not gyldigt_stelnummer(vaerdi):
            raise UgyldigIdentifikatorFejl(
                f"Ugyldigt stelnummer: {identifikator_vaerdi!r}. "
                "Et stelnummer skal bestå af 17 gyldige VIN-tegn."
            )
    else:
        vaerdi = normaliser_nummerplade(identifikator_vaerdi)
        if not gyldig_nummerplade(vaerdi):
            raise UgyldigIdentifikatorFejl(
                f"Ugyldig nummerplade: {identifikator_vaerdi!r}"
            )

    return NormaliseretIdentifikator(type=type_, vaerdi=vaerdi)


def normaliser_stelnummer(value: Any) -> str:
    """Fjern formatering og konvertér stelnummer til store ASCII-tegn."""
    return re.sub(r"[^A-Z0-9]", "", _ascii_store_bogstaver(value))


def normaliser_nummerplade(value: Any) -> str:
    """Fjern mellemrum, bindestreger og anden visuel formatering."""
    return re.sub(r"[^A-Z0-9]", "", _ascii_store_bogstaver(value))


def gyldigt_stelnummer(value: str) -> bool:
    """Kontrollér standardformatet for et stelnummer/VIN."""
    return bool(VIN_PATTERN.fullmatch(str(value or "")))


def gyldig_nummerplade(value: str) -> bool:
    """Kontrollér en normaliseret nummerplade uden at antage én pladetype."""
    tekst = str(value or "")
    min_laengde = _config_int("REGISTRATION_MIN_LENGTH", 5)
    max_laengde = _config_int("REGISTRATION_MAX_LENGTH", 9)

    if min_laengde > max_laengde:
        raise UgyldigKoeretoejslisteFejl(
            "REGISTRATION_MIN_LENGTH må ikke være større end "
            "REGISTRATION_MAX_LENGTH"
        )
    if not min_laengde <= len(tekst) <= max_laengde:
        return False
    if not REGISTRATION_ALLOWED_PATTERN.fullmatch(tekst):
        return False

    # Standard er mindst ét bogstav og mindst ét tal. Kravet kan deaktiveres,
    # hvis Insubiz indeholder særlige pladetyper med et andet format.
    if bool(getattr(config, "REGISTRATION_REQUIRE_LETTER_AND_DIGIT", True)):
        return any(tegn.isalpha() for tegn in tekst) and any(
            tegn.isdigit() for tegn in tekst
        )
    return True


# ---------------------------------------------------------------------------
# Match
# ---------------------------------------------------------------------------

def _matcher_row(
    row: Mapping[str, Any],
    identifikator: NormaliseretIdentifikator,
) -> bool:
    if not isinstance(row, Mapping):
        raise UgyldigKoeretoejslisteFejl(
            "Alle elementer i køretøjslisten skal være objekter/dictionaries"
        )

    if identifikator.type == STELNUMMER:
        row_vaerdi = _hent_stelnummer(row)
        return bool(row_vaerdi and row_vaerdi == identifikator.vaerdi)

    row_vaerdi = _hent_nummerplade(row)
    return bool(row_vaerdi and row_vaerdi == identifikator.vaerdi)


def _materialiser_koeretoejer(
    koeretoejer: Sequence[Mapping[str, Any]] | Iterable[Mapping[str, Any]],
) -> list[Mapping[str, Any]]:
    if koeretoejer is None:
        raise UgyldigKoeretoejslisteFejl("Køretøjslisten mangler")

    if isinstance(koeretoejer, (str, bytes, Mapping)):
        raise UgyldigKoeretoejslisteFejl(
            "Køretøjslisten skal være en liste eller anden iterable af objekter"
        )

    try:
        liste = list(koeretoejer)
    except TypeError as exc:
        raise UgyldigKoeretoejslisteFejl(
            "Køretøjslisten kan ikke gennemløbes"
        ) from exc

    if not liste:
        raise UgyldigKoeretoejslisteFejl("Køretøjslisten er tom")
    return liste


def _normaliser_koeretoej(
    row: Mapping[str, Any],
    *,
    index: int,
) -> dict[str, Any]:
    if not isinstance(row, Mapping):
        raise UgyldigKoeretoejslisteFejl(
            f"Køretøjsrække {index} er ikke et objekt/dictionary"
        )

    resultat = deepcopy(dict(row))
    stelnummer = _hent_stelnummer(row)
    nummerplade = _hent_nummerplade(row)
    afdelingsmail = _hent_afdelingsmail(row)

    if not stelnummer and not nummerplade:
        raise UgyldigKoeretoejslisteFejl(
            f"Det matchede køretøj på række {index} mangler både "
            "stelnummer og nummerplade"
        )

    resultat.update(
        {
            "stelnummer": stelnummer,
            "nummerplade": nummerplade,
            "afdelingsmail": afdelingsmail,
        }
    )
    return resultat


def _hent_stelnummer(row: Mapping[str, Any]) -> str | None:
    value = _hent_felt(
        row,
        "stelnummer",
        getattr(config, "INSUBIZ_CHASSIS_COLUMN", "ChassisNumber"),
        "ChassisNumber",
        "Stelnummer",
        "Stel nr",
        "VIN",
    )
    if value in (None, ""):
        return None
    normaliseret = normaliser_stelnummer(value)
    return normaliseret or None


def _hent_nummerplade(row: Mapping[str, Any]) -> str | None:
    value = _hent_felt(
        row,
        "nummerplade",
        getattr(config, "INSUBIZ_REGISTRATION_COLUMN", "Registration"),
        "Registration",
        "RegistrationNumber",
        "Registreringsnummer",
        "Nummerplade",
    )
    if value in (None, ""):
        return None
    normaliseret = normaliser_nummerplade(value)
    return normaliseret or None


def _hent_afdelingsmail(row: Mapping[str, Any]) -> str | None:
    value = _hent_felt(
        row,
        "afdelingsmail",
        getattr(config, "INSUBIZ_DEPARTMENT_EMAIL_COLUMN", "ContactPerson1"),
        "ContactPerson1",
        "Afdelingsmail",
        "Email",
        "E-mail",
    )
    return _normaliser_email(value)


def _hent_felt(row: Mapping[str, Any], *mulige_navne: str) -> Any:
    # Først anvendes direkte opslag. Derefter et robust opslag, hvor forskelle i
    # store/små bogstaver, mellemrum, tegn og danske bogstaver ignoreres.
    for navn in mulige_navne:
        if navn in row and row[navn] not in (None, ""):
            return row[navn]

    normaliserede_noegler = {
        _normaliser_feltnavn(noegle): noegle for noegle in row.keys()
    }
    for navn in mulige_navne:
        faktisk_noegle = normaliserede_noegler.get(_normaliser_feltnavn(navn))
        if faktisk_noegle is not None and row[faktisk_noegle] not in (None, ""):
            return row[faktisk_noegle]
    return None


# ---------------------------------------------------------------------------
# Entydighed
# ---------------------------------------------------------------------------

def _dedupliker_identiske_raekker(
    match: Sequence[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Saml kun teknisk identiske dubletter.

    To Insubiz-rækker regnes som samme køretøj, hvis stelnummer, nummerplade og
    afdelingsmail er ens efter normalisering. Hvis ét af felterne afviger,
    beholdes begge rækker, og opslaget afvises som tvetydigt.
    """
    unikke: dict[tuple[str | None, str | None, str | None], dict[str, Any]] = {}
    for row in match:
        noegle = (
            row.get("stelnummer"),
            row.get("nummerplade"),
            _email_sammenligningsvaerdi(row.get("afdelingsmail")),
        )
        unikke.setdefault(noegle, row)
    return list(unikke.values())


def _format_match_detaljer(match: Sequence[Mapping[str, Any]]) -> str:
    maks = _config_int("VEHICLE_MATCHER_ERROR_DETAIL_LIMIT", 5)
    detaljer = []
    for row in match[:maks]:
        detaljer.append(
            "[stelnummer={stel}, nummerplade={reg}, afdelingsmail={mail}]".format(
                stel=row.get("stelnummer") or "mangler",
                reg=row.get("nummerplade") or "mangler",
                mail=row.get("afdelingsmail") or "mangler",
            )
        )
    if len(match) > maks:
        detaljer.append(f"... og {len(match) - maks} yderligere match")
    return ", ".join(detaljer)


# ---------------------------------------------------------------------------
# Normalisering
# ---------------------------------------------------------------------------

def _normaliser_type(value: Any) -> str:
    tekst = _ascii_store_bogstaver(value).lower()
    tekst = re.sub(r"[^a-z0-9]", "", tekst)

    mapping = {
        "stelnummer": STELNUMMER,
        "stel": STELNUMMER,
        "stelnr": STELNUMMER,
        "chassisnumber": STELNUMMER,
        "vin": STELNUMMER,
        "nummerplade": NUMMERPLADE,
        "registreringsnummer": NUMMERPLADE,
        "registrering": NUMMERPLADE,
        "registreringsnr": NUMMERPLADE,
        "regnr": NUMMERPLADE,
        "registration": NUMMERPLADE,
        "registrationnumber": NUMMERPLADE,
    }
    type_ = mapping.get(tekst)
    if type_ not in TILLADE_IDENTIFIKATORTYPER:
        raise UgyldigIdentifikatorFejl(
            f"Ukendt identifikatortype: {value!r}. "
            "Tilladte værdier er stelnummer og nummerplade."
        )
    return type_


def _normaliser_email(value: Any) -> str | None:
    if value in (None, ""):
        return None
    email = str(value).strip()
    return email or None


def _email_sammenligningsvaerdi(value: Any) -> str | None:
    email = _normaliser_email(value)
    return email.casefold() if email else None


def _normaliser_feltnavn(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", _ascii_store_bogstaver(value).lower())


def _ascii_store_bogstaver(value: Any) -> str:
    tekst = unicodedata.normalize("NFKD", str(value or ""))
    return tekst.encode("ascii", "ignore").decode("ascii").upper()


def _config_int(navn: str, standard: int) -> int:
    value = getattr(config, navn, standard)
    try:
        resultat = int(value)
    except (TypeError, ValueError) as exc:
        raise UgyldigKoeretoejslisteFejl(
            f"Config-værdien {navn} skal være et heltal"
        ) from exc
    if resultat < 1:
        raise UgyldigKoeretoejslisteFejl(
            f"Config-værdien {navn} skal være større end 0"
        )
    return resultat
