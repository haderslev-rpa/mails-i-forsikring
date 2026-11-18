"""Forretningsregler for manuel behandling i mails-i-forsikring."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Final


MANUEL_ID_MANGLENDE_REGISTRERINGSNUMMER: Final[int] = 1
MANUEL_ID_MANGLENDE_STELNUMMER: Final[int] = 2
MANUEL_ID_INTET_MATCH: Final[int] = 3
MANUEL_ID_FLERE_MATCH: Final[int] = 4
MANUEL_ID_MANGLENDE_AFDELINGSMAIL: Final[int] = 5
MANUEL_ID_DIGITALPOST_BETALING: Final[int] = 6


@dataclass(frozen=True, slots=True)
class ManuelBehandling:
    """Resultat, når et kø-item skal afsluttes til manuel behandling."""

    manuel_id: int
    emne: str
    besked: str


_MANGLENDE_REGISTRERINGSNUMMER = ManuelBehandling(
    manuel_id=MANUEL_ID_MANGLENDE_REGISTRERINGSNUMMER,
    emne="Manuel behandling - Registreringsnummer mangler",
    besked=(
        "Robotten har fundet et køretøj, men køretøjets "
        "registreringsnummer mangler i Insubiz."
    ),
)

_MANGLENDE_STELNUMMER = ManuelBehandling(
    manuel_id=MANUEL_ID_MANGLENDE_STELNUMMER,
    emne="Manuel behandling - Stelnummer mangler",
    besked=(
        "Robotten har fundet et køretøj, men køretøjets "
        "stelnummer mangler i Insubiz."
    ),
)

_INTET_MATCH = ManuelBehandling(
    manuel_id=MANUEL_ID_INTET_MATCH,
    emne="Manuel behandling - Intet køretøjsmatch",
    besked=(
        "Robotten kunne ikke finde et køretøj i Insubiz, som matcher "
        "nummerpladen eller stelnummeret fra PDF-filen."
    ),
)

_FLERE_MATCH = ManuelBehandling(
    manuel_id=MANUEL_ID_FLERE_MATCH,
    emne="Manuel behandling - Flere køretøjsmatch",
    besked=(
        "Robotten fandt flere køretøjer i Insubiz, som matcher "
        "nummerpladen eller stelnummeret fra PDF-filen."
    ),
)

_MANGLENDE_AFDELINGSMAIL = ManuelBehandling(
    manuel_id=MANUEL_ID_MANGLENDE_AFDELINGSMAIL,
    emne="Manuel behandling - Afdelingsmail mangler",
    besked=(
        "Robotten har fundet køretøjet i Insubiz, men køretøjet "
        "har ingen gyldig afdelingsmail."
    ),
)


def digitalpost_betaling_til_manuel_behandling() -> ManuelBehandling:
    """Betalingsmailen må ikke automatisk videresendes eller slettes."""
    return ManuelBehandling(
        manuel_id=MANUEL_ID_DIGITALPOST_BETALING,
        emne="Manuel behandling - Digitalpost betaling",
        besked=(
            "Digitalpost-betalingsmailen er fundet, men dokumenttypen har endnu "
            "ingen godkendt PDF-parser. Kontrollér originalmailen i Forsikrings "
            "indbakke manuelt. Originalmailen er ikke flyttet eller slettet."
        ),
    )


def vurder_manuel_behandling(
    *,
    registreringsnummer: str | None,
    stelnummer: str | None,
    antal_fundne: int,
    afdelingsmail: str | None,
) -> ManuelBehandling | None:
    """Returnerer den første relevante årsag til manuel behandling.

    Match-antallet vurderes før køretøjsfelterne, fordi der ikke findes et
    entydigt køretøj at validere, når resultatet er nul eller større end ét.
    """
    if isinstance(antal_fundne, bool) or not isinstance(antal_fundne, int):
        raise TypeError("antal_fundne skal være et heltal.")
    if antal_fundne < 0:
        raise ValueError("antal_fundne må ikke være negativt.")

    if antal_fundne == 0:
        return _INTET_MATCH
    if antal_fundne > 1:
        return _FLERE_MATCH
    if not _har_tekst(registreringsnummer):
        return _MANGLENDE_REGISTRERINGSNUMMER
    if not _har_tekst(stelnummer):
        return _MANGLENDE_STELNUMMER
    if not _gyldig_email(afdelingsmail):
        return _MANGLENDE_AFDELINGSMAIL
    return None


def _har_tekst(value: str | None) -> bool:
    return isinstance(value, str) and bool(value.strip())


def _gyldig_email(value: str | None) -> bool:
    if not _har_tekst(value):
        return False
    email = str(value).strip()
    if email.count("@") != 1 or any(character.isspace() for character in email):
        return False
    local_part, domain = email.rsplit("@", 1)
    return bool(local_part and domain and "." in domain)


__all__ = [
    "MANUEL_ID_FLERE_MATCH",
    "MANUEL_ID_INTET_MATCH",
    "MANUEL_ID_MANGLENDE_AFDELINGSMAIL",
    "MANUEL_ID_MANGLENDE_REGISTRERINGSNUMMER",
    "MANUEL_ID_MANGLENDE_STELNUMMER",
    "ManuelBehandling",
    "vurder_manuel_behandling",
]
