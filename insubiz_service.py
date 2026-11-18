"""Læs køretøjslisten gennem q-insubiz med ATS-credentialen Q_INSUBIZ.

Eksporten caches i hukommelsen under samme proceskørsel. Ingen credentials
eller dokumentindhold logges eller gemmes lokalt.
"""
from __future__ import annotations

import asyncio
import logging
import re
import time
import unicodedata
from copy import deepcopy
from typing import Any, Iterable, Mapping, Sequence

from q_insubiz.api.auth_manager import InsubizAuthManager
from q_insubiz.api.client import InsubizApiClient
from q_insubiz.functionality.koeretoejer import KOERETOEJER_LISTE

import configuration as config

logger = logging.getLogger(__name__)


class InsubizFejl(RuntimeError):
    """Fælles fejl for Insubiz-opslag."""


class InsubizKonfigurationsfejl(InsubizFejl):
    """Ugyldig Insubiz-konfiguration."""


class InsubizDatafejl(InsubizFejl):
    """Eksporten har ugyldigt format."""


_cache: list[dict[str, Any]] | None = None
_cache_tid: float | None = None
_cache_lock = asyncio.Lock()


async def hent_koeretoejer_liste(
    *, force_refresh: bool = False
) -> list[dict[str, Any]]:
    """Returnér normaliserede køretøjer; login hentes af q-insubiz fra ATS."""
    global _cache, _cache_tid
    if not force_refresh and _cache_er_gyldig():
        return deepcopy(_cache)
    async with _cache_lock:
        if not force_refresh and _cache_er_gyldig():
            return deepcopy(_cache)
        auth_manager = InsubizAuthManager(headless=config.INSUBIZ_HEADLESS)
        api_client = InsubizApiClient(auth_manager=auth_manager)
        try:
            raekker = await KOERETOEJER_LISTE(
                api_client=api_client,
                customer_id=config.INSUBIZ_CUSTOMER_ID,
                show_tree_data=config.INSUBIZ_SHOW_TREE_DATA,
                active_only=config.INSUBIZ_ACTIVE_ONLY,
                columns=config.get_insubiz_vehicle_columns(),
            )
        finally:
            await api_client.close()
        if not isinstance(raekker, list):
            raise InsubizDatafejl("Insubiz returnerede ikke en liste")
        resultat = _normaliser_koeretoejer(raekker)
        _valider_dubletter(resultat)
        if config.INSUBIZ_CACHE_ENABLED:
            _cache, _cache_tid = deepcopy(resultat), time.monotonic()
        return deepcopy(resultat)


def ryd_cache() -> None:
    """Nulstil lokal køretøjscache."""
    global _cache, _cache_tid
    _cache, _cache_tid = None, None


def _cache_er_gyldig() -> bool:
    if not config.INSUBIZ_CACHE_ENABLED or _cache is None:
        return False
    ttl = config.INSUBIZ_CACHE_TTL_SECONDS
    if ttl is None:
        return True
    return _cache_tid is not None and time.monotonic() - _cache_tid < ttl


def _normaliser_koeretoejer(
    raekker: Sequence[Mapping[str, Any]],
) -> list[dict[str, Any]]:
    if not raekker:
        return []

    chassis_column = str(_cfg("INSUBIZ_CHASSIS_COLUMN", "ChassisNumber"))
    registration_column = str(_cfg("INSUBIZ_REGISTRATION_COLUMN", "Registration"))
    email_column = str(_cfg("INSUBIZ_DEPARTMENT_EMAIL_COLUMN", "ContactPerson1"))

    aliases = {
        chassis_column: [
            chassis_column,
            "ChassisNumber",
            "Stelnummer",
            "Stel nr",
            "VIN",
        ],
        registration_column: [
            registration_column,
            "Registration",
            "RegistrationNumber",
            "Registreringsnummer",
            "Nummerplade",
        ],
        email_column: [
            email_column,
            "ContactPerson1",
            "Afdelingsmail",
            "Email",
            "E-mail",
        ],
    }
    bruger_aliaser = _cfg("INSUBIZ_COLUMN_ALIASES", {}) or {}
    for canonical, values in bruger_aliaser.items():
        aliases.setdefault(str(canonical), [str(canonical)])
        aliases[str(canonical)].extend(str(value) for value in values)

    tilgaengelige = list(raekker[0].keys())
    kolonnemap = {
        canonical: _find_kolonne(tilgaengelige, kandidater)
        for canonical, kandidater in aliases.items()
    }

    mangler = [
        navn
        for navn in (chassis_column, registration_column, email_column)
        if kolonnemap.get(navn) is None
    ]
    ekstra_paakraevet = list(_cfg("INSUBIZ_REQUIRED_COLUMNS", []) or [])
    mangler.extend(
        navn for navn in ekstra_paakraevet if _find_kolonne(tilgaengelige, [navn]) is None
    )
    if mangler:
        raise InsubizDatafejl(
            "Insubiz-eksporten mangler obligatoriske kolonner: "
            + ", ".join(sorted(set(map(str, mangler))))
        )

    resultater: list[dict[str, Any]] = []
    for indeks, row in enumerate(raekker, start=2):
        stelnummer = _normaliser_identifikator(row.get(kolonnemap[chassis_column]))
        nummerplade = _normaliser_identifikator(row.get(kolonnemap[registration_column]))
        afdelingsmail = _normaliser_email(row.get(kolonnemap[email_column]))

        # Rækker uden begge køretøjsidentifikatorer kan ikke anvendes i processen.
        if not stelnummer and not nummerplade:
            logger.warning(
                "Springer Insubiz-række %s over: mangler stelnummer og nummerplade",
                indeks,
            )
            continue

        normaliseret = {str(key): _normaliser_celle(value) for key, value in row.items()}
        normaliseret.update(
            {
                "stelnummer": stelnummer,
                "nummerplade": nummerplade,
                "afdelingsmail": afdelingsmail,
            }
        )
        resultater.append(normaliseret)

    return resultater


def _valider_dubletter(koeretoejer: Sequence[Mapping[str, Any]]) -> None:
    policy = str(_cfg("INSUBIZ_DUPLICATE_POLICY", "keep")).strip().lower()
    if policy not in {"keep", "error"}:
        raise InsubizKonfigurationsfejl(
            "INSUBIZ_DUPLICATE_POLICY skal være keep eller error"
        )

    fundne: dict[tuple[str, str], int] = {}
    dubletter: set[tuple[str, str]] = set()
    for indeks, koeretoej in enumerate(koeretoejer):
        for type_, felt in (("stelnummer", "stelnummer"), ("nummerplade", "nummerplade")):
            vaerdi = koeretoej.get(felt)
            if not vaerdi:
                continue
            key = (type_, str(vaerdi))
            if key in fundne and fundne[key] != indeks:
                dubletter.add(key)
            else:
                fundne[key] = indeks

    if not dubletter:
        return

    visning = ", ".join(f"{type_}={vaerdi}" for type_, vaerdi in sorted(dubletter))
    if policy == "error":
        raise InsubizDatafejl(f"Insubiz-eksporten indeholder dubletter: {visning}")
    logger.warning(
        "Insubiz-eksporten indeholder dubletter; vehicle_matcher skal kræve entydigt match: %s",
        visning,
    )


def _find_kolonne(
    tilgaengelige: Iterable[str],
    kandidater: Iterable[str],
) -> str | None:
    opslag = {_kolonnenoegle(navn): navn for navn in tilgaengelige}
    for kandidat in kandidater:
        fundet = opslag.get(_kolonnenoegle(kandidat))
        if fundet is not None:
            return fundet
    return None


def _normaliser_headers(values: Iterable[Any]) -> list[str]:
    headers: list[str] = []
    brugte: dict[str, int] = {}
    for index, value in enumerate(values, start=1):
        navn = str(value).strip() if value not in (None, "") else f"kolonne_{index}"
        antal = brugte.get(navn, 0) + 1
        brugte[navn] = antal
        headers.append(navn if antal == 1 else f"{navn}_{antal}")
    return headers


def _normaliser_identifikator(value: Any) -> str | None:
    if value in (None, ""):
        return None
    tekst = _celle_til_tekst(value).upper()
    tekst = re.sub(r"[^A-Z0-9]", "", _ascii(tekst))
    return tekst or None


def _normaliser_email(value: Any) -> str | None:
    if value in (None, ""):
        return None
    email = _celle_til_tekst(value).strip()
    return email or None


def _normaliser_celle(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        return value.strip() or None
    return value


def _celle_til_tekst(value: Any) -> str:
    # openpyxl kan levere et heltal som float. Registreringsværdier skal ikke
    # ende som eksempelvis "12345.0".
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _kolonnenoegle(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", _ascii(str(value)).lower())


def _ascii(value: str) -> str:
    normaliseret = unicodedata.normalize("NFKD", value)
    return normaliseret.encode("ascii", "ignore").decode("ascii")


def _cfg(navn: str, standard: Any = None) -> Any:
    return getattr(config, navn, standard)
