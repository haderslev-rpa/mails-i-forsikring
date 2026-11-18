"""PDF-beregninger for processen mails-i-forsikring.

Modulet indeholder al logik til at udtrække præcis én identifikator fra en PDF:
- nummerplade
- stelnummer

Den fælles indgang er ``udtraek_stel_eller_nummerplade``. Funktionen vælger en
mailtypespecifik funktion, så behandel.py ikke indeholder PDF-beregninger.

Understøttede mailtyper:
- BOEDEKONTOR
- INDBETALINGSKORT
- VIGTIG_INFORMATION
- INDKALDELSE_PERIODISK_SYN
- PAAMINDELSE_PERIODISK_SYN

Afhængighed:
    pypdf
"""

from __future__ import annotations

import io
import re
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Callable, Iterable, Sequence

from pypdf import PdfReader

import configuration as config

class PdfBeregningFejl(RuntimeError):
    """Basisklasse for kontrollerede fejl under PDF-beregning."""


class PdfKanIkkeLaesesFejl(PdfBeregningFejl):
    """PDF-filen kunne ikke åbnes eller indeholdt ingen læsbar tekst."""


class IdentifikatorIkkeFundetFejl(PdfBeregningFejl):
    """Der blev ikke fundet et gyldigt stelnummer eller en nummerplade."""


class TvetydigIdentifikatorFejl(PdfBeregningFejl):
    """Flere forskellige, lige sandsynlige identifikatorer blev fundet."""


class UkendtMailtypeFejl(PdfBeregningFejl):
    """Mailtypen har ingen registreret PDF-beregning."""


@dataclass(frozen=True, slots=True)
class IdentifierResult:
    """Det eneste forretningsresultat fra PDF-beregningen."""

    type: str
    vaerdi: str


@dataclass(frozen=True, slots=True)
class PdfIndhold:
    """Internt, normaliseret indhold fra PDF-filen."""

    tekst: str
    xml_tekster: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class Kandidat:
    """Intern kandidat med score til entydigt valg."""

    type: str
    vaerdi: str
    score: int
    kilde: str


# VIN/stelnummer består normalt af 17 tegn. I, O og Q bruges ikke i VIN.
VIN_PATTERN = re.compile(r"(?<![A-Z0-9])[A-HJ-NPR-Z0-9]{17}(?![A-Z0-9])")

# Danske standardnummerplader samt kortere/særlige varianter. Den brede
# kandidat bruges kun tæt ved en nummerpladeetiket eller ved opslag i XML.
STANDARD_REG_PATTERN = re.compile(r"(?<![A-Z0-9])[A-Z]{2}[ -]?\d{4,5}(?![A-Z0-9])")
BRED_REG_PATTERN = re.compile(r"(?<![A-Z0-9])[A-Z0-9][A-Z0-9 -]{3,8}[A-Z0-9](?![A-Z0-9])")

VIN_LABELS = (
    "stelnummer",
    "stel nummer",
    "stelnr",
    "stel nr",
    "chassisnummer",
    "chassis number",
    "vin",
)

REG_LABELS = (
    "registreringsnummer",
    "registrerings nummer",
    "registreringsnr",
    "registrerings nr",
    "regnr",
    "reg nr",
    "nummerplade",
    "nummer plade",
)

XML_VIN_TAGS = {
    "koeretoejstelnummer",
    "koeretoejstelnr",
    "stelnummer",
    "stelnr",
    "chassisnumber",
    "vehicleidentificationnumber",
    "vin",
}

XML_REG_TAGS = {
    "koeretoejregistreringnummer",
    "koeretoejregistreringsnummer",
    "registreringnummer",
    "registreringsnummer",
    "registration",
    "registrationnumber",
    "nummerplade",
}


# ---------------------------------------------------------------------------
# Offentlig indgang
# ---------------------------------------------------------------------------

def udtraek_stel_eller_nummerplade(
    mail_type: str,
    pdf_bytes: bytes,
) -> IdentifierResult:
    """Vælg korrekt beregningsfunktion og returnér én identifikator.

    Raises:
        UkendtMailtypeFejl: Mailtypen er ikke understøttet.
        PdfKanIkkeLaesesFejl: PDF-filen kan ikke læses.
        IdentifikatorIkkeFundetFejl: Ingen gyldig identifikator findes.
        TvetydigIdentifikatorFejl: Resultatet er ikke entydigt.
    """
    normaliseret_mail_type = _normaliser_mail_type(mail_type)
    funktion = PDF_BEREGNINGER.get(normaliseret_mail_type)

    if funktion is None:
        raise UkendtMailtypeFejl(
            f"Ingen PDF-beregning er registreret for mailtypen: {mail_type!r}"
        )

    return funktion(pdf_bytes)


# ---------------------------------------------------------------------------
# Én funktion pr. dokumenttype
# ---------------------------------------------------------------------------

def boedekontor(pdf_bytes: bytes) -> IdentifierResult:
    """Udtræk nummerplade eller stelnummer fra en Bødekontor-PDF.

    Nummerplade prioriteres, fordi bødedokumenter typisk identificerer
    køretøjet ved registreringsnummer. Et eksplicit stelnummer accepteres som
    fallback.
    """
    indhold = _laes_pdf(pdf_bytes)
    kandidater = _kandidater_fra_xml(indhold.xml_tekster)
    kandidater += _kandidater_ved_labels(indhold.tekst, REG_LABELS, "nummerplade", 95)
    kandidater += _kandidater_ved_labels(indhold.tekst, VIN_LABELS, "stelnummer", 90)
    kandidater += _standard_nummerplader(indhold.tekst, score=55)
    kandidater += _frie_stelnumre(indhold.tekst, score=50)
    return _vaelg_entydig(kandidater, foretrukken_type="nummerplade")


def indbetalingskort(pdf_bytes: bytes) -> IdentifierResult:
    """Udtræk nummerplade eller stelnummer fra et indbetalingskort.

    Et eksplicit registreringsnummer prioriteres. XML-felter og et eksplicit
    stelnummer er sikre alternativer.
    """
    indhold = _laes_pdf(pdf_bytes)
    kandidater = _kandidater_fra_xml(indhold.xml_tekster)
    kandidater += _kandidater_ved_labels(indhold.tekst, REG_LABELS, "nummerplade", 100)
    kandidater += _kandidater_ved_labels(indhold.tekst, VIN_LABELS, "stelnummer", 95)
    kandidater += _standard_nummerplader(indhold.tekst, score=50)
    return _vaelg_entydig(kandidater, foretrukken_type="nummerplade")


def vigtig_information(pdf_bytes: bytes) -> IdentifierResult:
    """Udtræk identifikator fra PDF med emnet Vigtig information.

    Strukturerede XML-felter prioriteres. Derefter anvendes eksplicitte
    etiketter for stelnummer og registreringsnummer.
    """
    indhold = _laes_pdf(pdf_bytes)
    kandidater = _kandidater_fra_xml(indhold.xml_tekster)
    kandidater += _kandidater_ved_labels(indhold.tekst, VIN_LABELS, "stelnummer", 95)
    kandidater += _kandidater_ved_labels(indhold.tekst, REG_LABELS, "nummerplade", 95)
    kandidater += _frie_stelnumre(indhold.tekst, score=45)
    kandidater += _standard_nummerplader(indhold.tekst, score=40)
    return _vaelg_entydig(kandidater)


def indkaldelse_periodisk_syn(pdf_bytes: bytes) -> IdentifierResult:
    """Udtræk identifikator fra indkaldelse til periodisk syn."""
    indhold = _laes_pdf(pdf_bytes)
    kandidater = _kandidater_fra_xml(indhold.xml_tekster)
    kandidater += _kandidater_ved_labels(indhold.tekst, REG_LABELS, "nummerplade", 100)
    kandidater += _kandidater_ved_labels(indhold.tekst, VIN_LABELS, "stelnummer", 95)
    kandidater += _standard_nummerplader(indhold.tekst, score=45)
    kandidater += _frie_stelnumre(indhold.tekst, score=40)
    return _vaelg_entydig(kandidater, foretrukken_type="nummerplade")


def paamindelse_periodisk_syn(pdf_bytes: bytes) -> IdentifierResult:
    """Udtræk identifikator fra påmindelse om periodisk syn."""
    indhold = _laes_pdf(pdf_bytes)
    kandidater = _kandidater_fra_xml(indhold.xml_tekster)
    kandidater += _kandidater_ved_labels(indhold.tekst, REG_LABELS, "nummerplade", 100)
    kandidater += _kandidater_ved_labels(indhold.tekst, VIN_LABELS, "stelnummer", 95)
    kandidater += _standard_nummerplader(indhold.tekst, score=45)
    kandidater += _frie_stelnumre(indhold.tekst, score=40)
    return _vaelg_entydig(kandidater, foretrukken_type="nummerplade")


# ---------------------------------------------------------------------------
# PDF-læsning
# ---------------------------------------------------------------------------

def _laes_pdf(pdf_bytes: bytes) -> PdfIndhold:
    if not isinstance(pdf_bytes, (bytes, bytearray)) or not pdf_bytes:
        raise PdfKanIkkeLaesesFejl("PDF-data mangler eller har et ugyldigt format")

    try:
        reader = PdfReader(io.BytesIO(bytes(pdf_bytes)))
    except Exception as exc:
        raise PdfKanIkkeLaesesFejl("PDF-filen kunne ikke åbnes") from exc

    side_tekster: list[str] = []
    for side in reader.pages:
        try:
            tekst = side.extract_text() or ""
        except Exception:
            tekst = ""
        if tekst.strip():
            side_tekster.append(tekst)

    xml_tekster = tuple(_udtraek_xml_fra_attachments(reader))
    samlet_tekst = _normaliser_dokumenttekst("\n".join(side_tekster))

    if not samlet_tekst and not xml_tekster:
        raise PdfKanIkkeLaesesFejl(
            "PDF-filen indeholder ingen læsbar tekst eller XML-metadata"
        )

    return PdfIndhold(tekst=samlet_tekst, xml_tekster=xml_tekster)


def _udtraek_xml_fra_attachments(reader: PdfReader) -> Iterable[str]:
    """Læs indlejrede XML-filer, når PDF-biblioteket eksponerer attachments."""
    try:
        attachments = reader.attachments
    except Exception:
        return []

    resultater: list[str] = []
    for filnavn, indhold_liste in (attachments or {}).items():
        if not str(filnavn).lower().endswith(".xml"):
            continue
        if not isinstance(indhold_liste, (list, tuple)):
            indhold_liste = [indhold_liste]
        for indhold in indhold_liste:
            if isinstance(indhold, bytes):
                tekst = indhold.decode("utf-8", errors="replace")
            else:
                tekst = str(indhold)
            if tekst.strip():
                resultater.append(tekst)
    return resultater


# ---------------------------------------------------------------------------
# Kandidater
# ---------------------------------------------------------------------------

def _kandidater_fra_xml(xml_tekster: Sequence[str]) -> list[Kandidat]:
    kandidater: list[Kandidat] = []

    for xml_tekst in xml_tekster:
        try:
            root = ET.fromstring(xml_tekst)
        except ET.ParseError:
            continue

        for element in root.iter():
            tag = _lokalt_xml_navn(element.tag)
            vaerdi = (element.text or "").strip()
            if not vaerdi:
                continue

            if tag in XML_VIN_TAGS:
                normaliseret = _normaliser_stelnummer(vaerdi)
                if _gyldigt_stelnummer(normaliseret):
                    kandidater.append(
                        Kandidat("stelnummer", normaliseret, 120, f"xml:{tag}")
                    )

            if tag in XML_REG_TAGS:
                normaliseret = _normaliser_nummerplade(vaerdi)
                if _gyldig_nummerplade(normaliseret):
                    kandidater.append(
                        Kandidat("nummerplade", normaliseret, 120, f"xml:{tag}")
                    )

    return kandidater


def _kandidater_ved_labels(
    tekst: str,
    labels: Sequence[str],
    type_: str,
    score: int,
) -> list[Kandidat]:
    kandidater: list[Kandidat] = []

    for label in labels:
        # Der søges i et begrænset område efter etiketten for at undgå at tage
        # sagsnumre, datoer og betalingsreferencer andre steder i dokumentet.
        pattern = re.compile(
            rf"(?i)\b{re.escape(label)}\b\s*(?:nr\.?|nummer)?\s*[:#-]?\s*"
            rf"([^\n\r]{{1,45}})"
        )
        for match in pattern.finditer(tekst):
            udsnit = match.group(1)
            kandidat = _identifikator_fra_udsnit(type_, udsnit)
            if kandidat:
                kandidater.append(Kandidat(type_, kandidat, score, f"label:{label}"))

    return kandidater


def _identifikator_fra_udsnit(type_: str, udsnit: str) -> str | None:
    if type_ == "stelnummer":
        for match in VIN_PATTERN.finditer(udsnit.upper()):
            kandidat = _normaliser_stelnummer(match.group(0))
            if _gyldigt_stelnummer(kandidat):
                return kandidat
        return None

    # Ved en eksplicit nummerpladeetiket tillades flere danske pladetyper end
    # standardmønstret. Første token stoppes ved almindelige feltseparatorer.
    renset = re.split(r"[;,|\n\r]", udsnit, maxsplit=1)[0]
    match = BRED_REG_PATTERN.search(renset.upper())
    if not match:
        return None
    kandidat = _normaliser_nummerplade(match.group(0))
    return kandidat if _gyldig_nummerplade(kandidat) else None


def _frie_stelnumre(tekst: str, score: int) -> list[Kandidat]:
    resultater: list[Kandidat] = []
    for match in VIN_PATTERN.finditer(tekst.upper()):
        kandidat = _normaliser_stelnummer(match.group(0))
        if _gyldigt_stelnummer(kandidat):
            resultater.append(Kandidat("stelnummer", kandidat, score, "fri-tekst"))
    return resultater


def _standard_nummerplader(tekst: str, score: int) -> list[Kandidat]:
    resultater: list[Kandidat] = []
    for match in STANDARD_REG_PATTERN.finditer(tekst.upper()):
        kandidat = _normaliser_nummerplade(match.group(0))
        if _gyldig_nummerplade(kandidat):
            resultater.append(Kandidat("nummerplade", kandidat, score, "fri-tekst"))
    return resultater


# ---------------------------------------------------------------------------
# Entydigt valg og validering
# ---------------------------------------------------------------------------

def _vaelg_entydig(
    kandidater: Sequence[Kandidat],
    foretrukken_type: str | None = None,
) -> IdentifierResult:
    gyldige = _dedupliker_kandidater(kandidater)
    if not gyldige:
        raise IdentifikatorIkkeFundetFejl(
            "PDF-beregningen fandt hverken et gyldigt stelnummer eller en nummerplade"
        )

    hoejeste_score = max(kandidat.score for kandidat in gyldige)
    bedste = [k for k in gyldige if k.score == hoejeste_score]

    if foretrukken_type:
        foretrukne = [k for k in bedste if k.type == foretrukken_type]
        if foretrukne:
            bedste = foretrukne

    unikke = {(k.type, k.vaerdi) for k in bedste}
    if len(unikke) != 1:
        fundne = ", ".join(f"{type_}={vaerdi}" for type_, vaerdi in sorted(unikke))
        raise TvetydigIdentifikatorFejl(
            f"PDF-beregningen fandt flere lige sandsynlige identifikatorer: {fundne}"
        )

    type_, vaerdi = next(iter(unikke))
    return IdentifierResult(type=type_, vaerdi=vaerdi)


def _dedupliker_kandidater(kandidater: Sequence[Kandidat]) -> list[Kandidat]:
    bedste_pr_vaerdi: dict[tuple[str, str], Kandidat] = {}
    for kandidat in kandidater:
        key = (kandidat.type, kandidat.vaerdi)
        nuvaerende = bedste_pr_vaerdi.get(key)
        if nuvaerende is None or kandidat.score > nuvaerende.score:
            bedste_pr_vaerdi[key] = kandidat
    return list(bedste_pr_vaerdi.values())


def _normaliser_stelnummer(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", _ascii_store_bogstaver(value))


def _normaliser_nummerplade(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", _ascii_store_bogstaver(value))


def _gyldigt_stelnummer(value: str) -> bool:
    return bool(re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", value))


def _gyldig_nummerplade(value: str) -> bool:
    min_laengde = int(getattr(config, "REGISTRATION_MIN_LENGTH", 5))
    max_laengde = int(getattr(config, "REGISTRATION_MAX_LENGTH", 9))
    if not min_laengde <= len(value) <= max_laengde:
        return False
    if not value.isalnum():
        return False
    # En plade skal indeholde mindst ét bogstav og mindst ét tal. Det reducerer
    # falske fund fra datoer, beløb, OCR-støj og rene sagsnumre.
    return any(tegn.isalpha() for tegn in value) and any(tegn.isdigit() for tegn in value)


# ---------------------------------------------------------------------------
# Hjælpere
# ---------------------------------------------------------------------------

def _normaliser_mail_type(mail_type: str) -> str:
    return re.sub(r"\s+", "_", str(mail_type or "").strip().upper())


def _normaliser_dokumenttekst(tekst: str) -> str:
    tekst = unicodedata.normalize("NFKC", tekst or "")
    tekst = tekst.replace("\u00ad", "")
    tekst = tekst.replace("\u00a0", " ")
    tekst = re.sub(r"[ \t]+", " ", tekst)
    tekst = re.sub(r"\r\n?", "\n", tekst)
    tekst = re.sub(r"\n{3,}", "\n\n", tekst)
    return tekst.strip()


def _ascii_store_bogstaver(value: str) -> str:
    tekst = unicodedata.normalize("NFKD", str(value or ""))
    return tekst.encode("ascii", "ignore").decode("ascii").upper()


def _lokalt_xml_navn(tag: str) -> str:
    lokalt_navn = str(tag).split("}")[-1].split(":")[-1]
    return re.sub(r"[^a-z0-9]", "", _ascii_store_bogstaver(lokalt_navn).lower())


# Konfigurationskonstanter anvendes, hvis de findes. Fallbacks gør modulet
# testbart isoleret og svarer til de aftalte mailtypenavne.
MAIL_TYPE_BOEDEKONTOR = getattr(config, "MAIL_TYPE_BOEDEKONTOR", "BOEDEKONTOR")
MAIL_TYPE_INDBETALINGSKORT = getattr(
    config, "MAIL_TYPE_INDBETALINGSKORT", "INDBETALINGSKORT"
)
MAIL_TYPE_VIGTIG_INFORMATION = getattr(
    config, "MAIL_TYPE_VIGTIG_INFORMATION", "VIGTIG_INFORMATION"
)
MAIL_TYPE_INDKALDELSE_PERIODISK_SYN = getattr(
    config,
    "MAIL_TYPE_INDKALDELSE_PERIODISK_SYN",
    "INDKALDELSE_PERIODISK_SYN",
)
MAIL_TYPE_PAAMINDELSE_PERIODISK_SYN = getattr(
    config,
    "MAIL_TYPE_PAAMINDELSE_PERIODISK_SYN",
    "PAAMINDELSE_PERIODISK_SYN",
)

PDF_BEREGNINGER: dict[str, Callable[[bytes], IdentifierResult]] = {
    _normaliser_mail_type(MAIL_TYPE_BOEDEKONTOR): boedekontor,
    _normaliser_mail_type(MAIL_TYPE_INDBETALINGSKORT): indbetalingskort,
    _normaliser_mail_type(MAIL_TYPE_VIGTIG_INFORMATION): vigtig_information,
    _normaliser_mail_type(
        MAIL_TYPE_INDKALDELSE_PERIODISK_SYN
    ): indkaldelse_periodisk_syn,
    _normaliser_mail_type(
        MAIL_TYPE_PAAMINDELSE_PERIODISK_SYN
    ): paamindelse_periodisk_syn,
}
