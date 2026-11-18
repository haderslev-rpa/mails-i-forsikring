
"""Central konfiguration for processen mails-i-forsikring.

Proces-skabelonen bevares:
- populate_queue.py opretter ATS-kø-items.
- behandel.py behandler ét kø-item ad gangen.
- servicefilerne håndterer kun deres respektive integrationer.

Outlook-authentication håndteres internt af q_outlook_api. Derfor findes der
ingen Outlook-credentials, tokens eller client secrets i denne fil.
"""

from __future__ import annotations

import os
from typing import Final


# ============================================================
# PROCES OG AUTOMATION SERVER
# ============================================================

PROCESS_NAME: Final[str] = "mails-i-forsikring"
QUEUE_NAME: Final[str] = "mails-i-forsikring"
SCHEMA_VERSION: Final[int] = 1
BOX_SCHEMA_VERSION: Final[int] = 1
TIMEZONE: Final[str] = "Europe/Copenhagen"
INSUBIZ_CREDENTIAL_NAME: Final[str] = "Q_INSUBIZ"
CLEAR_NEW_QUEUE_ITEMS_BEFORE_POPULATE: Final[bool] = False


# ============================================================
# ITEM-STATUSSER OG STATES
# ============================================================

STATUS_COMPLETED: Final[str] = "Completed"
STATUS_MANUAL: Final[str] = "Manuel"
STATUS_MANUEL: Final[str] = STATUS_MANUAL
STATUS_BUSINESS_EXCEPTION: Final[str] = "BusinessException"
STATUS_SYSTEM_EXCEPTION: Final[str] = "SystemException"

STATE_MANGLENDE_AFDELINGSMAIL: Final[str] = (
    "5.0 Manuel behandling - Afdelingsmail mangler på køretøj"
)


# ============================================================
# OUTLOOK VIA q_outlook_api
# ============================================================

# Teknisk postkasseadresse. q_outlook_api bruger værdien som user_mail.
OUTLOOK_MAILBOX_ADDRESS: Final[str] = "forsikring@haderslev.dk"
OUTLOOK_MAILBOX_USER_ID: Final[str] = OUTLOOK_MAILBOX_ADDRESS

# Læsbart forretningsnavn og Graphs sprogneutrale standardmappe.
OUTLOOK_MAILBOX_NAME: Final[str] = "Forsikring"
OUTLOOK_FOLDER_NAME: Final[str] = "Indbakke"
OUTLOOK_DEFAULT_FOLDER: Final[str] = "inbox"

# Modtager ved manglende Afdelingsmail.
FORSIKRINGSTEAM_EMAIL: Final[str] = "forsikring@haderslev.dk"
MISSING_VEHICLE_EMAIL_RECIPIENT: Final[str] = FORSIKRINGSTEAM_EMAIL

# Hentning.
OUTLOOK_UNREAD_ONLY: Final[bool] = False
OUTLOOK_PAGE_SIZE: Final[int] = 100
OUTLOOK_ATTACHMENT_PAGE_SIZE: Final[int] = 100
OUTLOOK_MAX_ATTACHMENT_SIZE_BYTES: Final[int] = 25 * 1024 * 1024
OUTLOOK_TIMEOUT_SECONDS: Final[float] = 120.0
OUTLOOK_MAX_RETRIES: Final[int] = 3
OUTLOOK_RETRY_DELAY_SECONDS: Final[float] = 2.0

OUTLOOK_SELECT_FIELDS: Final[tuple[str, ...]] = (
    "id",
    "internetMessageId",
    "receivedDateTime",
    "subject",
    "hasAttachments",
)

OUTLOOK_MESSAGE_SELECT_FIELDS: Final[tuple[str, ...]] = (
    "id",
    "internetMessageId",
    "conversationId",
    "receivedDateTime",
    "sentDateTime",
    "subject",
    "body",
    "bodyPreview",
    "from",
    "sender",
    "toRecipients",
    "ccRecipients",
    "bccRecipients",
    "replyTo",
    "hasAttachments",
    "categories",
    "importance",
    "isRead",
    "webLink",
)

PDF_CONTENT_TYPES: Final[frozenset[str]] = frozenset(
    {"application/pdf", "application/x-pdf"}
)

OUTLOOK_FORWARD_COMMENT: Final[str] = (
    "Mailen videresendes automatisk vedrørende køretøjet."
)
OUTLOOK_FORWARD_MODE: Final[str] = "comment"
OUTLOOK_MANUAL_SUBJECT: Final[str] = (
    "Manuel behandling - Afdelingsmail mangler på køretøj"
)
OUTLOOK_SAVE_TO_SENT_ITEMS: Final[bool] = True

# Kompatibilitetsnavne til eksisterende test-, populate- og hjælpefiler.
# MAILBOX_NAME skal være den tekniske adresse, fordi q_outlook_api bruger
# værdien som user_mail. OUTLOOK_MAILBOX_NAME er det læsbare navn.
MAILBOX_NAME: Final[str] = OUTLOOK_MAILBOX_ADDRESS
MAILBOX_ADDRESS: Final[str] = OUTLOOK_MAILBOX_ADDRESS
MAILBOX_USER_ID: Final[str] = OUTLOOK_MAILBOX_USER_ID

# Rettelsen til den viste ImportError: begge navne peger på samme mappe.
MAIL_FOLDER_NAME: Final[str] = OUTLOOK_FOLDER_NAME
FOLDER_NAME: Final[str] = OUTLOOK_FOLDER_NAME
PAGE_SIZE: Final[int] = OUTLOOK_PAGE_SIZE
UNREAD_ONLY: Final[bool] = OUTLOOK_UNREAD_ONLY


# ============================================================
# MAILTYPER OG EMNEREGLER
# ============================================================

MAIL_TYPE_BOEDEKONTOR: Final[str] = "BOEDEKONTOR"
MAIL_TYPE_INDBETALINGSKORT: Final[str] = "INDBETALINGSKORT"
MAIL_TYPE_VIGTIG_INFORMATION: Final[str] = "VIGTIG_INFORMATION"
MAIL_TYPE_INDKALDELSE_PERIODISK_SYN: Final[str] = (
    "INDKALDELSE_PERIODISK_SYN"
)
MAIL_TYPE_PAAMINDELSE_PERIODISK_SYN: Final[str] = (
    "PAAMINDELSE_PERIODISK_SYN"
)

SUBJECT_BOEDEKONTOR: Final[str] = "Bødekontor"
SUBJECT_INDBETALINGSKORT: Final[str] = "Indbetalingskort"
SUBJECT_VIGTIG_INFORMATION: Final[str] = "Vigtig information"
SUBJECT_INDKALDELSE_PERIODISK_SYN: Final[str] = (
    "Indkaldelse til periodisk syn"
)
SUBJECT_PAAMINDELSE_PERIODISK_SYN: Final[str] = (
    "Påmindelse om periodisk syn af køretøj"
)

# Kun denne Digitalpost-afsender og denne konkrete betalingstype.
MAIL_TYPE_DIGITALPOST_BETALING: Final[str] = "DIGITALPOST_BETALING"
DIGITALPOST_BETALING_AFSENDER_ID: Final[str] = "29189757"
DIGITALPOST_BETALING_TEKST: Final[str] = "Du skal betale et beløb"
DIGITALPOST_BETALING_EMNE_PATTERN: Final[str] = (
    rf"^Digitalpost\s*\({DIGITALPOST_BETALING_AFSENDER_ID}\)\s*:\s*"
    rf"{DIGITALPOST_BETALING_TEKST}\b"
)
SUBJECT_RULES: Final[tuple[tuple[str, str], ...]] = (
    (
        MAIL_TYPE_PAAMINDELSE_PERIODISK_SYN,
        SUBJECT_PAAMINDELSE_PERIODISK_SYN,
    ),
    (
        MAIL_TYPE_INDKALDELSE_PERIODISK_SYN,
        SUBJECT_INDKALDELSE_PERIODISK_SYN,
    ),
    (MAIL_TYPE_DIGITALPOST_BETALING, "Digitalpost (29189757): Du skal betale et beløb"),
    (MAIL_TYPE_INDBETALINGSKORT, SUBJECT_INDBETALINGSKORT),
    (MAIL_TYPE_BOEDEKONTOR, SUBJECT_BOEDEKONTOR),
    (MAIL_TYPE_VIGTIG_INFORMATION, SUBJECT_VIGTIG_INFORMATION),
)

SUBJECT_TO_MAIL_TYPE: Final[dict[str, str]] = {
    subject.casefold(): mail_type for mail_type, subject in SUBJECT_RULES
}

ALLOWED_MAIL_TYPES: Final[frozenset[str]] = frozenset(
    mail_type for mail_type, _ in SUBJECT_RULES
)

SUBJECT_PREFIXES: Final[tuple[str, ...]] = (
    "SV:",
    "VS:",
    "RE:",
    "FW:",
    "FWD:",
)


# ============================================================
# KØ-REFERENCE
# ============================================================

QUEUE_REFERENCE_PREFIX: Final[str] = "outlook"
QUEUE_REFERENCE_DATETIME_FORMAT: Final[str] = "%Y-%m-%d - %H:%M:%S"
QUEUE_REFERENCE_MAIL_ID_HASH_LENGTH = 12
QUEUE_REFERENCE_TIME_FORMAT: Final[str] = "%Y%m%dT%H%M%S%f"
QUEUE_REFERENCE_MILLISECOND_DIGITS: Final[int] = 3
QUEUE_REFERENCE_MAX_MESSAGE_ID_LENGTH: Final[int] = 160
QUEUE_REFERENCE_SEPARATOR: Final[str] = " - "


# ============================================================
# ITEM DATA
# ============================================================

ITEM_KEY_MAIL_ID: Final[str] = "mail_id"
ITEM_KEY_INTERNET_MESSAGE_ID: Final[str] = "internet_message_id"
ITEM_KEY_RECEIVED_AT: Final[str] = "received_at"
ITEM_KEY_SUBJECT: Final[str] = "subject"
ITEM_KEY_MAIL_TYPE: Final[str] = "mail_type"

QUEUE_DATA_MAIL_ID: Final[str] = ITEM_KEY_MAIL_ID
QUEUE_DATA_INTERNET_MESSAGE_ID: Final[str] = ITEM_KEY_INTERNET_MESSAGE_ID
QUEUE_DATA_RECEIVED_AT: Final[str] = ITEM_KEY_RECEIVED_AT
QUEUE_DATA_RECEIVED_AT_MILLISECONDS: Final[str] = "received_at_milliseconds"
QUEUE_DATA_SUBJECT: Final[str] = ITEM_KEY_SUBJECT
QUEUE_DATA_MAIL_TYPE: Final[str] = ITEM_KEY_MAIL_TYPE
QUEUE_DATA_MAILBOX: Final[str] = "mailbox"
QUEUE_DATA_FOLDER: Final[str] = "folder"
QUEUE_DATA_HAS_ATTACHMENTS: Final[str] = "has_attachments"
QUEUE_DATA_SCHEMA_VERSION: Final[str] = "schema_version"

REQUIRED_ITEM_DATA_FIELDS: Final[tuple[str, ...]] = (
    ITEM_KEY_MAIL_ID,
    ITEM_KEY_RECEIVED_AT,
    ITEM_KEY_SUBJECT,
    ITEM_KEY_MAIL_TYPE,
)

REQUIRED_QUEUE_DATA_FIELDS: Final[tuple[str, ...]] = (
    QUEUE_DATA_MAIL_ID,
    QUEUE_DATA_RECEIVED_AT,
    QUEUE_DATA_SUBJECT,
    QUEUE_DATA_MAIL_TYPE,
)


# ============================================================
# BOX DATA OG CHECKPOINTS
# ============================================================

BOX_KEY_SCHEMA_VERSION: Final[str] = "schema_version"
BOX_KEY_IDENTIFIER: Final[str] = "identifikator"
BOX_KEY_VEHICLE: Final[str] = "koeretoej"
BOX_KEY_DELIVERY: Final[str] = "levering"
BOX_KEY_IDENTIFIER_TYPE: Final[str] = "type"
BOX_KEY_IDENTIFIER_VALUE: Final[str] = "vaerdi"
BOX_KEY_ATTACHMENT_NAME: Final[str] = "vedhaeftning"
BOX_KEY_VEHICLE_FOUND: Final[str] = "fundet"
BOX_KEY_REGISTRATION: Final[str] = "nummerplade"
BOX_KEY_CHASSIS_NUMBER: Final[str] = "stelnummer"
BOX_KEY_DEPARTMENT_EMAIL: Final[str] = "afdelingsmail"
BOX_KEY_SENT_TO_DEPARTMENT: Final[str] = "sendt_til_afdelingsmail"
BOX_KEY_DEPARTMENT_SENT_AT: Final[str] = "afdelingsmail_sendt_tidspunkt"
BOX_KEY_MAIL_DELETED: Final[str] = "mail_slettet"
BOX_KEY_MANUAL_EMAIL_SENT: Final[str] = "manuel_mail_sendt"
BOX_KEY_MANUAL_EMAIL_SENT_AT: Final[str] = "manuel_mail_sendt_tidspunkt"
BOX_KEY_MANUAL_ID: Final[str] = "manuel_id"
STATE_MANUEL_BEHANDLING: Final[str] = "Manuel behandling"

# Engelske aliases fra den tidligere konfiguration.
BOX_IDENTIFIER_KEY: Final[str] = BOX_KEY_IDENTIFIER
BOX_IDENTIFIER_TYPE_KEY: Final[str] = BOX_KEY_IDENTIFIER_TYPE
BOX_IDENTIFIER_VALUE_KEY: Final[str] = BOX_KEY_IDENTIFIER_VALUE
BOX_IDENTIFIER_ATTACHMENT_NAME_KEY: Final[str] = BOX_KEY_ATTACHMENT_NAME
BOX_VEHICLE_KEY: Final[str] = BOX_KEY_VEHICLE
BOX_VEHICLE_MATCHED_KEY: Final[str] = BOX_KEY_VEHICLE_FOUND
BOX_VEHICLE_REGISTRATION_KEY: Final[str] = BOX_KEY_REGISTRATION
BOX_VEHICLE_CHASSIS_NUMBER_KEY: Final[str] = BOX_KEY_CHASSIS_NUMBER
BOX_VEHICLE_EMAIL_KEY: Final[str] = BOX_KEY_DEPARTMENT_EMAIL
BOX_DELIVERY_KEY: Final[str] = BOX_KEY_DELIVERY
BOX_FORWARD_SENT_KEY: Final[str] = BOX_KEY_SENT_TO_DEPARTMENT
BOX_FORWARDED_AT_KEY: Final[str] = BOX_KEY_DEPARTMENT_SENT_AT
BOX_MAIL_DELETED_KEY: Final[str] = BOX_KEY_MAIL_DELETED


# ============================================================
# PDF-BEREGNINGER OG VEHICLE MATCHER
# ============================================================

ALLOWED_ATTACHMENT_EXTENSIONS: Final[frozenset[str]] = frozenset({".pdf"})
PDF_MIME_TYPES: Final[frozenset[str]] = PDF_CONTENT_TYPES
PDF_REQUIRE_EXACTLY_ONE_ATTACHMENT: Final[bool] = True
PDF_TEXT_MAX_CHARACTERS: Final[int] = 250_000
PDF_MAX_PAGES: Final[int] = 100

VIN_EXPECTED_LENGTH: Final[int] = 17
VIN_LENGTH: Final[int] = VIN_EXPECTED_LENGTH
VIN_PATTERN: Final[str] = r"^[A-HJ-NPR-Z0-9]{17}$"

REGISTRATION_MIN_LENGTH: Final[int] = 5
REGISTRATION_MAX_LENGTH: Final[int] = 9
REGISTRATION_PATTERN: Final[str] = r"^[A-Z0-9]{5,9}$"
REGISTRATION_REQUIRE_LETTER_AND_DIGIT: Final[bool] = True
VEHICLE_MATCHER_ERROR_DETAIL_LIMIT: Final[int] = 5

PDF_PARSER_BY_MAIL_TYPE: Final[dict[str, str]] = {
    MAIL_TYPE_BOEDEKONTOR: "boedekontor",
    MAIL_TYPE_INDBETALINGSKORT: "indbetalingskort",
    MAIL_TYPE_VIGTIG_INFORMATION: "vigtig_information",
    MAIL_TYPE_INDKALDELSE_PERIODISK_SYN: "indkaldelse_periodisk_syn",
    MAIL_TYPE_PAAMINDELSE_PERIODISK_SYN: "paamindelse_periodisk_syn",
}


# ============================================================
# INSUBIZ
# ============================================================

# q-insubiz henter login fra ATS-credentialen Q_INSUBIZ.
# URL og adgangskode skal ikke ligge i denne proces eller i .env.
INSUBIZ_HEADLESS: Final[bool] = True
INSUBIZ_DEBUG: Final[bool] = False
INSUBIZ_CUSTOMER_ID: Final[int] = 0
INSUBIZ_SHOW_TREE_DATA: Final[bool] = False
INSUBIZ_ACTIVE_ONLY: Final[bool] = False

INSUBIZ_CACHE_ENABLED: Final[bool] = True
INSUBIZ_CACHE_TTL_SECONDS: Final[int | None] = None
CACHE_VEHICLE_LIST_PER_RUN: Final[bool] = True
INSUBIZ_SHEET_NAME: Final[int | str] = 0
INSUBIZ_CSV_SEPARATOR: Final[str | None] = None

INSUBIZ_CHASSIS_COLUMN: Final[str] = "ChassisNumber"
INSUBIZ_REGISTRATION_COLUMN: Final[str] = "Registration"
INSUBIZ_DEPARTMENT_EMAIL_COLUMN: Final[str] = "ContactPerson1"
INSUBIZ_EMAIL_COLUMN: Final[str] = INSUBIZ_DEPARTMENT_EMAIL_COLUMN

INSUBIZ_VEHICLE_COLUMNS: Final[tuple[str, ...]] = (
    INSUBIZ_CHASSIS_COLUMN,
    INSUBIZ_DEPARTMENT_EMAIL_COLUMN,
    INSUBIZ_REGISTRATION_COLUMN,
)

INSUBIZ_REQUIRED_COLUMNS: Final[tuple[str, ...]] = ()

INSUBIZ_COLUMN_ALIASES: Final[dict[str, tuple[str, ...]]] = {
    INSUBIZ_CHASSIS_COLUMN: ("Stelnr.", "Stelnummer", "Stel nr", "VIN"),
    INSUBIZ_REGISTRATION_COLUMN: (
        "Reg.nr.",
        "RegistrationNumber",
        "Registreringsnummer",
        "Nummerplade",
    ),
    INSUBIZ_DEPARTMENT_EMAIL_COLUMN: (
        "Afdelingsmail",
        "Email",
        "E-mail",
    ),
}

INSUBIZ_DUPLICATE_POLICY: Final[str] = "keep"



# ============================================================
# FORRETNINGSTEKSTER OG FEJL
# ============================================================

FORWARD_BODY_HTML: Final[str] = (
    "<p>Denne mail er videresendt automatisk af robotten "
    "mails-i-forsikring.</p>"
    "<p>Mailen er matchet med køretøjets registreringsnummer "
    "eller stelnummer i Insubiz.</p>"
)

MISSING_EMAIL_SUBJECT_PREFIX: Final[str] = (
    "Manglende kontaktmail på køretøj"
)

MISSING_EMAIL_BODY_TEMPLATE: Final[str] = (
    "<p>Robotten har fundet køretøjet i Insubiz, men køretøjet "
    "har ingen gyldig Afdelingsmail.</p>"
    "<p>Originalmailen er ikke flyttet eller slettet og skal "
    "behandles manuelt.</p>"
)

SAVE_SENT_MAILS: Final[bool] = True
DELETE_MAIL_AFTER_CONFIRMED_FORWARD: Final[bool] = True
MOVE_MAIL_WITH_MISSING_RECIPIENT: Final[bool] = False
IGNORE_MISSING_MAIL_WHEN_DELETING: Final[bool] = True

BUSINESS_ERROR_NO_PDF: Final[str] = "Mailen indeholder ingen PDF."
BUSINESS_ERROR_MULTIPLE_PDFS: Final[str] = (
    "Mailen indeholder flere PDF-filer."
)
BUSINESS_ERROR_IDENTIFIER_NOT_FOUND: Final[str] = (
    "Stelnummer eller nummerplade kunne ikke udtrækkes."
)
BUSINESS_ERROR_VEHICLE_NOT_FOUND: Final[str] = (
    "Køretøjet blev ikke fundet i Insubiz."
)
BUSINESS_ERROR_MULTIPLE_VEHICLES: Final[str] = (
    "Flere køretøjer matcher identifikatoren."
)
BUSINESS_ERROR_VEHICLE_EMAIL_MISSING: Final[str] = (
    "Køretøjet har ingen gyldig Afdelingsmail."
)
SYSTEM_ERROR_MAIL_NOT_FOUND: Final[str] = (
    "Mailen kunne ikke genfindes i Outlook."
)


# ============================================================
# HJÆLPEFUNKTIONER OG VALIDERING
# ============================================================

def get_subject_rules() -> tuple[tuple[str, str], ...]:
    return SUBJECT_RULES


def get_pdf_parser_name(mail_type: str) -> str:
    normalized = str(mail_type or "").strip().upper()
    try:
        return PDF_PARSER_BY_MAIL_TYPE[normalized]
    except KeyError as error:
        raise ValueError(
            f"Mailtypen {mail_type!r} har ingen konfigureret PDF-parser."
        ) from error


def get_insubiz_vehicle_columns() -> list[str]:
    return list(INSUBIZ_VEHICLE_COLUMNS)


def create_initial_box_data() -> dict[str, object]:
    return {
        BOX_KEY_SCHEMA_VERSION: BOX_SCHEMA_VERSION,
        BOX_KEY_IDENTIFIER: {},
        BOX_KEY_VEHICLE: {BOX_KEY_VEHICLE_FOUND: False},
        BOX_KEY_DELIVERY: {
            BOX_KEY_SENT_TO_DEPARTMENT: False,
            BOX_KEY_MAIL_DELETED: False,
        },
    }


def validate_queue_data(data: dict[str, object]) -> None:
    if not isinstance(data, dict):
        raise TypeError("Queue data skal være en dictionary.")

    missing = [
        field for field in REQUIRED_ITEM_DATA_FIELDS if not data.get(field)
    ]
    if missing:
        raise ValueError(
            "Queue data mangler obligatoriske felter: " + ", ".join(missing)
        )

    mail_type = str(data[ITEM_KEY_MAIL_TYPE]).strip().upper()
    if mail_type not in ALLOWED_MAIL_TYPES:
        raise ValueError(f"Mailtypen understøttes ikke: {mail_type!r}.")


def validate_config(*, require_insubiz: bool = False) -> None:
    errors: list[str] = []

    if not PROCESS_NAME.strip():
        errors.append("PROCESS_NAME mangler")

    if not QUEUE_NAME.strip():
        errors.append("QUEUE_NAME mangler")

    if not OUTLOOK_MAILBOX_ADDRESS.strip():
        errors.append("OUTLOOK_MAILBOX_ADDRESS mangler")

    if not OUTLOOK_FOLDER_NAME.strip():
        errors.append("OUTLOOK_FOLDER_NAME mangler")

    if not FORSIKRINGSTEAM_EMAIL.strip():
        errors.append("FORSIKRINGSTEAM_EMAIL mangler")

    if REGISTRATION_MIN_LENGTH > REGISTRATION_MAX_LENGTH:
        errors.append(
            "REGISTRATION_MIN_LENGTH må ikke være større end "
            "REGISTRATION_MAX_LENGTH"
        )

    if OUTLOOK_FORWARD_MODE not in {"comment", "formatted"}:
        errors.append(
            "OUTLOOK_FORWARD_MODE skal være comment eller formatted"
        )

    if INSUBIZ_DUPLICATE_POLICY not in {"keep", "error"}:
        errors.append(
            "INSUBIZ_DUPLICATE_POLICY skal være keep eller error"
        )

    # q-insubiz validerer selv Q_INSUBIZ-credentialen ved login.

    if errors:
        raise RuntimeError(
            "Ugyldig configuration.py: " + "; ".join(errors)
        )


# Outlook-konfiguration valideres ved import; Insubiz-login sker ved opslag.
validate_config(require_insubiz=False)
