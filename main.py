"""Entrypoint for Automation Server-processen mails-i-forsikring.

--queue henter mails og opretter ATS-items.
Uden --queue behandles eksisterende ATS-items et ad gangen.
Produceren behandler aldrig mails; workeren opretter aldrig kø-items.
"""
from __future__ import annotations

import argparse
import asyncio
import inspect
import logging
import sys
from contextlib import nullcontext
from typing import Any, Callable, Mapping

from dotenv import load_dotenv

import configuration as config
from behandel import Forretningsfejl, Systemfejl, process_item
from populate_queue import get_relevant_queue_items

try:
    from automation_server_client import AutomationServer, Workqueue
except ImportError as error:
    raise RuntimeError(
        "automation-server-client er ikke installeret. Kør 'uv sync'."
    ) from error

logger = logging.getLogger(__name__)


def _positive_int(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("Værdien skal være et heltal.") from error
    if parsed <= 0:
        raise argparse.ArgumentTypeError("Værdien skal være større end 0.")
    return parsed


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog=config.PROCESS_NAME,
        description="Automation Server-proces til købaseret behandling af mails i Forsikring.",
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--queue", action="store_true", help="Hent relevante mails og opret ATS-kø-items.")
    mode.add_argument("--process", action="store_true", help="Behandl eksisterende kø-items. Standardtilstand.")
    parser.add_argument("--max-items", type=_positive_int, default=None,
                        help="Maksimalt antal kø-items i denne kørsel.")
    parser.add_argument("--debug", action="store_true", help="Aktivér debug-logging.")
    args = parser.parse_args(argv)
    if args.queue and args.max_items is not None:
        parser.error("--max-items kan kun bruges ved itembehandling, ikke med --queue.")
    return args


def main(argv: list[str] | None = None) -> int:
    """Kør producer eller worker. Returnér proces-exitkode."""
    load_dotenv()
    args = parse_args(argv)
    _configure_logging(debug=args.debug)
    logger.info("Starter proces: %s", config.PROCESS_NAME)
    try:
        config.validate_config(require_insubiz=not args.queue)
        automation_server = _create_automation_server()
        workqueue = _get_workqueue(automation_server)
        if args.queue:
            antal = populate_workqueue(workqueue)
            logger.info("Kø-populering afsluttet. Nye items: %s", antal)
            return 0
        antal = process_workqueue(workqueue, max_items=args.max_items)
        logger.info("Itembehandling afsluttet. Behandlede items: %s", antal)
        return 0
    except KeyboardInterrupt:
        logger.warning("Processen blev afbrudt fra terminalen.")
        return 130
    except Exception:
        logger.exception("Processen fejlede.")
        return 1


def populate_workqueue(workqueue: Any) -> int:
    """Opret ATS-kø-items fra populate_queue uden at behandle mails."""
    queue_items = get_relevant_queue_items()
    if not isinstance(queue_items, list):
        raise TypeError("get_relevant_queue_items() skal returnere en liste.")
    oprettede = 0
    dubletter = 0
    for index, queue_item in enumerate(queue_items, start=1):
        reference, data = _validate_queue_item(queue_item, index=index)
        try:
            _add_item(workqueue=workqueue, reference=reference, data=data)
        except Exception as error:
            if _is_duplicate_error(error):
                dubletter += 1
                logger.info("Kø-item findes allerede. reference=%s", reference)
                continue
            raise
        oprettede += 1
        logger.info("Kø-item oprettet. reference=%s, mail_id=%s",
                    reference, data.get(config.QUEUE_DATA_MAIL_ID))
    logger.info("Populate-resultat. Relevant=%s, oprettet=%s, dubletter=%s",
                len(queue_items), oprettede, dubletter)
    return oprettede


def _validate_queue_item(queue_item: Any, *, index: int) -> tuple[str, dict[str, Any]]:
    if not isinstance(queue_item, Mapping):
        raise TypeError(f"Kø-element nr. {index} skal være en dictionary.")
    reference = queue_item.get("reference")
    data = queue_item.get("data")
    if not isinstance(reference, str) or not reference.strip():
        raise ValueError(f"Kø-element nr. {index} mangler en gyldig reference.")
    if not isinstance(data, Mapping):
        raise TypeError(f"Kø-element nr. {index} mangler dictionary-feltet data.")
    normalized_data = dict(data)
    config.validate_queue_data(normalized_data)
    return reference.strip(), normalized_data


def process_workqueue(workqueue: Any, *, max_items: int | None = None) -> int:
    """Behandl ATS-items sekventielt gennem behandel.process_item()."""
    behandlet = 0
    while max_items is None or behandlet < max_items:
        item = _get_next_item(workqueue)
        if item is None:
            logger.info("Ingen flere kø-items er klar til behandling.")
            break
        reference = _item_reference(item)
        logger.info("Behandler kø-item. reference=%s", reference)
        try:
            with _item_context(item) as active_item:
                result = asyncio.run(process_item(active_item))
                _complete_item_if_needed(active_item, result=result)
        except Forretningsfejl as error:
            logger.warning("Forretningsfejl. reference=%s, fejl=%s", reference, error)
            _fail_item_if_needed(item, error=error,
                                 status=config.STATUS_BUSINESS_EXCEPTION)
        except Systemfejl as error:
            logger.exception("Systemfejl. reference=%s", reference)
            _fail_item_if_needed(item, error=error,
                                 status=config.STATUS_SYSTEM_EXCEPTION)
        except Exception as error:
            logger.exception("Uventet fejl. reference=%s", reference)
            _fail_item_if_needed(item, error=error,
                                 status=config.STATUS_SYSTEM_EXCEPTION)
        behandlet += 1
    return behandlet


def _create_automation_server() -> Any:
    """Initialisér Automation Server fra miljøet."""
    factory = getattr(AutomationServer, "from_environment", None)
    if not callable(factory):
        raise RuntimeError("AutomationServer mangler from_environment().")
    return _resolve_sync(factory())


def _get_workqueue(automation_server: Any) -> Any:
    """Hent processens workqueue."""
    values = {"name": config.QUEUE_NAME,
              "queue_name": config.QUEUE_NAME,
              "workqueue_name": config.QUEUE_NAME}
    for method_name in ("get_workqueue", "workqueue"):
        method = getattr(automation_server, method_name, None)
        if callable(method):
            return _resolve_sync(_call_with_known_arguments(method, values))
    for method_name in ("get_workqueue", "get_by_name", "from_name"):
        method = getattr(Workqueue, method_name, None)
        if callable(method):
            return _resolve_sync(_call_with_known_arguments(method, values))
    try:
        return Workqueue(config.QUEUE_NAME)
    except TypeError as error:
        raise RuntimeError("Kunne ikke oprette workqueue med den installerede klient.") from error


def _add_item(*, workqueue: Any, reference: str, data: dict[str, Any]) -> Any:
    method = getattr(workqueue, "add_item", None)
    if not callable(method):
        raise TypeError("Workqueue mangler add_item().")
    result = _call_with_known_arguments(
        method, {"reference": reference, "data": data,
                 "item_data": data, "payload": data}
    )
    return _resolve_sync(result)


# ============================================================
# START: _get_next_item
# ============================================================


def _get_next_item(workqueue: Any) -> Any | None:
    """
    Henter næste tilgængelige item fra Automation Server-køen.

    Workqueue fra automation-server-client v0.3.0 er iterable.
    Iteratoren henter og reserverer det næste tilgængelige item.

    Output:
        Et WorkItem, hvis køen indeholder et tilgængeligt item.
        None, hvis køen er tom.
    """
    if workqueue is None:
        raise TypeError("workqueue skal angives.")

    try:
        return next(iter(workqueue))
    except StopIteration:
        return None


# ============================================================
# SLUT: _get_next_item
# ============================================================

def _item_context(item: Any) -> Any:
    if hasattr(item, "__enter__") and hasattr(item, "__exit__"):
        return item
    return nullcontext(item)


# START: _complete_item_if_needed()

def _complete_item_if_needed(
    item: Any,
    *,
    result: Any,
) -> None:
    """Overfør det gemte forretningsresultat til teknisk ATS-status."""
    from automation_server_client import WorkItemStatus

    # Box er allerede gemt af behandel.py.
    # Returværdien må ikke erstatte hele item.data.
    del result

    teknisk_status = getattr(item, "status", None)
    teknisk_vaerdi = getattr(
        teknisk_status,
        "value",
        teknisk_status,
    )

    terminale_statusser = {
        WorkItemStatus.COMPLETED.value,
        WorkItemStatus.FAILED.value,
        WorkItemStatus.PENDING_USER_ACTION.value,
    }

    if teknisk_vaerdi in terminale_statusser:
        return

    data = item.data

    if not isinstance(data, Mapping):
        raise TypeError("ATS-itemets data skal være en dictionary.")

    status_data = data.get("status")

    if not isinstance(status_data, Mapping):
        raise RuntimeError(
            "behandel.py har ikke gemt et forretningsresultat."
        )

    forretningsstatus = status_data.get("status")
    besked = str(
        status_data.get("status_code")
        or forretningsstatus
        or ""
    ).strip()

    if forretningsstatus == config.STATUS_MANUAL:
        item.update_status(
            WorkItemStatus.PENDING_USER_ACTION,
            besked,
        )

        logger.info(
            "ATS-item afventer manuel behandling. reference=%s",
            item.reference,
        )
        return

    if forretningsstatus == config.STATUS_COMPLETED:
        item.complete(besked)

        logger.info(
            "ATS-item afsluttet. reference=%s",
            item.reference,
        )
        return

    raise RuntimeError(
        "Itemet kan ikke afsluttes med forretningsstatus "
        f"{forretningsstatus!r}."
    )

# SLUT: _complete_item_if_needed()


def _fail_item_if_needed(item: Any, *, error: Exception, status: str) -> None:
    if _item_has_terminal_status(item):
        return
    for method_name in ("fail", "set_failed", "set_status", "update_status"):
        method = getattr(item, method_name, None)
        if not callable(method):
            continue
        values = {"status": status, "error": error, "exception": error,
                  "message": str(error), "error_message": str(error)}
        _resolve_sync(_call_with_known_arguments(method, values))
        return


def _item_has_terminal_status(item: Any) -> bool:
    value = _read_item_value(item, "status")
    if value is None:
        return False
    normalized = str(value).strip().casefold()
    terminal = {
        config.STATUS_COMPLETED.casefold(), config.STATUS_MANUAL.casefold(),
        config.STATUS_BUSINESS_EXCEPTION.casefold(),
        config.STATUS_SYSTEM_EXCEPTION.casefold(), "done", "failed",
    }
    return normalized in terminal


def _item_reference(item: Any) -> str:
    value = (_read_item_value(item, "reference")
             or _read_item_value(item, "id") or "[ukendt reference]")
    return str(value)


def _read_item_value(item: Any, name: str) -> Any:
    if isinstance(item, Mapping):
        return item.get(name)
    return getattr(item, name, None)


def _call_with_known_arguments(
    function: Callable[..., Any], values: Mapping[str, Any]
) -> Any:
    """Send kun argumenter, som den installerede klientfunktion accepterer."""
    signature = inspect.signature(function)
    kwargs: dict[str, Any] = {}
    missing: list[str] = []
    for name, parameter in signature.parameters.items():
        if parameter.kind in {inspect.Parameter.VAR_POSITIONAL,
                              inspect.Parameter.VAR_KEYWORD}:
            continue
        if name in values and values[name] is not None:
            kwargs[name] = values[name]
        elif parameter.default is inspect.Parameter.empty:
            missing.append(name)
    if missing:
        raise RuntimeError(
            f"{function.__qualname__} kræver ukendte obligatoriske "
            f"argumenter: {missing!r}. Signatur: {signature}."
        )
    return function(**kwargs)


def _resolve_sync(value: Any) -> Any:
    """Afvikl awaitable returværdier i synkron main."""
    if not inspect.isawaitable(value):
        return value
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return asyncio.run(value)
    raise RuntimeError("Asynkron ATS-metode blev kaldt fra en aktiv event loop.")


def _is_duplicate_error(error: Exception) -> bool:
    response = getattr(error, "response", None)
    status_code = getattr(response, "status_code", None)
    if status_code == 409 or getattr(error, "status_code", None) == 409:
        return True
    message = str(error).casefold()
    return any(marker in message for marker in (
        "duplicate", "already exists", "reference already",
        "findes allerede", "allerede oprettet",
    ))


def _configure_logging(*, debug: bool) -> None:
    """Skriv proceslog og fejltraceback til terminalen, også ved eksisterende handlers."""
    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        stream=sys.stderr,
        force=True,
    )
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("requests").setLevel(logging.WARNING)
    logging.getLogger("httpx").setLevel(logging.WARNING)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
