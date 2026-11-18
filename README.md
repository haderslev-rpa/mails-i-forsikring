# Manuel behandling til mails-i-forsikring

ZIP-filen indeholder komplette erstatningsfiler og nye filer.

## Erstat

- `behandel.py`
- `config.py`
- `outlook_service.py`

## Tilføj

- `manuel_behandling.py`
- `tests/test_manuel_behandling.py`

## Manuel-id'er

1. Manglende registreringsnummer
2. Manglende stelnummer
3. Intet køretøjsmatch
4. Flere køretøjsmatch
5. Manglende afdelingsmail

Match-antallet vurderes først. Ved præcis ét match valideres derefter
registreringsnummer, stelnummer og afdelingsmail i denne rækkefølge.

Originalmailen røres ikke ved manuel behandling. Informationsmailen sendes kun
én gang via et checkpoint i Box data.

## Kontrol

```bash
uvx ruff format behandel.py config.py outlook_service.py manuel_behandling.py tests/test_manuel_behandling.py
uvx ruff check behandel.py config.py outlook_service.py manuel_behandling.py tests/test_manuel_behandling.py
uv run python -m py_compile behandel.py config.py outlook_service.py manuel_behandling.py tests/test_manuel_behandling.py
uv run pytest tests/test_manuel_behandling.py -q