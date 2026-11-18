# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

The goal is to post a message in a friends' WhatsApp group whenever there is a noteworthy event, agenda change or birthday. The pipeline is:

1. A daily GitHub Actions job runs `main.py`, which writes three text files (Dutch-language) to the repo root: today's Google Calendar events (`today_events.txt`), events added/modified in the last 24 hours (`modified_events.txt`), and birthday messages (`birthday_messages.txt`).
2. Every day, Apple Shortcuts on the owner's iPhone read these files. If a file is empty, nothing happens. If it has content, another Shortcut sends it to the WhatsApp group.

There are no tests or linter configured.

## Commands

```bash
pip install -r requirements.txt
python main.py
```

`main.py` reads three environment variables: `CALENDAR_API_KEY`, `CALENDAR_ID`, `BIRTHDAY_URL` (a Google Drive link to the birthdays xlsx). It does **not** import `config.py`. That file is git-ignored and holds the same three values for local use, so export them yourself first (e.g. `set -a; source config.py; set +a`, which works if it's in `KEY=value` form).

## Architecture

- **Deployment is the GitHub Actions workflow** (`.github/workflows/main.yml`): it runs `main.py` daily at 00:30 UTC (and on manual dispatch) with the three values as repo secrets, then commits any changed files with the message "Update TXT files" and pushes to `master`. Because of that, the output files are tracked in git and the commit history is mostly bot commits.
- **Outputs are the three tracked `.txt` files** in the repo root (`today_events.txt`, `modified_events.txt`, `birthday_messages.txt`). They are overwritten on every run. **An empty file is meaningful**: it signals "send nothing", so never write placeholder or "no events" text into them. The iPhone Shortcuts consume them, and the text is sent verbatim as a WhatsApp message, hence the WhatsApp markup (`*bold*` titles, `_italic_` lines). Keep the file names, the empty-means-silent behavior and that format stable.
- **Error handling is nested try/except in `main()`**: calendar auth, birthdays, and events each fail independently and only print the error. The script always exits 0, so a failed run in CI looks successful and just leaves stale or empty files.
- **Birthdays** come from `birthdays.xlsx` (downloaded by `gdown`, git-ignored). The required columns are `birth_month`, `birth_date`, `birth_year`, `name`, `congrats_to`, `custom_message`. Message precedence is `custom_message`, then the "ALL" song text, then the "Gefeliciteerd ..." text. The age suffix is added when the birth year is known.
- **Calendar access uses an API key** (`developerKey`), not OAuth, so the calendar must be publicly readable. Event times are filtered in UTC, and all-day events (`start.date`) are formatted differently from timed ones (`start.dateTime`). Multi-day events are rendered as `start t/m end`.
- `print_calendar_events` and `print_birthday_messages` are unused debug helpers. Only the `write_*` functions produce the output files.

## Gotchas

- CI pins Python 3.8, so don't use newer syntax (e.g. `match`, `X | Y` type unions).
- `python-dateutil` is imported (`from dateutil import parser`) but not listed in `requirements.txt`. It currently arrives transitively via pandas.
