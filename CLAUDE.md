# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

This project posts a daily message in a friends' WhatsApp group when there is something noteworthy: events today, events added or changed in the last 24 hours, and birthdays. The pipeline:

1. The local n8n (Docker container `n8n-n8n-1`, http://localhost:5678) runs `vrienden-calendar_Daily-Message` at 06:00 Europe/Amsterdam. It reads the shared Google Calendar and the birthday Google Sheet with the owner's Google login, and a Gemini agent writes the message in Dutch.
2. It overwrites a file in a secret GitHub gist with `{"date": "YYYY-MM-DD", "message": "..."}`. When there is nothing to say, the message is `""`.
3. An iPhone Shortcut fetches the gist at the owner's chosen time. It sends `message` to the WhatsApp group only when `date` is today and `message` is not empty.

The laptop is not reachable from the internet, which is why the phone pulls from the gist. GitHub Actions, the Python script and the public txt files of the old version are gone, so don't reintroduce them.

## Commands

```bash
node n8n/scripts/push.mjs   # files -> n8n (import), publishes the daily workflow when config is complete
node n8n/scripts/pull.mjs   # n8n -> files, after a change in the n8n editor
```

There are no dependencies, tests or build: the scripts use only Node built-ins and the n8n CLI via `docker exec`.

## How the pieces fit

- **Workflow files hold no environment data.** `push.mjs` (`n8n/scripts/lib.mjs:forN8n`) injects three things:
  - the values of `n8n/config/config.json` (git-ignored, template `config/example.json`) into the node named `Config`;
  - `n8n/prompts/system.md` into the agent's system message;
  - credential ids per credential type from `config.credentials`.

  `pull.mjs` (`forFiles`) strips them out again, so a prompt edited in the n8n editor is lost on pull: edit `prompts/system.md`. When you add a config key or credential type, add it to `CONFIG_KEYS` / `CREDENTIAL_TYPES` in `lib.mjs` and to `example.json`.
- **Workflow and node names are referenced by name.** Expressions use `$('Config')`, `$('Today events')` and so on, and the `Facts` Code node reads the three fetch nodes by name. Renaming a node means updating those references.
- **Daily flow:** `Config` → `Today events` → `Changed events` → `Birthday sheet` → `Facts` → IF → agent (`Write message`, Gemini model) or `No message` → `Write gist`.
  - The fetch nodes run in sequence with `alwaysOutputData` and `executeOnce`. `Facts` drops the empty placeholder items (no `id`).
  - `Facts` turns all-day end dates (exclusive in Google) into the inclusive last day.
- **The empty message is a feature.** `Write gist` always writes, also `""`, so a stale message is never re-sent. On failure the gist keeps the old date (the Shortcut stays silent) and `vrienden-calendar_Error-Mail` (the daily workflow's `errorWorkflow`) mails the owner.
- **Message rules live in the prompt**, not in code: Dutch, WhatsApp markup (`*bold*`, `_italic_`), never invent facts, and the birthday rules (`custom_message` literally > `congrats_to = ALL` song > "Gefeliciteerd X met Y", plus the age when known).
- **Fixed workflow ids** (`vriendenCalDaily`, `vriendenCalError`), so import overwrites instead of duplicating. Import deactivates. `push.mjs` re-publishes the daily one with `n8n publish:workflow`, and only when config has no empty values.

The owner's iPhone side is documented in `docs/iphone-shortcut.md`.
