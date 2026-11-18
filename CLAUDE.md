# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

This project posts a daily message in a friends' WhatsApp group when there is something noteworthy: events today, events added or changed in the last 24 hours, and birthdays. The pipeline:

1. The local n8n (Docker container `n8n-n8n-1`, http://localhost:5678) runs `vrienden-calendar_Daily-Message` at 06:00 Europe/Amsterdam. It reads the shared Google Calendar and the birthday Google Sheet with a Google service account (both are shared with its email), and a Gemini agent (API key) writes the message in Dutch.
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
- **Daily flow:** `Config` → `Today events` → `Changed events` → `Birthday sheet` → `Recent messages` → `Facts` → IF → agent (`Write message`, Gemini model) → `Message` (output check), or `No message` → `Write gist`.
  - **Memory:** `Recent messages` reads the gist. `Facts` passes the last 7 messages from the gist file `botolas-history.json` to the agent, so it doesn't repeat itself. `Write gist` appends today's message to that file, but only on days that have a message.
  - **Prompt injection:** calendar and sheet text is untrusted. `Facts` clips each field and strips `<>`, the agent gets it inside `<feiten>` tags, and the `Message` Code node throws (gist stays stale, error mail) when the message lacks the "Jullie maat Botolas" sign-off or exceeds 2500 characters.
  - `Today events` and `Changed events` are HTTP Request nodes that call the Calendar API directly, because n8n's Google Calendar node only supports OAuth, not a service account. The `googleApi` credential therefore needs "Set up for use in HTTP Request node" with the `calendar.readonly` scope. Each returns one item with the events in `items`.
  - The fetch nodes run in sequence with `executeOnce` (and `alwaysOutputData` on the sheet). `Facts` drops empty placeholders (no `id`).
  - `Facts` turns all-day end dates (exclusive in Google) into the inclusive last day.
- **The empty message is a feature.** `Write gist` always writes, also `""`, so a stale message is never re-sent. On failure the gist keeps the old date (the Shortcut stays silent) and `vrienden-calendar_Error-Mail` (the daily workflow's `errorWorkflow`) mails the owner.
- **Message rules live in the prompt**, not in code: the Botolas persona (opening line "🤖 Botolas zegt:", varying good-morning greeting, useless fact or joke, a bridge to the news, bold headings per category, sign-off "Jullie maat Botolas"), Dutch, WhatsApp markup (`*bold*`, `_italic_`), and never inventing events. Birthdays are written freely: `custom_message` is a hint, not text to copy.
- **Fixed workflow ids** (`vriendenCalDaily`, `vriendenCalError`), so import overwrites instead of duplicating. Import deactivates. `push.mjs` re-publishes the daily one with `n8n publish:workflow`, and only when config has no empty values.

The owner's iPhone side is documented in `docs/iphone-shortcut.md`.
