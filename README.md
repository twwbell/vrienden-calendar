# vrienden-calendar

Daily WhatsApp message for the friends' group: what's on the shared Google Calendar today, what was added or changed in the last 24 hours, and whose birthday it is. A Gemini agent writes the text.

```
n8n (local, 06:00)                        iPhone (your send time)
  Google Calendar ─┐                        Shortcut
  Google Sheet  ───┼─ Facts ─ Gemini ─▶ secret gist ─▶ date = today and message ≠ "" ─▶ WhatsApp
                   └─ nothing? → message ""
```

## Files

| Path | What |
|---|---|
| `n8n/workflows/daily-message.json` | `vrienden-calendar_Daily-Message`: the daily run |
| `n8n/workflows/error-mail.json` | `vrienden-calendar_Error-Mail`: mails you when the daily run fails |
| `n8n/prompts/system.md` | The agent's system prompt (tone, WhatsApp markup, birthday rules) |
| `n8n/config/example.json` | Template for `n8n/config/config.json` (git-ignored): ids, addresses, n8n credential ids |
| `n8n/scripts/push.mjs`, `pull.mjs` | Files → n8n, and n8n → files |
| `docs/iphone-shortcut.md` | How to build the Shortcut that sends the message |

## Commands

```bash
node n8n/scripts/push.mjs   # import both workflows; publish the daily one once config is complete
node n8n/scripts/pull.mjs   # after editing in the n8n editor, before editing the files
```

The scripts run the n8n CLI in the container `n8n-n8n-1` (override with `N8N_CONTAINER`).

## One-time setup

1. In n8n (http://localhost:5678), create these credentials and put their ids (from the credential URL) in `n8n/config/config.json` under `credentials`:
   - **Google Calendar OAuth2 API**, with the Google account that can see the friends' calendar. The calendar no longer needs to be public. The same OAuth client as the existing Google Sheets credential works.
   - **Google Sheets OAuth2 API**: the existing one, if that account can read the birthday sheet.
   - **Google Gemini (PaLM) API**: an API key from Google AI Studio.
   - **GitHub API**: a token that can write gists (classic `gist` scope or fine-grained *Gists: read and write*).
   - **SMTP**: for the error mail. Uses the existing SMTP account.
2. Create a secret gist with a file named like `gistFile` (any content) and put its id in `gistId`.
3. `node n8n/scripts/push.mjs`, open the daily workflow and run it by hand.
4. Build the Shortcut: `docs/iphone-shortcut.md`.

## Birthday sheet

The Google Sheet is read from the tab with gid `birthdaySheetGid`. The columns are `name`, `birth_month`, `birth_date`, `birth_year` (may be empty), `congrats_to` (`ALL` = a group member), and `custom_message` (used literally when filled).
