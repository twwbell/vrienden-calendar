# iPhone Shortcut

n8n writes the day's message to a secret GitHub gist every morning at 06:00:

```json
{
  "date": "2026-11-18",
  "message": "*Vandaag in de vriendenagenda* ..."
}
```

`message` is the exact WhatsApp text, or `""` when there is nothing to say. When a run fails, the file keeps the previous date. The Shortcut sends only when `date` is today and `message` is not empty, so a failed or missed run never sends twice.

## Build the Shortcut

Personal Automation → **Time of Day** (your send time) → **Run Immediately**, with these actions:

1. **Get Contents of URL**
   - URL: `https://api.github.com/gists/<gistId>`
   - Headers: `Accept: application/vnd.github+json`. Optionally add `Authorization: Bearer <token>` with a fine-grained token that has only *Gists: read*.
2. **Get Dictionary Value** `files` → **Get Dictionary Value** `vrienden-calendar.json` (the `gistFile` from the config) → **Get Dictionary Value** `content`.
3. **Get Dictionary from Input** (parses the content).
4. **Format Date** (Current Date, custom format `yyyy-MM-dd`).
5. **If** dictionary value `date` *is* the formatted date
   - **If** dictionary value `message` *has any value*
     - **Send Message** via WhatsApp to the group with `message`. Alternatively, copy it to the clipboard and open the group chat, if Send Message can't reach the group.

## Note on privacy

A secret gist is unlisted, not private: anyone who knows its ID can read it. The ID is only in `n8n/config/config.json` (git-ignored) and in the Shortcut.
