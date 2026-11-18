# iPhone Shortcut

n8n writes the day's message to a secret GitHub gist every morning at 06:00. The gist holds two files:

- `vrienden-calendar.json` (the `gistFile` from the config) is the file the Shortcut reads:

  ```json
  {
    "date": "2026-11-18",
    "message": "🤖 Botolas zegt:\nMogge eikels! ..."
  }
  ```

- `botolas-history.json` is Botolas' memory of recent messages. The Shortcut ignores it.

`message` is the exact WhatsApp text, or `""` when there is nothing to say. When a run fails, the file keeps the previous date. The Shortcut sends only when `date` is today and `message` is not empty, so a failed or missed run never sends twice.

## Create the automation

1. Open **Shortcuts** → tab **Automation** → **+** (or **New Automation**).
2. Choose **Time of Day**.
   - Time: your send time, after 06:00 (for example 08:00). The laptop must have been on at 06:00 for n8n to run.
   - Repeat: **Daily**.
   - Select **Run Immediately**, and turn **Notify When Run** off.
3. Tap **Next** → **New Blank Automation**, and add the actions below in this exact order.

Shortcuts calls the output of every "Get Dictionary Value" action *Dictionary Value*, which makes them impossible to tell apart. That's why the steps store the important values in named variables with **Set Variable**. To pick a variable in a field, tap the field and choose it from the variable bar above the keyboard, or long-press the field → **Select Variable**.

## Actions

1. **Get Contents of URL**
   - URL: `https://api.github.com/gists/<gistId>`. Use the `gistId` from `n8n/config/config.json`; it's only the ID, not the whole gist URL.
   - Tap **Show More**:
     - Method: **GET**
     - Headers → **Add new header**: Key `Accept`, Value `application/vnd.github+json`.
     - Optional second header: Key `Authorization`, Value `Bearer <token>`, with a fine-grained GitHub token that has only *Gists: read*. Without it the request also works, but GitHub then limits unauthenticated requests to 60 per hour.

2. **Get Dictionary Value**
   - Get: **Value**
   - Key: `files`
   - Dictionary: **Contents of URL** (the output of action 1)

3. **Get Dictionary Value**
   - Get: **Value**
   - Key: `vrienden-calendar.json`
   - Dictionary: **Dictionary Value** (the output of action 2)
   - Use a separate action for this step: because the file name contains a dot, a key path like `files.vrienden-calendar.json` doesn't work.

4. **Get Dictionary Value**
   - Get: **Value**
   - Key: `content`
   - Dictionary: **Dictionary Value** (the output of action 3)
   - The output is the file content as text (JSON), not yet a dictionary.

5. **Get Dictionary from Input**
   - Input: **Dictionary Value** (the output of action 4)
   - This turns the JSON text into a dictionary with `date` and `message`.

6. **Set Variable**
   - Variable Name: `Gist`
   - Input: **Dictionary** (the output of action 5)

7. **Get Dictionary Value**
   - Get: **Value**
   - Key: `date`
   - Dictionary: variable **Gist**

8. **Set Variable**
   - Variable Name: `GistDatum`
   - Input: **Dictionary Value** (the output of action 7)

9. **Format Date**
   - Date: **Current Date** (tap the field → **Select Variable** → *Current Date*)
   - Date Format: **Custom**
   - Format String: `yyyy-MM-dd` (lowercase `yyyy` and `dd`, uppercase `MM`)
   - Set the Time Format to **None** if it's shown.

10. **Set Variable**
    - Variable Name: `Vandaag`
    - Input: **Formatted Date** (the output of action 9)

11. **Get Dictionary Value**
    - Get: **Value**
    - Key: `message`
    - Dictionary: variable **Gist**

12. **Set Variable**
    - Variable Name: `Bericht`
    - Input: **Dictionary Value** (the output of action 11)

13. **Count**
    - Count: **Characters**
    - In: variable **Bericht**
    - This guards against an empty message; "has any value" isn't reliable for empty text.

14. **If**
    - Tap **Add Condition** so that the If action has two conditions, and set it to **All** are true.
    - Condition 1: variable **GistDatum** · **is** · variable **Vandaag**
    - Condition 2: **Count** (the output of action 13) · **is greater than** · `0`
    - Shortcuts automatically adds **Otherwise** and **End If**. Leave Otherwise empty.

15. **Copy to Clipboard** (inside the If, before Otherwise)
    - Content: variable **Bericht**

16. **Send Message** (inside the If)
    - Message: variable **Bericht**
    - Recipients: the WhatsApp group. Tap the arrow (**>**) to choose **WhatsApp** as the app instead of Messages.
    - If WhatsApp doesn't offer the group as a recipient (groups often aren't supported), replace this action with **Open App** → App: **WhatsApp**. Then open the group yourself and paste the message, which action 15 put on the clipboard.

## Testing

- Tap the **▶︎** button in the shortcut editor to run it right away.
- If nothing happens, add a temporary **Show Result** or **Quick Look** after action 5 (input: **Dictionary**) to see what the gist returned.
- If there is no message for today, the If condition is false and nothing is sent, as intended. To test anyway, temporarily set condition 1 to **GistDatum is not empty**, or remove it.

## Note on privacy

A secret gist is unlisted, not private: anyone who knows its ID can read it. The ID is only in `n8n/config/config.json` (git-ignored) and in the Shortcut.
