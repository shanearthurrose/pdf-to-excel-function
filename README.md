# PDF-to-Excel Azure Function

Wraps the existing `pdf_extractor.py` logic (unchanged) in an HTTP-triggered
Azure Function so Power Automate can call it directly.

## Files
- `pdf_extractor.py` — your original extraction/Excel-writing logic, as an
  importable module (CLI `main()` kept for local testing only).
- `function_app.py` — the HTTP trigger. Takes raw PDF bytes in the request
  body, returns raw .xlsx bytes in the response.
- `requirements.txt`, `host.json`, `local.settings.json` — standard Azure
  Functions Python project files.

## 1. Test locally (optional)
```bash
pip install -r requirements.txt
func start
```
Then POST a PDF file's bytes to `http://localhost:7071/api/pdf-to-excel`.

## 2. Deploy
Same as your existing Azure Functions setup (GitHub + Deployment Center):
push this folder to the repo/branch connected to your Function App and let
Deployment Center build/deploy it — no changes needed there.

## 3. Get the callable URL + key
In the Azure Portal → your Function App → Functions → `pdf_to_excel` →
**Get Function URL**. It'll look like:
```
https://<your-app>.azurewebsites.net/api/pdf-to-excel?code=<function-key>
```
Keep this key private — anyone with it can call the function.

## 4. Wire it up in Power Automate

**Trigger:** "When a new email arrives" (Outlook), filtered to the
Anglicare mailbox/subject, with attachments included.

**Step — HTTP action** (call the function):
- Method: `POST`
- URI: the function URL from step 3
- Headers:
  - `Content-Type`: `application/pdf`
  - `X-File-Name`: `@{triggerBody()?['Attachments'][0]['Name']}`
- Body:
  `@{base64ToBinary(triggerBody()?['Attachments'][0]['ContentBytes'])}`

  (If you loop over multiple attachments with "For each", reference the
  current item instead of index `[0]`.)

**Step — Create file** (SharePoint or OneDrive action):
- File content: `@{body('HTTP')}` — the HTTP action's raw response body
  *is* the .xlsx file, no decoding needed.
- File name: something like
  `@{replace(triggerBody()?['Attachments'][0]['Name'], '.pdf', '.xlsx')}`

**Handling errors:** if extraction fails, the function returns a JSON body
`{"error": "..."}` with a 4xx/5xx status instead of xlsx bytes. Add a
"Configure run after" on the HTTP action (run after is successful/failed)
so you can branch — e.g. post the error message to a Teams channel or send
a notification email — instead of trying to save the error JSON as an xlsx.

## Notes
- The function has a 5-minute timeout (`host.json` → `functionTimeout`),
  which should be comfortable for typical Scope of Works PDFs. Bump it if
  you start feeding it much larger multi-hundred-page files.
- `auth_level=FUNCTION` means the function key is required on every call —
  matches how you'd normally lock down a function only Power Automate should
  reach.
