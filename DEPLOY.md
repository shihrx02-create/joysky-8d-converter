# Deploy 8D Converter to Render

This deploys the 8D Corrective Action Converter so it can be opened from any computer with a browser.

## What users will see

A webpage with:

- Access Password
- Customer No.
- Lister dropdown: Grace Shih / Rita Lin / Joy Lin
- Chinese 8D Report upload
- Generate button

The English template is built in. Users do not upload it.

## One-time setup

### 1. Put this folder in GitHub

Create a private GitHub repo, for example:

```text
joysky-8d-converter
```

Upload/push the contents of:

```text
/Users/grace/8d-converter
```

Important files:

```text
app.py
Dockerfile
render.yaml
requirements.txt
templates_docx/corrective_action_template.docx
```

Do not upload `.env` or API keys.

### 2. Create Render service

1. Go to Render.
2. Click **New +**.
3. Choose **Blueprint**.
4. Connect the GitHub repo.
5. Render will detect `render.yaml`.
6. Confirm creation.

### 3. Set environment variables in Render

In the Render service settings, set:

```text
OPENAI_API_KEY = your OpenAI API key
APP_PASSWORD = internal password for Grace/Rita/Joy
OPENAI_MODEL = gpt-4o-mini
OPENAI_BASE_URL = https://api.openai.com/v1
USE_AI = 1
```

`OPENAI_API_KEY` and `APP_PASSWORD` must be secret values.

### 4. Deploy

Click deploy. Render will build Docker and start the web service.

When finished, Render gives a URL like:

```text
https://joysky-8d-converter.onrender.com
```

Share that URL plus the internal password with users.

## Verify after deployment

1. Open the Render URL.
2. Enter the access password.
3. Enter Customer No., e.g. `001044`.
4. Select lister.
5. Upload a Chinese 8D `.docx`.
6. Download the generated English Word.
7. Check that these fields are filled:
   - Date = today
   - Customer No. = user input
   - JS P/N / PO NO. / Part NO. / SC NO. / QTY parsed from Chinese file
   - Lister = dropdown selection
   - Analyst / Strategist / authority concern names blank
   - Content / Analysis / Solution / Confirm / Instructions in English

## Notes

- The current app stores uploads and outputs in temporary local container folders. On Render free instances, files can disappear after restart. This is okay because users download immediately.
- If AI fails, the app uses rule-based fallback so it still returns a Word file.
- For stronger security later, add user login or deploy only inside the company network.
