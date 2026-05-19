# 8D Corrective Action Converter

A small web tool that converts a Chinese Joysky 8D Word report into an English Corrective Action Word file using a built-in template.

## User workflow

1. Open the website URL
2. Enter the internal access password if configured
3. Enter `Customer No.`
4. Select `Lister` (`Grace Shih`, `Rita Lin`, or `Joy Lin`)
5. Upload the Chinese 8D `.docx`
6. Download the generated English Corrective Action `.docx`

Users do **not** upload the English template; it is built in at:

```text
templates_docx/corrective_action_template.docx
```

## Field rules

- `Date` = today
- `No.` = blank
- `Customer No.` = user input
- `Lister` = user dropdown choice
- `JS P/N` = Chinese `型號` outside parentheses
- `Part NO.` = Chinese `型號` parentheses
- `PO NO.` = `訂單號碼` second parentheses
- `SC NO.` = `訂單號碼` first parentheses with prefix stripped
- `DRAWING` = `JS P/N`
- `QT'Y` = Chinese `數量`
- `malformed Type` = Complain
- Analyst / Strategist / authority concern names = blank

## AI translation / rewrite

If `OPENAI_API_KEY` is set, the app uses an OpenAI-compatible Chat Completions API to rewrite the Chinese 8D sections into professional customer-facing English.

If AI is not configured or the API fails, the app falls back to the rule-based English generator so the tool still works.

Environment variables:

```bash
OPENAI_API_KEY=your_key_here
OPENAI_MODEL=gpt-4o-mini
OPENAI_BASE_URL=https://api.openai.com/v1
USE_AI=1
APP_PASSWORD=internal_password_optional
PORT=8787
```

Do not commit `.env` or API keys.

## Run locally

```bash
cd /Users/grace/8d-converter
OPENAI_API_KEY='your_key_here' APP_PASSWORD='your_password' python3 app.py --serve --port 8787
```

Open:

```text
http://127.0.0.1:8787
```

## Run on LAN / cloud

For any-computer access, run on a server/cloud platform and bind to `0.0.0.0`:

```bash
OPENAI_API_KEY='your_key_here' APP_PASSWORD='your_password' python3 app.py --serve --host 0.0.0.0 --port ${PORT:-8787}
```

## Docker

```bash
cd /Users/grace/8d-converter
docker build -t joysky-8d-converter .
docker run --rm -p 8787:8787 \
  -e OPENAI_API_KEY='your_key_here' \
  -e APP_PASSWORD='your_password' \
  joysky-8d-converter
```

Open:

```text
http://127.0.0.1:8787
```

## Render deployment

This folder includes `render.yaml`.

1. Push this folder to a GitHub repository.
2. Log in to Render.
3. Choose **New +** → **Blueprint**.
4. Connect the GitHub repository.
5. Render reads `render.yaml` automatically.
6. Set secret environment variables:
   - `OPENAI_API_KEY`
   - `APP_PASSWORD`
7. Deploy.
8. Share the generated Render URL with users.

## CLI test

```bash
python3 app.py \
  --input /path/to/chinese_8d.docx \
  --customer-no 001044 \
  --lister 'Grace Shih' \
  --output outputs/test.docx
```
