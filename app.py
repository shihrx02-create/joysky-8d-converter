#!/usr/bin/env python3
"""Small local 8D → English Corrective Action converter.

No third-party dependencies. It reads/writes .docx by editing Word XML.
Prototype notes:
- The English template is built in: templates_docx/corrective_action_template.docx
- Users upload only the Chinese 8D Word file, enter Customer No., and select Lister.
- Translation/polishing is rule-based for the current Joysky 8D format; later we can swap in an AI API.
"""

from __future__ import annotations

import cgi
import html
import json
import os
import re
import secrets
import shutil
import tempfile
import uuid
import urllib.error
import urllib.request
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, List, Tuple
from urllib.parse import parse_qs, urlparse
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates_docx" / "corrective_action_template.docx"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]
ET.register_namespace("w", NS["w"])
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")
ET.register_namespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")

LISTER_OPTIONS = ["Grace Shih", "Rita Lin", "Joy Lin"]
APP_PASSWORD = os.environ.get("APP_PASSWORD", "").strip()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini").strip()
USE_AI = os.environ.get("USE_AI", "1").strip().lower() not in {"0", "false", "no"}


def cell_text(el: ET.Element) -> str:
    parts: List[str] = []
    for node in el.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag == W + "tab":
            parts.append("\t")
        elif node.tag == W + "br":
            parts.append("\n")
    return "".join(parts).strip()


def set_cell_text(tc: ET.Element, text: str) -> None:
    """Replace cell contents while preserving cell properties."""
    tc_pr = tc.find("w:tcPr", NS)
    for child in list(tc):
        if child is not tc_pr:
            tc.remove(child)
    p = ET.SubElement(tc, W + "p")
    for pi, paragraph in enumerate(str(text).split("\n")):
        if pi:
            r_br = ET.SubElement(p, W + "r")
            ET.SubElement(r_br, W + "br")
        r = ET.SubElement(p, W + "r")
        t = ET.SubElement(r, W + "t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = paragraph


def read_docx_tables(path: Path) -> List[List[List[str]]]:
    with ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    tables = []
    for tbl in root.findall(".//w:tbl", NS):
        rows = []
        for tr in tbl.findall("w:tr", NS):
            rows.append([cell_text(tc).replace("\n", " | ") for tc in tr.findall("w:tc", NS)])
        tables.append(rows)
    return tables


def clean(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\t", " ")).strip()


def extract_between_label(rows: List[List[str]], starts_with: str) -> str:
    for row in rows:
        if row and row[0].startswith(starts_with):
            return row[1] if len(row) > 1 else ""
    return ""


def parse_chinese_8d(path: Path) -> Dict[str, str]:
    tables = read_docx_tables(path)
    if not tables:
        raise ValueError("找不到 Word 表格，請確認是中文 8D Word 檔。")
    rows = tables[0]

    # Header row: 訂單號碼 | (221-20240430003) (0041193) | 型號 | 3-... (DSG-...) | 數量 | 22,233 PCS
    order_raw = rows[0][1] if len(rows) > 0 and len(rows[0]) > 1 else ""
    model_raw = rows[0][3] if len(rows) > 0 and len(rows[0]) > 3 else ""
    qty_raw = rows[0][5] if len(rows) > 0 and len(rows[0]) > 5 else ""
    order_parens = re.findall(r"\(([^)]+)\)", order_raw)
    model_parens = re.findall(r"\(([^)]+)\)", model_raw)
    order_first = order_parens[0] if len(order_parens) >= 1 else ""
    po_no = order_parens[1] if len(order_parens) >= 2 else ""
    sc_no = order_first.split("-", 1)[1] if "-" in order_first else order_first
    js_pn = clean(re.sub(r"\([^)]*\)", "", model_raw))
    part_no = model_parens[0] if model_parens else ""
    qty = clean(qty_raw).replace("PCS", "").replace("pcs", "").strip()

    d1 = extract_between_label(rows, "D1")
    d3 = extract_between_label(rows, "D3")
    d4 = extract_between_label(rows, "D4")
    d5 = extract_between_label(rows, "D5")
    d6 = extract_between_label(rows, "D6")
    d7 = extract_between_label(rows, "D7")
    d8 = extract_between_label(rows, "D8")

    return {
        "po_no": po_no,
        "sc_no": sc_no,
        "js_pn": js_pn,
        "part_no": part_no,
        "drawing": js_pn,
        "qty": qty,
        "d1": clean(d1),
        "d3": clean(d3),
        "d4": clean(d4),
        "d5": clean(d5),
        "d6": clean(d6),
        "d7": clean(d7),
        "d8": clean(d8),
    }


def extract_defect_rate(d1: str) -> str:
    m = re.search(r"不良率[:：]\s*([0-9.]+\s*%)", d1)
    return m.group(1).replace(" ", "") if m else ""


def extract_spec_range(d1: str) -> str:
    m = re.search(r"([0-9.]+[”\"]?\s*[~－–-]\s*[0-9.]+[”\"]?)\s*\(([^)]+)\)", d1)
    if m:
        return f"{m.group(1).replace('~', '–')} ({m.group(2)})"
    return "the specified range"


def extract_counts(d1: str) -> Tuple[str, str]:
    total = ""
    bad = ""
    m_total = re.search(r"([0-9,]+)\s*pcs", d1, flags=re.I)
    if m_total:
        total = m_total.group(1)
    m_bad = re.search(r"大約有\s*([0-9,]+)\s*pcs", d1, flags=re.I)
    if m_bad:
        bad = m_bad.group(1)
    return total, bad


def english_sections(data: Dict[str, str]) -> Dict[str, str]:
    """Rule-based English rewrite for Joysky corrective-action wording.

    This is intentionally conservative for the prototype. It recognizes the common 8D text
    patterns Grace showed and produces formal English. Later we can replace this function
    with an AI call while keeping the same UI/template code.
    """
    d1, d3, d4, d5, d6, d7, d8 = (data.get(k, "") for k in ["d1", "d3", "d4", "d5", "d6", "d7", "d8"])
    defect_rate = extract_defect_rate(d1)
    spec_range = extract_spec_range(d1)
    total, bad = extract_counts(d1)

    if "開口" in d1 and ("過大" in d1 or "橢圓" in d1):
        count_phrase = ""
        if bad and total:
            count_phrase = f", approximately {bad} pieces out of {total} screws had oversized openings"
        elif bad:
            count_phrase = f", approximately {bad} pieces had oversized openings"
        content = (
            f"The customer reported that during assembly over the past two weeks{count_phrase}. "
            f"The opening dimension was significantly larger than the specified range of {spec_range}, allowing the screws to move freely in and out of the button. "
            "The customer also observed that the openings were oval-shaped rather than perfectly round, and considered this to be the root cause of the issue."
        )
        if defect_rate:
            content += f" The defect rate was approximately {defect_rate}."
    else:
        content = "The customer reported a nonconformity during assembly/inspection. The reported condition has been reviewed and summarized based on the submitted 8D information."

    if "夾持偏擺" in d3 or "定位點" in d3:
        analysis = (
            "The issue was determined to be caused by improper placement of the workpiece on the fixture locating points. "
            "This resulted in clamping misalignment during machining, causing the opening to become oversized and oval-shaped. "
            "The underlying cause was improper manual operation, which was related to insufficient operator experience and incomplete training."
        )
    else:
        analysis = "The root cause was reviewed based on the submitted 8D analysis. The issue was related to process control and operator execution, requiring corrective and preventive actions to reduce recurrence risk."

    if "防呆" in d6 or "限位" in d6 or "輔助治具" in d6:
        solution = (
            "An auxiliary fixture with positioning stops and mistake-proofing features will be designed to ensure that the workpiece can only be placed in a consistent orientation and position. "
            "This will reduce the risk of human positioning errors."
        )
    else:
        solution = "Permanent corrective actions will be implemented to strengthen process control and prevent recurrence."
    if "SOP" in d6 or "SOP" in d7:
        solution += " In addition, the machining SOP will be revised to include the allowable deformation tolerance for the inner diameter."

    confirm_parts = []
    if "定時抽樣" in d5 or "抽樣" in d5:
        confirm_parts.append("Periodic sampling will be implemented during production, especially after fixture replacement, repositioning, or batch changeover, to verify clamping stability.")
    if "偏擺" in d5:
        confirm_parts.append("Runout data will be recorded to monitor process consistency.")
    if "清潔" in d5 or "異物" in d5:
        confirm_parts.append("Before machining, the contact surfaces of the fixture, workpiece, and measuring tools will be thoroughly cleaned and inspected to prevent false runout caused by foreign particles.")
    if "三批" in d8:
        confirm_parts.append("The effectiveness of the corrective actions will be verified in the next production/shipment batch, and three consecutive batches will be continuously monitored to confirm that the same issue does not recur.")
    elif d8:
        confirm_parts.append("The effectiveness of the corrective actions will be verified in the next production/shipment batch and continuously monitored to confirm that the same issue does not recur.")
    confirm = " ".join(confirm_parts) or "The effectiveness of the corrective actions will be verified through follow-up production and inspection records."

    instruction_parts = []
    if "教育" in d4 or "訓練" in d4:
        instruction_parts.append("Immediate on-site training and communication will be conducted to reinforce the standard for part placement.")
    if "重新檢查" in d4:
        instruction_parts.append("All products on site will be re-inspected.")
    if "監督" in d4:
        instruction_parts.append("On-site supervision will be strengthened for all clamping and positioning steps.")
    if "教育" in d7 or "訓練" in d7:
        instruction_parts.append("Additional re-education and hands-on training will be provided to all operators and inspectors to ensure stable clamping and consistent part alignment.")
    instructions = " ".join(instruction_parts) or "Relevant operators and inspectors will receive training and on-site reminders to ensure consistent execution of the updated process requirements."

    return {
        "content": content,
        "analysis": analysis,
        "solution": solution,
        "confirm": confirm,
        "instructions": instructions,
    }


def ai_english_sections(data: Dict[str, str]) -> Dict[str, str]:
    """Use an OpenAI-compatible API to translate/rewrite the 8D fields.

    Required env vars:
    - OPENAI_API_KEY
    Optional:
    - OPENAI_MODEL, default gpt-4o-mini
    - OPENAI_BASE_URL, default https://api.openai.com/v1
    """
    if not USE_AI or not OPENAI_API_KEY:
        raise RuntimeError("AI is not configured; set OPENAI_API_KEY to enable AI rewriting.")

    prompt = {
        "task": "Translate and rewrite a Chinese Joysky 8D report into formal English corrective-action fields.",
        "rules": [
            "Return JSON only, no markdown.",
            "Keys must be exactly: content, analysis, solution, confirm, instructions.",
            "Use professional manufacturing/customer corrective-action English.",
            "Do not invent facts, dates, quantities, root causes, customer names, or commitments not present in the Chinese source.",
            "Keep the wording concise but complete enough for a customer-facing corrective action form.",
            "Content = D1 problem description.",
            "Analysis = D3 root cause analysis.",
            "Solution = D6 permanent corrective action + D7 recurrence prevention, without duplicating training details unless needed.",
            "Confirm = D5 corrective action verification + D8 closure/effectiveness tracking.",
            "Instructions = D4 temporary containment + education/training/on-site supervision from D7.",
        ],
        "source": {
            "D1_problem_description": data.get("d1", ""),
            "D3_root_cause_analysis": data.get("d3", ""),
            "D4_temporary_action": data.get("d4", ""),
            "D5_corrective_action_verification": data.get("d5", ""),
            "D6_permanent_corrective_action": data.get("d6", ""),
            "D7_prevent_recurrence": data.get("d7", ""),
            "D8_closure_confirmation": data.get("d8", ""),
        },
    }
    payload = {
        "model": OPENAI_MODEL,
        "messages": [
            {"role": "system", "content": "You are a manufacturing quality engineer writing customer-facing English corrective action reports."},
            {"role": "user", "content": json.dumps(prompt, ensure_ascii=False)},
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"},
    }
    req = urllib.request.Request(
        f"{OPENAI_BASE_URL}/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {OPENAI_API_KEY}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        raw = resp.read().decode("utf-8")
    result = json.loads(raw)
    content = result["choices"][0]["message"]["content"]
    sections = json.loads(content)
    required = ["content", "analysis", "solution", "confirm", "instructions"]
    missing = [k for k in required if not clean(str(sections.get(k, "")))]
    if missing:
        raise ValueError(f"AI response missing required fields: {', '.join(missing)}")
    return {k: clean(str(sections[k])) for k in required}


def build_english_sections(data: Dict[str, str]) -> Dict[str, str]:
    """Prefer AI translation/rewrite; fall back to deterministic rules if AI is unavailable."""
    try:
        sections = ai_english_sections(data)
        sections["translation_source"] = f"AI ({OPENAI_MODEL})"
        return sections
    except Exception as exc:
        sections = english_sections(data)
        sections["translation_source"] = f"Rule fallback ({exc})"
        return sections


def fill_template(chinese_docx: Path, customer_no: str, lister: str, out_path: Path) -> Dict[str, str]:
    data = parse_chinese_8d(chinese_docx)
    sections = build_english_sections(data)
    today = datetime.now().strftime("%Y/%m/%d")
    lister = lister if lister in LISTER_OPTIONS else LISTER_OPTIONS[0]

    with ZipFile(TEMPLATE_PATH, "r") as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        tbl = root.findall(".//w:tbl", NS)[0]
        rows = tbl.findall("w:tr", NS)

        # Row 2: No. blank + Date
        cells = rows[2].findall("w:tc", NS)
        if len(cells) >= 10:
            set_cell_text(cells[1], "")
            set_cell_text(cells[9], today)

        # Row 3 is a merged info cell in this template.
        info = f"Customer No.:   {customer_no}            JS P/N:  {data['js_pn']}           PO NO.:  {data['po_no']}           Part NO.: {data['part_no']}           SC NO.:  {data['sc_no']}"
        cells = rows[3].findall("w:tc", NS)
        if cells:
            set_cell_text(cells[0], info)

        # Row 4: DRAWING / QTY
        cells = rows[4].findall("w:tc", NS)
        if len(cells) >= 4:
            set_cell_text(cells[1], data["drawing"])
            set_cell_text(cells[3], data["qty"])

        # Main English fields.
        for row_idx, key in [(6, "content"), (8, "analysis"), (10, "solution"), (12, "confirm"), (14, "instructions")]:
            cells = rows[row_idx].findall("w:tc", NS)
            if len(cells) >= 2:
                set_cell_text(cells[1], sections[key])

        # Lister name goes under the lister label; other names remain blank.
        cells = rows[7].findall("w:tc", NS)
        if len(cells) >= 3:
            set_cell_text(cells[2], lister)
        for idx in [9, 11, 13]:
            cells = rows[idx].findall("w:tc", NS)
            if len(cells) >= 3:
                set_cell_text(cells[2], "")

        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with ZipFile(out_path, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data_bytes = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data_bytes = new_xml
                zout.writestr(item, data_bytes)

    return {**data, **sections, "date": today, "customer_no": customer_no, "lister": lister}


PAGE = """<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>8D Corrective Action Converter</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background:#f6f7fb; color:#1f2937; margin:0; }
    .wrap { max-width: 760px; margin: 48px auto; padding: 0 20px; }
    .card { background:white; border-radius:18px; box-shadow:0 12px 40px rgba(15,23,42,.09); padding:32px; }
    h1 { margin:0 0 8px; font-size:28px; }
    .sub { color:#6b7280; margin-bottom:28px; }
    label { display:block; font-weight:700; margin:18px 0 8px; }
    input, select { width:100%; box-sizing:border-box; border:1px solid #d1d5db; border-radius:12px; padding:13px 14px; font-size:16px; background:white; }
    button { margin-top:26px; width:100%; border:0; border-radius:14px; padding:15px 18px; background:#2563eb; color:white; font-size:17px; font-weight:800; cursor:pointer; }
    button:hover { background:#1d4ed8; }
    .hint { font-size:13px; color:#6b7280; margin-top:8px; }
    .rules { background:#f9fafb; border:1px solid #e5e7eb; border-radius:14px; padding:16px; margin-top:22px; font-size:14px; }
    .err { background:#fef2f2; color:#991b1b; padding:14px; border-radius:12px; margin-bottom:16px; }
  </style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <h1>8D Corrective Action Converter</h1>
    <div class="sub">上傳中文 8D，系統會用內建英文模板產出 Corrective Action Word。</div>
    {error}
    <form method="post" action="/convert" enctype="multipart/form-data">
      <label>Access Password</label>
      <input name="password" type="password" placeholder="公司內部密碼；如果未設定可留空">

      <label>Customer No.</label>
      <input name="customer_no" placeholder="例如：001044" required>

      <label>Lister</label>
      <select name="lister">
        <option>Grace Shih</option>
        <option>Rita Lin</option>
        <option>Joy Lin</option>
      </select>

      <label>Chinese 8D Report (.docx)</label>
      <input type="file" name="report" accept=".docx" required>
      <div class="hint">不需要上傳英文模板；模板已內建。</div>

      <button type="submit">Generate English Corrective Action</button>
    </form>
    <div class="rules">
      <b>AI 翻譯：</b>如果伺服器有設定 OPENAI_API_KEY，會用 AI 產生正式英文；如果沒有設定或 AI 失敗，會自動改用規則式備援，避免工具不能用。<br>
      <b>自動規則：</b>Date=今天、No.=空白、Type=Complain、DRAWING=JS P/N；PO NO/JS P/N/Part NO/SC NO/QTY 從中文 8D 自動抓。
    </div>
  </div>
</div>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_html(PAGE.replace("{error}", ""))
            return
        if parsed.path.startswith("/download/"):
            name = Path(parsed.path).name
            path = OUTPUT_DIR / name
            if not path.exists() or path.suffix.lower() != ".docx":
                self.send_error(HTTPStatus.NOT_FOUND, "File not found")
                return
            data = path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f'attachment; filename="{name}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        if self.path != "/convert":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        try:
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")})
            if APP_PASSWORD:
                supplied_password = str(form.getfirst("password", "")).strip()
                if not secrets.compare_digest(supplied_password, APP_PASSWORD):
                    raise ValueError("Access password 不正確。")
            customer_no = str(form.getfirst("customer_no", "")).strip()
            lister = str(form.getfirst("lister", LISTER_OPTIONS[0])).strip()
            item = form["report"] if "report" in form else None
            if not customer_no or item is None or not getattr(item, "filename", ""):
                raise ValueError("請填 Customer No. 並上傳中文 8D Word 檔。")
            safe_upload = UPLOAD_DIR / f"upload_{uuid.uuid4().hex}.docx"
            with safe_upload.open("wb") as f:
                shutil.copyfileobj(item.file, f)
            out_name = f"Corrective_Action_{customer_no}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            out_path = OUTPUT_DIR / out_name
            fill_template(safe_upload, customer_no, lister, out_path)
            self.send_response(303)
            self.send_header("Location", f"/download/{out_name}")
            self.end_headers()
        except Exception as e:
            self.send_html(PAGE.replace("{error}", f'<div class="err">{html.escape(str(e))}</div>'), status=400)

    def log_message(self, fmt: str, *args) -> None:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {fmt % args}")

    def send_html(self, body: str, status: int = 200) -> None:
        data = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> None:
    import argparse
    parser = argparse.ArgumentParser(description="8D Corrective Action Converter")
    parser.add_argument("--serve", action="store_true", help="start local web server")
    parser.add_argument("--host", default="127.0.0.1", help="host to bind, use 0.0.0.0 for cloud/LAN")
    parser.add_argument("--port", type=int, default=int(os.environ.get("PORT", "8787")))
    parser.add_argument("--input", type=Path, help="Chinese 8D .docx for CLI conversion")
    parser.add_argument("--customer-no", default="001044")
    parser.add_argument("--lister", default="Grace Shih", choices=LISTER_OPTIONS)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    if args.serve:
        server = ThreadingHTTPServer((args.host, args.port), Handler)
        print(f"Open http://{args.host}:{args.port}")
        print(f"AI mode: {'enabled' if (USE_AI and OPENAI_API_KEY) else 'fallback only'}")
        server.serve_forever()
    if not args.input:
        parser.error("use --serve or provide --input")
    out = args.output or OUTPUT_DIR / f"Corrective_Action_{args.customer_no}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    info = fill_template(args.input, args.customer_no, args.lister, out)
    print(out)
    print("Parsed:")
    for k in ["date", "customer_no", "lister", "js_pn", "po_no", "part_no", "sc_no", "drawing", "qty"]:
        print(f"  {k}: {info.get(k, '')}")


if __name__ == "__main__":
    main()
