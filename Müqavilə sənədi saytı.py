#!/usr/bin/env python3
import base64
import html
import json
import os
import re
import ssl
import tempfile
import urllib.request
import zipfile
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs
import xml.etree.ElementTree as ET

APP_DIR = Path(__file__).resolve().parent

# Vercel-de fayl sistemi YALNIZ /tmp qovluguna yazila biler (read-only deployment bundle).
# tempfile.gettempdir() lokalda da, Vercel-de de isleyir.
OUTPUT_DIR = Path(tempfile.gettempdir()) / "generated"
UPLOAD_DIR = Path(tempfile.gettempdir()) / "uploads"

# Numune Word fayli GitHub repo-ya "templates" qovluguna qoyulmalidir (Path.home()/Downloads
# Vercel serverinde movcud deyil -- server senin lokal kompyuterine cixis ede bilmir).
TEMPLATE_DIR = APP_DIR / "templates"


def find_default_template():
    if TEMPLATE_DIR.exists():
        docs = sorted(TEMPLATE_DIR.glob("*.docx"))
        if docs:
            return docs[0]
    return None


DEFAULT_TEMPLATE = find_default_template()

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ET.register_namespace("w", NS["w"])

DEFAULT_REPLACEMENTS = {
    "Yusubov Elxan Yusif": "{person_short}",
    "Yusubov Elxan Yusif oğlu": "{person_full}",
    "Yusubov Elxan": "{person_signature}",
    "1807306302": "{voen}",
    "“Kapital bank” ASC Bakıxanov filialı": "{bank_name}",
    "Kod: 200219": "Kod: {bank_code}",
    "H/H:A081AIIB41040019447692493121": "H/H:{account}",
    "H/H: A081AIIB41040019447692493121": "H/H: {account}",
    "A081AIIB41040019447692493121": "{account}",
    "110 FC 581": "{vehicle_plate}",
    "Doosan DX210 WA": "{vehicle_model}",
    "Təkərli Eksakovator": "{vehicle_type}",
    "Təkərli ekskovator": "{vehicle_type_lower}",
    "20 23": "{vehicle_year_spaced}",
    "2023": "{vehicle_year}",
    "Narıncı": "{vehicle_color}",
    "7500": "{price}",
    "Yeddi min beş yüz manat  00 qəpik": "{price_words}",
    "27.04.2026-cı il imzalanmış XM № 05/26": "__CONTRACT_DATE_NUMERIC__ imzalanmış __XM_NUMBER__",
    "XM 05/26 №-li xidmət müqaviləsinə 27.04.2026-cı il tarixli": "__XM_NUMBER__ №-li xidmət müqaviləsinə __CONTRACT_DATE_NUMERIC__ tarixli",
    "«27» aprel 2026-cı il": "«__CONTRACT_DATE__»",
    "«26» aprel 2027-ci ilə qədər qüvvədədir.": "«__CONTRACT_END_DATE__» ilə qədər qüvvədədir.",
    "XM № 05/26": "__XM_NUMBER__",
    "XM 05/26": "__XM_NUMBER__",
}

DEFAULT_DATA = {
    "person_full": "ŞADMANOV ABBASƏLİ İMAMVERDİ OĞLU",
    "person_short": "Şadmanov Abbasəli İmamverdi",
    "person_signature": "Şadmanov Abbasəli",
    "voen": "1008729172",
    "bank_name": "KapitalBank ASC Xırdalan filialı",
    "bank_code": "201597",
    "bank_voen": "9900003611",
    "correspondent_account": "AZ37NABZ01350100000000001944",
    "swift": "AIIBAZ2XXX",
    "account": "AZ90AIIB410400H9449120681249",
    "address": "Abşeron rayonu, Xırdalan şəhəri, Heydər Əliyev prospekti 31-ci məhəllə, ev 3",
    "phone": "",
    "vehicle_plate": "AZ01BA186",
    "vehicle_model": "HYUNDAI R 210 LC 7A",
    "vehicle_type": "Ekskavator",
    "vehicle_type_lower": "ekskavator",
    "vehicle_year": "2010",
    "vehicle_year_spaced": "2010",
    "vehicle_color": "Sarı",
    "price": "7500",
    "price_words": "Yeddi min beş yüz manat 00 qəpik",
    "xm_number": "XM № 05/26",  # İstifadəçi daxil edir
    "contract_date": "27 aprel 2026-cı il",
    "contract_end_date": "26 aprel 2027-ci il",
    "contract_date_numeric": "27.04.2026-cı il",
}

FIELDS = [
    ("xm_number", "XM Nömrəsi"),
    ("contract_date", "Müqavilə tarixi (mətn: 27 aprel 2026-cı il)"),
    ("contract_end_date", "Qüvvədə olan tarix (mətn: 26 aprel 2027-ci il)"),
    ("contract_date_numeric", "Müqavilə tarixi (rəqəm: 27.04.2026-cı il)"),
    ("person_full", "İcraçı tam adı"),
    ("person_short", "İcraçı adı"),
    ("person_signature", "İmza adı"),
    ("voen", "VÖEN"),
    ("bank_name", "Bank"),
    ("bank_code", "Bank kodu"),
    ("swift", "SWIFT"),
    ("account", "Hesab"),
    ("vehicle_plate", "Dövlət qeydiyyat nişanı"),
    ("vehicle_model", "Marka / model"),
    ("vehicle_type", "Nəqliyyat tipi"),
    ("vehicle_type_lower", "İşin adı üçün tip"),
    ("vehicle_year", "Buraxılış ili"),
    ("vehicle_color", "Rəngi"),
    ("price", "Qiymət"),
    ("price_words", "Qiymət yazı ilə"),
]


def page(title, body):
    return f"""<!doctype html>
<html lang="az">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    :root {{
      --ink:#1f2933; --muted:#64748b; --line:#d6dde8; --soft:#f5f7fb;
      --accent:#0f766e; --accent-dark:#115e59; --danger:#b42318;
    }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif; color:var(--ink); background:#ffffff; }}
    header {{ border-bottom:1px solid var(--line); background:#fff; position:sticky; top:0; z-index:2; }}
    .bar {{ max-width:1180px; margin:0 auto; padding:18px 20px; display:flex; align-items:center; justify-content:space-between; gap:16px; }}
    h1 {{ font-size:22px; margin:0; font-weight:700; }}
    main {{ max-width:1180px; margin:0 auto; padding:22px 20px 40px; }}
    form {{ display:grid; grid-template-columns:280px 1fr; gap:20px; align-items:start; }}
    aside {{ border-right:1px solid var(--line); padding-right:20px; display:grid; gap:14px; }}
    section {{ min-width:0; }}
    .panel {{ border:1px solid var(--line); border-radius:8px; padding:16px; background:#fff; margin-bottom:16px; }}
    .panel h2 {{ font-size:15px; margin:0 0 12px; }}
    .grid {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }}
    label {{ display:grid; gap:6px; font-size:12px; color:var(--muted); font-weight:600; }}
    input {{ width:100%; min-height:38px; border:1px solid var(--line); border-radius:6px; padding:8px 10px; font-size:14px; color:var(--ink); background:#fff; }}
    input:focus {{ outline:2px solid rgba(15,118,110,.2); border-color:var(--accent); }}
    .full {{ grid-column:1 / -1; }}
    button, .button {{ min-height:40px; border:0; border-radius:6px; background:var(--accent); color:white; padding:0 14px; font-weight:700; cursor:pointer; text-decoration:none; display:inline-flex; align-items:center; justify-content:center; }}
    button:hover, .button:hover {{ background:var(--accent-dark); }}
    .secondary {{ background:#e8eef6; color:var(--ink); }}
    .secondary:hover {{ background:#dce5f1; }}
    .note {{ color:var(--muted); font-size:13px; line-height:1.45; margin:0; }}
    .success {{ border-color:#9bd3bd; background:#f0fbf6; }}
    .error {{ border-color:#f3b4ad; background:#fff4f2; color:var(--danger); }}
    .actions {{ display:flex; gap:10px; flex-wrap:wrap; align-items:center; }}
    @media (max-width: 820px) {{
      form {{ grid-template-columns:1fr; }}
      aside {{ border-right:0; border-bottom:1px solid var(--line); padding-right:0; padding-bottom:16px; }}
      .grid {{ grid-template-columns:1fr; }}
    }}
  </style>
</head>
<body>
  <header><div class="bar"><h1>Word Müqavilə Forması</h1></div></header>
  <main>{body}</main>
</body>
</html>"""


class UploadedFile:
    def __init__(self, filename, data):
        self.filename = filename
        self.data = data


class SimpleForm:
    def __init__(self):
        self.items = {}

    def add(self, key, value):
        if key in self.items:
            if isinstance(self.items[key], list):
                self.items[key].append(value)
            else:
                self.items[key] = [self.items[key], value]
        else:
            self.items[key] = value

    def getfirst(self, key, default=""):
        value = self.items.get(key, default)
        if isinstance(value, list):
            value = value[0] if value else default
        if isinstance(value, UploadedFile):
            return default
        return value

    def __contains__(self, key):
        return key in self.items

    def __getitem__(self, key):
        return self.items[key]


def text_value(form, key):
    value = form.getfirst(key, "")
    return value.strip() if isinstance(value, str) else ""


def fill_derived(data):
    data.setdefault("person_full", "")
    data.setdefault("person_short", data["person_full"])
    data.setdefault("person_signature", " ".join(data["person_short"].split()[:2]))
    data.setdefault("vehicle_type_lower", data.get("vehicle_type", "").lower())
    data.setdefault("vehicle_year_spaced", data.get("vehicle_year", ""))
    return data


def replace_in_text(text, data):
    new = text
    # First pass: replace template doc strings with placeholders
    for old, placeholder in sorted(DEFAULT_REPLACEMENTS.items(), key=lambda item: len(item[0]), reverse=True):
        new = new.replace(old, placeholder)
    # Second pass: replace placeholders with actual data values
    placeholder_map = {
        "__XM_NUMBER__": data.get("xm_number", ""),
        "__CONTRACT_DATE__": data.get("contract_date", ""),
        "__CONTRACT_END_DATE__": data.get("contract_end_date", ""),
        "__CONTRACT_DATE_NUMERIC__": data.get("contract_date_numeric", ""),
    }
    # Also replace {key} style placeholders (for other fields)
    for key, val in data.items():
        new = new.replace(f"{{{key}}}", val)
    for ph, val in placeholder_map.items():
        new = new.replace(ph, val)
    return re.sub(r"Swift: AIIBAZ2X(?!X)", f"Swift: {data.get('swift', '')}", new)


def update_docx(template, output, data):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        with zipfile.ZipFile(template) as zf:
            zf.extractall(tmp)
        xml_path = tmp / "word" / "document.xml"
        tree = ET.parse(xml_path)
        root = tree.getroot()
        containers = list(root.findall(".//w:p", NS)) + list(root.findall(".//w:txbxContent", NS))
        for container in containers:
            text_nodes = container.findall(".//w:t", NS)
            if not text_nodes:
                continue
            original = "".join(node.text or "" for node in text_nodes)
            replaced = replace_in_text(original, data)
            if replaced == original:
                continue
            text_nodes[0].text = replaced
            for node in text_nodes[1:]:
                node.text = ""
        tree.write(xml_path, xml_declaration=True, encoding="UTF-8")
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in sorted(tmp.rglob("*")):
                if file.is_file():
                    zf.write(file, file.relative_to(tmp).as_posix())


def image_to_data_url(path):
    mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
    return f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


def extract_with_openai(image_paths, api_key, model):
    prompt = """Extract Azerbaijani contract data from these document images.
Return only JSON with keys: person_full, person_short, person_signature, voen,
bank_name, bank_code, swift, account, vehicle_plate, vehicle_model,
vehicle_type, vehicle_type_lower, vehicle_year, vehicle_color, price, price_words.
Use empty string for missing values."""
    content = [{"type": "input_text", "text": prompt}]
    content += [{"type": "input_image", "image_url": image_to_data_url(p)} for p in image_paths]
    payload = {
        "model": model or "gpt-4.1-mini",
        "input": [{"role": "user", "content": content}],
        "text": {"format": {"type": "json_object"}},
    }
    req = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=180) as res:
            body = json.loads(res.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        if exc.code == 429:
            raise RuntimeError(
                "OpenAI API 429 qaytardı: limit və ya balans problemi var. "
                "Bir az sonra yenidən yoxla, API billing/balans aktivdir deyə platform.openai.com hesabında kontrol et, "
                "ya da məlumatları əl ilə doldurub Word yarat."
            ) from exc
        raise RuntimeError(f"OpenAI API xətası {exc.code}: {detail[:500]}") from exc
    except urllib.error.URLError as exc:
        if "CERTIFICATE_VERIFY_FAILED" not in str(exc):
            raise
        context = ssl._create_unverified_context()
        with urllib.request.urlopen(req, timeout=180, context=context) as res:
            body = json.loads(res.read().decode("utf-8"))
    text = body.get("output_text", "")
    if not text:
        parts = []
        for item in body.get("output", []):
            for c in item.get("content", []):
                if c.get("text"):
                    parts.append(c["text"])
        text = "\n".join(parts)
    return json.loads(text)


def form_html(data=None, message="", error=""):
    data = {**DEFAULT_DATA, **(data or {})}
    fields = []
    for key, label in FIELDS:
        fields.append(
            f'<label><span>{html.escape(label)}</span>'
            f'<input name="{key}" value="{html.escape(data.get(key, ""))}"></label>'
        )
    status = ""
    if message:
        status = f'<div class="panel success"><p class="note">{message}</p></div>'
    if error:
        status = f'<div class="panel error"><p>{html.escape(error)}</p></div>'
    template_note = "Tapıldı" if DEFAULT_TEMPLATE and DEFAULT_TEMPLATE.exists() else "Tapılmadı, aşağıdan nümunə Word seç"
    template_path_label = str(DEFAULT_TEMPLATE) if DEFAULT_TEMPLATE else "templates/ qovluğunda .docx tapılmadı"
    return page("Word Müqavilə Forması", f"""
{status}
<form method="post" enctype="multipart/form-data" action="/generate">
  <aside>
    <div class="panel">
      <h2>Nümunə Word</h2>
      <p class="note">Default: {html.escape(template_path_label)}<br>Status: {template_note}</p>
      <label><span>Başqa nümunə seç</span><input type="file" name="template" accept=".docx"></label>
    </div>
    <div class="panel">
      <h2>Şəkildən oxuma</h2>
      <p class="note">İstəsən şəkilləri və API açarını verib datanı avtomatik doldura bilərsən.</p>
      <label><span>OpenAI API açarı</span><input name="api_key" type="password" placeholder="sk-..."></label>
      <label><span>Şəkillər</span><input type="file" name="images" accept="image/*" multiple></label>
      <label><span>Model</span><input name="model" value="gpt-4.1-mini"></label>
      <button class="secondary" formaction="/extract" formmethod="post">Şəkillərdən doldur</button>
    </div>
    <div class="actions">
      <button type="submit">Word yarat</button>
    </div>
  </aside>
  <section>
    <div class="panel">
      <h2>Məlumatlar</h2>
      <div class="grid">{''.join(fields)}</div>
    </div>
  </section>
</form>
""")


class Handler(BaseHTTPRequestHandler):
    def send_html(self, content, status=200):
        body = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.startswith("/download/"):
            filename = Path(self.path.split("/download/", 1)[1]).name
            target = OUTPUT_DIR / filename
            if not target.exists():
                self.send_error(404)
                return
            data = target.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_html(form_html())

    def read_form(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length)
        content_type = self.headers.get("Content-Type", "")
        form = SimpleForm()
        if content_type.startswith("multipart/form-data"):
            message = BytesParser(policy=default).parsebytes(
                b"Content-Type: " + content_type.encode("utf-8") + b"\r\n\r\n" + raw
            )
            for part in message.iter_parts():
                name = part.get_param("name", header="content-disposition")
                if not name:
                    continue
                filename = part.get_filename()
                payload = part.get_payload(decode=True) or b""
                if filename:
                    form.add(name, UploadedFile(filename, payload))
                else:
                    charset = part.get_content_charset() or "utf-8"
                    form.add(name, payload.decode(charset, errors="replace"))
            return form
        for key, values in parse_qs(raw.decode("utf-8", errors="replace")).items():
            for value in values:
                form.add(key, value)
        return form

    def save_upload(self, item, fallback):
        if item is None or not getattr(item, "filename", ""):
            return fallback
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        safe = Path(item.filename).name
        target = UPLOAD_DIR / safe
        target.write_bytes(item.data)
        return target

    def collect_data(self, form):
        return fill_derived({key: text_value(form, key) for key, _ in FIELDS})

    def do_POST(self):
        form = self.read_form()
        if self.path == "/extract":
            data = self.collect_data(form)
            try:
                api_key = text_value(form, "api_key") or os.environ.get("OPENAI_API_KEY", "")
                if not api_key:
                    raise RuntimeError("API açarı yazılmayıb.")
                items = form["images"] if "images" in form else []
                if not isinstance(items, list):
                    items = [items]
                image_paths = [self.save_upload(item, None) for item in items if getattr(item, "filename", "")]
                if not image_paths:
                    raise RuntimeError("Şəkil seçilməyib.")
                extracted = extract_with_openai(image_paths, api_key, text_value(form, "model"))
                data.update({k: v for k, v in extracted.items() if v})
                self.send_html(form_html(fill_derived(data), "Şəkillərdən data oxundu. Yoxla, sonra Word yarat."))
            except Exception as exc:
                self.send_html(form_html(data, error=str(exc)))
            return

        if self.path == "/generate":
            data = self.collect_data(form)
            try:
                template = self.save_upload(form["template"] if "template" in form else None, DEFAULT_TEMPLATE)
                if not template or not template.exists():
                    raise RuntimeError(
                        "Nümunə Word faylı tapılmadı. Ya formadan .docx seç, ya da GitHub repo-da "
                        "'templates' qovluğuna nümunə Word faylını əlavə et."
                    )
                OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                output = OUTPUT_DIR / "hazir_muqavile.docx"
                update_docx(template, output, data)
                # Faylı birbaşa cavab kimi göndəririk (Vercel-də hər sorğu fərqli, müvəqqəti
                # konteynerdə işləyə bilər deyə diskdə saxlayıb ayrıca /download sorğusuna
                # güvənmək etibarsızdır).
                file_bytes = output.read_bytes()
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                self.send_header("Content-Disposition", 'attachment; filename="hazir_muqavile.docx"')
                self.send_header("Content-Length", str(len(file_bytes)))
                self.end_headers()
                self.wfile.write(file_bytes)
            except Exception as exc:
                self.send_html(form_html(data, error=str(exc)))
            return
        self.send_error(404)


app = Handler  # Vercel (legacy @vercel/python builder) bu dəyişəni axtarır


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    port = int(os.environ.get("PORT", 8765))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Sayt hazırdır: http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
