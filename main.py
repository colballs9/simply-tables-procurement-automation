#!/usr/bin/env python3
"""
Procurement Quote Generator — Cloud Run Service
Accepts POST: { "driveFolderUrl": "...", "quoteSheetUrl": "..." }
Finds PDF in Vietnam Quote subfolder, generates procurement tabs in quote sheet.
"""

import os
import re
import io
import json
import base64
import tempfile
import anthropic
import fitz
from flask import Flask, request, jsonify
from PIL import Image
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.auth import default as google_default
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload, MediaIoBaseDownload

app = Flask(__name__)

# ── CONFIG ────────────────────────────────────────────────────────────────────
MARKUP            = 2.0
TARIFF            = 0.20
TRUC_COMMISSION   = 0.07
TEMPLATE_SHEET_ID = "1ZJsT7jAv2t3sIzGY5dPTDXrgX2R6MIqfYlIgQgggOJU"
PE_TAB            = "Procurement Price Entry"
PQ_TAB            = "Procurement Quote"
VIETNAM_FOLDER    = "Vietnam Quote"
APPS_SCRIPT_URL   = os.environ.get("APPS_SCRIPT_URL", "")

# Protected rows in summary block — never written to
PROTECTED_PE_ROWS = {12, 13}  # Shipping Cost, Domestic Shipping

# Template structure constants (2-item template)
ITEMS_START_ROW   = 3   # First item row in Price Entry
ITEMS_START_ROW_Q = 14  # First item row in Quote tab
TEMPLATE_ITEM_ROWS = 2  # Template has 2 item rows (row 3 = first, row 4 = last)

# Summary block offsets relative to last item row
# In the 2-item template, last item = row 4, summary starts at row 6 (gap of 1)
SUMMARY_GAP       = 1   # blank row between last item and summary

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SCRIPT_DIR       = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "credentials.json")
TOKEN_FILE       = os.path.join(SCRIPT_DIR, "token.json")


# ── AUTH ──────────────────────────────────────────────────────────────────────
def get_services():
    """Get authenticated Sheets and Drive services.
    Uses Application Default Credentials on Cloud Run,
    falls back to OAuth credentials.json for local dev.
    """
    try:
        creds, _ = google_default(scopes=SCOPES)
    except Exception:
        creds = None
        if os.path.exists(TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(TOKEN_FILE, "w") as f:
                f.write(creds.to_json())

    sheets = build("sheets", "v4", credentials=creds)
    drive  = build("drive",  "v3", credentials=creds)
    return sheets, drive


# ── URL PARSING ───────────────────────────────────────────────────────────────
def extract_sheet_id(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not match:
        raise ValueError(f"Could not extract sheet ID from: {url}")
    return match.group(1)


def extract_folder_id(url: str) -> str:
    match = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    if not match:
        raise ValueError(f"Could not extract folder ID from: {url}")
    return match.group(1)


# ── FIND PDF IN DRIVE ─────────────────────────────────────────────────────────
def find_vietnam_folder_id(parent_folder_id: str, drive) -> str:
    """Return the ID of the Vietnam Quote subfolder."""
    result = drive.files().list(
        q=f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='{VIETNAM_FOLDER}' and trashed=false",
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    folders = result.get("files", [])
    if not folders:
        raise FileNotFoundError(f"No '{VIETNAM_FOLDER}' subfolder found")
    return folders[0]["id"]


def find_pdf_in_vietnam_folder(parent_folder_id: str, drive) -> tuple:
    """Find Vietnam Quote subfolder, then find PDF inside it.
    Returns (pdf_bytes, filename).
    """
    # Find Vietnam Quote subfolder
    result = drive.files().list(
        q=f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='{VIETNAM_FOLDER}' and trashed=false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()

    folders = result.get("files", [])
    if not folders:
        raise FileNotFoundError(f"No '{VIETNAM_FOLDER}' subfolder found in project Drive folder")

    viet_folder_id = folders[0]["id"]

    # Find PDF inside it
    pdf_result = drive.files().list(
        q=f"'{viet_folder_id}' in parents and mimeType='application/pdf' and trashed=false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()

    pdfs = pdf_result.get("files", [])
    if not pdfs:
        raise FileNotFoundError(f"No PDF found in '{VIETNAM_FOLDER}' folder")

    pdf_file = pdfs[0]
    print(f"  Found PDF: {pdf_file['name']}")

    # Download PDF bytes
    request_dl = drive.files().get_media(
        fileId=pdf_file["id"],
        supportsAllDrives=True
    )
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request_dl)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    return buf.getvalue(), pdf_file["name"]


# ── STEP 1: EXTRACT DATA FROM PDF ─────────────────────────────────────────────
def extract_pdf_data(pdf_bytes: bytes) -> dict:
    print("Sending to Claude for extraction...")
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")
    client  = anthropic.Anthropic()

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}
                },
                {
                    "type": "text",
                    "text": """Extract all line items from this manufacturer quote PDF.
Return ONLY a JSON object with this exact structure, no markdown, no explanation:
{
  "project": "project name",
  "vendor": "manufacturer/vendor name",
  "date": "quote date as MM/DD/YYYY",
  "items": [
    {
      "number": "item number e.g. 1 or C-UD-100 A",
      "name": "short item name",
      "width": 30,
      "depth": 28,
      "height": 30,
      "description": "full description text",
      "material": "material and finishing info",
      "qty": 8,
      "unit_cost": 260.00,
      "crating_cost": 0,
      "yards_per_unit": 0,
      "shipping_cost": 0,
      "page": 1,
      "image_count": 2
    }
  ],
  "lead_time": "lead time if stated",
  "origin": "country of origin",
  "payment_terms": "payment terms"
}

For "page": which page of the PDF this item appears on (1, 2, 3...).
For "image_count": count the actual product photographs in the Picture column for this row.
  - Count furniture photos AND fabric swatches
  - Do NOT count logos, text, or decorative elements
  - Most items have 1 image; some have 2 (e.g. chair + fabric swatch)

For crating_cost: per-item crating charge if listed, else 0.
For shipping_cost: any shipping charge listed per item, else 0.
For yards_per_unit: total yards divided by quantity. Use 0 if not listed.

DIMENSION UNITS — IMPORTANT:
All dimensions in the output must be in INCHES. Use this logic:
1. Check the actual values against typical furniture sizes:
   - Seat heights: 14–36 inches (350–900mm)
   - Table heights: 28–42 inches (700–1070mm)
   - Widths/depths: 12–120 inches (300–3000mm)
2. If the values are consistent with inches (e.g. seat height 18.5, width 62), keep them as-is.
3. If the values are only consistent with mm (e.g. seat height 470, width 1574), divide by 25.4.
4. IGNORE the unit label in the PDF header — manufacturers sometimes label columns "MM" but enter inch values. Trust the numbers, not the label."""
                }
            ]
        }]
    )

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())



# ── STEP 2: EXTRACT IMAGES FROM PDF ──────────────────────────────────────────
def extract_images(pdf_bytes: bytes, items: list) -> dict:
    """
    Simple approach — Claude already returned image_count and page per item
    in extract_pdf_data. PyMuPDF collects images per page in order.
    Python distributes them using item page + image_count fields.
    Returns {item_number: [(img_bytes, 'png'), ...]}
    """
    print("Extracting images from PDF...")
    doc   = fitz.open(stream=pdf_bytes, filetype="pdf")
    found = {item["number"].strip(): [] for item in items}

    # Collect all images per page with PyMuPDF
    images_by_page = {}
    for page_num in range(len(doc)):
        page     = doc[page_num]
        mat      = fitz.Matrix(2, 2)
        pix      = page.get_pixmap(matrix=mat)
        w, h     = pix.width, pix.height
        page_img = Image.frombytes("RGB", [w, h], pix.samples)
        blocks   = page.get_text("dict")["blocks"]

        img_blocks = []
        for b in blocks:
            if b.get("type") != 1:
                continue
            bw = b["bbox"][2] - b["bbox"][0]
            bh = b["bbox"][3] - b["bbox"][1]
            if bw < 15 or bh < 15:
                continue
            y_center = (b["bbox"][1] + b["bbox"][3]) / 2
            if page_num == 0 and y_center < page.rect.height * 0.10:
                continue  # skip logo on page 1
            img_blocks.append(b)

        img_blocks.sort(key=lambda b: b["bbox"][1])

        page_imgs = []
        for b in img_blocks:
            x0, y0, x1, y1 = [int(c * 2) for c in b["bbox"]]
            x0, y0 = max(0, x0), max(0, y0)
            x1, y1 = min(w, x1), min(h, y1)
            crop = page_img.crop((x0, y0, x1, y1))
            buf  = io.BytesIO()
            crop.save(buf, format="PNG")
            page_imgs.append(buf.getvalue())

        images_by_page[page_num] = page_imgs
        print(f"  Page {page_num+1}: {len(page_imgs)} images found")

    doc.close()

    # Distribute images using Claude's page + image_count per item
    page_idx = {}  # current position in each page's image pool

    for item in items:
        item_num = item["number"].strip()
        page_num = int(item.get("page", 1)) - 1  # 0-indexed
        count    = int(item.get("image_count", 1))
        pool     = images_by_page.get(page_num, [])
        idx      = page_idx.get(page_num, 0)

        for _ in range(count):
            if idx >= len(pool):
                print(f"  Warning: ran out of images on page {page_num+1} at item {item_num}")
                break
            found[item_num].append((pool[idx], "png"))
            idx += 1

        page_idx[page_num] = idx

    total = sum(len(v) for v in found.values())
    print(f"  Assigned {total} images across {sum(1 for v in found.values() if v)} items")
    return found



# ── STEP 3: UPLOAD IMAGES TO DRIVE ────────────────────────────────────────────
def upload_images(images_by_item: dict, project_name: str, drive, folder_id: str) -> dict:
    """Upload images into the Vietnam Quote folder in Shared Drive.
    Upscales small images to minimum 300px on longest side before uploading.
    Returns {item_number: [file_id, ...]}"""
    print("Uploading images...")
    file_ids = {}
    for item_num, img_list in images_by_item.items():
        file_ids[item_num] = []
        for i, (img_bytes, ext) in enumerate(img_list):
            # Upscale small images — PDF thumbnails are often only 25-50px
            img     = Image.open(io.BytesIO(img_bytes))
            w, h    = img.size
            longest = max(w, h)
            if longest < 300:
                scale = 300 / longest
                img   = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
                buf   = io.BytesIO()
                img.save(buf, format="PNG")
                img_bytes = buf.getvalue()
                ext = "png"

            media  = MediaInMemoryUpload(img_bytes, mimetype=f"image/{ext}")
            result = drive.files().create(
                body={
                    "name": f"{project_name}_{item_num.replace(' ','_')}_{i+1}.{ext}",
                    "parents": [folder_id]
                },
                media_body=media,
                fields="id",
                supportsAllDrives=True
            ).execute()
            fid = result.get("id")
            drive.permissions().create(
                fileId=fid,
                body={"type": "anyone", "role": "reader"},
                supportsAllDrives=True
            ).execute()
            file_ids[item_num].append(fid)
            # Small delay to avoid Drive API rate limits
            import time
            time.sleep(0.3)
    total = sum(len(v) for v in file_ids.values())
    print(f"  Uploaded {total} images")
    return file_ids


# ── STEP 4: COPY TEMPLATE TABS ────────────────────────────────────────────────
def copy_template_tabs(sheet_id: str, sheets) -> tuple:
    """Copy formatted template tabs into target sheet. Returns (pe_id, pq_id)."""
    print("Copying template tabs...")
    template = sheets.spreadsheets().get(spreadsheetId=TEMPLATE_SHEET_ID).execute()
    tmpl_ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in template["sheets"]}

    target  = sheets.spreadsheets().get(spreadsheetId=sheet_id).execute()
    deletes = [
        {"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}
        for s in target["sheets"]
        if s["properties"]["title"] in [PE_TAB, PQ_TAB]
    ]
    if deletes:
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id, body={"requests": deletes}
        ).execute()

    new_ids = {}
    for tab in [PE_TAB, PQ_TAB]:
        result = sheets.spreadsheets().sheets().copyTo(
            spreadsheetId=TEMPLATE_SHEET_ID,
            sheetId=tmpl_ids[tab],
            body={"destinationSpreadsheetId": sheet_id}
        ).execute()
        new_id = result["sheetId"]
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={"requests": [{"updateSheetProperties": {
                "properties": {"sheetId": new_id, "title": tab},
                "fields": "title"
            }}]}
        ).execute()
        new_ids[tab] = new_id
        print(f"  Copied: {tab}")

    return new_ids[PE_TAB], new_ids[PQ_TAB]


# ── STEP 5: WRITE DATA ────────────────────────────────────────────────────────
def write_data(sheet_id: str, pe_id: int, pq_id: int,
               data: dict, file_ids_by_item: dict, sheets, drive, viet_folder_id: str):
    """
    Write all item data and summary block with dynamic row count.
    Template has 2 item rows (3, 4). For N items we:
      - Use row 3 for item 1
      - Insert N-2 rows after row 3 for items 2 to N-1
      - Use the final row for item N
      - Write summary block starting at last_item_row + SUMMARY_GAP + 1
    """
    print("Writing data...")
    items = data["items"]
    n     = len(items)

    # ── Insert extra rows if N > 2 ──
    if n > 2:
        # Insert n-2 rows after row 3 (0-indexed row 3 = index 3)
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={"requests": [{
                "insertDimension": {
                    "range": {
                        "sheetId": pe_id,
                        "dimension": "ROWS",
                        "startIndex": 3,
                        "endIndex": 3 + (n - 2)
                    },
                    "inheritFromBefore": True
                }
            }]}
        ).execute()

        # Same for Quote tab — insert after row 14 (0-indexed 14)
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={"requests": [{
                "insertDimension": {
                    "range": {
                        "sheetId": pq_id,
                        "dimension": "ROWS",
                        "startIndex": 14,
                        "endIndex": 14 + (n - 2)
                    },
                    "inheritFromBefore": True
                }
            }]}
        ).execute()

        # Copy merged cells from template item row to all inserted rows
        # Template row 3 (index 2) has B:C merged — copy to rows 4 through N-1
        merge_requests = []
        for row_idx in range(3, 2 + n):  # rows 4 to last_item_row (0-indexed 3 to n+1)
            merge_requests.append({
                "mergeCells": {
                    "range": {
                        "sheetId": pe_id,
                        "startRowIndex": row_idx,
                        "endRowIndex": row_idx + 1,
                        "startColumnIndex": 1,  # column B
                        "endColumnIndex": 3,    # through C
                    },
                    "mergeType": "MERGE_ALL"
                }
            })
            merge_requests.append({
                "mergeCells": {
                    "range": {
                        "sheetId": pq_id,
                        "startRowIndex": row_idx + 11,  # quote tab items start at row 14 (index 13)
                        "endRowIndex": row_idx + 12,
                        "startColumnIndex": 1,
                        "endColumnIndex": 3,
                    },
                    "mergeType": "MERGE_ALL"
                }
            })
        if merge_requests:
            sheets.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body={"requests": merge_requests}
            ).execute()

    # ── Calculate dynamic row numbers ──
    last_pe_row  = ITEMS_START_ROW + n - 1          # e.g. 3+8-1 = 10 for 8 items
    last_pq_row  = ITEMS_START_ROW_Q + n - 1        # e.g. 14+8-1 = 21 for 8 items
    summary_start = last_pe_row + SUMMARY_GAP + 1   # e.g. 10+1+1 = 12 for 8 items

    # Summary row positions (dynamic)
    r_total_cost  = summary_start       # Total Cost (formula)
    r_tariff_pct  = summary_start + 1   # Tariff %
    r_tariff_cost = summary_start + 2   # Tariff Cost (formula)
    r_truc_pct    = summary_start + 3   # Truc Commission %
    r_truc_cost   = summary_start + 4   # Truc Commission (formula)
    r_ship_intl   = summary_start + 6   # Shipping Cost 🔒
    r_ship_dom    = summary_start + 7   # Domestic Shipping 🔒
    r_other1      = summary_start + 8   # Other Cost 1
    r_other2      = summary_start + 9   # Other Cost 2
    r_other3      = summary_start + 10  # Other Cost 3
    r_rep_pct     = summary_start + 12  # Rep Percent
    r_rep_cost    = summary_start + 13  # Rep Commission (formula)
    r_total_all   = summary_start + 15  # Total Cost all-in (formula)
    r_ship_cust   = summary_start + 17  # Shipping charge to customer
    r_total_item  = summary_start + 18  # Total Item Price (formula)
    r_total_sale  = summary_start + 20  # Total Sale (formula)
    r_margin      = summary_start + 23  # Margin (formula)

    # ── Collect all crating/shipping costs from items ──
    total_crating  = sum(item.get("crating_cost", 0) for item in items)
    total_shipping = sum(item.get("shipping_cost", 0) for item in items)
    # Any item-level shipping goes to Other Cost 1 (never rows 12/13)
    other_cost_1   = total_crating + total_shipping

    # ── Build Price Entry batch data ──
    batch = []

    # Item rows
    for i, item in enumerate(items):
        r = ITEMS_START_ROW + i
        batch.append({
            "range": f"'{PE_TAB}'!D{r}:R{r}",
            "values": [[
                f"{item['number']} — {item['name']}",
                item.get("width", 0),
                item.get("depth", 0),
                item.get("height", 0),
                item.get("description", ""),
                item.get("material", ""),
                "PE foam + 5-layer carton",
                item.get("qty", 1),
                item.get("unit_cost", 0),
                f"=K{r}*L{r}",
                MARKUP,
                f"=L{r}*N{r}",
                f"=K{r}*O{r}",
                item.get("yards_per_unit", 0),
                f"=Q{r}*K{r}",
            ]]
        })

    # Summary block — formulas reference dynamic row numbers
    item_range = f"M{ITEMS_START_ROW}:M{last_pe_row}"
    p_range    = f"P{ITEMS_START_ROW}:P{last_pe_row}"

    batch += [
        {"range": f"'{PE_TAB}'!M{r_total_cost}",  "values": [[f"=SUM({item_range})"]]},
        {"range": f"'{PE_TAB}'!M{r_tariff_pct}",  "values": [[TARIFF]]},
        {"range": f"'{PE_TAB}'!M{r_tariff_cost}", "values": [[f"=M{r_tariff_pct}*M${r_total_cost}"]]},
        {"range": f"'{PE_TAB}'!M{r_truc_pct}",    "values": [[TRUC_COMMISSION]]},
        {"range": f"'{PE_TAB}'!M{r_truc_cost}",   "values": [[f"=M{r_truc_pct}*M${r_total_cost}"]]},
        # Shipping rows 🔒 — never touched, left as template defaults
        {"range": f"'{PE_TAB}'!M{r_other1}",      "values": [[other_cost_1]]},
        {"range": f"'{PE_TAB}'!M{r_rep_pct}",     "values": [[0.08]]},
        {"range": f"'{PE_TAB}'!M{r_rep_cost}",    "values": [[f"=M{r_total_item}*M{r_rep_pct}"]]},
        {"range": f"'{PE_TAB}'!M{r_total_all}",   "values": [[
            f"=SUM(M{r_ship_intl}:M{r_other3},M{r_truc_cost},M{r_tariff_cost},M{r_total_cost},M{r_rep_cost})"
        ]]},
        {"range": f"'{PE_TAB}'!M{r_total_item}",  "values": [[f"=SUM({p_range})"]]},
        {"range": f"'{PE_TAB}'!M{r_total_sale}",  "values": [[f"=SUM(M{r_ship_cust}:M{r_total_item})"]]},
        {"range": f"'{PE_TAB}'!M{r_margin}",      "values": [[f"=M{r_total_sale}-M{r_total_all}"]]},
    ]

    # Quote tab — project header
    batch += [
        {"range": f"'{PQ_TAB}'!F2",  "values": [[data.get("project", "")]]},
        {"range": f"'{PQ_TAB}'!F7",  "values": [[data.get("date", "")]]},
    ]

    # Quote tab — item rows (pull from Price Entry)
    for i, item in enumerate(items):
        pe_r = ITEMS_START_ROW + i
        pq_r = ITEMS_START_ROW_Q + i
        batch.append({
            "range": f"'{PQ_TAB}'!D{pq_r}:O{pq_r}",
            "values": [[
                f"='{PE_TAB}'!D{pe_r}",
                f"='{PE_TAB}'!E{pe_r}",
                f"='{PE_TAB}'!F{pe_r}",
                f"='{PE_TAB}'!G{pe_r}",
                f"='{PE_TAB}'!H{pe_r}",
                f"='{PE_TAB}'!I{pe_r}",
                f"='{PE_TAB}'!J{pe_r}",
                f"='{PE_TAB}'!K{pe_r}",
                f"='{PE_TAB}'!O{pe_r}",
                f"=K{pq_r}*L{pq_r}",
                item.get("yards_per_unit", 0),
                f"=N{pq_r}*K{pq_r}",
            ]]
        })

    # Quote tab footer row numbers
    # Template has: row 16=Price Condition, row 17=Shipping, row 19=Total
    # After inserting N-2 rows after row 14, these shift to:
    # Price Condition = 16+(N-2) = 14+N
    # Shipping        = 17+(N-2) = 15+N
    # Total           = 19+(N-2) = 17+N
    pq_ship_row  = 15 + n
    pq_total_row = 17 + n

    batch += [
        {"range": f"'{PQ_TAB}'!M{pq_ship_row}",  "values": [[f"='{PE_TAB}'!M{r_ship_cust}"]]},
        {"range": f"'{PQ_TAB}'!M{pq_total_row}",  "values": [[
            f"=SUM(M{pq_ship_row},M{ITEMS_START_ROW_Q}:M{last_pq_row})"
        ]]},
    ]

    # Write all data
    sheets.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": batch}
    ).execute()

    # ── Insert images via =IMAGE() formula — instant, reliable ──
    if file_ids_by_item:
        img_requests = []
        image_metadata = []  # saved for Polish Images button

        for i, item in enumerate(items):
            item_num = item["number"].strip()
            fids     = file_ids_by_item.get(item_num, [])
            pe_r     = ITEMS_START_ROW + i
            pq_r     = ITEMS_START_ROW_Q + i
            if fids:
                fid     = fids[0]
                formula = f'=IMAGE("https://drive.google.com/uc?id={fid}")'
                for sheet_id_inner, row_idx in [(pe_id, pe_r - 1), (pq_id, pq_r - 1)]:
                    img_requests.append({
                        "updateCells": {
                            "range": {
                                "sheetId": sheet_id_inner,
                                "startRowIndex": row_idx,
                                "endRowIndex": row_idx + 1,
                                "startColumnIndex": 1,
                                "endColumnIndex": 2,
                            },
                            "rows": [{"values": [{"userEnteredValue": {"formulaValue": formula}}]}],
                            "fields": "userEnteredValue"
                        }
                    })
                # Store metadata for all images (including extras)
                for j, fid in enumerate(fids):
                    image_metadata.append({
                        "tabName": PE_TAB, "row": pe_r, "col": 2, "fileId": fid
                    })
                    image_metadata.append({
                        "tabName": PQ_TAB, "row": pq_r, "col": 2, "fileId": fid
                    })

        if img_requests:
            sheets.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body={"requests": img_requests}
            ).execute()
            print(f"  Inserted {len(img_requests)//2} images via =IMAGE() formula")

        # Save metadata JSON to Vietnam Quote folder for Polish Images button
        if image_metadata:
            try:
                meta_json = json.dumps({
                    "sheetId": sheet_id,
                    "images": image_metadata
                }, ensure_ascii=True).encode("utf-8")
                meta_media = MediaInMemoryUpload(meta_json, mimetype="application/json")
                drive.files().create(
                    body={"name": "_image_metadata.json", "parents": [viet_folder_id]},
                    media_body=meta_media,
                    fields="id",
                    supportsAllDrives=True
                ).execute()
                print(f"  Saved image metadata for Polish Images button")
            except Exception as e:
                print(f"  Could not save image metadata: {e}")

    print("  Done")


SECRET_TOKEN = os.environ.get("SECRET_TOKEN", "")


def check_auth(req):
    """Verify secret token in request header."""
    if not SECRET_TOKEN:
        return True  # no token configured, allow all (dev mode)
    token = req.headers.get("X-Secret-Token", "")
    return token == SECRET_TOKEN


# ── FLASK ROUTES ──────────────────────────────────────────────────────────────
@app.route("/", methods=["GET"])
def health():
    if not check_auth(request):
        return jsonify({"error": "Unauthorized"}), 403
    return jsonify({"status": "Procurement Quote Service running"})


@app.route("/build", methods=["POST"])
def build_quote():
    if not check_auth(request):
        return jsonify({"success": False, "error": "Unauthorized"}), 403
    try:
        payload         = request.get_json()
        drive_folder_url = payload.get("driveFolderUrl", "")
        quote_sheet_url  = payload.get("quoteSheetUrl", "")

        if not drive_folder_url or not quote_sheet_url:
            return jsonify({"success": False, "error": "Missing driveFolderUrl or quoteSheetUrl"}), 400

        folder_id = extract_folder_id(drive_folder_url)
        sheet_id  = extract_sheet_id(quote_sheet_url)

        print(f"Building procurement quote — sheet: {sheet_id}, folder: {folder_id}")

        sheets, drive = get_services()

        # Find and download PDF
        print("Finding PDF in Drive...")
        pdf_bytes, pdf_name = find_pdf_in_vietnam_folder(folder_id, drive)
        print(f"  Downloaded: {pdf_name}")

        # Extract data
        data = extract_pdf_data(pdf_bytes)
        print(f"  Found {len(data['items'])} items — {data.get('project', 'Unknown')}")

        # Extract images
        images_by_item = extract_images(pdf_bytes, data["items"])

        # Upload images into Vietnam Quote folder
        viet_folder_id = find_vietnam_folder_id(folder_id, drive)
        file_ids = upload_images(images_by_item, data.get("project", "project"), drive, viet_folder_id) if any(images_by_item.values()) else {}

        # Copy template tabs
        pe_id, pq_id = copy_template_tabs(sheet_id, sheets)

        # Write data
        write_data(sheet_id, pe_id, pq_id, data, file_ids, sheets, drive, viet_folder_id)

        return jsonify({
            "success": True,
            "project": data.get("project"),
            "items": len(data["items"]),
            "sheetUrl": f"https://docs.google.com/spreadsheets/d/{sheet_id}"
        })

    except FileNotFoundError as e:
        return jsonify({"success": False, "error": str(e)}), 404
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({"success": False, "error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)