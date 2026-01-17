# converter.py
from pathlib import Path
from datetime import datetime, date
import re
import pdfplumber
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# =====================
# Regex definitions
# =====================
DATE_RE = re.compile(r"^\d{4}/\d{1,2}/\d{1,2}$")
TXN_RE = re.compile(r"^\d{4}$")
ORDER_LONG_RE = re.compile(r"^\d{10,}(?:_\d+)?$")
NUM_RE = re.compile(r"^\d+(?:\.\d{1,2})?$")

TOTAL_USD_RE = re.compile(
    r"Total\s*\(\s*USD\s*\)\s*:\s*([$]?\s*[\d,]+(?:\.\d+)?)",
    re.IGNORECASE,
)
TOTAL_AMOUNT_RE = re.compile(
    r"TOTAL\s+AMOUNT\s*[$]?\s*([\d,]+(?:\.\d+)?)",
    re.IGNORECASE,
)

SINGLE_COUNTRIES = {
    "canada", "australia", "mexico", "china", "france", "germany",
    "italy", "spain", "japan", "singapore", "uae", "pakistan", "uk"
}

# =====================
# Helpers
# =====================
def clean(t: str) -> str:
    return (t or "").strip().strip("[]{}(),;:")

def parse_ymd_to_date(s: str) -> date:
    y, m, d = s.split("/")
    return date(int(y), int(m), int(d))

def excel_serial_from_date(dt: date) -> int:
    excel_epoch = date(1899, 12, 30)
    return (dt - excel_epoch).days

def to_dutch_text(num):
    if num is None:
        return ""
    return f"{float(num):.2f}".replace(".", ",")

def is_country_at(tokens, i):
    t = clean(tokens[i]).lower()

    if t == "united" and i + 1 < len(tokens):
        if clean(tokens[i + 1]).lower() == "states":
            return True, i + 1

    if t in SINGLE_COUNTRIES:
        return True, i

    return False, -1

def is_row_start(tokens, i):
    if i + 2 >= len(tokens):
        return False
    return (
        ORDER_LONG_RE.match(clean(tokens[i])) and
        TXN_RE.match(clean(tokens[i + 1])) and
        DATE_RE.match(clean(tokens[i + 2]))
    )

def qty_before_country(tokens, row_start, country_pos):
    for j in range(country_pos - 1, row_start, -1):
        t = clean(tokens[j])
        if t.isdigit():
            v = int(t)
            if 1 <= v <= 999:
                return v
    return None

def looks_like_price(tok: str):
    t = clean(tok)

    if TXN_RE.match(t):
        return False
    if not NUM_RE.match(t):
        return False
    if ORDER_LONG_RE.match(t):
        return False

    try:
        v = float(t)
    except ValueError:
        return False

    # avoid tiny integers like 1,2,3 being mistaken as price
    if "." not in t and v < 10:
        return False
    if v <= 0 or v > 100000:
        return False

    return True

def last_price_after_country(tokens, start, row_end):
    last = None
    for j in range(start, row_end):
        if looks_like_price(tokens[j]):
            last = float(clean(tokens[j]))
    return last

def stop_at_total_usd_token_index(tokens):
    # Stop tokens when totals section begins (saves parsing work)
    for i, tok in enumerate(tokens):
        low = clean(tok).lower()
        if low.startswith("total(usd") or low.startswith("total") or low.startswith("total_amount"):
            return i
    return len(tokens)

# =====================
# Memory-optimized extraction
# =====================
def extract_tokens_only(pdf_path: Path):
    """
    Extract tokens only (no full_text build).
    Using txt.split() is lighter than regex split.
    """
    tokens = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
            if txt:
                tokens.extend(txt.split())
    return tokens

def extract_total_usd_stream(pdf_path: Path):
    """
    Extract totals by scanning page text one page at a time.
    Avoids building a huge full_text string in memory.
    """
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
            if not txt:
                continue

            m = TOTAL_USD_RE.search(txt) or TOTAL_AMOUNT_RE.search(txt)
            if not m:
                continue

            raw = (
                m.group(1)
                .replace("$", "")
                .replace(" ", "")
                .replace(",", "")
            )
            try:
                return float(raw)
            except ValueError:
                return None
    return None

# =====================
# Core parsing logic (same idea, but no full_text)
# =====================
def parse_pdf_tokens(tokens, file_name):
    rows = []
    i = 0
    n = len(tokens)

    # Optional: limit parsing once totals section starts
    n = min(n, stop_at_total_usd_token_index(tokens))

    while i < n:
        if not is_row_start(tokens, i):
            i += 1
            continue

        txn_s = clean(tokens[i + 1])
        date_s = clean(tokens[i + 2])

        order4 = int(txn_s) if TXN_RE.match(txn_s) else None
        dt = parse_ymd_to_date(date_s) if DATE_RE.match(date_s) else None
        date_serial = excel_serial_from_date(dt) if dt else None

        # Find end of this row chunk
        row_end = min(n, i + 320)
        j = i + 3
        while j < row_end:
            if is_row_start(tokens, j):
                row_end = j
                break
            j += 1

        # Find country in this chunk
        country_pos = -1
        country_end = -1
        k = i + 3
        while k < row_end:
            ok, cend = is_country_at(tokens, k)
            if ok:
                country_pos = k
                country_end = cend
                break
            k += 1

        qty = None
        cost = None

        if country_pos != -1:
            qty = qty_before_country(tokens, i, country_pos)
            cost = last_price_after_country(tokens, country_end + 1, row_end)

        if order4 and date_serial and cost is not None:
            rows.append({
                "File Name": file_name,
                "DateSerial": date_serial,
                "Order #": order4,
                "Qty": int(qty) if qty else None,
                "Cost": float(cost),
                "Cost_NL": to_dutch_text(cost),
            })

        i = row_end

    return rows

# =====================
# Excel helpers (write_only safe)
# =====================
def set_column_widths(ws, widths):
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

# =====================
# Main converter (STREAMING EXCEL)
# =====================
def convert_pdfs_to_excel(pdf_paths, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_xlsx = output_dir / f"COGS_{ts}.xlsx"

    # Streaming workbook (very low memory)
    wb = Workbook(write_only=True)

    ws_cogs = wb.create_sheet("COGS")
    ws_tot = wb.create_sheet("InvoiceTotals")
    ws_log = wb.create_sheet("Log")

    # Headers
    ws_cogs.append(["File Name", "DateSerial", "Order #", "Qty", "Cost", "Cost_NL"])
    ws_tot.append(["File Name", "Total_USD_Extracted", "COGS_Sum", "Diff", "Match"])
    ws_log.append(["File", "Tokens", "Rows", "Status", "Error"])

    # Optional column widths (nice output)
    set_column_widths(ws_cogs, [38, 12, 10, 8, 12, 12])
    set_column_widths(ws_tot, [38, 18, 12, 12, 10])
    set_column_widths(ws_log, [38, 10, 10, 10, 60])

    extracted_any = False

    for pdf in pdf_paths:
        try:
            tokens = extract_tokens_only(pdf)
            rows = parse_pdf_tokens(tokens, pdf.name)

            cogs_sum = 0.0
            for r in rows:
                extracted_any = True
                cost = float(r["Cost"]) if r["Cost"] is not None else 0.0
                cogs_sum += cost
                ws_cogs.append([
                    r["File Name"],
                    r["DateSerial"],
                    r["Order #"],
                    r["Qty"],
                    cost,
                    r["Cost_NL"],
                ])

            total_usd = extract_total_usd_stream(pdf)
            diff = (cogs_sum - total_usd) if (total_usd is not None) else None
            match = "OK" if (diff is not None and abs(diff) < 0.01) else ("CHECK" if total_usd is not None else "")

            ws_tot.append([
                pdf.name,
                total_usd,
                cogs_sum,
                diff,
                match,
            ])

            ws_log.append([
                pdf.name,
                len(tokens),
                len(rows),
                "OK",
                "",
            ])

            # Free memory ASAP
            del tokens
            del rows

        except Exception as e:
            ws_tot.append([pdf.name, None, None, None, "ERROR"])
            ws_log.append([pdf.name, "", "", "ERROR", str(e)])

    if not extracted_any:
        # Still write file (helpful for debugging), but you can return None if you prefer
        print("[WARNING] No rows extracted from any PDF.")

    wb.save(out_xlsx)
    print(f"[INFO] Excel file created: {out_xlsx}")
    return out_xlsx
