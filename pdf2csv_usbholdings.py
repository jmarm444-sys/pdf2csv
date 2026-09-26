"""
Extract U.S. Bancorp holdings-statement PDFs (print-to-PDF / image pages) to CSV.

Pipeline: render each PDF page to an image (PyMuPDF), OCR it with the
built-in Windows OCR engine (winocr), cluster the recognized words into
table rows/columns by position, crop-retry suspicious cells, repair values
with price*quantity=market value and market value-cost basis=gain/loss,
and write one CSV row per holding.

The statement table has eight data columns. The CSV keeps three:
    sym, qty, per-share
where per-share is cost basis divided by quantity (half-up to the cent).
Example: AAPL, quantity 3, cost basis 945.39 -> aapl,3,315.13

Requirements (Windows 10/11):
    pip install pymupdf winocr pillow

Run with no arguments for the Browse/Convert window.
Run with two arguments for the command line:
    python pdf2csv_usbholdings.py input.pdf output.csv
"""

import csv
import os
import re
import sys
import threading
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

# Configure TCL/TK environment when frozen as an executable
if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
    for tcl_sub in ["tcl/tcl8.6", "tcl8.6", "tcl"]:
        cand = os.path.join(base_dir, tcl_sub)
        if os.path.exists(cand):
            os.environ["TCL_LIBRARY"] = cand
            break
    for tk_sub in ["tcl/tk8.6", "tk8.6", "tk"]:
        cand = os.path.join(base_dir, tk_sub)
        if os.path.exists(cand):
            os.environ["TK_LIBRARY"] = cand
            break

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz  # PyMuPDF
import winocr
from PIL import Image

# Pixel thresholds below are tuned for this DPI.
DPI = 200

# Column bins as fractions of page width (left edge, right edge).
# Measured from the current U.S. Bank Investments > Holdings print layout:
# Symbol | Price | Price change | Quantity | Market value | Day change | Cost basis | Gain/loss
COLUMN_BINS = {
    "symbol":       (0.000, 0.150),
    "price":        (0.150, 0.258),
    "price_change": (0.258, 0.395),
    "quantity":     (0.395, 0.515),
    "market_value": (0.515, 0.630),
    "day_change":   (0.630, 0.730),
    "cost_basis":   (0.730, 0.845),
    "gain_loss":    (0.845, 0.955),
}

# Same-line tokens sit within this many pixels of the row anchor.
# The company name and the percent line begin further down (~40px and ~50px).
ROW_BAND = 22
# Distinct holdings are much farther apart than a split OCR token.
ROW_CLUSTER = 40

STOP_WORDS = {"Disclosures", "Insurance", "FINRA"}
MONEY_COLS = ("price", "market_value", "cost_basis", "gain_loss")
ANCHOR_COLS = ("market_value", "cost_basis")

CSV_HEADER = ["sym", "qty", "per-share"]

CENT = Decimal("0.01")
QTY_STEP = Decimal("0.0001")


def D(value):
    """Parse a cleaned number string to Decimal, or None."""
    if value is None or value == "":
        return None
    try:
        return Decimal(value)
    except (InvalidOperation, ValueError):
        return None


def money_str(value):
    """Format a Decimal as a signed money string with 2 decimal places."""
    q = value.quantize(CENT, rounding=ROUND_HALF_UP)
    return f"{q:.2f}"


def qty_str(value):
    """Format a share count: integers without a decimal, otherwise up to 4 places."""
    q = value.quantize(QTY_STEP, rounding=ROUND_HALF_UP)
    if q == q.to_integral_value():
        return str(q.to_integral_value())
    text = f"{q:.4f}".rstrip("0").rstrip(".")
    return text


def per_share_str(cost, qty):
    """Cost basis per share, half-up to the cent."""
    value = (cost / qty).quantize(CENT, rounding=ROUND_HALF_UP)
    return f"{value:.2f}"


def product_close(price, qty, market):
    """True when price * quantity matches market value within rounding slack."""
    if price is None or qty is None or market is None or qty == 0:
        return False
    tol = max(Decimal("0.08"), Decimal("0.006") * abs(qty))
    return abs(price * qty - market) <= tol


def values_close(a, b):
    """Money identity on this statement is exact to the cent."""
    if a is None or b is None:
        return False
    return abs(a - b) <= Decimal("0.02")


def is_blank_cell(text):
    """True when the cell is empty or an intentional dash, not a misread number."""
    if not text:
        return True
    t = text.strip().replace(" ", "")
    return t in {"", "-", "–", "—", "−", "--"}


def ocr_words(img):
    """OCR a PIL image, return [(y, x1, x2, text)] sorted by position."""
    result = winocr.recognize_pil_sync(img, "en-US")
    words = []
    for line in result["lines"]:
        for w in line["words"]:
            br = w["bounding_rect"]
            words.append((br["y"], br["x"], br["x"] + br["width"], w["text"]))
    return sorted(words)


def ocr_crop_texts(img, x1, y1, x2, y2):
    """Yield raw OCR text of one cell at several zoom levels."""
    x1, y1 = max(0, int(x1)), max(0, int(y1))
    x2, y2 = min(img.width, int(x2)), min(img.height, int(y2))
    if x2 - x1 < 4 or y2 - y1 < 4:
        return
    for scale in (2, 3, 1):
        crop = img.crop((x1, y1, x2, y2)).resize(
            ((x2 - x1) * scale, (y2 - y1) * scale), Image.LANCZOS)
        text = "".join(t for _, _, _, t in ocr_words(crop))
        if text:
            yield text


def clean_number(text):
    """Normalize an OCR'd numeric cell to a plain number string ('' if blank)."""
    if not text:
        return ""
    t = text.replace(" ", "")
    t = t.replace("O", "0").replace("o", "0")
    t = t.replace("I", "1").replace("l", "1").replace("S", "5")
    neg = t.startswith("-") or t.startswith("(") or ("(" in t and ")" in t)
    t = t.replace(",", "").replace("$", "").replace("+", "")
    t = re.sub(r"[^0-9.]", "", t)
    if not t or t == ".":
        return ""
    if t.count(".") > 1:
        head, _, tail = t.rpartition(".")
        t = head.replace(".", "") + "." + tail
    if t.startswith("."):
        t = "0" + t
    if t.endswith("."):
        t = t[:-1]
    if not t:
        return ""
    return ("-" if neg else "") + t


def clean_symbol(text):
    t = text.replace("'", "").replace(" ", "")
    t = re.sub(r"[^A-Za-z0-9.]", "", t)
    t = t.strip(".")
    return t.lower()


def plausible_symbol(text):
    return bool(text) and bool(re.fullmatch(r"[a-z][a-z0-9.]{0,5}", text))


def trim_to_ink(crop):
    """Crop away surrounding whitespace; None if the region is blank."""
    gray = crop.convert("L").point(lambda p: 255 if p < 160 else 0)
    bbox = gray.getbbox()
    if bbox is None:
        return None
    return crop.crop(bbox)


def best_decimal_match(raw, target):
    """Re-place the decimal in an OCR digit string so it lands nearest target.

    Returns (error, number_string, Decimal) or None when raw has no digits.
    """
    if not raw or target is None:
        return None
    neg = "-" in raw or "(" in raw
    prepared = raw.replace("O", "0").replace("o", "0")
    digits = re.sub(r"\D", "", prepared)
    if not digits:
        return None
    target_abs = abs(target)
    best = None
    limit = min(len(digits), 6)
    for dec in range(0, limit + 1):
        if dec == 0:
            body = digits
        elif dec >= len(digits):
            body = "0." + "0" * (dec - len(digits)) + digits
        else:
            body = digits[:-dec] + "." + digits[-dec:]
        val = Decimal(body)
        err = abs(val - target_abs)
        if best is None or err < best[0]:
            shown = ("-" if neg else "") + body
            best = (err, shown, -val if neg else val)
    return best


def ocr_symbol_with_context(img, width, y0):
    """Fallback for tickers the engine drops (1-2 letter words).

    Windows OCR silently discards short isolated words, but keeps them when
    they closely follow a confidently-recognized token. OCR the symbol cell
    with the row's market-value cell pasted in front of it, then keep the
    letter token.
    """
    lo, hi = COLUMN_BINS["market_value"]
    num = trim_to_ink(img.crop((
        int(lo * width), int(y0 - 10), int(hi * width), int(y0 + 34))))
    sym = trim_to_ink(img.crop((
        10, int(y0 - 8), int(0.145 * width), int(y0 + 30))))
    if sym is None:
        return ""
    if num is None:
        big = sym.resize((sym.width * 3, sym.height * 3), Image.LANCZOS)
        tokens = [t for _, _, _, t in ocr_words(big)]
    else:
        gap, margin = 35, 30
        h = max(num.height, sym.height)
        canvas = Image.new(
            "RGB",
            (num.width + sym.width + gap + 2 * margin, h + 2 * margin),
            "white",
        )
        canvas.paste(num, (margin, margin + (h - num.height) // 2))
        canvas.paste(sym, (margin + num.width + gap, margin + (h - sym.height) // 2))
        big = canvas.resize((canvas.width * 2, canvas.height * 2), Image.LANCZOS)
        tokens = [t for _, _, _, t in ocr_words(big)]

    letters = []
    for t in tokens:
        c = clean_symbol(t)
        if plausible_symbol(c):
            letters.append(c)
    # The market-value crop contributes no ticker; keep the last letter token,
    # which is the symbol pasted on the right.
    return letters[-1] if letters else ""


def col_of(x_center, width):
    f = x_center / width
    for name, (lo, hi) in COLUMN_BINS.items():
        if lo <= f < hi:
            return name
    return None


def bin_px(col, width):
    lo, hi = COLUMN_BINS[col]
    return lo * width, hi * width


def has_digit(text):
    return any(ch.isdigit() for ch in text)


def read_number_cell(img, width, y0, col, raw, money):
    """Return a cleaned number string, crop-retrying a doubtful full-page read."""
    if is_blank_cell(raw):
        return ""
    val = clean_number(raw) if raw else ""
    suspicious = (not val) or (money and "$" not in raw)
    if not suspicious:
        return val
    reread = reread_number(img, width, y0, col, money)
    return reread or val


def reread_number(img, width, y0, col, money):
    """OCR just this cell. Money cells must include a dollar sign."""
    lo, hi = bin_px(col, width)
    fallback = ""
    for cand in ocr_crop_texts(img, lo, y0 - 10, hi, y0 + 36):
        v = clean_number(cand)
        if not v:
            continue
        if money and "$" in cand:
            return v
        if not money:
            return v
        fallback = fallback or v
    return "" if money else fallback


def infer_qty(market, price):
    """Pick a share count that makes price * qty match market value."""
    raw = market / price
    for places in (0, 1, 2, 3, 4, 5):
        q = raw.quantize(Decimal(10) ** -places, rounding=ROUND_HALF_UP)
        if q != 0 and product_close(price, q, market):
            return q
    return raw.quantize(QTY_STEP, rounding=ROUND_HALF_UP)


def repair_price_qty_mv(price, qty, market, price_raw, qty_raw, mv_raw):
    """Make price * quantity = market value when a dropped decimal explains the miss."""
    if product_close(price, qty, market):
        return price, qty, market

    if qty is not None and market is not None and qty != 0 and price_raw:
        target = market / qty
        match = best_decimal_match(price_raw, target)
        if match and product_close(match[2], qty, market):
            price = match[2]

    if not product_close(price, qty, market) and price and market and price != 0 and qty_raw:
        target = market / price
        match = best_decimal_match(qty_raw, target)
        if match and product_close(price, match[2], market):
            qty = match[2]

    if not product_close(price, qty, market) and price and qty and mv_raw:
        match = best_decimal_match(mv_raw, price * qty)
        if match and product_close(price, qty, match[2]):
            market = match[2]

    if qty is None and price and market and price != 0:
        qty = infer_qty(market, price)
    if price is None and qty and market and qty != 0:
        price = (market / qty).quantize(CENT, rounding=ROUND_HALF_UP)
    if market is None and price and qty:
        market = (price * qty).quantize(CENT, rounding=ROUND_HALF_UP)
    return price, qty, market


# Single-digit swaps Windows OCR commonly makes inside a money amount.
DIGIT_SWAPS = {
    "0": "6",
    "1": "7",
    "3": "5",
    "4": "9",
    "5": "3",
    "6": "8",
    "7": "1",
    "8": "6",
    "9": "4",
}


def digit_variants(value):
    """Yield (edit_count, Decimal) for a money amount and one confusable-digit flip."""
    if value is None:
        return
    text = money_str(value)
    yield 0, value
    chars = list(text)
    for i, ch in enumerate(chars):
        swap = DIGIT_SWAPS.get(ch)
        if not swap:
            continue
        trial = chars[:]
        trial[i] = swap
        body = "".join(trial)
        if body.endswith("."):
            continue
        yield 1, Decimal(body)


def repair_cost_basis(cost, gain, market, cost_raw, reread_cost, reread_gain):
    """Return (cost Decimal or None, warning or '').

    Identity: market value - cost basis = gain/loss, exact to the cent.
    A closer crop of the cell is preferred when the full-page read breaks
    that identity. A one-digit confusion (3/5, 8/6, ...) is tried next.
    """
    if market is None or gain is None:
        return cost, ""

    def fits(c, g):
        return c is not None and g is not None and values_close(market - c, g)

    if fits(cost, gain) or fits(cost, -gain):
        return cost, ""

    # A tight crop often fixes one digit the full-page pass misread.
    if fits(reread_cost, gain) or fits(reread_cost, -gain):
        return reread_cost, "repaired cost basis from a closer read of the cell"
    if fits(cost, reread_gain):
        return cost, ""
    if fits(reread_cost, reread_gain):
        return reread_cost, "repaired cost basis from a closer read of the cell"

    best = None
    cost_choices = list(digit_variants(cost))
    if reread_cost is not None:
        cost_choices.extend((1, v) for _, v in digit_variants(reread_cost))
    gain_choices = list(digit_variants(gain))
    if reread_gain is not None:
        gain_choices.extend((1, v) for _, v in digit_variants(reread_gain))
    gain_choices.append((1, -gain if gain is not None else None))
    for cost_edits, cost_try in cost_choices:
        for gain_edits, gain_try in gain_choices:
            if gain_try is None or not fits(cost_try, gain_try):
                continue
            score = cost_edits + gain_edits
            if best is None or score < best[0]:
                best = (score, cost_try)
    if best is not None:
        chosen = best[1]
        if cost is not None and values_close(chosen, cost):
            return cost, ""
        return chosen, "repaired cost basis digit from market value - gain/loss"

    target = market - gain
    match = best_decimal_match(cost_raw, target) if cost_raw else None
    if match and values_close(match[2], target):
        return match[2], "repaired cost basis decimal from gain/loss"
    if cost is None:
        return target.quantize(CENT, rounding=ROUND_HALF_UP), "filled cost basis from market value - gain/loss"
    return cost, "cost basis does not match market value - gain/loss"


def parse_page(img, pageno, records, warnings):
    width = img.width
    words = ocr_words(img)
    if not words:
        return

    min_y = 0
    for y, _x1, _x2, t in words:
        if t == "Symbol":
            min_y = max(min_y, y + 24)

    max_y = float("inf")
    for y, _x1, _x2, t in words:
        if t in STOP_WORDS and y > min_y:
            max_y = min(max_y, y - 12)

    words = [w for w in words if min_y < w[0] < max_y]

    anchors = []
    for y, x1, x2, t in words:
        col = col_of((x1 + x2) / 2, width)
        if col in ANCHOR_COLS and has_digit(t):
            anchors.append(y)
    anchors.sort()

    row_ys = []
    for y in anchors:
        if row_ys and y - row_ys[-1] <= ROW_CLUSTER:
            continue
        row_ys.append(y)

    for y0 in row_ys:
        grouped = {name: [] for name in COLUMN_BINS}
        for y, x1, x2, t in words:
            if abs(y - y0) > ROW_BAND:
                continue
            col = col_of((x1 + x2) / 2, width)
            if col is None:
                continue
            grouped[col].append((x1, t))

        def joined(col):
            return "".join(t for _, t in sorted(grouped[col]))

        sym_parts = []
        for x1, t in sorted(grouped["symbol"]):
            c = clean_symbol(t)
            if c and re.search(r"[a-z]", c):
                sym_parts.append(c)
        if not sym_parts:
            symbol = ""
        else:
            joined_sym = "".join(sym_parts)
            if len(joined_sym) > 6 and len(sym_parts) > 1:
                symbol = sym_parts[0]
            else:
                symbol = joined_sym
        if not plausible_symbol(symbol):
            symbol = ocr_symbol_with_context(img, width, y0)
        if not plausible_symbol(symbol):
            symbol = ""

        price_raw = joined("price")
        qty_raw = joined("quantity")
        mv_raw = joined("market_value")
        cb_raw = joined("cost_basis")
        gl_raw = joined("gain_loss")

        price_s = read_number_cell(img, width, y0, "price", price_raw, money=True)
        qty_s = read_number_cell(img, width, y0, "quantity", qty_raw, money=False)
        mv_s = read_number_cell(img, width, y0, "market_value", mv_raw, money=True)
        cb_s = read_number_cell(img, width, y0, "cost_basis", cb_raw, money=True)
        gl_s = read_number_cell(img, width, y0, "gain_loss", gl_raw, money=True)

        # Crop retry replaces the raw text when the full-page read was empty.
        if not price_raw and price_s:
            price_raw = price_s
        if not qty_raw and qty_s:
            qty_raw = qty_s
        if not mv_raw and mv_s:
            mv_raw = mv_s
        if not cb_raw and cb_s:
            cb_raw = cb_s

        price = D(price_s)
        qty = D(qty_s)
        market = D(mv_s)
        cost = D(cb_s)
        gain = D(gl_s)

        if price is None and qty is None and market is None and cost is None:
            continue

        row_notes = []
        price0, cost0, gain0 = price, cost, gain
        price, qty, market = repair_price_qty_mv(
            price, qty, market, price_raw, qty_raw, mv_raw)

        # Cash and bank-sweep rows print a dash for price and cost basis.
        # Market value equals quantity, so each unit is $1.
        if (price0 is None and cost0 is None and gain0 is None
                and qty is not None and market is not None and qty != 0
                and abs(market - qty) <= Decimal("0.05")):
            cost = market
            cost_note = ""
        else:
            reread_cost = None
            reread_gain = None
            if (market is not None and gain is not None and cost is not None
                    and not (values_close(market - cost, gain)
                             or values_close(market - cost, -gain))):
                reread_cost = D(reread_number(img, width, y0, "cost_basis", money=True))
                reread_gain = D(reread_number(img, width, y0, "gain_loss", money=True))
            cost, cost_note = repair_cost_basis(
                cost, gain, market, cb_raw or cb_s, reread_cost, reread_gain)
        if cost_note:
            row_notes.append(cost_note)

        if qty is None or qty == 0 or cost is None:
            label = symbol or "???"
            warnings.append(
                f"page {pageno} row at y={y0:.0f} {label}: skipped, quantity or cost basis unreadable")
            continue

        if not symbol:
            symbol = "???"
            row_notes.append("symbol not readable")

        if price is not None and not product_close(price, qty, market):
            row_notes.append("price*qty != market value, check this row")

        for note in row_notes:
            warnings.append(f"page {pageno} {symbol}: {note}")

        records.append({
            "sym": symbol,
            "qty": qty_str(qty),
            "per-share": per_share_str(cost, qty),
            "_page": pageno,
            "_price": price,
            "_market": market,
            "_cost": cost,
        })


def parse_pdf(pdf_path, progress_callback=None):
    doc = fitz.open(pdf_path)
    records = []
    warnings = []
    total_pages = len(doc)
    try:
        for pageno, page in enumerate(doc, start=1):
            if progress_callback:
                progress_callback(pageno, total_pages)
            pix = page.get_pixmap(dpi=DPI)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            parse_page(img, pageno, records, warnings)
    finally:
        doc.close()
    return records, warnings


def validate(records):
    """Flag rows whose price * quantity still misses market value after repair."""
    issues = []
    for r in records:
        if not product_close(r["_price"], D(r["qty"]), r["_market"]):
            if r["_price"] is None or r["_market"] is None:
                continue
            calc = r["_price"] * D(r["qty"])
            issues.append(
                f"{r['sym']}: price*qty={calc:.2f} but market value={r['_market']:.2f}")
    return issues


def process_conversion(pdf_path, csv_path, progress_callback=None):
    """Convert a holdings PDF to a 3-column CSV. Returns (records, warnings, issues)."""
    records, warnings = parse_pdf(pdf_path, progress_callback)
    issues = validate(records)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADER)
        writer.writeheader()
        for rec in records:
            writer.writerow({key: rec[key] for key in CSV_HEADER})
    return records, warnings, issues


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("USB Holdings PDF to CSV Converter")
        self.geometry("640x360")
        self.minsize(550, 300)
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_lbl = ttk.Label(
            main_frame,
            text="USB Holdings PDF to CSV Converter",
            font=("Segoe UI", 12, "bold"),
        )
        title_lbl.pack(pady=(0, 15))

        pdf_frame = ttk.Frame(main_frame)
        pdf_frame.pack(fill=tk.X, pady=5)
        pdf_lbl = ttk.Label(pdf_frame, text="Input PDF file:", width=16, anchor="w")
        pdf_lbl.pack(side=tk.LEFT)
        self.pdf_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_var)
        self.pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        pdf_btn = ttk.Button(pdf_frame, text="Browse...", command=self.browse_pdf)
        pdf_btn.pack(side=tk.RIGHT)

        csv_frame = ttk.Frame(main_frame)
        csv_frame.pack(fill=tk.X, pady=5)
        csv_lbl = ttk.Label(csv_frame, text="Output CSV name:", width=16, anchor="w")
        csv_lbl.pack(side=tk.LEFT)
        self.csv_var = tk.StringVar()
        self.csv_entry = ttk.Entry(csv_frame, textvariable=self.csv_var)
        self.csv_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        csv_btn = ttk.Button(csv_frame, text="Browse...", command=self.browse_csv)
        csv_btn.pack(side=tk.RIGHT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)
        self.convert_btn = ttk.Button(
            btn_frame,
            text="Convert PDF to CSV",
            command=self.start_conversion,
            width=22,
        )
        self.convert_btn.pack()

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        self.status_var = tk.StringVar(value="Ready. Enter PDF and CSV file paths or click Browse.")
        self.status_lbl = ttk.Label(main_frame, textvariable=self.status_var, wraplength=580, justify="left")
        self.status_lbl.pack(anchor="w", fill=tk.X)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="Select Input PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if filename:
            self.pdf_var.set(filename)
            if not self.csv_var.get().strip():
                base, _ = os.path.splitext(filename)
                self.csv_var.set(f"{base}.csv")

    def browse_csv(self):
        filename = filedialog.asksaveasfilename(
            title="Select Output CSV File",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if filename:
            self.csv_var.set(filename)

    def start_conversion(self):
        pdf_path = self.pdf_var.get().strip()
        csv_path = self.csv_var.get().strip()

        if not pdf_path:
            messagebox.showwarning("Input Required", "Please specify the input PDF file.")
            self.pdf_entry.focus_set()
            return

        if not os.path.exists(pdf_path):
            messagebox.showerror("File Not Found", f"The input PDF file was not found:\n{pdf_path}")
            return

        if not csv_path:
            messagebox.showwarning("Input Required", "Please specify the output CSV file name.")
            self.csv_entry.focus_set()
            return

        self.convert_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.status_var.set("Processing PDF... Please wait.")

        threading.Thread(
            target=self.run_conversion_worker,
            args=(pdf_path, csv_path),
            daemon=True,
        ).start()

    def run_conversion_worker(self, pdf_path, csv_path):
        def update_progress(current, total):
            pct = (current / total) * 100
            self.after(0, lambda: self.set_progress(pct, current, total))

        try:
            records, warnings, issues = process_conversion(pdf_path, csv_path, update_progress)
            self.after(0, lambda: self.on_conversion_success(records, csv_path, warnings, issues))
        except Exception as e:
            self.after(0, lambda: self.on_conversion_error(str(e)))

    def set_progress(self, pct, current, total):
        self.progress_var.set(pct)
        self.status_var.set(f"Processing page {current} of {total}...")

    def on_conversion_success(self, records, csv_path, warnings, issues):
        self.progress_var.set(100)
        self.convert_btn.config(state=tk.NORMAL)

        msg = f"Successfully converted and wrote {len(records)} holdings to:\n{csv_path}"
        if warnings:
            msg += f"\n\nWarnings: {len(warnings)} note(s) during processing."
        if issues:
            msg += f"\nIssues: {len(issues)} row(s) failed arithmetic validation."

        self.status_var.set(f"Done. Wrote {len(records)} holdings to {os.path.basename(csv_path)}.")
        messagebox.showinfo("Conversion Complete", msg)

    def on_conversion_error(self, error_message):
        self.progress_var.set(0)
        self.convert_btn.config(state=tk.NORMAL)
        self.status_var.set(f"Error: {error_message}")
        messagebox.showerror("Conversion Failed", f"An error occurred during conversion:\n\n{error_message}")


def main():
    if len(sys.argv) == 3:
        pdf_path, csv_path = sys.argv[1], sys.argv[2]
        records, warnings, issues = process_conversion(pdf_path, csv_path)
        print(f"Wrote {len(records)} holdings to {csv_path}")
        for w in warnings:
            print("WARNING:", w)
        if issues:
            print(f"\n{len(issues)} rows failed the price x quantity check (verify by eye):")
            for i in issues:
                print(" ", i)
        else:
            print("All rows passed the price x quantity = market value check.")
    else:
        app = App()
        app.mainloop()


if __name__ == "__main__":
    main()
