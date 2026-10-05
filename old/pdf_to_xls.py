"""
Convert Raport kasowy PDF to Excel XLS (Excel 97-2003).

Usage:
    python pdf_to_xls.py "rk kwiecien.PDF"
    python pdf_to_xls.py "rk kwiecien.PDF" output.xls

Requirements:
    pip install pdfplumber xlwt
"""

import sys
import os
import re
import pdfplumber
import xlwt


HEADER_KEYWORDS = {"lp.", "data", "klient", "tytułem", "wpłata", "wypłata"}

COLUMNS = [
    "Lp.",
    "Data",
    "Nr. KP/KW",
    "Klient",
    "Tytułem",
    "F",
    "Nr Faktury",
    "Wpłata",
    "Wypłata",
]


def is_header_row(row):
    text = " ".join(c.lower() for c in row if c)
    return sum(kw in text for kw in HEADER_KEYWORDS) >= 3


FOOTER_KEYWORDS = {
    "księgowano", "dzień rej.", "sporządził", "zatwierdził",
    "ilość załączników", "stan początkowy", "stan końcowy",
    "w tym fiskalizowane",
}

def is_footer_row(row):
    text = " ".join(c.lower() for c in row if c)
    if any(kw in text for kw in FOOTER_KEYWORDS):
        return True
    lp    = row[0] if row else ""
    date  = row[1] if len(row) > 1 else ""
    kp_kw = row[2] if len(row) > 2 else ""
    # A genuine data row needs (numeric Lp. + DD-MM-YYYY date) or a KP/KW ref
    lp_ok   = bool(re.match(r"^\d+$", lp.strip()))
    date_ok = bool(re.match(r"\d{2}-\d{2}-\d{4}", date.strip()))
    kpkw_ok = bool(re.match(r"K[PW]/\d", kp_kw.strip()))
    has_id  = (lp_ok and date_ok) or kpkw_ok
    # Grand-total row: no id but contains Polish-formatted amounts
    has_amounts = any(re.match(r"^\d[\d\s]*,\d{2}$", c.strip()) for c in row if c)
    if not has_id and has_amounts:
        return True
    # Admin rows: no id and no amounts (e.g. "KP | KW" or "346 | 32")
    if not has_id and not has_amounts:
        return True
    return False


def extract_summary(pdf_path):
    """
    Extract the 3 footer summary fields from the last page using word positions.
    Returns a list of (label_str, numeric_value) tuples in document order.
    """
    targets = [
        ("stan początkowy",    "Stan początkowy :"),
        ("stan końcowy",       "Stan końcowy :"),
        ("w tym fiskalizowane","W tym fiskalizowane :"),
    ]

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[-1]
        words = page.extract_words()

    # Group words into logical lines (±3 pt tolerance on 'top')
    lines = {}
    for w in words:
        top = w["top"]
        matched = next((k for k in lines if abs(k - top) <= 3), None)
        if matched is None:
            matched = top
            lines[matched] = []
        lines[matched].append(w)

    results = []
    for line_top in sorted(lines):
        line_words = sorted(lines[line_top], key=lambda w: w["x0"])
        line_text  = " ".join(w["text"] for w in line_words).lower()

        for keyword, label in targets:
            if keyword in line_text:
                # Value words appear after the ":" separator
                colon_x = next((w["x0"] for w in line_words if w["text"] == ":"), None)
                if colon_x is not None:
                    val_words = [w["text"] for w in line_words if w["x0"] > colon_x]
                else:
                    val_words = []
                value = clean_amount(" ".join(val_words))
                results.append((label, value))
                break

    return results


def clean_amount(value):
    """Convert Polish-formatted number string to float, or return as-is."""
    if not value:
        return value
    cleaned = value.replace("\n", " ").strip()
    # Remove thousands spaces: "1 234,56" -> "1234,56"
    match = re.match(r"^-?\d[\d\s]*,\d{2}$", cleaned)
    if match:
        return float(cleaned.replace(" ", "").replace(",", "."))
    return cleaned


def extract_rows(pdf_path):
    """Extract all data rows from the PDF, skipping repeated headers.

    Returns (rows, pdf_wplata, pdf_wyplata) where the last two values are the
    grand-total amounts printed in the PDF (None if not found).
    """
    rows = []
    pdf_wplata = None
    pdf_wyplata = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table_settings = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 3,
            }
            tables = page.extract_tables(table_settings)
            if not tables:
                tables = page.extract_tables()

            for table in tables:
                for row in table:
                    if row is None:
                        continue
                    cleaned = [" ".join(c.split()) if c else "" for c in row]
                    if all(c == "" for c in cleaned):
                        continue
                    if is_header_row(cleaned):
                        continue

                    # Detect the grand-total row: 7 empty cols + 2 amounts
                    if (len(cleaned) >= 9
                            and all(not c for c in cleaned[:7])
                            and cleaned[7] and cleaned[8]):
                        w = clean_amount(cleaned[7])
                        y = clean_amount(cleaned[8])
                        if isinstance(w, float) and isinstance(y, float):
                            pdf_wplata, pdf_wyplata = w, y
                            continue  # don't add to data rows

                    if is_footer_row(cleaned):
                        continue
                    rows.append(cleaned)

    return rows, pdf_wplata, pdf_wyplata


def normalize_rows(raw_rows):
    """
    The PDF table has 9 logical columns but pdfplumber may return varying
    column counts depending on page layout.  Map whatever we get onto the
    9 canonical columns as best we can.
    """
    normalized = []
    for row in raw_rows:
        n = len(row)
        if n == 9:
            normalized.append(row)
        elif n == 8:
            # "F" column is often merged with the adjacent cell — insert empty F
            normalized.append(row[:5] + [""] + row[5:])
        elif n > 9:
            # Merge extra cells into Tytułem (column index 4)
            merged = row[:4] + [" ".join(row[4 : n - 4])] + row[n - 4 :]
            normalized.append(merged[:9])
        else:
            # Pad short rows
            normalized.append(row + [""] * (9 - n))
    return normalized


def validate_rows(rows):
    """Return statistics dict: row count, Lp. continuity, column sums."""
    lp_nums = []
    wplata_sum  = 0.0
    wyplata_sum = 0.0

    for row in rows:
        lp = row[0].strip() if row[0] else ""
        if re.match(r"^\d+$", lp):
            lp_nums.append(int(lp))
        w = clean_amount(row[7]) if len(row) > 7 else None
        y = clean_amount(row[8]) if len(row) > 8 else None
        if isinstance(w, (int, float)):
            wplata_sum  += w
        if isinstance(y, (int, float)):
            wyplata_sum += y

    lp_nums.sort()
    expected = list(range(lp_nums[0], lp_nums[-1] + 1)) if lp_nums else []
    gaps = sorted(set(expected) - set(lp_nums))

    return {
        "row_count":   len(rows),
        "lp_min":      lp_nums[0]  if lp_nums else None,
        "lp_max":      lp_nums[-1] if lp_nums else None,
        "lp_count":    len(lp_nums),
        "lp_gaps":     gaps,
        "wplata_sum":  round(wplata_sum,  2),
        "wyplata_sum": round(wyplata_sum, 2),
    }



def save_xls(rows, summary, stats, output_path):
    """Write XLS (Excel 97-2003) using xlwt with computed sums."""
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Raport kasowy")

    # --- style helpers ---
    def make_style(bold=False, size=10, bg=None, align="left",
                   num_fmt=None, border=True):
        style = xlwt.XFStyle()

        font = xlwt.Font()
        font.bold = bold
        font.height = size * 20       # xlwt uses 1/20 pt
        style.font = font

        if bg:
            pat = xlwt.Pattern()
            pat.pattern = xlwt.Pattern.SOLID_PATTERN
            # Convert hex "RRGGBB" → xlwt colour index via closest match
            # xlwt has a fixed palette; use predefined indices for grey shades
            pat.pattern_fore_colour = bg
            style.pattern = pat

        al = xlwt.Alignment()
        al.wrap = xlwt.Alignment.WRAP_AT_RIGHT
        al.vert = xlwt.Alignment.VERT_CENTER
        if align == "center":
            al.horz = xlwt.Alignment.HORZ_CENTER
        elif align == "right":
            al.horz = xlwt.Alignment.HORZ_RIGHT
        else:
            al.horz = xlwt.Alignment.HORZ_LEFT
        style.alignment = al

        if border:
            brd = xlwt.Borders()
            brd.left   = xlwt.Borders.THIN
            brd.right  = xlwt.Borders.THIN
            brd.top    = xlwt.Borders.THIN
            brd.bottom = xlwt.Borders.THIN
            style.borders = brd

        if num_fmt:
            style.num_format_str = num_fmt

        return style

    # xlwt grey palette indices (closest to D9D9D9 / F2F2F2)
    GREY_DARK  = 0x16   # grey 25%  → header / totals
    GREY_LIGHT = 0x02A  # white     → summary (xlwt light grey approximation)

    s_default   = make_style()
    s_header    = make_style(bold=True, bg=GREY_DARK, align="center")
    s_center    = make_style(align="center")
    s_number    = make_style(align="right", num_fmt="#,##0.00")
    s_totals    = make_style(bold=True, bg=GREY_DARK,  align="center")
    s_tot_num   = make_style(bold=True, bg=GREY_DARK,  align="right", num_fmt="#,##0.00")
    s_summary   = make_style(bold=True, bg=GREY_LIGHT, align="right")
    s_sum_num   = make_style(bold=True, bg=GREY_LIGHT, align="right", num_fmt="#,##0.00")

    # Column widths (in character units × 256)
    col_widths = [6, 12, 16, 18, 45, 4, 20, 14, 14]
    for i, w in enumerate(col_widths):
        ws.col(i).width = w * 256

    # Row 0: title (merged A–I via repeated write is not possible in xlwt;
    #         write to col 0 and leave others blank)
    title_style = make_style(bold=True, size=12, align="center", border=False)
    ws.write(0, 0, "HURTOWNIA PAPIEROSOW TOMASZ MICHALSKI - Raport kasowy",
             title_style)

    # Row 1: source note
    src_style = make_style(size=9, align="center", border=False)
    ws.write(1, 0, os.path.basename(output_path).replace(".xls", ".pdf"),
             src_style)

    # Row 2: column headers
    for col_idx, col_name in enumerate(COLUMNS):
        ws.write(2, col_idx, col_name, s_header)

    # Data rows starting at row 3
    for r_idx, row in enumerate(rows, 3):
        for c_idx, value in enumerate(row):
            if c_idx in (7, 8):        # Wpłata / Wypłata
                numeric = clean_amount(value)
                ws.write(r_idx, c_idx, numeric if isinstance(numeric, (int, float)) else "", s_number)
            elif c_idx in (0, 1, 5):   # Lp., Data, F
                ws.write(r_idx, c_idx, value, s_center)
            else:
                ws.write(r_idx, c_idx, value, s_default)

    # Totals row
    totals_row = 3 + len(rows)
    ws.write(totals_row, 0, "RAZEM", s_totals)
    for c in range(1, 7):
        ws.write(totals_row, c, "", s_totals)
    ws.write(totals_row, 7, stats["wplata_sum"],  s_tot_num)
    ws.write(totals_row, 8, stats["wyplata_sum"], s_tot_num)

    # Blank separator row then summary fields
    first_summary = totals_row + 2
    for s_idx, (label, value) in enumerate(summary):
        s_row = first_summary + s_idx
        # Merge-like effect: write label in col 0, leave 1-7 blank, value in col 8
        ws.write(s_row, 0, label, s_summary)
        for c in range(1, 8):
            ws.write(s_row, c, "", s_summary)
        ws.write(s_row, 8,
                 value if isinstance(value, (int, float)) else "",
                 s_sum_num)

    wb.save(output_path)


def fmt(value):
    """Format a number in Polish style for display."""
    if isinstance(value, (int, float)):
        return f"{value:>16,.2f} PLN".replace(",", " ").replace(".", ",")
    return str(value) if value is not None else "—"


def main():
    if len(sys.argv) < 2:
        print("Usage: python pdf_to_xls.py <input.pdf> [output.xlsx]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"File not found: {pdf_path}")
        sys.exit(1)

    base = os.path.splitext(pdf_path)[0]
    output_path = sys.argv[2] if len(sys.argv) > 2 else base + ".xls"

    raw_rows, pdf_wplata, pdf_wyplata = extract_rows(pdf_path)
    rows    = normalize_rows(raw_rows)
    summary = extract_summary(pdf_path)
    stats   = validate_rows(rows)

    save_xls(rows, summary, stats, output_path)

    W = stats["wplata_sum"]
    Y = stats["wyplata_sum"]
    w_ok = pdf_wplata  is not None and abs(W - pdf_wplata)  < 0.01
    y_ok = pdf_wyplata is not None and abs(Y - pdf_wyplata) < 0.01

    SEP = "-" * 54
    print(SEP)
    print(f"  Plik wejsciowy : {os.path.basename(pdf_path)}")
    print(f"  Plik wyjsciowy : {os.path.basename(output_path)}")
    print(SEP)

    # Row count & Lp. continuity
    print(f"  Liczba wierszy : {stats['row_count']}")
    if stats["lp_gaps"]:
        gaps_str = ", ".join(str(g) for g in stats["lp_gaps"][:10])
        if len(stats["lp_gaps"]) > 10:
            gaps_str += f" ... ({len(stats['lp_gaps'])} laczen)"
        print(f"  Kolumna Lp.    : {stats['lp_min']}-{stats['lp_max']}"
              f"  BRAKUJACE: {gaps_str}")
    else:
        print(f"  Kolumna Lp.    : {stats['lp_min']}-{stats['lp_max']}"
              f"  ({stats['lp_count']} pozycji, ciag bez luk)")

    print(SEP)

    # Wpłata
    print(f"  Wplata  (obliczona) :{fmt(W)}")
    if pdf_wplata is not None:
        mark = "OK" if w_ok else "ROZNICA!"
        print(f"  Wplata  (z PDF)     :{fmt(pdf_wplata)}  [{mark}]")
    else:
        print(f"  Wplata  (z PDF)     :  brak danych w PDF")

    print()

    # Wypłata
    print(f"  Wyplata (obliczona) :{fmt(Y)}")
    if pdf_wyplata is not None:
        mark = "OK" if y_ok else "ROZNICA!"
        print(f"  Wyplata (z PDF)     :{fmt(pdf_wyplata)}  [{mark}]")
    else:
        print(f"  Wyplata (z PDF)     :  brak danych w PDF")

    print(SEP)

    # Summary fields
    for label, value in summary:
        # Strip trailing " :" from label, then transliterate for safe console
        safe = label.rstrip(" :")
        safe = (safe.replace("ą","a").replace("ę","e")
                    .replace("ó","o").replace("ś","s")
                    .replace("ł","l").replace("ż","z")
                    .replace("ź","z").replace("ć","c")
                    .replace("ń","n"))
        print(f"  {safe:<22}:{fmt(value)}")

    print(SEP)
    print(f"  Zapisano XLS : {output_path}")
    print(SEP)


if __name__ == "__main__":
    main()
