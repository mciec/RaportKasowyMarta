#!/usr/bin/env python3
"""
XLS/PDF Raport Kasowy -> Symfonia FK import converter.

Usage: python rk_to_fk.py <input.xls|input.pdf> [output.txt]

The input format is chosen by file extension.

If output file is not specified, the output is placed next to the input file
with a .txt extension.
"""

import re
import sys
from datetime import datetime
from pathlib import Path

_PL_CHARS = str.maketrans(
    'ąćęłńóśźżĄĆĘŁŃÓŚŹŻ',
    'acelnoszzACELNOSZZ'
)

def ascii_pl(s):
    """Replace Polish diacritics with plain ASCII equivalents."""
    return str(s).translate(_PL_CHARS)


def parse_polish_decimal(val):
    """Parse Polish-formatted decimal (non-breaking space as thousands, comma as decimal)."""
    if val is None or val == '':
        return None
    if isinstance(val, (int, float)):
        f = float(val)
        return f if f != 0.0 else None
    s = str(val).strip()
    if not s:
        return None
    s = s.replace('\xa0', '').replace(' ', '').replace(' ', '')
    s = s.replace(',', '.')
    try:
        f = float(s)
        return f if f != 0.0 else None
    except ValueError:
        return None


def parse_date_ddmmyyyy(val):
    """Parse dd-MM-yyyy or dd.MM.yyyy date string to yyyy-MM-dd."""
    if not val:
        return None
    s = str(val).strip()
    for fmt in ('%d-%m-%Y', '%d.%m.%Y'):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass
    return s


def parse_end_date_from_period(text):
    """Extract end date from 'za okres : dd.MM.yyyy - dd.MM.yyyy' style text."""
    if not text:
        return None
    matches = re.findall(r'\d{2}[.\-]\d{2}[.\-]\d{4}', str(text))
    if matches:
        d = matches[-1].replace('-', '.')
        try:
            dt = datetime.strptime(d, '%d.%m.%Y')
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass
    return None


def format_decimal(val):
    """Format float as decimal string with 2 decimal places."""
    if val is None:
        return ''
    return f'{val:.2f}'


def build_merge_map(ws):
    """Map every (row, col) in a merged range to its top-left (row, col)."""
    merge_map = {}
    for (r1, r2, c1, c2) in ws.merged_cells:
        for r in range(r1, r2):
            for c in range(c1, c2):
                merge_map[(r, c)] = (r1, c1)
    return merge_map


def get_cell(ws, merge_map, row, col):
    """Return the effective value of a cell, resolving merged ranges to the top-left."""
    key = (row, col)
    if key in merge_map:
        r0, c0 = merge_map[key]
        return ws.cell_value(r0, c0)
    return ws.cell_value(row, col)


def find_raport_kasowy(ws, merge_map):
    """Locate the cell containing 'Raport kasowy' (search in header area only)."""
    for r in range(min(20, ws.nrows)):
        for c in range(ws.ncols):
            val = str(get_cell(ws, merge_map, r, c)).strip()
            if val == 'Raport kasowy':
                return r, c
    return None, None


def is_data_row_lp(lp_val):
    """Return (True, lp_int) if the Lp. cell contains a positive integer, else (False, 0)."""
    try:
        lp = int(float(str(lp_val).strip()))
        return (lp > 0, lp)
    except (ValueError, TypeError):
        return False, 0


def _need(module, pip_name):
    print(f"ERROR: {module} is required. Install with: pip install {pip_name}", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# XLS reader
# ---------------------------------------------------------------------------

def read_xls(input_file):
    """Return a document dict: header fields, balances and a list of row dicts."""
    try:
        import xlrd
    except ImportError:
        _need('xlrd', 'xlrd')

    wb = xlrd.open_workbook(input_file, formatting_info=True)
    ws = wb.sheets()[0]
    merge_map = build_merge_map(ws)

    def gcv(row, col):
        return get_cell(ws, merge_map, row, col)

    # --- Header: locate 'Raport kasowy' and read document metadata ---
    rk_row, rk_col = find_raport_kasowy(ws, merge_map)
    if rk_row is None:
        print("ERROR: could not find 'Raport kasowy' cell in the sheet", file=sys.stderr)
        sys.exit(1)

    company_name = ascii_pl('HURTOWNIA PAPIEROSÓW TOMASZ MICHALSKI')
    fk_nazwa = ascii_pl(gcv(rk_row + 1, rk_col)).strip()
    period_text = str(gcv(rk_row + 2, rk_col)).strip()

    # --- Column layout (top-left column of each merged group) ---
    COL_LP = 0
    COL_DATA = 1      # merged with col 2
    COL_NR_KPKW = 3   # merged with col 4
    COL_KLIENT = 5    # merged with col 6
    COL_TYTULEM = 7   # merged with cols 8-13
    # col 14: F (flag, unused)
    # cols 15-17: Nr Faktury (unused in output)
    COL_WPLATA = 18   # standalone
    COL_WYPLATA = 19  # merged with col 20

    # --- Collect data rows (rows where Lp. is a positive integer) ---
    rows = []
    last_data_row = -1
    for r in range(ws.nrows):
        ok, lp = is_data_row_lp(gcv(r, COL_LP))
        if ok:
            last_data_row = r
            rows.append({
                'lp': lp,
                'data': parse_date_ddmmyyyy(gcv(r, COL_DATA)),
                'nr': ascii_pl(gcv(r, COL_NR_KPKW)).strip(),
                'klient': ascii_pl(gcv(r, COL_KLIENT)).strip(),
                'tytulem': ascii_pl(gcv(r, COL_TYTULEM)).strip(),
                'wplata': parse_polish_decimal(gcv(r, COL_WPLATA)),
                'wyplata': parse_polish_decimal(gcv(r, COL_WYPLATA)),
            })

    # --- Summary section: scan rows after last data row for balance fields ---
    saldo_prk = None
    saldo_zrk = None
    for r in range(last_data_row + 1, ws.nrows):
        for c in range(ws.ncols):
            label = str(gcv(r, c)).strip()
            if 'Stan pocz' in label:
                saldo_prk = parse_polish_decimal(gcv(r, COL_WYPLATA))
            elif 'Stan ko' in label and 'cowy' in label:
                saldo_zrk = parse_polish_decimal(gcv(r, COL_WYPLATA))

    return {
        'symbol_fk': f'{company_name} {fk_nazwa}',
        'fk_nazwa': fk_nazwa,
        'data_wystawienia': parse_end_date_from_period(period_text),
        'saldo_prk': saldo_prk,
        'saldo_zrk': saldo_zrk,
        'rows': rows,
    }


# ---------------------------------------------------------------------------
# PDF reader
# ---------------------------------------------------------------------------

# The PDF font has no usable Unicode map for Polish letters, so pdfplumber
# yields "(cid:NNN)". Glyph ids below were inferred from names in the sample.
_CID_MAP = {260: 'A', 265: 'A', 224: 'L', 275: 'N', 297: 'Z', 295: 'Z', 285: 'Z'}
_unknown_cids = set()


def _resolve_cid(match):
    n = int(match.group(1))
    if n in _CID_MAP:
        return _CID_MAP[n]
    if 3 <= n <= 97:            # plain ASCII glyphs are numbered ascii-29
        return chr(n + 29)
    _unknown_cids.add(n)
    return ''


def clean_pdf_text(s):
    """Resolve (cid:N) glyphs, collapse line breaks, strip diacritics."""
    s = re.sub(r'\(cid:(\d+)\)', _resolve_cid, s or '')
    return ascii_pl(' '.join(s.split()))


def read_pdf(input_file):
    """Return a document dict from a Symfonia 'Raport kasowy' PDF."""
    try:
        import pdfplumber
    except ImportError:
        _need('pdfplumber', 'pdfplumber')

    rows = []
    with pdfplumber.open(input_file) as pdf:
        first_text = pdf.pages[0].extract_text() or ''
        last_text = pdf.pages[-1].extract_text() or ''

        for page in pdf.pages:
            for table in page.extract_tables():
                for cells in table:
                    if len(cells) < 6:
                        continue
                    # Lp. may be wrapped over two lines ("16\n1" -> 161)
                    lp_txt = ''.join((cells[0] or '').split())
                    if not lp_txt.isdigit():
                        continue      # header row, page banner, footer
                    rows.append({
                        'lp': int(lp_txt),
                        'data': parse_date_ddmmyyyy(clean_pdf_text(cells[2])),
                        'nr': clean_pdf_text(cells[1]),
                        'klient': clean_pdf_text(cells[7]) if len(cells) > 7 else '',
                        'tytulem': clean_pdf_text(cells[3]),
                        'wplata': parse_polish_decimal(cells[4]),   # Wn
                        'wyplata': parse_polish_decimal(cells[5]),  # Ma
                    })

    lines = [l.strip() for l in first_text.split('\n') if l.strip()]
    company = clean_pdf_text(lines[0]) if lines else ''
    m = re.search(r'Nr raportu kasowego\s+(.+)', first_text)
    if not m:
        print("ERROR: could not find 'Nr raportu kasowego' in the PDF", file=sys.stderr)
        sys.exit(1)
    fk_nazwa = clean_pdf_text(m.group(1))

    m = re.search(r'za okres od\s+(\S+)\s+do\s+(\S+)', first_text)
    data_wystawienia = parse_date_ddmmyyyy(m.group(2)) if m else None

    m = re.search(r'Z poprzedniego raportu\s+([\d \xa0,.\-]+)', last_text)
    saldo_prk = parse_polish_decimal(m.group(1)) if m else None
    m = re.search(r'Stan kasy\s+([\d \xa0,.\-]+)', last_text)
    saldo_zrk = parse_polish_decimal(m.group(1)) if m else None

    return {
        'symbol_fk': f'{company} {fk_nazwa}',
        'fk_nazwa': fk_nazwa,
        'data_wystawienia': data_wystawienia,
        'saldo_prk': saldo_prk,
        'saldo_zrk': saldo_zrk,
        'rows': rows,
    }


# ---------------------------------------------------------------------------
# Writer (shared by both readers)
# ---------------------------------------------------------------------------

def write_output(doc, output_file):
    out = []
    out.append('Dokument{')
    out.append(f"\tsymbol FK\t\t\t={doc['symbol_fk']}")
    out.append('\tobsluguj jak\t\t=RK')
    out.append(f"\tFK nazwa\t\t\t={doc['fk_nazwa']}")
    out.append(f"\tdataWystawienia\t\t={doc['data_wystawienia']}")
    out.append('\topis FK\t\t\t\t=Raport kasowy')
    out.append('\tKontoKasy\t\t\t=100')
    out.append(f"\tSaldoPRK\t\t\t={format_decimal(doc['saldo_prk'])}")
    out.append(f"\tSaldoZRK\t\t\t={format_decimal(doc['saldo_zrk'])}")
    out.append('')

    for row in doc['rows']:
        wplata, wyplata = row['wplata'], row['wyplata']
        kwota = wplata if wplata is not None else wyplata
        opis = f"{row['tytulem']}, {row['klient']}" if row['klient'] else row['tytulem']

        # Konto per pattern.txt:
        # WN Zapis: Wplata not empty -> "100", else "202-3-1-k"
        # MA Zapis: Wyplata not empty -> "100", else "201-3-1-k"
        konto_wn = '100' if wplata is not None else '202-3-1-k'
        konto_ma = '100' if wyplata is not None else '201-3-1-k'

        for strona, konto in (('WN', konto_wn), ('MA', konto_ma)):
            out.append('\tZapis{')
            out.append(f"\t\tPozycja\t\t\t\t={row['lp']}")
            out.append(f'\t\tkwota\t\t\t\t={format_decimal(kwota)}')
            out.append(f'\t\tstrona\t\t\t\t={strona}')
            out.append(f'\t\tKonto\t\t\t\t={konto}')
            out.append(f"\t\tNumerDok\t\t\t={row['nr']}")
            out.append(f'\t\tOpis\t\t\t\t={opis}')
            out.append(f"\t\tdataKPKW\t\t\t={row['data']}")
            out.append('\t}')

    out.append('}')

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out) + '\n')


def main(input_file, output_file):
    ext = Path(input_file).suffix.lower()
    if ext == '.xls':
        doc = read_xls(input_file)
    elif ext == '.pdf':
        doc = read_pdf(input_file)
    else:
        print(f'ERROR: unsupported input type "{ext}" (expected .xls or .pdf)', file=sys.stderr)
        sys.exit(1)

    write_output(doc, output_file)

    rows = doc['rows']
    print(f'Wrote {len(rows)} entries ({len(rows) * 2} Zapis blocks) to {output_file}')
    print(f"  symbol FK:       {doc['symbol_fk']}")
    print(f"  dataWystawienia: {doc['data_wystawienia']}")
    print(f"  SaldoPRK:        {format_decimal(doc['saldo_prk'])}")
    print(f"  SaldoZRK:        {format_decimal(doc['saldo_zrk'])}")

    lps = [r['lp'] for r in rows]
    if lps != list(range(1, len(lps) + 1)):
        print('WARNING: Lp. values are not a continuous 1..N sequence - rows may be missing or duplicated',
              file=sys.stderr)
    if _unknown_cids:
        print(f'WARNING: unresolved PDF glyphs dropped from text: {sorted(_unknown_cids)}', file=sys.stderr)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(f'Usage: python {sys.argv[0]} <input.xls|input.pdf> [output.txt]')
        sys.exit(1)

    input_path = sys.argv[1]
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        output_path = str(Path(input_path).with_suffix('.txt'))

    main(input_path, output_path)
