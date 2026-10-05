#!/usr/bin/env python3
"""Creates a minimal sample XLS matching the Raport Kasowy structure for testing."""

import xlwt

wb = xlwt.Workbook(encoding='cp1250')
ws = wb.add_sheet('Sheet1')

def mwrite(ws, row, col, value, col_end=None, row_end=None):
    ws.write(row, col, value)
    if col_end is not None or row_end is not None:
        r2 = (row_end or row) + 1
        c2 = (col_end or col) + 1
        ws.merge(row, r2 - 1, col, c2 - 1)

# --- Row 0: company name (merged A1:I3) and "Raport kasowy" header (merged N1:U1) ---
mwrite(ws, 0, 0,  'HURTOWNIA PAPIEROSÓW TOMASZ MICHALSKI', col_end=8, row_end=2)
mwrite(ws, 0, 13, 'Raport kasowy', col_end=20)

# --- Row 1: FK nazwa (N2:U2) ---
mwrite(ws, 1, 13, 'RK/TEST/2026', col_end=20)

# --- Row 2-3: period text (N3:U4) ---
mwrite(ws, 2, 13, '  za okres : 01.05.2026 - 31.05.2026 ', col_end=20, row_end=3)

# --- Row 4-5: page number (col 20, merged rows 4-5) ---
mwrite(ws, 4, 20, 'Strona 1 z 1', row_end=5)

# --- Row 5-6: repeated RK id (cols 0-4, merged) ---
mwrite(ws, 5, 0, 'RK/TEST/2026', col_end=4, row_end=6)

# --- Row 7: blank ---

# --- Row 8: column headers ---
ws.write(8, 0,  'Lp.')
mwrite(ws, 8,  1, 'Data',      col_end=2)
mwrite(ws, 8,  3, 'Nr. KP/KW', col_end=4)
mwrite(ws, 8,  5, 'Klient',    col_end=6)
mwrite(ws, 8,  7, 'Tytułem',   col_end=13)
ws.write(8, 14, 'F')
mwrite(ws, 8, 15, 'Nr Faktury', col_end=17)
ws.write(8, 18, 'Wpłata')
mwrite(ws, 8, 19, 'Wypłata',   col_end=20)

# --- Data rows ---
def data_row(ws, r, lp, data, nr, klient, tytulem, nr_fakt, wplata, wyplata):
    ws.write(r, 0, str(lp))
    mwrite(ws, r,  1, data,    col_end=2)
    mwrite(ws, r,  3, nr,      col_end=4)
    mwrite(ws, r,  5, klient,  col_end=6)
    mwrite(ws, r,  7, tytulem, col_end=13)
    ws.write(r, 14, '')
    mwrite(ws, r, 15, nr_fakt, col_end=17)
    ws.write(r, 18, wplata)
    mwrite(ws, r, 19, wyplata, col_end=20)

#           row lp  data           nr_kpkw        klient     tytulem                              nr_fakt            wplata         wyplata
data_row(ws,  9, 1, '01-05-2026', 'KP/100/2026', 'KOWALSKI', 'Wpłata FV/1/2026/05  (1 500,00)', 'FV/1/2026/05',   '1\xa0500,00', '')
data_row(ws, 10, 2, '02-05-2026', 'KP/101/2026', 'NOWAK',    'Wpłata FV/2/2026/05  (250,50)',   'FV/2/2026/05',   '250,50',      '')
data_row(ws, 11, 3, '03-05-2026', 'KW/50/2026',  'BANK',     'Wypłata WPŁATA BANKOWA  (500,00)','WPŁATA BANKOWA', '',            '500,00')

# --- Summary row (after last data row) ---
# Row 12: blank
# Row 13: totals (Wplata sum in col 18, Wyplata sum in col 19-20)
mwrite(ws, 13, 18, '1\xa0750,50')
mwrite(ws, 13, 19, '500,00', col_end=20)

# Row 14: "KSIEGOWANO" | "Stan poczatkowy" label (cols 17-18) | value (cols 19-20)
mwrite(ws, 14,  0, 'KSIĘGOWANO', col_end=7)
mwrite(ws, 14, 10, 'ILOŚĆ ZAŁĄCZNIKÓW', col_end=14)
mwrite(ws, 14, 17, 'Stan początkowy : ', col_end=18)
mwrite(ws, 14, 19, '1\xa0000,00', col_end=20)

# Row 15: "Stan końcowy" label (cols 17-18) | value (cols 19-20)
mwrite(ws, 15, 17, 'Stan końcowy : ', col_end=18)
mwrite(ws, 15, 19, '1\xa0750,50', col_end=20)

wb.save('rk_sample.xls')
print('Created rk_sample.xls')
