---
name: transform-raport-kasowy
description: Transform a monthly "Raport kasowy" (cash report) XLS or PDF file into the Symfonia FK import text format. Use when the user asks to transform, convert or process a raport kasowy / RK file (e.g. "transform rk10.XLS", "run the script for <file>").
---

# Transform Raport kasowy to Symfonia format

The converter is `rk_to_fk.py` in the project root (the folder containing `.claude`). It reads an `.xls` or `.pdf` cash report and writes the Symfonia FK import file (`Dokument{ ... Zapis{ ... } }`, values unquoted, pure ASCII, UTF-8).

## Steps

1. Resolve the input file in the project root (the user usually gives just a name such as `rk10.XLS`; extension case doesn't matter). If it doesn't exist, list candidates with Glob rather than guessing.
2. Run (from the project root):
   ```
   python rk_to_fk.py "<input file>" [output.txt]
   ```
   Without an output argument the result is `<input name>.txt` next to the input. Do not overwrite an existing, unrelated file - check first if the user named an output.
3. Report the printed summary: number of entries (Zapis blocks = 2x), `symbol FK`, `dataWystawienia`, `SaldoPRK`, `SaldoZRK`, and the output path.
4. Sanity-check and mention any `WARNING` lines the script prints:
   - Lp. not a continuous 1..N sequence -> rows missing/duplicated (page-break or merged-cell problem).
   - Unresolved PDF glyphs (`(cid:N)`) -> a new Polish-letter glyph id; add it to `_CID_MAP` in the script (letters are transliterated to ASCII).
5. Optional extra check: `SaldoPRK + sum(Wplata) - sum(Wyplata) == SaldoZRK`.

## Rules baked into the script (don't re-derive; see pattern.txt)

- 2 `Zapis` per source row: WN and MA, same Pozycja/kwota/NumerDok/Opis/dataKPKW.
- kwota = Wplata if present, else Wyplata (PDF: Wn = Wplata, Ma = Wyplata).
- Konto WN = `100` if Wplata else `202-3-1-k`; Konto MA = `100` if Wyplata else `201-3-1-k`.
- Opis = Tytulem + ", " + Klient (Kontrahent in PDF); comma omitted when the client is empty.
- XLS: `symbol FK` = constant "HURTOWNIA PAPIEROSOW TOMASZ MICHALSKI" + space + FK nazwa; dates `dd-MM-yyyy` -> `yyyy-MM-dd`; page-number/header/summary rows skipped (only rows whose Lp. is an integer count).
- PDF: `symbol FK` = company line (first line of page 1) + space + report number; Lp. split over two lines (e.g. `16`/`1`) is joined to 161; NumerDok keeps a space (`KW 1/02/01/24`).
- No Polish diacritics in output; numbers use `.` decimal, no thousands separator.

## Requirements

Python with `xlrd` (for .xls) and `pdfplumber` (for .pdf). If one is missing the script says which `pip install` to run.

If the user changes a rule (accounts, field formulas), edit `rk_to_fk.py` (`write_output` for output rules, `read_xls` / `read_pdf` for extraction) and update this file and `pattern.txt` accordingly.
