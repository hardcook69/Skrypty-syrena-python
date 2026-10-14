#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Nanosi DECYZJE usług (z arkusza Niezgodnosci_usluga_DECYZJE.xlsx) na główny
Excel. Dla każdej pozycji decyzyjnej docelowa usługa = kolumna „Twoja usługa"
(jeśli wypełniona), inaczej „PROPONOWANA usługa". Dopasowuje wiersze w bazowym
Excelu po (DPS, Zadanie pracownik, Obecna usługa) i ustawia nową „Usługa
pracownik" oraz odpowiadającą „Usługa mieszkaniec" (ta sama baza + rola +
'(mieszkaniec)'). Rola/stanowisko bez zmian. Zapisuje arkusz „Zmiany usług".

Tylko zapis do nowego pliku. Użycie:
  python zastosuj_decyzje_uslug.py --xlsx Zdania_grupowe_POPRAWIONY.xlsx \
      --decyzje Niezgodnosci_usluga_DECYZJE.xlsx --out Zdania_grupowe_FINALNY.xlsx
"""
import argparse
import re
import sys
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    print("pip install openpyxl"); sys.exit(1)

import przygotuj_korekte_opiekun_medyczny as K


def _usluga_base(u):
    m = re.match(r"^(.*?)\s*<", u or "")
    return (m.group(1).strip() if m else (u or "").split(" - DPS")[0].strip())


def flatten_all(ws):
    """Rozscala WSZYSTKIE scalenia i wypełnia każdy wiersz wartością bloku —
    kolumny zadania/usług bywają scalone w 2-wierszowe bloki, więc dopasowanie
    i edycja per wiersz wymaga pełnego spłaszczenia."""
    for rng in list(ws.merged_cells.ranges):
        top = ws.cell(row=rng.min_row, column=rng.min_col).value
        ws.unmerge_cells(str(rng.coord))
        if rng.min_col == rng.max_col:
            for rr in range(rng.min_row, rng.max_row + 1):
                ws.cell(row=rr, column=rng.min_col).value = top


def read_decisions(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    out = []
    for r in range(3, ws.max_row + 1):
        dps = (ws.cell(r, 2).value or "")
        zad = (ws.cell(r, 3).value or "")
        cur = (ws.cell(r, 4).value or "")
        prop = (ws.cell(r, 6).value or "")
        own = (ws.cell(r, 7).value or "")
        if not (str(zad).strip() and str(cur).strip()):
            continue
        target = str(own).strip() if str(own).strip() else str(prop).strip()
        out.append({"dps": str(dps).strip(), "zad": str(zad).strip(),
                    "cur": str(cur).strip(), "target": target})
    return out


def apply(base_xlsx, sheet, decisions, out_path, cols):
    c_org, c_up, c_zp, c_um = cols
    wb = openpyxl.load_workbook(base_xlsx)
    ws = wb[sheet]
    flatten_all(ws)

    # Dopasowujemy po (DPS, Zadanie pracownik) — kolumna „Obecna usługa" w
    # pliku decyzyjnym bywa edytowana, więc jej NIE używamy. Fallback: zawieranie
    # (np. decyzja 'Ocena…' vs baza 'Okresowa Ocena…'), w obrębie tego DPS.
    idx = defaultdict(list)          # (org, zad) -> [wiersze]
    by_org = defaultdict(list)       # org -> [(zad, wiersz)]
    for r in range(2, ws.max_row + 1):
        org = str(ws.cell(r, c_org).value or "").strip()
        zp = str(ws.cell(r, c_zp).value or "").strip()
        idx[(org, zp)].append(r)
        by_org[org].append((zp, r))

    changes = []
    skipped = []
    nomatch = []
    for d in decisions:
        if "<…>" in d["target"] or "<…>" in d["target"]:
            skipped.append(d); continue
        rows = idx.get((d["dps"], d["zad"]), [])
        if not rows:
            z = d["zad"]
            rows = [r for (zz, r) in by_org.get(d["dps"], [])
                    if zz and (z in zz or zz in z)]
        if not rows:
            nomatch.append(d); continue
        tgt_base = _usluga_base(d["target"])
        for r in rows:
            up_old = str(ws.cell(r, c_up).value or "")
            ws.cell(r, c_up).value = d["target"]
            um_old = str(ws.cell(r, c_um).value or "")
            um_new = ""
            if um_old.strip():
                um_new = K._rewrite_usluga(um_old, new_base=tgt_base, new_role=K._role_of(um_old))
                ws.cell(r, c_um).value = um_new
            changes.append((r, d["dps"], d["zad"], up_old, d["target"], um_old, um_new))

    log = wb.create_sheet("Zmiany usług")
    hf = Font(bold=True, color="FFFFFF"); hfill = PatternFill("solid", fgColor="4472C4")
    heads = ["Wiersz", "DPS", "Zadanie pracownik", "Usługa PRZED", "Usługa PO",
             "Usł. miesz. PRZED", "Usł. miesz. PO"]
    widths = [8, 16, 48, 38, 38, 38, 38]
    for j, (h, w) in enumerate(zip(heads, widths), start=1):
        c = log.cell(row=1, column=j, value=h); c.font = hf; c.fill = hfill
        c.alignment = Alignment(wrap_text=True, vertical="center")
        log.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    log.freeze_panes = "A2"
    wrap = Alignment(wrap_text=True, vertical="top")
    for i, row in enumerate(changes, start=2):
        for j, v in enumerate(row, start=1):
            log.cell(row=i, column=j, value=v).alignment = wrap

    wb.save(out_path)
    return changes, skipped, nomatch


def main():
    ap = argparse.ArgumentParser(description="Zastosuj decyzje usług na główny Excel.")
    ap.add_argument("--xlsx", required=True, help="Baza (Zdania_grupowe_POPRAWIONY.xlsx).")
    ap.add_argument("--decyzje", required=True, help="Odesłany Niezgodnosci_usluga_DECYZJE.xlsx.")
    ap.add_argument("--sheet", default="Arkusz1")
    ap.add_argument("--out", default="Zdania_grupowe_FINALNY.xlsx")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    if args.sheet not in wb.sheetnames:
        raise SystemExit(f"✗ Brak arkusza '{args.sheet}'. Dostępne: {wb.sheetnames}")
    ws = wb[args.sheet]
    cols = (K._find_col(ws, "Nazwa organizacji"), K._find_col(ws, "Usługa pracownik"),
            K._find_col(ws, "Zadanie pracownik"), K._find_col(ws, "Usługa mieszkaniec"))

    decisions = read_decisions(args.decyzje)
    changes, skipped, nomatch = apply(args.xlsx, args.sheet, decisions, args.out, cols)
    print(f"✓ Zapisano: {args.out}")
    print(f"  Decyzji wczytanych: {len(decisions)}")
    print(f"  Zmienionych wierszy w Excelu: {len(changes)}")
    print(f"  Pominięto (cel bez roli '<…>', do ręcznego): {len(skipped)}")
    for d in skipped:
        print(f"     - [{d['dps']}] {d['zad'][:50]}  (cel: {d['target']})")
    if nomatch:
        print(f"  ⚠ Decyzje BEZ dopasowania w Excelu: {len(nomatch)}")
        for d in nomatch:
            print(f"     - [{d['dps']}] {d['zad'][:50]} | obecna: {d['cur'][:40]}")


if __name__ == "__main__":
    main()
