#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Poprawia GŁÓWNY Excel importowy (Zdania grupowe … uzupelnione.xlsx):

1) KOREKTA „opiekun medyczny" (Opcja B — rola + usługa): zadania będące
   czynnością opiekuńczą (kąpiel, toaleta, posiłki, karmienie, dopajanie,
   pościel, naczynia, strzyżenie/paznokcie) przypisane do roli medycznej
   (<OPIEKUN MEDYCZNY>/<PIELĘGNIARKA>) dostają rolę OPIEKUN oraz poprawną
   usługę wg treści (kąpiel→"kąpiel", toaleta/ubieranie→"Toalety",
   posiłki/karmienie/dopajanie/naczynia→"Wyżywienie", pościel→"Sprzątanie").
   Zmienia: „Usługa pracownik", „Usługa mieszkaniec", kolumnę „Stanowisko"
   (na kanoniczny zestaw opiekunów). Zapisuje arkusz „Zmiany korekty".

   Kolumny usług/stanowiska bywają SCALONE pionowo — przed edycją rozscalamy
   te trzy kolumny i wypełniamy wartością per wiersz (import i tak to czyta).

2) SKAN innych PODEJRZANYCH niezgodności (NIE poprawia — tylko raport do
   decyzji): m.in. zadanie z adnotacją „(Opiekun…)" przy roli medycznej,
   zadania medyczne przy roli nie-medycznej, eskortowanie przy psychologu/
   fizjoterapeucie, terapia/rehabilitacja/psychologia pomieszane z rolą.

Tylko odczyt źródła; wynik zapisywany do nowych plików. Użycie:
  python popraw_glowny_excel.py --xlsx "…uzupelnione.xlsx"
  python popraw_glowny_excel.py --xlsx plik.xlsx --sheet Arkusz1 \
      --out Zdania_grupowe_POPRAWIONY.xlsx --raport Podejrzane_niezgodnosci.xlsx
"""
import argparse
import re
import sys

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    print("pip install openpyxl"); sys.exit(1)

import przygotuj_korekte_opiekun_medyczny as K  # reużycie reguł/pomocników

CANON_OPIEKUN = ("Młodszy opiekun, Opiekun, Opiekun kwalifikowany w domu pomocy "
                 "społecznej, Starszy opiekun, Starszy opiekun kwalifikowany w "
                 "domu pomocy społecznej")

ROLE_RE = re.compile(r"<([^>]+)>")


# ── 1. KOREKTA ────────────────────────────────────────────────────────────────
def plan_corrections(ws_v, mm, cols):
    """Zwraca {wiersz: dict} zmian dla czynności opiekuńczych pod rolą medyczną."""
    c_org, c_up, c_zp, c_um, c_st = cols
    corr = {}
    for r in range(2, ws_v.max_row + 1):
        up = K._cell(ws_v, mm, r, c_up)
        zp = K._cell(ws_v, mm, r, c_zp)
        if not up or not zp:
            continue
        role = K._role_of(up)
        if role not in K.MED_ROLES or not K._is_care(zp):
            continue
        base_b = K._usluga_base_B(zp)
        up_new = K._rewrite_usluga(up, new_base=base_b, new_role="OPIEKUN")
        um = K._cell(ws_v, mm, r, c_um)
        um_new = K._rewrite_usluga(um, new_base=base_b, new_role="OPIEKUN") if um else ""
        st = K._cell(ws_v, mm, r, c_st)
        corr[r] = {
            "org": K._cell(ws_v, mm, r, c_org), "zp": zp,
            "up_old": up, "up_new": up_new,
            "um_old": um, "um_new": um_new,
            "st_old": st, "st_new": CANON_OPIEKUN,
        }
    return corr


def flatten_all(ws):
    """Rozscala WSZYSTKIE scalenia i wypełnia każdy wiersz wartością bloku —
    po tym każdy wiersz jest samodzielny (kolumny zadania/usług bywają scalone
    w 2-wierszowe bloki, więc edycja per wiersz wymaga pełnego spłaszczenia)."""
    for rng in list(ws.merged_cells.ranges):
        top = ws.cell(row=rng.min_row, column=rng.min_col).value
        ws.unmerge_cells(str(rng.coord))
        if rng.min_col == rng.max_col:   # pionowe (dane) — wypełnij kolumnę
            for rr in range(rng.min_row, rng.max_row + 1):
                ws.cell(row=rr, column=rng.min_col).value = top


def apply_corrections(src, sheet, out, cols):
    """Spłaszcza arkusz i nanosi korektę ról per wiersz. Zwraca listę zmian."""
    c_org, c_up, c_zp, c_um, c_st = cols
    wb = openpyxl.load_workbook(src)
    ws = wb[sheet]
    flatten_all(ws)

    changes = []
    for r in range(2, ws.max_row + 1):
        up = str(ws.cell(r, c_up).value or "")
        zp = str(ws.cell(r, c_zp).value or "")
        if not up.strip() or not zp.strip():
            continue
        role = K._role_of(up)
        if role not in K.MED_ROLES or not K._is_care(zp):
            continue
        base_b = K._usluga_base_B(zp)
        up_new = K._rewrite_usluga(up, new_base=base_b, new_role="OPIEKUN")
        um = str(ws.cell(r, c_um).value or "")
        um_new = K._rewrite_usluga(um, new_base=base_b, new_role="OPIEKUN") if um.strip() else ""
        st_old = str(ws.cell(r, c_st).value or "")
        org = str(ws.cell(r, c_org).value or "")
        ws.cell(row=r, column=c_up).value = up_new
        if um.strip():
            ws.cell(row=r, column=c_um).value = um_new
        ws.cell(row=r, column=c_st).value = CANON_OPIEKUN
        changes.append({"r": r, "org": org, "zp": zp, "up_old": up, "up_new": up_new,
                        "um_old": um, "um_new": um_new, "st_old": st_old, "st_new": CANON_OPIEKUN})

    log = wb.create_sheet("Zmiany korekty")
    hf = Font(bold=True, color="FFFFFF"); hfill = PatternFill("solid", fgColor="4472C4")
    heads = ["Wiersz", "DPS", "Zadanie pracownik", "Usługa PRZED", "Usługa PO",
             "Usługa miesz. PRZED", "Usługa miesz. PO", "Stanowisko PRZED", "Stanowisko PO"]
    widths = [8, 16, 46, 36, 36, 36, 36, 34, 34]
    for j, (h, w) in enumerate(zip(heads, widths), start=1):
        c = log.cell(row=1, column=j, value=h); c.font = hf; c.fill = hfill
        c.alignment = Alignment(wrap_text=True, vertical="center")
        log.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    log.freeze_panes = "A2"
    wrap = Alignment(wrap_text=True, vertical="top")
    for rr, d in enumerate(changes, start=2):
        vals = [d["r"], d["org"], d["zp"], d["up_old"], d["up_new"],
                d["um_old"] or "—", d["um_new"] or "—", d["st_old"], d["st_new"]]
        for j, v in enumerate(vals, start=1):
            cell = log.cell(row=rr, column=j, value=v); cell.alignment = wrap
    wb.save(out)
    return changes


# ── 2. SKAN PODEJRZANYCH ──────────────────────────────────────────────────────
MED_KW = ["podawanie lek", "podanie lek", "opatrun", "saturacj", "glikemi",
          "ciśnien", "cisnien", "iniekcj", "zastrzyk", "parametr", "pomiar ciśnien",
          "monitorowanie stanu"]
ESCORT_KW = ["doprowadzanie", "odprowadzanie", "sprowadzanie", "przygotowanie mieszkańca do wyjazdu",
             "wyjazd z mieszkańcem", "sprowadzenie do wyjścia"]
# UWAGA: bez "ipwm" — IPWM (Indywidualny Plan Wsparcia Mieszkańca) to spotkanie
# międzydyscyplinarne (instruktorzy, psycholog, opiekun), więc nie jest samo w
# sobie oznaką błędu roli.
PSYCH_KW = ["psycholog", "agresja mieszkańca", "obserwacja psychologiczna"]
REHAB_KW = ["rehabilitacj", "gimnastyk", "fizjoterap", "sala rehabilitacyjna"]
TERAPIA_KW = ["ludoterapia", "muzykoterapia", "biblioterapia", "hortiterapia",
              "filmoterapia", "sylwoterapia", "terapia zajęciowa", "zajęcia plastyczne",
              "zajęcia teatralne", "zajęcia kulinarne", "zajęcia tematyczne"]


def scan_suspicious(ws_v, mm, cols, already):
    """Zwraca listę (kategoria, DPS, rola, zadanie, usługa) — bez korekty 'care'."""
    c_org, c_up, c_zp, c_um, c_st = cols
    out = []
    seen = set()
    for r in range(2, ws_v.max_row + 1):
        if r in already:
            continue
        up = K._cell(ws_v, mm, r, c_up); zp = K._cell(ws_v, mm, r, c_zp)
        if not up or not zp:
            continue
        role = K._role_of(up) or "(brak)"
        z = zp.lower()
        cat = None
        # adnotacja (Opiekun ...) w treści, a rola nie-OPIEKUN
        paren = " ".join(re.findall(r"\(([^)]*)\)", z))
        if "opiekun" in paren and role != "OPIEKUN":
            cat = "treść mówi '(Opiekun…)', a rola ≠ OPIEKUN"
        elif any(k in z for k in MED_KW) and role not in K.MED_ROLES:
            cat = "zadanie MEDYCZNE przy roli nie-medycznej"
        elif any(k in z for k in ESCORT_KW) and role in ("PSYCHOLOG", "FIZJOTERAPEUTA"):
            cat = "eskortowanie przy psychologu/fizjoterapeucie (zwykle OPIEKUN)"
        elif any(k in z for k in PSYCH_KW) and role != "PSYCHOLOG":
            cat = "zadanie PSYCHOLOGICZNE przy roli ≠ PSYCHOLOG"
        elif any(k in z for k in REHAB_KW) and role != "FIZJOTERAPEUTA":
            cat = "zadanie REHABILITACYJNE przy roli ≠ FIZJOTERAPEUTA"
        elif any(k in z for k in TERAPIA_KW) and not role.startswith("INSTRUKTOR"):
            cat = "zadanie TERAPII ZAJĘCIOWEJ przy roli ≠ INSTRUKTOR"
        if cat:
            key = (cat, K._cell(ws_v, mm, r, c_org), up, zp)
            if key in seen:
                continue
            seen.add(key)
            out.append((cat, K._cell(ws_v, mm, r, c_org), role, zp, up))
    return out


def write_report(rows, path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Podejrzane niezgodności"
    hf = Font(bold=True, color="FFFFFF"); hfill = PatternFill("solid", fgColor="C00000")
    heads = ["Kategoria podejrzenia", "DPS", "Rola (usługa)", "Zadanie pracownik", "Obecna usługa"]
    widths = [44, 16, 22, 52, 40]
    for j, (h, w) in enumerate(zip(heads, widths), start=1):
        c = ws.cell(row=1, column=j, value=h); c.font = hf; c.fill = hfill
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = "A2"
    wrap = Alignment(wrap_text=True, vertical="top")
    for i, (cat, org, role, zp, up) in enumerate(sorted(rows), start=2):
        for j, v in enumerate([cat, org, role, zp, up], start=1):
            ws.cell(row=i, column=j, value=v).alignment = wrap
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="Korekta głównego Excela + skan podejrzanych niezgodności.")
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--sheet", default="Arkusz1")
    ap.add_argument("--out", default="Zdania_grupowe_POPRAWIONY.xlsx")
    ap.add_argument("--raport", default="Podejrzane_niezgodnosci.xlsx")
    args = ap.parse_args()

    wb_v = openpyxl.load_workbook(args.xlsx, data_only=True)
    if args.sheet not in wb_v.sheetnames:
        raise SystemExit(f"✗ Brak arkusza '{args.sheet}'. Dostępne: {wb_v.sheetnames}")
    ws_v = wb_v[args.sheet]
    mm = K._build_merge_map(ws_v)
    cols = (K._find_col(ws_v, "Nazwa organizacji"), K._find_col(ws_v, "Usługa pracownik"),
            K._find_col(ws_v, "Zadanie pracownik"), K._find_col(ws_v, "Usługa mieszkaniec"),
            K._find_col(ws_v, "Stanowisko"))

    corr = plan_corrections(ws_v, mm, cols)   # zbiór wierszy „care" do pominięcia w skanie
    changes = apply_corrections(args.xlsx, args.sheet, args.out, cols)
    print(f"✓ Poprawiono i zapisano: {args.out}")
    print(f"  Skorygowanych wierszy (opiekuńcze spod roli medycznej → OPIEKUN): {len(changes)}")

    susp = scan_suspicious(ws_v, mm, cols, set(corr))
    write_report(susp, args.raport)
    print(f"✓ Raport podejrzanych: {args.raport}  ({len(susp)} pozycji — NIE poprawione, do decyzji)")
    from collections import Counter
    c = Counter(cat for cat, *_ in susp)
    for cat, n in c.most_common():
        print(f"   {n:3d}  {cat}")


if __name__ == "__main__":
    main()
