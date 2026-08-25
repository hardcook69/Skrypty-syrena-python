#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Wprowadzanie usług pracowników (słownik 51 — "Usługi") z pliku Excel.

Na razie obsługuje TYLKO usługi pracowników (Typ usługi = "Dla pracownika").
Usługi mieszkańców — do zrobienia później, jeśli będzie potrzebne.

Przepływ (dokładnie taki jak w przechwyconym ruchu przeglądarki, "Dodawanie
usługi"):
  1. GET  :5000/api/dictionary-structure-element/by-dictionary-id/51
         -> pola wymagane dla usługi (nazwa, attributeKind/attributeType,
            i do jakiego słownika odsyła każde pole) — pobierane DYNAMICZNIE,
            nie hardkodowane, żeby nie polegać na numerach kind, które mogą
            się różnić między środowiskami/wersjami.
  2. Dla każdego wymaganego pola: GET :5000/api/dictionary-value/by-dictionary-id/<id>
         -> lista wartości do wyboru w tym polu; dopasowanie po treści (content),
            niewrażliwe na wielkość liter.
  3. GET  :5000/api/dictionary-value/by-dictionary-id/51
         -> istniejące usługi (do sprawdzenia duplikatów i wyliczenia displayOrder)
  4. Dla każdej NOWEJ usługi z Excela:

     POST :5000/api/dictionary-value
       {id:0, rowVersion:0, isDeleted:false, dictionaryId:51,
        content:<nazwa usługi z Excela>, displayOrder:<kolejny numer>,
        parentDictionaryValueId:null,
        valueAttributes:[
          {id:0, rowVersion:0, isDeleted:false,
           attributeKind:<z kroku 1>, attributeType:<z kroku 1>,
           value:<id wybranej wartości ze słownika z kroku 2>},
          ... (jedno na każde wymagane pole)
        ]}

Reguły wyboru wartości pól (potwierdzone z użytkownikiem):
  - Typ usługi        -> zawsze "Dla pracownika" (--typ-uslugi żeby zmienić)
  - Kategoria usługi   -> z kolumny w Excelu (per usługa)
  - Pochodzenie usługi -> zawsze "Wewnętrzna" (--pochodzenie żeby zmienić)
  - jeśli w przyszłości pojawi się jeszcze jedno wymagane pole, którego
    powyższe reguły nie obsługują — skrypt PRZERYWA z czytelnym błędem,
    zamiast zgadywać.

Excel: domyślnie arkusz "Arkusz1", kolumny dopasowywane po nagłówku
(nie po literze kolumny), więc kolejność kolumn może się zmienić bez
psucia skryptu:
  - "Usługa pracownik"                       -> nazwa usługi (content)
  - "Kategoria usługi(...) (pracownik)"      -> kategoria usługi
Komórki połączone (scalone) w Excelu są poprawnie odczytywane (wartość
z scalenia jest propagowana na wszystkie wiersze, które ono obejmuje).

BEZPIECZEŃSTWO: ten skrypt PISZE do systemu. Domyślnie działa w trybie
--dry-run (tylko pokazuje co by zrobił, nic nie wysyła). Żeby faktycznie
dodać usługi: --execute, i mimo to trzeba potwierdzić [tak] w konsoli
tuż przed wysyłką. Usługi, które już istnieją (taka sama treść w słowniku
51) są pomijane — bezpiecznie uruchomić skrypt wielokrotnie na tym samym
pliku.

Uruchomienie:
  python dodaj_uslugi_pracownikow.py --xlsx plik.xlsx --org-id 1
      # dry-run, arkusz "Arkusz1"
  python dodaj_uslugi_pracownikow.py --xlsx plik.xlsx --org-id 1 --execute
      # jw., ale faktycznie dodaje po potwierdzeniu
  python dodaj_uslugi_pracownikow.py --xlsx plik.xlsx --sheet Arkusz2 --execute
"""
import sys
import argparse
import getpass
import logging

try:
    import requests
except ImportError:
    print("pip install requests"); sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("pip install openpyxl"); sys.exit(1)


ENVS = {
    "prod": "apedps01.bzmw.gov.pl",
    "test": "ttapedps01.bzmw.gov.pl",
}

TIMEOUT = 20
USLUGI_DICTIONARY_ID = 51
DEFAULT_SHEET = "Arkusz1"
DEFAULT_TYP_USLUGI = "Dla pracownika"
DEFAULT_POCHODZENIE = "Wewnętrzna"

# Nagłówek kolumny z nazwą usługi — dopasowanie dokładne (po obcięciu białych znaków).
COL_USLUGA_HEADER = "Usługa pracownik"
# Nagłówek kolumny z kategorią usługi jest wieloliniowy (instrukcja wyboru
# wewnątrz nawiasu) — dopasowanie po początku i końcu, nie całości.
COL_KATEGORIA_PREFIX = "Kategoria usługi("
COL_KATEGORIA_SUFFIX = "(pracownik)"

logger = logging.getLogger("dodaj_uslugi_pracownikow")


class Client:
    def __init__(self, host):
        self.servers = {
            "auth": f"http://{host}:5010",
            "employee": f"http://{host}:5000",
            "beneficiary": f"http://{host}:5020",
            "shift": f"http://{host}:5070",
        }
        self.s = requests.Session()
        self.s.headers.update({"Content-Type": "application/json", "Accept": "application/json"})
        self._creds = None
        self._current_org_id = None

    def login(self, username, password):
        for key in ("login", "userName", "username"):
            try:
                r = self.s.post(f"{self.servers['auth']}/api/authentication/token",
                                 json={key: username, "password": password}, timeout=TIMEOUT)
                if r.status_code == 200:
                    d = r.json()
                    tok = d.get("token") or d.get("access_token")
                    if tok:
                        self.s.headers["Authorization"] = f"Bearer {tok}"
                        self._creds = (username, password)
                        logger.info(f"Zalogowano: {d.get('userName')}")
                        return True
            except Exception as e:
                logger.error(f"Błąd logowania (klucz={key}): {e}")
        return False

    def select_organization(self, org_id):
        try:
            r = self.s.post(f"{self.servers['auth']}/api/authentication/token",
                             json={"organizationId": org_id}, timeout=TIMEOUT)
        except Exception as e:
            logger.error(f"Błąd przełączania organizacji org_id={org_id}: {e}")
            return False
        if r.status_code != 200:
            logger.error(f"Przełączenie na organizationId={org_id} nieudane: HTTP {r.status_code}")
            return False
        try:
            d = r.json()
        except Exception:
            d = {}
        tok = d.get("token") or d.get("access_token")
        if tok:
            self.s.headers["Authorization"] = f"Bearer {tok}"
        self._current_org_id = org_id
        logger.info(f"Przełączono na organizationId={org_id}")
        return True

    def _relogin(self):
        if not self._creds:
            return False
        if self.login(*self._creds):
            if self._current_org_id is not None:
                self.select_organization(self._current_org_id)
            return True
        return False

    def _get(self, url, params=None, _retry=True):
        logger.debug(f"GET {url} params={params}")
        try:
            r = self.s.get(url, params=params, timeout=TIMEOUT)
        except Exception as e:
            logger.error(f"Wyjątek przy GET {url}: {e}")
            return None, str(e)
        if r.status_code == 401 and _retry and self._relogin():
            return self._get(url, params=params, _retry=False)
        return r, None

    def _post(self, url, json_body=None, _retry=True):
        logger.debug(f"POST {url} body={json_body}")
        try:
            r = self.s.post(url, json=json_body, timeout=TIMEOUT)
        except Exception as e:
            logger.error(f"Wyjątek przy POST {url}: {e}")
            return None, str(e)
        if r.status_code == 401 and _retry and self._relogin():
            return self._post(url, json_body=json_body, _retry=False)
        return r, None


# ── Słownik 51 (usługi) i jego struktura ──────────────────────────────────────
def fetch_structure_elements(client, dictionary_id):
    r, err = client._get(f"{client.servers['employee']}/api/dictionary-structure-element/by-dictionary-id/{dictionary_id}")
    if err or r is None or r.status_code != 200:
        logger.error(f"Błąd pobierania struktury słownika {dictionary_id}: {err or (r and r.status_code)}")
        return []
    data = r.json()
    items = data if isinstance(data, list) else (data.get("items") or data.get("data") or [])
    logger.info(f"Słownik {dictionary_id}: {len(items)} pól struktury")
    return items


def fetch_dictionary_values(client, dictionary_id):
    r, err = client._get(f"{client.servers['employee']}/api/dictionary-value/by-dictionary-id/{dictionary_id}")
    if err or r is None or r.status_code != 200:
        logger.error(f"Błąd pobierania wartości słownika {dictionary_id}: {err or (r and r.status_code)}")
        return []
    data = r.json()
    items = data if isinstance(data, list) else (data.get("items") or data.get("data") or [])
    return items


def find_value_id_by_content(values, wanted_content, field_label):
    wanted_norm = wanted_content.strip().lower()
    for v in values:
        if (v.get("content") or "").strip().lower() == wanted_norm:
            return v.get("id")
    available = ", ".join(repr(v.get("content")) for v in values)
    raise SystemExit(
        f"✗ Nie znaleziono wartości '{wanted_content}' dla pola '{field_label}'. "
        f"Dostępne opcje: {available}")


def resolve_value_attributes(client, structure_elements, kategoria_uslugi_content,
                              typ_uslugi_content, pochodzenie_content, dict_value_cache=None):
    """Buduje listę valueAttributes na podstawie struktury słownika 51 (kroki 1-2
    z docstringu modułu). Dla każdego wymaganego pola stosuje regułę wyboru
    (Typ usługi / Kategoria usługi / Pochodzenie usługi) — pole, którego nazwa
    nie pasuje do żadnej znanej reguły, przerywa działanie skryptu (nie zgadujemy).
    `dict_value_cache` można podać z zewnątrz, żeby nie odpytywać API osobno dla
    każdej usługi o te same słowniki (Typ/Pochodzenie usługi się nie zmieniają)."""
    attrs = []
    if dict_value_cache is None:
        dict_value_cache = {}
    for elem in structure_elements:
        name = (elem.get("name") or "").strip()
        linked_dict_id = elem.get("elementDictionaryId")
        attr_kind = elem.get("elementKind")
        attr_type = elem.get("elementType")

        if name.lower() == "typ usługi":
            wanted = typ_uslugi_content
        elif name.lower() == "kategoria usługi":
            wanted = kategoria_uslugi_content
        elif name.lower() == "pochodzenie usługi":
            wanted = pochodzenie_content
        else:
            raise SystemExit(
                f"✗ Słownik {USLUGI_DICTIONARY_ID} wymaga pola '{name}', którego ten skrypt "
                f"jeszcze nie obsługuje (obsługiwane: Typ usługi, Kategoria usługi, "
                f"Pochodzenie usługi). Zgłoś to — trzeba dopisać regułę wyboru wartości.")

        if linked_dict_id not in dict_value_cache:
            dict_value_cache[linked_dict_id] = fetch_dictionary_values(client, linked_dict_id)
        value_id = find_value_id_by_content(dict_value_cache[linked_dict_id], wanted, name)

        attrs.append({
            "id": 0, "rowVersion": 0, "isDeleted": False,
            "attributeKind": attr_kind, "attributeType": attr_type, "value": value_id,
        })
    return attrs


def post_service(client, content, display_order, value_attributes):
    payload = {
        "id": 0, "rowVersion": 0, "isDeleted": False,
        "dictionaryId": USLUGI_DICTIONARY_ID,
        "content": content, "displayOrder": display_order,
        "parentDictionaryValueId": None,
        "valueAttributes": value_attributes,
    }
    r, err = client._post(f"{client.servers['employee']}/api/dictionary-value", json_body=payload)
    if err:
        return False, err
    if r.status_code not in (200, 201, 204):
        try:
            detail = r.text[:300]
        except Exception:
            detail = ""
        return False, f"HTTP {r.status_code}: {detail}"
    return True, None


# ── Excel ─────────────────────────────────────────────────────────────────────
def _build_merge_map(ws):
    """Komórki scalone: openpyxl zwraca wartość tylko w lewym-górnym rogu
    scalenia, reszta to None. Ta mapa pozwala odczytać "efektywną" wartość
    dla każdej komórki wewnątrz scalenia."""
    m = {}
    for rng in ws.merged_cells.ranges:
        top_val = ws.cell(row=rng.min_row, column=rng.min_col).value
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                m[(row, col)] = top_val
    return m


def _cell_value(ws, merge_map, row, col):
    v = ws.cell(row=row, column=col).value
    if v is None:
        v = merge_map.get((row, col))
    return v


def find_header_column(ws, header_row, predicate, label):
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=header_row, column=col).value
        if header and predicate(str(header)):
            return col
    raise SystemExit(f"✗ Nie znaleziono kolumny '{label}' w wierszu nagłówka arkusza '{ws.title}'.")


def read_employee_services_from_excel(xlsx_path, sheet_name):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise SystemExit(f"✗ Arkusz '{sheet_name}' nie istnieje w {xlsx_path}. Dostępne: {wb.sheetnames}")
    ws = wb[sheet_name]
    merge_map = _build_merge_map(ws)

    col_usluga = find_header_column(
        ws, 1, lambda h: h.strip() == COL_USLUGA_HEADER, COL_USLUGA_HEADER)
    col_kategoria = find_header_column(
        ws, 1, lambda h: h.startswith(COL_KATEGORIA_PREFIX) and h.rstrip().endswith(COL_KATEGORIA_SUFFIX),
        "Kategoria usługi (pracownik)")

    seen = {}   # content_norm -> (content, kategoria, pierwszy_wiersz)
    order = []
    for row in range(2, ws.max_row + 1):
        usluga = _cell_value(ws, merge_map, row, col_usluga)
        kategoria = _cell_value(ws, merge_map, row, col_kategoria)
        if not usluga or not str(usluga).strip():
            continue
        usluga = str(usluga).strip()
        kategoria = str(kategoria).strip() if kategoria else ""
        if not kategoria:
            raise SystemExit(f"✗ Wiersz {row}: usługa '{usluga}' nie ma podanej kategorii usługi.")
        norm = usluga.lower()
        if norm in seen:
            prev_content, prev_kategoria, prev_row = seen[norm]
            if prev_kategoria.lower() != kategoria.lower():
                raise SystemExit(
                    f"✗ Usługa '{usluga}' ma różne kategorie w wierszach {prev_row} "
                    f"('{prev_kategoria}') i {row} ('{kategoria}') — popraw Excel.")
            continue
        seen[norm] = (usluga, kategoria, row)
        order.append(norm)

    return [seen[n] for n in order]   # [(content, kategoria, wiersz), ...]


# ─────────────────────────────────────────────────────────────────────────────
def parse_args():
    p = argparse.ArgumentParser(description="Wprowadzanie usług pracowników (słownik 51) z pliku Excel")
    p.add_argument("--login")
    p.add_argument("--password")
    p.add_argument("--env", choices=["prod", "test"], default="test")
    p.add_argument("--org-id", type=int, default=1)
    p.add_argument("--xlsx", required=True, help="Ścieżka do pliku Excel.")
    p.add_argument("--sheet", default=DEFAULT_SHEET, help=f"Nazwa arkusza (domyślnie '{DEFAULT_SHEET}').")
    p.add_argument("--typ-uslugi", default=DEFAULT_TYP_USLUGI,
                    help=f"Wartość pola 'Typ usługi' dla wszystkich wpisów (domyślnie '{DEFAULT_TYP_USLUGI}').")
    p.add_argument("--pochodzenie", default=DEFAULT_POCHODZENIE,
                    help=f"Wartość pola 'Pochodzenie usługi' dla wszystkich wpisów (domyślnie '{DEFAULT_POCHODZENIE}').")
    p.add_argument("--execute", action="store_true", help="Bez tego flaga: tylko podgląd (dry-run), nic nie wysyła.")
    p.add_argument("--log-file", default="dodaj_uslugi_pracownikow.log")
    return p.parse_args()


def main():
    args = parse_args()
    logging.basicConfig(filename=args.log_file, level=logging.DEBUG,
                        format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("%(message)s"))
    logging.getLogger().addHandler(console)

    host = ENVS[args.env]
    client = Client(host)

    print("=" * 60)
    print("  Wprowadzanie usług pracowników (słownik 51) — API SYRENA")
    print(f"  Środowisko: {args.env} ({host})  |  organizationId={args.org_id}")
    print(f"  Excel: {args.xlsx}  |  arkusz: {args.sheet}")
    if not args.execute:
        print("  TRYB: DRY-RUN (nic nie zostanie wysłane — dodaj --execute)")
    print("=" * 60)

    services_from_excel = read_employee_services_from_excel(args.xlsx, args.sheet)
    if not services_from_excel:
        print("Brak usług do wprowadzenia w tym arkuszu — kończę."); sys.exit(0)
    print(f"\n→ znaleziono {len(services_from_excel)} unikalnych usług w Excelu:")
    for content, kategoria, row in services_from_excel:
        print(f"   [wiersz {row}] {content}  (kategoria: {kategoria})")

    u = args.login or input("\nLogin (email): ").strip()
    pw = args.password or getpass.getpass(f"Hasło dla {u}: ")
    if not client.login(u, pw):
        print("✗ Logowanie nieudane"); sys.exit(1)
    if not client.select_organization(args.org_id):
        print(f"✗ Nie udało się przełączyć na organizationId={args.org_id}"); sys.exit(1)

    structure_elements = fetch_structure_elements(client, USLUGI_DICTIONARY_ID)
    if not structure_elements:
        print(f"✗ Nie udało się pobrać struktury słownika {USLUGI_DICTIONARY_ID} — kończę."); sys.exit(1)

    existing = fetch_dictionary_values(client, USLUGI_DICTIONARY_ID)
    existing_contents = {(v.get("content") or "").strip().lower() for v in existing}
    next_display_order = len(existing)

    to_create = []
    skipped_existing = []
    for content, kategoria, row in services_from_excel:
        if content.strip().lower() in existing_contents:
            skipped_existing.append(content)
            continue
        to_create.append((content, kategoria, row))

    if skipped_existing:
        print(f"\n→ pomijam {len(skipped_existing)} usług, które już istnieją w słowniku {USLUGI_DICTIONARY_ID}:")
        for c in skipped_existing:
            print(f"   - {c}")

    if not to_create:
        print("\nWszystkie usługi z Excela już istnieją — nic do zrobienia."); sys.exit(0)

    print(f"\n{'='*60}\nPODGLĄD: {len(to_create)} nowych usług do utworzenia")
    print(f"  Typ usługi (wszystkie): {args.typ_uslugi}")
    print(f"  Pochodzenie usługi (wszystkie): {args.pochodzenie}")
    print(f"{'='*60}")
    for content, kategoria, row in to_create:
        print(f"   {content}  (kategoria: {kategoria})")

    if not args.execute:
        print("\nTryb dry-run — nic nie wysłano. Dodaj --execute, żeby faktycznie dodać usługi.")
        sys.exit(0)

    confirm = input(f"\nPotwierdź dodanie {len(to_create)} usług do {args.env} [tak/nie]: ").strip().lower()
    if confirm not in ("tak", "t", "yes", "y"):
        print("Anulowano — nic nie wysłano."); sys.exit(0)

    ok = fail = 0
    dict_value_cache = {}
    for i, (content, kategoria, row) in enumerate(to_create, 1):
        try:
            value_attributes = resolve_value_attributes(
                client, structure_elements, kategoria,
                args.typ_uslugi, args.pochodzenie, dict_value_cache)
        except SystemExit as e:
            logger.error(f"[{i}/{len(to_create)}] ✗ {content} — {e}")
            fail += 1
            continue
        success, err = post_service(client, content, next_display_order, value_attributes)
        if success:
            ok += 1
            next_display_order += 1
            logger.info(f"[{i}/{len(to_create)}] ✓ {content} — utworzono")
        else:
            fail += 1
            logger.error(f"[{i}/{len(to_create)}] ✗ {content} — {err}")

    print(f"\n{'='*60}\nZakończono: {ok}/{len(to_create)} utworzono, {fail} błędów")
    print(f"Log: {args.log_file}")


if __name__ == "__main__":
    main()
