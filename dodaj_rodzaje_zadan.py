#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Wprowadzanie RODZAJÓW zadań grupowych (słownik 53 — "Rodzaj zadania", inaczej
"Kategoria zadania" w interfejsie) z pliku Excel.

WAŻNE — to NIE jest to samo co zadania_ad_hoc.py / zadania_ad_hoc_gui.py.
Te skrypty tworzą KONKRETNE zadania (instancje, POST :5070/api/task) dla
mieszkańców na podstawie już istniejącego rodzaju zadania. Ten skrypt tworzy
sam RODZAJ zadania (wpis w katalogu/słowniku 53, POST :5000/api/dictionary-value)
— czyli pozycję, którą potem można wybrać np. w zadania_ad_hoc_gui.py na
liście "Rodzaj zadania (słownik 53)".

Mechanizm grupowania pracownik <-> mieszkaniec (potwierdzone z użytkownikiem
i przechwyconym ruchem przeglądarki, "Dodawanie zadań"):
  - Zadanie mieszkańca to "Zadanie grupowe (mieszkaniec)" (attributeKind 39 = 225).
  - Zadanie pracownika to "Grupujące otwarte (pracownik)" (39=226) albo
    "Grupujące zamknięte (pracownik)" (39=224) — rozpoznawane po słowie
    "otwarte"/"zamknięte" w nazwie zadania z Excela.
  - Zadanie pracownika grupowe MUSI wskazywać "Rodzaj zadania podrzędnego"
    (attributeKind 40) = ID odpowiadającego zadania mieszkańca. Dlatego
    kolejność jest ważna: NAJPIERW tworzymy wszystkie zadania mieszkańca,
    POTEM zadania pracownika z linkiem do właściwego ID (nie osobną edycją
    po fakcie — łączenie dzieje się w tym samym POST co tworzenie).
  - Zadania pracownika bez sparowanego zadania mieszkańca w Excelu (np. same
    zajęcia grupowe otwarte bez odpowiednika po stronie mieszkańca) po prostu
    nie dostają pola "Rodzaj zadania podrzędnego" (opcjonalne).

Pole "Rodzaj usługi" (attributeKind 15, opcjonalne, słownik 51):
  - dla zadania pracownika -> usługa pracownika z kolumny "Usługa pracownik"
    (ta sama, którą tworzy dodaj_uslugi_pracownikow.py — MUSI już istnieć).
  - dla zadania mieszkańca -> usługa mieszkańca z kolumny "Usługa mieszkaniec"
    (już istnieje w systemie z wcześniejszej konfiguracji — ten skrypt jej
    NIE tworzy, tylko wyszukuje po treści; brak = błąd dla tego wiersza).

Pozostałe pola słownika 53 wypełniane z Excela (dopasowanie po nazwie pola
z /api/dictionary-structure-element/by-dictionary-id/53, DYNAMICZNIE, nie
hardkodowane numery attributeKind):
  Kategoria zadania, Czas normatywny (z "Nominalny czas realizacji", parsowane
  "30 minut"/"1 godzina"/"2 godziny" -> minuty), Stanowisko, Ilość pracowników,
  Priorytet zadania, Kanał komunikacyjny.

Pola zawsze wypełniane wartością domyślną (potwierdzone z użytkownikiem, nie
ma ich w Excelu): Długotrwałe=false, Przeznaczony dla procesu=false.
Pola zawsze pomijane w payloadzie (opcjonalne, brak danych w Excelu, brak w
przechwyconym przykładzie): Podstawa prawna, Umiejętności specjalne, Opis
czynności, Horyzont czasowy, Cel zadania.

Jeśli struktura słownika 53 kiedyś zyska nowe wymagane pole, którego powyższe
reguły nie obsługują — skrypt PRZERYWA dla tego wiersza z czytelnym błędem
zamiast zgadywać.

BEZPIECZEŃSTWO: ten skrypt PISZE do systemu. Domyślnie działa w trybie
--dry-run. Żeby faktycznie dodać: --execute + potwierdzenie w konsoli.
Rodzaje zadań, które już istnieją (taka sama treść w słowniku 53) są
pomijane — bezpiecznie uruchomić wielokrotnie na tym samym pliku.

Uruchomienie:
  python dodaj_rodzaje_zadan.py --xlsx plik.xlsx --org-id 1
      # dry-run, arkusz "Arkusz1"
  python dodaj_rodzaje_zadan.py --xlsx plik.xlsx --org-id 1 --execute
"""
import sys
import re
import argparse
import getpass
import logging
import difflib

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
ZADANIA_DICTIONARY_ID = 53
DEFAULT_SHEET = "Arkusz1"

GROUPING_MIESZKANIEC = "Zadanie grupowe (mieszkaniec)"
GROUPING_PRACOWNIK_OTWARTE = "Grupujące otwarte (pracownik)"
GROUPING_PRACOWNIK_ZAMKNIETE = "Grupujące zamknięte (pracownik)"

logger = logging.getLogger("dodaj_rodzaje_zadan")

_OMIT = object()   # sentinel: pole opcjonalne bez wartości -> nie wysyłaj go wcale


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


# ── Słownik 53 (rodzaje zadań) i jego struktura ───────────────────────────────
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
    return data if isinstance(data, list) else (data.get("items") or data.get("data") or [])


FUZZY_MATCH_CUTOFF = 0.85


def find_value_id_by_content(values, wanted_content, field_label):
    """Zwraca (id, błąd, notatka). Najpierw dokładne dopasowanie (bez
    uwzględniania wielkości liter). Jeśli go brak, próbuje dopasowania
    przybliżonego (literówki typu brakująca litera na końcu) — jeśli
    znajdzie jednoznacznego kandydata, zwraca go razem z notatką do
    wypisania w podglądzie/logu (NIGDY po cichu). Bez dopasowania: błąd
    z pełną listą dostępnych opcji."""
    wanted_norm = wanted_content.strip().lower()
    for v in values:
        if (v.get("content") or "").strip().lower() == wanted_norm:
            return v.get("id"), None, None
    by_lower = {(v.get("content") or "").strip().lower(): v for v in values}
    close = difflib.get_close_matches(wanted_norm, list(by_lower.keys()), n=1, cutoff=FUZZY_MATCH_CUTOFF)
    if close:
        matched = by_lower[close[0]]
        note = (f"pole '{field_label}': '{wanted_content.strip()}' → "
               f"'{matched.get('content')}' (dopasowanie przybliżone — sprawdź czy to na pewno o to chodziło)")
        return matched.get("id"), None, note
    available = ", ".join(repr(v.get("content")) for v in values)
    return None, f"Nie znaleziono wartości '{wanted_content}' dla pola '{field_label}'. Dostępne opcje: {available}", None


# ── Grupy zawodowe stanowisk ────────────────────────────────────────────────
# Jedno zadanie może wykonywać wiele stanowisk, ale tylko z tej samej grupy:
# gdy Excel wskazuje jedno konkretne stanowisko (np. "Opiekun"), zadanie ma
# być wykonywalne przez WSZYSTKIE stanowiska z tej samej grupy zawodowej
# (np. też "Starszy opiekun" i "Młodszy opiekun"), nie tylko przez dokładnie
# wpisane. Grupy pochodzą ze słownika 2 (Stanowisko), pole "Grupa zawodowa"
# (eksport Stanowisko_1.xlsx), z trzema ręcznymi poprawkami ustalonymi z
# użytkownikiem 2026-08-28: "opiekun" NIE obejmuje opiekuna medycznego
# (osobna grupa), opiekun medyczny jest połączony z pielęgniarkami, a
# pokojowa/starsza pokojowa zostają bez zmian (tak jak już były w słowniku).
# Ratownik medyczny / starszy ratownik medyczny nie zostały przypisane do
# żadnej z tych trzech grup w rozmowie z użytkownikiem, więc tworzą własną,
# osobną parę zamiast trafiać do grupy "opiekun". Stanowiska spoza tej mapy
# (np. kadry, księgowość, kierownictwo) dopasowują się tylko dokładnie do
# siebie -- patrz expand_stanowisko_ids().
STANOWISKO_GROUPS = {
    "opiekun": {"OPIEKUN", "STARSZY OPIEKUN", "MŁODSZY OPIEKUN"},
    "pielegniarka": {"PIELĘGNIARKA", "STARSZA PIELĘGNIARKA", "OPIEKUN MEDYCZNY", "STARSZY OPIEKUN MEDYCZNY"},
    "pokojowa": {"POKOJOWA", "STARSZA POKOJOWA"},
    "ratownik_medyczny": {"RATOWNIK MEDYCZNY", "STARSZY RATOWNIK MEDYCZNY"},
    "administracja_jednostka": {"STARSZY INSPEKTOR DS.  ADMINISTRACYJNYCH", "INSPEKTOR DS. ADMINISTRACYJNYCH", "PODINSPEKTOR DS. ADMINISTRACYJNYCH"},
    "administracja_syrena": {"TECHNIK", "STARSZY TECHNIK"},
    "bezpieczenstwo_it": {"INFORMATYK", "STARSZY INFORMATYK"},
    "bhp": {"SPECJALISTA DS. BEZPIECZEŃSTWA I HIGIENY PRACY", "INSPEKTOR DS. BEZPIECZEŃSTWA I HIGIENY PRACY", "STARSZY INSPEKTOR DS.  BEZPIECZEŃSTWA I HIGIENY PRACY"},
    "dietetyk": {"STARSZY DIETETYK", "DIETETYK"},
    "finanse_ksiegowosc": {"STARSZY INSPEKTOR DS. FINANSOWO-KSIĘGOWYCH", "INSPEKTOR DS. FINANSOWO-KSIĘGOWYCH", "PODINSPEKTOR DS. FINANSOWO-KSIĘGOWYCH"},
    "gospodarcze_obsluga": {"STARSZY RECEPCJONISTA", "RECEPCJONISTA", "KIEROWNIK DZIAŁU ADMINISTRACYJNO-GOSPODARCZEGO"},
    "kadry": {"INSPEKTOR DS. KADR", "PODINSPEKTOR DS. KADR", "STARSZY INSPEKTOR DS. KADR"},
    "kancelaria": {"INSPEKTOR DS. KANCELARYJNYCH", "SEKRETARKA", "PODINSPEKTOR DS. KANCELARYJNYCH", "STARSZY INSPEKTOR DS.  KANCELARYJNYCH"},
    "kuchnia_kierowanie": {"SZEF KUCHNI"},
    "kierowanie_zespolem": {"KIEROWNIK ZESPOŁU PIELĘGNIAREK", "KIEROWNIK ZESPOŁU"},
    "krawiectwo": {"KRAWIEC", "SZWACZKA"},
    "kulturalno_oswiatowe": {"INSTRUKTOR DS. KULTURALNO-OŚWIATOWYCH", "STARSZY INSTRUKTOR DS. KULTURALNO-OŚWIATOWYCH"},
    "magazyn": {"MAGAZYNIER", "STARSZY MAGAZYNIER"},
    "ochrona_zdrowia": {"STARSZY LEKARZ"},
    "ppoz": {"STARSZY INSPEKTOR PPOŻ.", "INSPEKTOR DS. PPOŻ", "STARSZY INSPEKTOR DS. PPOŻ."},
    "pomoc_kuchenna": {"POMOC KUCHENNA"},
    "praca_socjalna": {"PRACOWNIK SOCJALNY", "STARSZY PRACOWNIK SOCJALNY", "SPECJALISTA PRACY SOCJALNEJ", "STARSZY SPECJALISTA PRACY SOCJALNEJ"},
    "pralnia": {"PRACZKA"},
    "prawne": {"RADCA PRAWNY"},
    "rehabilitacja_ruchowa": {"FIZJOTERAPEUTA", "TECHNIK FIZJOTERAPII", "STARSZY TECHNIK FIZJOTERAPII", "STARSZY FIZJOTERAPEUTA", "TECHNIK MASAŻYSTA", "STARSZY TECHNIK MASAŻYSTA"},
    "religijne": {"KAPELAN"},
    "rzemieslnicze": {"KONSERWATOR", "ROBOTNIK GOSPODARCZY", "STARSZY KONSERWATOR"},
    "terapia_zajeciowa": {"STARSZY TERAPEUTA ZAJĘCIOWY", "TERAPEUTA ZAJĘCIOWY"},
    "psychologiczno_terapeutyczne": {"PSYCHOLOG", "STARSZY TERAPEUTA", "TERAPEUTA"},
    "zamowienia_publiczne": {"INSPEKTOR DS. ZAMÓWIEŃ PUBLICZNYCH", "PODINSPEKTOR DS.ZAMÓWIEŃ PUBLICZNYCH", "STARSZY INSPEKTOR DS. ZAMÓWIEŃ PUBLICZNYCH"},
    "transport": {"KIEROWCA SAMOCHODU OSOBOWEGO"},
    "zywienie": {"STARSZY KUCHARZ", "KUCHARZ"},
    "kierowanie_wtz": {"KIEROWNIK WARSZTATU TERAPII ZAJĘCIOWEJ"},
    "kierowanie_dzialem": {"GŁÓWNY KSIĘGOWY", "KIEROWNIK DZIAŁU OPIEKUŃCZO-TERAPETYCZNEGO", "KIEROWNIK DZIAŁU MEDYCZNO-TERAPEUTYCZNEGO", "ZASTĘPCA KIEROWNIKA DZIAŁU OPIEKUŃCZO-TERAPEUTYCZNEGO"},
    "kierowanie_warsztatem": {"KIEROWNIK WARSZTATU"},
    "opiekun_kwalifikowany": {"LEKARZ", "STARSZY OPIEKUN KWALIFIKOWANY W DOMU POMOCY SPOŁECZNEJ", "OPIEKUN KWALIFIKOWANY W DOMU POMOCY SPOŁECZNEJ"},
    "dyrektor": {"DYREKTOR"},
    "zastepca_dyrektora": {"ZASTĘPCA DYREKTORA"},
}

STANOWISKO_GROUP_BY_POSITION = {pos: key for key, members in STANOWISKO_GROUPS.items() for pos in members}


def expand_stanowisko_ids(values, matched_id):
    """Rozszerza jedno dopasowane stanowisko na wszystkie stanowiska z tej
    samej grupy zawodowej (patrz STANOWISKO_GROUPS) -- zwraca listę id bez
    duplikatów. Stanowiska spoza mapy zwracają tylko swoje własne id (bez
    rozszerzania)."""
    matched = next((v for v in values if v.get("id") == matched_id), None)
    if matched is None:
        return [matched_id]
    content = (matched.get("content") or "").strip().upper()
    group_key = STANOWISKO_GROUP_BY_POSITION.get(content)
    if group_key is None:
        return [matched_id]
    members = STANOWISKO_GROUPS[group_key]
    ids, seen = [], set()
    for v in values:
        if (v.get("content") or "").strip().upper() in members and v.get("id") not in seen:
            seen.add(v.get("id"))
            ids.append(v.get("id"))
    return ids or [matched_id]


_DURATION_RANGE_RE = re.compile(r"\d+(?:[.,]\d+)?\s*-\s*\d+(?:[.,]\d+)?\s*(minut|min\.?\b|godzin|h\b)")
_DURATION_VALUE_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(minut\w*|min\.?\b|godzin\w*|h\b)")


def parse_duration_minutes(text):
    """'30 minut' -> 30, '1 godzina' -> 60, '2 godziny' -> 120, '30 min' -> 30,
    '1 h' -> 60. Szuka liczby z jednostką GDZIEKOLWIEK w tekście (nie tylko na
    początku wiersza) i rozpoznaje skróty 'min'/'h' -- w Excelu czas bywa
    zapisany jako np. 'DOT B – gr IV – 50 minut' albo '30 min dla całej grupy
    mieszkańców'. Jeśli w tekście jest ZAKRES (np. '1-1,5 h, 2x dziennie') nie
    zgadujemy, który koniec zakresu wybrać, i zwracamy None -- tak samo jak
    dla tekstu bez żadnej jawnej liczby+jednostki (np. 'W zależności od
    potrzeb')."""
    t = (text or "").strip().lower()
    if not t:
        return None
    if _DURATION_RANGE_RE.search(t):
        return None
    m = _DURATION_VALUE_RE.search(t)
    if not m:
        return None
    n = float(m.group(1).replace(",", "."))
    is_hours = m.group(2).startswith("godzin") or m.group(2) == "h"
    return int(round(n * 60 if is_hours else n))


def resolve_task_value_attributes(client, structure_elements, row, side, dict_value_cache,
                                  subordinate_id=None):
    """Buduje valueAttributes dla jednego wpisu w słowniku 53 (jedno "zadanie
    pracownika" albo jedno "zadanie mieszkańca"). `row` to słownik pól
    wczytanych z jednego wiersza Excela (patrz read_tasks_from_excel).
    `side` to "pracownik" albo "mieszkaniec". Zwraca (attrs, błąd, notatki)
    — `notatki` to lista tekstów o zastosowanych dopasowaniach przybliżonych
    (literówki w Excelu), pusta jeśli wszystko dopasowało się dokładnie."""
    attrs = []
    notes = []
    for elem in structure_elements:
        name = (elem.get("name") or "").strip().lower()
        linked_dict_id = elem.get("elementDictionaryId")
        attr_kind = elem.get("elementKind")
        attr_type = elem.get("elementType")

        def dict_values(dict_id=linked_dict_id):
            if dict_id not in dict_value_cache:
                dict_value_cache[dict_id] = fetch_dictionary_values(client, dict_id)
            return dict_value_cache[dict_id]

        if name == "kategoria zadania":
            wanted = row["kategoria_zadania_pracownik"] if side == "pracownik" else row["kategoria_zadania_mieszkaniec"]
            value, err, note = find_value_id_by_content(dict_values(), wanted, elem.get("name"))
            if err: return None, err, notes
            if note: notes.append(note)

        elif name == "czas normatywny":
            value = parse_duration_minutes(row["czas_realizacji"])
            if value is None:
                return None, f"Nie rozpoznano czasu trwania '{row['czas_realizacji']}' (oczekiwano np. '30 minut', '1 godzina').", notes

        elif name == "stanowisko":
            names = [x.strip() for x in (row["stanowisko"] or "").split(",") if x.strip()]
            if not names:
                return None, "Brak wartości w polu 'Stanowisko'.", notes
            ids, seen_ids = [], set()
            for n in names:
                vid, err, note = find_value_id_by_content(dict_values(), n, elem.get("name"))
                if err: return None, err, notes
                if note: notes.append(note)
                for expanded_id in expand_stanowisko_ids(dict_values(), vid):
                    if expanded_id not in seen_ids:
                        seen_ids.add(expanded_id)
                        ids.append(expanded_id)
            value = ids

        elif name == "ilość pracowników":
            try:
                value = int(str(row["ilosc_pracownikow"]).strip())
            except (ValueError, TypeError):
                return None, f"'{row['ilosc_pracownikow']}' nie jest liczbą całkowitą (Ilość pracowników).", notes

        elif name == "priorytet zadania":
            value, err, note = find_value_id_by_content(dict_values(), row["priorytet"], elem.get("name"))
            if err: return None, err, notes
            if note: notes.append(note)

        elif name == "podstawa prawna":
            value = _OMIT

        elif name == "długotrwałe":
            value = False

        elif name == "rodzaj usługi":
            usluga = row["usluga_pracownik"] if side == "pracownik" else row["usluga_mieszkaniec"]
            if not usluga:
                value = _OMIT
            else:
                value, err, note = find_value_id_by_content(dict_values(), usluga, elem.get("name"))
                if err: return None, err, notes
                if note: notes.append(note)

        elif name == "kanał komunikacyjny":
            names = [x.strip() for x in (row["kanal"] or "").split(",") if x.strip()]
            if not names:
                value = _OMIT
            else:
                ids = []
                for n in names:
                    vid, err, note = find_value_id_by_content(dict_values(), n, elem.get("name"))
                    if err: return None, err, notes
                    if note: notes.append(note)
                    ids.append(vid)
                value = ids

        elif name == "umiejętności specjalne":
            value = _OMIT

        elif name == "opis czynności":
            value = _OMIT

        elif name == "horyzont czasowy":
            value = _OMIT

        elif name == "przeznaczony dla procesu":
            value = False

        elif name == "rodzaj zadania grupowego":
            if side == "mieszkaniec":
                wanted = GROUPING_MIESZKANIEC
            else:
                task_name_lower = row["zadanie_pracownik"].lower()
                if "otwarte" in task_name_lower:
                    wanted = GROUPING_PRACOWNIK_OTWARTE
                elif "zamkni" in task_name_lower:
                    wanted = GROUPING_PRACOWNIK_ZAMKNIETE
                else:
                    return None, (f"Nazwa zadania '{row['zadanie_pracownik']}' nie zawiera słowa "
                                  f"'otwarte' ani 'zamknięte' — nie wiem jaki 'Rodzaj zadania "
                                  f"grupowego' ustawić."), notes
            value, err, note = find_value_id_by_content(dict_values(), wanted, elem.get("name"))
            if err: return None, err, notes
            if note: notes.append(note)

        elif name == "rodzaj zadania podrzędnego":
            if side == "mieszkaniec" or subordinate_id is None:
                value = _OMIT
            else:
                value = subordinate_id

        elif name == "cel zadania":
            value = _OMIT

        else:
            return None, (f"Słownik {ZADANIA_DICTIONARY_ID} wymaga pola '{elem.get('name')}', którego "
                          f"ten skrypt jeszcze nie obsługuje."), notes

        if value is _OMIT:
            continue
        attrs.append({
            "id": 0, "rowVersion": 0, "isDeleted": False,
            "attributeKind": attr_kind, "attributeType": attr_type, "value": value,
        })
    return attrs, None, notes


def post_task_kind(client, content, display_order, value_attributes):
    payload = {
        "id": 0, "rowVersion": 0, "isDeleted": False,
        "dictionaryId": ZADANIA_DICTIONARY_ID,
        "content": content, "displayOrder": display_order,
        "parentDictionaryValueId": None,
        "valueAttributes": value_attributes,
    }
    r, err = client._post(f"{client.servers['employee']}/api/dictionary-value", json_body=payload)
    if err:
        return None, err
    if r.status_code not in (200, 201, 204):
        try:
            detail = r.text[:300]
        except Exception:
            detail = ""
        return None, f"HTTP {r.status_code}: {detail}"
    # Potwierdzone z rzeczywistego ruchu: odpowiedź to gołe id (np. "176960"),
    # NIE obiekt {"id": ...} jak przy innych POST-ach w tym repo -- stąd
    # osobna obsługa obu kształtów zamiast zakładania jednego.
    try:
        data = r.json()
    except Exception:
        data = None
    if isinstance(data, dict):
        new_id = data.get("id")
    elif isinstance(data, int) and not isinstance(data, bool):
        new_id = data
    else:
        new_id = None
    return new_id, None


# ── Excel ─────────────────────────────────────────────────────────────────────
def _build_merge_map(ws):
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


COLUMN_SPECS = {
    "usluga_pracownik":            lambda h: h.strip() == "Usługa pracownik",
    "kategoria_usluga_pracownik":  lambda h: h.startswith("Kategoria usługi(") and h.rstrip().endswith("(pracownik)"),
    "zadanie_pracownik":           lambda h: h.strip() == "Zadanie pracownik",
    "kategoria_zadania_pracownik": lambda h: h.startswith("Kategoria zadania(") and h.rstrip().lower().endswith("(pracownik)"),
    "zadanie_mieszkaniec":         lambda h: h.strip() == "Zadanie mieszkaniec",
    "kategoria_zadania_mieszkaniec": lambda h: h.startswith("Kategoria zadania(") and h.rstrip().lower().endswith("(mieszkaniec)"),
    "usluga_mieszkaniec":          lambda h: h.strip() == "Usługa mieszkaniec",
    "czas_realizacji":             lambda h: h.strip() == "Nominalny czas realizacji",
    "stanowisko":                  lambda h: h.strip() == "Stanowisko",
    "ilosc_pracownikow":           lambda h: h.startswith("Ilość pracowników"),
    "priorytet":                   lambda h: h.startswith("Piorytet zadania"),
    "kanal":                       lambda h: h.startswith("Kanał komunikacyjny"),
}


def read_tasks_from_excel(xlsx_path, sheet_name):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise SystemExit(f"✗ Arkusz '{sheet_name}' nie istnieje w {xlsx_path}. Dostępne: {wb.sheetnames}")
    ws = wb[sheet_name]
    merge_map = _build_merge_map(ws)

    cols = {}
    for key, predicate in COLUMN_SPECS.items():
        cols[key] = find_header_column(ws, 1, predicate, key)

    rows = []
    seen_pracownik = set()
    for r in range(2, ws.max_row + 1):
        zadanie_pracownik = _cell_value(ws, merge_map, r, cols["zadanie_pracownik"])
        if not zadanie_pracownik or not str(zadanie_pracownik).strip():
            continue
        zadanie_pracownik = str(zadanie_pracownik).strip()
        if zadanie_pracownik.lower() in seen_pracownik:
            continue   # ten sam wiersz-grupa (np. dodatkowe wiersze "Południe"/"Wieczór")
        seen_pracownik.add(zadanie_pracownik.lower())

        def get(key):
            v = _cell_value(ws, merge_map, r, cols[key])
            return str(v).strip() if v is not None and str(v).strip() else None

        zadanie_mieszkaniec = get("zadanie_mieszkaniec")
        row = {
            "row": r,
            "usluga_pracownik": get("usluga_pracownik"),
            "kategoria_usluga_pracownik": get("kategoria_usluga_pracownik"),
            "zadanie_pracownik": zadanie_pracownik,
            "kategoria_zadania_pracownik": get("kategoria_zadania_pracownik"),
            "zadanie_mieszkaniec": zadanie_mieszkaniec,
            "kategoria_zadania_mieszkaniec": get("kategoria_zadania_mieszkaniec") if zadanie_mieszkaniec else None,
            "usluga_mieszkaniec": get("usluga_mieszkaniec") if zadanie_mieszkaniec else None,
            "czas_realizacji": get("czas_realizacji"),
            "stanowisko": get("stanowisko"),
            "ilosc_pracownikow": get("ilosc_pracownikow"),
            "priorytet": get("priorytet"),
            "kanal": get("kanal"),
        }
        missing = [k for k in ("czas_realizacji", "stanowisko", "ilosc_pracownikow", "priorytet")
                  if not row[k]]
        if missing:
            raise SystemExit(f"✗ Wiersz {r} ('{zadanie_pracownik}'): brak wartości w polach: {', '.join(missing)}.")
        if zadanie_mieszkaniec and not row["kategoria_zadania_mieszkaniec"]:
            raise SystemExit(f"✗ Wiersz {r}: '{zadanie_mieszkaniec}' nie ma podanej kategorii zadania (mieszkaniec).")
        rows.append(row)

    return rows


# ─────────────────────────────────────────────────────────────────────────────
def parse_args():
    p = argparse.ArgumentParser(description="Wprowadzanie rodzajów zadań grupowych (słownik 53) z pliku Excel")
    p.add_argument("--login")
    p.add_argument("--password")
    p.add_argument("--env", choices=["prod", "test"], default="test")
    p.add_argument("--org-id", type=int, default=1)
    p.add_argument("--xlsx", required=True, help="Ścieżka do pliku Excel.")
    p.add_argument("--sheet", default=DEFAULT_SHEET, help=f"Nazwa arkusza (domyślnie '{DEFAULT_SHEET}').")
    p.add_argument("--execute", action="store_true", help="Bez tego flaga: tylko podgląd (dry-run), nic nie wysyła.")
    p.add_argument("--log-file", default="dodaj_rodzaje_zadan.log")
    return p.parse_args()


def _plan(rows, existing_by_content):
    """Wylicza co trzeba utworzyć: najpierw zadania mieszkańca, potem
    zadania pracownika (z linkiem do zadania mieszkańca, jeśli jest w wierszu)."""
    resident_planned = {}   # content_lower -> row (kolejność tworzenia)
    employee_planned = []
    skipped = []
    for row in rows:
        if row["zadanie_mieszkaniec"]:
            key = row["zadanie_mieszkaniec"].lower()
            if key not in existing_by_content and key not in resident_planned:
                resident_planned[key] = row
        content_lower = row["zadanie_pracownik"].lower()
        if content_lower in existing_by_content:
            skipped.append(row["zadanie_pracownik"])
            continue
        employee_planned.append(row)
    return list(resident_planned.values()), employee_planned, skipped


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
    print("  Wprowadzanie rodzajów zadań grupowych (słownik 53) — API SYRENA")
    print(f"  Środowisko: {args.env} ({host})  |  organizationId={args.org_id}")
    print(f"  Excel: {args.xlsx}  |  arkusz: {args.sheet}")
    if not args.execute:
        print("  TRYB: DRY-RUN (nic nie zostanie wysłane — dodaj --execute)")
    print("=" * 60)

    rows = read_tasks_from_excel(args.xlsx, args.sheet)
    if not rows:
        print("Brak zadań do wprowadzenia w tym arkuszu — kończę."); sys.exit(0)
    print(f"\n→ znaleziono {len(rows)} zadań pracownika w Excelu "
         f"({sum(1 for r in rows if r['zadanie_mieszkaniec'])} z powiązanym zadaniem mieszkańca)")

    u = args.login or input("\nLogin (email): ").strip()
    pw = args.password or getpass.getpass(f"Hasło dla {u}: ")
    if not client.login(u, pw):
        print("✗ Logowanie nieudane"); sys.exit(1)
    if not client.select_organization(args.org_id):
        print(f"✗ Nie udało się przełączyć na organizationId={args.org_id}"); sys.exit(1)

    structure = fetch_structure_elements(client, ZADANIA_DICTIONARY_ID)
    if not structure:
        print(f"✗ Nie udało się pobrać struktury słownika {ZADANIA_DICTIONARY_ID} — kończę."); sys.exit(1)

    existing = fetch_dictionary_values(client, ZADANIA_DICTIONARY_ID)
    existing_by_content = {(v.get("content") or "").strip().lower(): v.get("id") for v in existing}
    next_display_order = len(existing)

    resident_to_create, employee_to_create, skipped = _plan(rows, existing_by_content)

    if skipped:
        print(f"\n→ pomijam {len(skipped)} zadań pracownika, które już istnieją w słowniku {ZADANIA_DICTIONARY_ID}:")
        for c in skipped:
            print(f"   - {c}")

    print(f"\n{'='*60}\nPODGLĄD")
    print(f"  Zadania mieszkańca do utworzenia: {len(resident_to_create)}")
    for row in resident_to_create:
        print(f"   [wiersz {row['row']}] {row['zadanie_mieszkaniec']}  (kategoria: {row['kategoria_zadania_mieszkaniec']}, usługa: {row['usluga_mieszkaniec']})")
    print(f"  Zadania pracownika do utworzenia: {len(employee_to_create)}")
    for row in employee_to_create:
        link = f" -> podrzędne: {row['zadanie_mieszkaniec']}" if row["zadanie_mieszkaniec"] else ""
        print(f"   [wiersz {row['row']}] {row['zadanie_pracownik']}  (kategoria: {row['kategoria_zadania_pracownik']}, usługa: {row['usluga_pracownik']}){link}")
    print(f"{'='*60}")

    if not resident_to_create and not employee_to_create:
        print("\nWszystko już istnieje — nic do zrobienia."); sys.exit(0)

    # Walidacja "na sucho" -- sprawdza, czy da się zbudować atrybuty dla każdego
    # wiersza (Stanowisko/Kategoria/itd. istnieją w słownikach), żeby błędy i
    # dopasowania przybliżone (literówki w Excelu) było widać już w podglądzie,
    # a nie dopiero w ścianie błędów po --execute.
    dict_value_cache = {}
    validation_errors = []
    fuzzy_notes = []
    for row in resident_to_create:
        _, e, notes = resolve_task_value_attributes(client, structure, row, "mieszkaniec", dict_value_cache)
        if e:
            validation_errors.append(f"{row['zadanie_mieszkaniec']}: {e}")
        fuzzy_notes.extend(notes)
    for row in employee_to_create:
        fake_subordinate = 1 if row["zadanie_mieszkaniec"] else None
        _, e, notes = resolve_task_value_attributes(client, structure, row, "pracownik", dict_value_cache,
                                                     subordinate_id=fake_subordinate)
        if e:
            validation_errors.append(f"{row['zadanie_pracownik']}: {e}")
        fuzzy_notes.extend(notes)

    fuzzy_notes = sorted(set(fuzzy_notes))
    if fuzzy_notes:
        print(f"\n⚠ Dopasowania przybliżone (literówka w Excelu?) — SPRAWDŹ zanim wyślesz:")
        for n in fuzzy_notes:
            print(f"   {n}")
    if validation_errors:
        print(f"\n✗ Błędy walidacji (te wiersze NIE zostaną utworzone):")
        for e in validation_errors:
            print(f"   {e}")

    if not args.execute:
        print("\nTryb dry-run — nic nie wysłano. Dodaj --execute, żeby faktycznie dodać.")
        sys.exit(0)

    total = len(resident_to_create) + len(employee_to_create)
    confirm = input(f"\nPotwierdź dodanie {total} rodzajów zadań do {args.env} [tak/nie]: ").strip().lower()
    if confirm not in ("tak", "t", "yes", "y"):
        print("Anulowano — nic nie wysłano."); sys.exit(0)

    ok = fail = 0
    resident_id_by_content = dict(existing_by_content)   # zawiera już istniejące + będzie uzupełniane

    print("\n— Zadania mieszkańca —")
    for i, row in enumerate(resident_to_create, 1):
        attrs, err, notes = resolve_task_value_attributes(client, structure, row, "mieszkaniec", dict_value_cache)
        for note in notes:
            logger.warning(f"  ⚠ {note}")
        if err:
            fail += 1
            logger.error(f"[mieszkaniec {i}/{len(resident_to_create)}] ✗ {row['zadanie_mieszkaniec']} — {err}")
            continue
        new_id, err = post_task_kind(client, row["zadanie_mieszkaniec"], next_display_order, attrs)
        next_display_order += 1
        if new_id is None:
            fail += 1
            logger.error(f"[mieszkaniec {i}/{len(resident_to_create)}] ✗ {row['zadanie_mieszkaniec']} — {err}")
        else:
            ok += 1
            resident_id_by_content[row["zadanie_mieszkaniec"].lower()] = new_id
            logger.info(f"[mieszkaniec {i}/{len(resident_to_create)}] ✓ {row['zadanie_mieszkaniec']} — utworzono (id={new_id})")

    print("\n— Zadania pracownika —")
    for i, row in enumerate(employee_to_create, 1):
        subordinate_id = None
        if row["zadanie_mieszkaniec"]:
            subordinate_id = resident_id_by_content.get(row["zadanie_mieszkaniec"].lower())
            if subordinate_id is None:
                fail += 1
                logger.error(f"[pracownik {i}/{len(employee_to_create)}] ✗ {row['zadanie_pracownik']} — "
                             f"brak utworzonego zadania mieszkańca '{row['zadanie_mieszkaniec']}' (nie powiodło się wcześniej?)")
                continue
        attrs, err, notes = resolve_task_value_attributes(client, structure, row, "pracownik", dict_value_cache,
                                                           subordinate_id=subordinate_id)
        for note in notes:
            logger.warning(f"  ⚠ {note}")
        if err:
            fail += 1
            logger.error(f"[pracownik {i}/{len(employee_to_create)}] ✗ {row['zadanie_pracownik']} — {err}")
            continue
        new_id, err = post_task_kind(client, row["zadanie_pracownik"], next_display_order, attrs)
        next_display_order += 1
        if new_id is None:
            fail += 1
            logger.error(f"[pracownik {i}/{len(employee_to_create)}] ✗ {row['zadanie_pracownik']} — {err}")
        else:
            ok += 1
            logger.info(f"[pracownik {i}/{len(employee_to_create)}] ✓ {row['zadanie_pracownik']} — utworzono (id={new_id})")

    print(f"\n{'='*60}\nZakończono: {ok}/{total} utworzono, {fail} błędów")
    print(f"Log: {args.log_file}")


if __name__ == "__main__":
    main()
