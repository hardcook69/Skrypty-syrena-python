#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Syrena — Wprowadzanie rodzajów zadań grupowych z Excela (GUI)
Domyślnie config_testowy.json obok skryptu; inny plik jako argument CLI
(np. `python dodaj_rodzaje_zadan_gui.py config_produkcja.json` dla
produkcji). Ten sam mechanizm co dodaj_rodzaje_zadan.py (wersja CLI) —
patrz tam po szczegóły przepływu API, kolejności tworzenia i kształtu
payloadu POST :5000/api/dictionary-value.

WAŻNE — to NIE jest to samo co zadania_ad_hoc.py / zadania_ad_hoc_gui.py.
Te skrypty tworzą KONKRETNE zadania (instancje) dla mieszkańców. To GUI
tworzy sam RODZAJ zadania (wpis w katalogu/słowniku 53) — pozycję, którą
potem można wybrać np. w zadania_ad_hoc_gui.py na liście "Rodzaj zadania".

BEZPIECZEŃSTWO: to skrypt PISZĄCY do systemu. "Pokaż podgląd" nic nie
wysyła. Wysyłka wymaga jawnego potwierdzenia w oknie dialogowym (i
dodatkowym ostrzeżeniem, gdy środowisko to PRODUKCJA). Rodzaje zadań,
które już istnieją (taka sama treść w słowniku 53), są pomijane.
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import threading, requests, json, os, re, sys, traceback, logging, difflib

try:
    import ttkbootstrap as ttkb
    HAS_TTKB = True
except ImportError:
    HAS_TTKB = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ── Konfiguracja ─────────────────────────────────────────────────────────────
SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
CONFIG_NAME     = sys.argv[1] if len(sys.argv) > 1 else "config_testowy.json"
CONFIG_PATH     = os.path.join(SCRIPT_DIR, CONFIG_NAME)
CRASH_LOG_PATH  = os.path.join(SCRIPT_DIR, "dodaj_rodzaje_zadan_gui_crash.log")
LAST_LOGIN_PATH = os.path.join(SCRIPT_DIR, "dodaj_rodzaje_zadan_gui_last_login.json")


def _fatal_startup_error(title, message, exc=None):
    detail = f"{message}\n\n{traceback.format_exc()}" if exc else message
    try:
        with open(CRASH_LOG_PATH, "a", encoding="utf-8") as f:
            from datetime import datetime as _dt
            f.write(f"\n[{_dt.now().strftime('%Y-%m-%d %H:%M:%S')}] {title}\n{detail}\n")
    except Exception:
        pass
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, f"{message}\n\nSzczegóły zapisane w:\n{CRASH_LOG_PATH}")
        root.destroy()
    except Exception:
        print(f"{title}: {detail}", file=sys.stderr)
    sys.exit(1)


if not HAS_TTKB:
    _fatal_startup_error(
        "Brak biblioteki ttkbootstrap",
        "To GUI wymaga pakietu 'ttkbootstrap'.\nZainstaluj go poleceniem:\n\n"
        "    pip install ttkbootstrap")

if not HAS_OPENPYXL:
    _fatal_startup_error(
        "Brak biblioteki openpyxl",
        "To GUI wymaga pakietu 'openpyxl' (odczyt plików Excel).\nZainstaluj go poleceniem:\n\n"
        "    pip install openpyxl")


try:
    with open(CONFIG_PATH, encoding="utf-8") as _f:
        CFG = json.load(_f)
except FileNotFoundError:
    _fatal_startup_error(
        "Brak pliku konfiguracyjnego",
        f"Nie znaleziono {CONFIG_PATH}.\nUpewnij się, że {CONFIG_NAME} "
        f"leży w tym samym folderze co ten skrypt.")
except json.JSONDecodeError as e:
    _fatal_startup_error(
        f"Błędny format {CONFIG_NAME}",
        f"Plik {CONFIG_PATH} nie jest poprawnym JSON-em: {e}", exc=e)

try:
    LOG_PATH = os.path.join(SCRIPT_DIR, CFG.get("log_file", "dodaj_rodzaje_zadan_gui.log"))
    logging.basicConfig(filename=LOG_PATH, level=logging.DEBUG,
                        format="%(asctime)s %(levelname)s %(message)s",
                        encoding="utf-8")

    AUTH_URL       = CFG["auth_url"]
    EMP_URL        = CFG.get("employee_api_url", AUTH_URL.replace(":5010", ":5000"))
    ORG_ID_DEFAULT = CFG.get("organization_id", 1)
    ENV_LABEL      = CFG.get("env_label", "TEST")
except KeyError as e:
    _fatal_startup_error(
        f"Brak klucza w {CONFIG_NAME}",
        f"W pliku konfiguracyjnym brakuje wymaganego klucza: {e}\nWymagane: auth_url.", exc=e)


def _load_last_login():
    try:
        with open(LAST_LOGIN_PATH, encoding="utf-8") as f:
            return (json.load(f).get("last_login") or "").strip()
    except Exception:
        return ""

def _save_last_login(login_value):
    try:
        with open(LAST_LOGIN_PATH, "w", encoding="utf-8") as f:
            json.dump({"last_login": login_value}, f)
    except Exception as e:
        logging.warning(f"Nie udało się zapisać ostatniego loginu: {e}")


# ── Styl (ttkbootstrap) ───────────────────────────────────────────────────────
IS_PROD = ENV_LABEL.strip().upper() in ("PROD", "PRODUKCJA", "PRODUCTION")
if IS_PROD:
    ACCENT_STYLE = "danger"
elif ENV_LABEL == "TEST":
    ACCENT_STYLE = "primary"
else:
    ACCENT_STYLE = "info"
THEME_NAME = "united"

FONT_TITLE   = ("Segoe UI", 12, "bold")
FONT_SECTION = ("Segoe UI", 10, "bold")
FONT_BASE    = ("Segoe UI", 10)
FONT_SMALL   = ("Segoe UI", 9)
FONT_TINY    = ("Segoe UI", 8)


def _card(parent):
    return ttkb.Labelframe(parent, text="", bootstyle="secondary")


# ── Backend (ten sam mechanizm co dodaj_rodzaje_zadan.py) ────────────────────
TIMEOUT = 20
ZADANIA_DICTIONARY_ID = 53
DEFAULT_SHEET = "Arkusz1"

GROUPING_MIESZKANIEC = "Zadanie grupowe (mieszkaniec)"
GROUPING_PRACOWNIK_OTWARTE = "Grupujące otwarte (pracownik)"
GROUPING_PRACOWNIK_ZAMKNIETE = "Grupujące zamknięte (pracownik)"

_OMIT = object()


class Client:
    def __init__(self):
        self.servers = {"auth": AUTH_URL, "employee": EMP_URL}
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
                        logging.info(f"Zalogowano: {d.get('userName')}")
                        return True
            except Exception as e:
                logging.error(f"Błąd logowania (klucz={key}): {e}")
        return False

    def select_organization(self, org_id):
        try:
            r = self.s.post(f"{self.servers['auth']}/api/authentication/token",
                             json={"organizationId": org_id}, timeout=TIMEOUT)
        except Exception as e:
            logging.error(f"Błąd przełączania organizacji org_id={org_id}: {e}")
            return False
        if r.status_code != 200:
            logging.error(f"Przełączenie na organizationId={org_id} nieudane: HTTP {r.status_code}")
            return False
        try:
            d = r.json()
        except Exception:
            d = {}
        tok = d.get("token") or d.get("access_token")
        if tok:
            self.s.headers["Authorization"] = f"Bearer {tok}"
        self._current_org_id = org_id
        logging.info(f"Przełączono na organizationId={org_id}")
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
        logging.debug(f"GET {url} params={params}")
        try:
            r = self.s.get(url, params=params, timeout=TIMEOUT)
        except Exception as e:
            logging.error(f"Wyjątek przy GET {url}: {e}")
            return None, str(e)
        logging.debug(f"-> HTTP {r.status_code} {url}: {r.text[:500]}")
        if r.status_code == 401 and _retry and self._relogin():
            return self._get(url, params=params, _retry=False)
        return r, None

    def _post(self, url, json_body=None, _retry=True):
        logging.debug(f"POST {url} body={json_body}")
        try:
            r = self.s.post(url, json=json_body, timeout=TIMEOUT)
        except Exception as e:
            logging.error(f"Wyjątek przy POST {url}: {e}")
            return None, str(e)
        logging.debug(f"-> HTTP {r.status_code} {url}: {r.text[:500]}")
        if r.status_code == 401 and _retry and self._relogin():
            return self._post(url, json_body=json_body, _retry=False)
        return r, None


def fetch_structure_elements(client, dictionary_id):
    r, err = client._get(f"{client.servers['employee']}/api/dictionary-structure-element/by-dictionary-id/{dictionary_id}")
    if err or r is None or r.status_code != 200:
        logging.error(f"Błąd pobierania struktury słownika {dictionary_id}: {err or (r and r.status_code)}")
        return []
    data = r.json()
    items = data if isinstance(data, list) else (data.get("items") or data.get("data") or [])
    logging.info(f"Słownik {dictionary_id}: {len(items)} pól struktury")
    return items


def fetch_dictionary_values(client, dictionary_id):
    r, err = client._get(f"{client.servers['employee']}/api/dictionary-value/by-dictionary-id/{dictionary_id}")
    if err or r is None or r.status_code != 200:
        logging.error(f"Błąd pobierania wartości słownika {dictionary_id}: {err or (r and r.status_code)}")
        return []
    data = r.json()
    return data if isinstance(data, list) else (data.get("items") or data.get("data") or [])


FUZZY_MATCH_CUTOFF = 0.85


def find_value_id_by_content(values, wanted_content, field_label):
    """Zwraca (id, błąd, notatka). Najpierw dokładne dopasowanie (bez
    uwzględniania wielkości liter). Jeśli go brak, próbuje dopasowania
    przybliżonego (literówki typu brakująca litera na końcu) — jeśli
    znajdzie jednoznacznego kandydata, zwraca go razem z notatką do
    pokazania w podglądzie (NIGDY po cichu — użytkownik ma to zobaczyć
    przed wysyłką). Bez dopasowania: błąd z pełną listą dostępnych opcji."""
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
# Poprawka 2026-09-21 (ustalona z użytkownikiem, na podstawie Stanowisko_2.xlsx
# i Grupa_zawodowa.xlsx): "opiekun" obejmuje TERAZ wszystkie niemedyczne
# warianty łącznie z "kwalifikowanymi" (STARSZY OPIEKUN KWALIFIKOWANY W DOMU
# POMOCY SPOŁECZNEJ, OPIEKUN KWALIFIKOWANY W DOMU POMOCY SPOŁECZNEJ) -- dawna
# grupa "opiekun_kwalifikowany" zawierająca je razem z LEKARZ została
# rozbita: LEKARZ ma teraz własną grupę "lekarz". Terapeuci połączeni w
# jedną grupę "terapia_zajeciowa" (TERAPEUTA, STARSZY TERAPEUTA, TERAPEUTA
# ZAJĘCIOWY, STARSZY TERAPEUTA ZAJĘCIOWY), a PSYCHOLOG wydzielony do własnej
# grupy "psycholog" (nie jest terapeutą). Pokojowa/pielęgniarka/fizjoterapia
# sprawdzone ponownie względem tych plików -- bez zmian, brak w nich
# dodatkowych wariantów spoza już zgrupowanych.
STANOWISKO_GROUPS = {
    "opiekun": {"OPIEKUN", "STARSZY OPIEKUN", "MŁODSZY OPIEKUN", "STARSZY OPIEKUN KWALIFIKOWANY W DOMU POMOCY SPOŁECZNEJ", "OPIEKUN KWALIFIKOWANY W DOMU POMOCY SPOŁECZNEJ"},
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
    "terapia_zajeciowa": {"STARSZY TERAPEUTA ZAJĘCIOWY", "TERAPEUTA ZAJĘCIOWY", "STARSZY TERAPEUTA", "TERAPEUTA"},
    "psycholog": {"PSYCHOLOG"},
    "zamowienia_publiczne": {"INSPEKTOR DS. ZAMÓWIEŃ PUBLICZNYCH", "PODINSPEKTOR DS.ZAMÓWIEŃ PUBLICZNYCH", "STARSZY INSPEKTOR DS. ZAMÓWIEŃ PUBLICZNYCH"},
    "transport": {"KIEROWCA SAMOCHODU OSOBOWEGO"},
    "zywienie": {"STARSZY KUCHARZ", "KUCHARZ"},
    "kierowanie_wtz": {"KIEROWNIK WARSZTATU TERAPII ZAJĘCIOWEJ"},
    "kierowanie_dzialem": {"GŁÓWNY KSIĘGOWY", "KIEROWNIK DZIAŁU OPIEKUŃCZO-TERAPETYCZNEGO", "KIEROWNIK DZIAŁU MEDYCZNO-TERAPEUTYCZNEGO", "ZASTĘPCA KIEROWNIKA DZIAŁU OPIEKUŃCZO-TERAPEUTYCZNEGO"},
    "kierowanie_warsztatem": {"KIEROWNIK WARSZTATU"},
    "lekarz": {"LEKARZ"},
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


# Etykiety kolumny "Grupa zawodowa" z Excela (kolumna opcjonalna, obok
# "Stanowisko") -- osobna, grubsza taksonomia niż STANOWISKO_GROUPS, ale
# mapowalna na te same klucze grup. Używana jako DRUGA linia dopasowania
# dla pola 'Stanowisko': gdy dosłowna/przybliżona treść ze Stanowiska nie
# pasuje do ŻADNEJ wartości w słowniku, szukamy najlepszego dopasowania
# WŚRÓD stanowisk należących do wskazanej tu grupy zawodowej zamiast od
# razu zgłaszać błąd -- ustalone z użytkownikiem 2026-08-28 ("stanowiska na
# pewno są w systemie, znajdź najlepsze odpowiednie"). Zawsze z notatką do
# przejrzenia (NIGDY po cichu).
EXCEL_GRUPA_ZAWODOWA_TO_STANOWISKO_GROUPS = {
    "PERSONEL OPIEKUŃCZY": ["opiekun"],
    "PERSONEL OPIEKUŃCZO-MEDYCZNY": ["pielegniarka"],
    "PERSONEL PIELĘGNIARSKI": ["pielegniarka"],
    "PERSONEL SPRZĄTAJĄCY": ["pokojowa"],
    "PERSONEL SOCJALNY": ["praca_socjalna"],
    "PERSONEL TERAPEUTYCZNY": ["rehabilitacja_ruchowa", "kulturalno_oswiatowe", "terapia_zajeciowa", "psycholog"],
}


def find_value_id_by_grupa_zawodowa(values, wanted_content, grupa_zawodowa_text, field_label):
    """Druga linia dopasowania dla pola 'Stanowisko', wywoływana TYLKO gdy
    find_value_id_by_content nie znalazła nic (ani dokładnie, ani
    przybliżenie). Z kolumny 'Grupa zawodowa' w Excelu bierze wszystkie
    wymienione tam etykiety, mapuje na grupy z STANOWISKO_GROUPS i szuka
    NAJLEPSZEGO dopasowania treści ze Stanowiska wśród stanowisk z tych
    grup -- bez progu odcięcia (użytkownik: "na pewno są w systemie", więc
    zawsze wybieramy najbliższe), ale ZAWSZE z notatką do przejrzenia.
    Zwraca (id, notatka) albo (None, None), jeśli kolumna pusta/etykieta
    nierozpoznana albo brak kandydatów w słowniku."""
    if not grupa_zawodowa_text:
        return None, None
    labels = [x.strip().upper() for x in grupa_zawodowa_text.split(",") if x.strip()]
    candidate_positions = set()
    for label in labels:
        for group_key in EXCEL_GRUPA_ZAWODOWA_TO_STANOWISKO_GROUPS.get(label, []):
            candidate_positions |= STANOWISKO_GROUPS[group_key]
    if not candidate_positions:
        return None, None
    by_upper = {(v.get("content") or "").strip().upper(): v for v in values
                if (v.get("content") or "").strip().upper() in candidate_positions}
    if not by_upper:
        return None, None
    close = difflib.get_close_matches(wanted_content.strip().upper(), list(by_upper.keys()), n=1, cutoff=0)
    if not close:
        return None, None
    matched = by_upper[close[0]]
    note = (f"pole '{field_label}': '{wanted_content.strip()}' nie znaleziono w słowniku -- "
           f"dopasowano po grupie zawodowej z Excela ('{grupa_zawodowa_text}') do "
           f"'{matched.get('content')}' -- SPRAWDŹ, czy to na pewno o to stanowisko chodziło.")
    return matched.get("id"), note


# Znane niedopasowania z Excela, które regularnie się powtarzają w różnych
# plikach i mają jednoznaczne, wcześniej rozstrzygnięte z użytkownikiem
# dopasowanie do grupy zawodowej -- działa NIEZALEŻNIE od tego, czy w danym
# pliku jest (jeszcze) kolumna "Grupa zawodowa". Rozstrzygnięte 2026-08-28:
# "Instruktor terapii zajęciowej" -> grupa terapii zajęciowej (nie istnieje
# jako osobne stanowisko w słowniku).
STANOWISKO_ALIASES = {
    "INSTRUKTOR TERAPII ZAJĘCIOWEJ": "terapia_zajeciowa",
}


def find_value_id_by_alias(values, wanted_content, field_label):
    """Dopasowanie dla 'Stanowisko' przez znane, wcześniej rozstrzygnięte
    niedopasowania z Excela (patrz STANOWISKO_ALIASES) -- próbowane PRZED
    find_value_id_by_grupa_zawodowa, bo nie zależy od obecności kolumny
    'Grupa zawodowa' w pliku. Zawsze z notatką do przejrzenia. Zwraca (id,
    notatka) albo (None, None), jeśli treść nie jest znanym aliasem albo w
    słowniku brakuje choćby jednego stanowiska z tej grupy."""
    group_key = STANOWISKO_ALIASES.get(wanted_content.strip().upper())
    if group_key is None:
        return None, None
    members = STANOWISKO_GROUPS[group_key]
    by_upper = {(v.get("content") or "").strip().upper(): v for v in values
                if (v.get("content") or "").strip().upper() in members}
    if not by_upper:
        return None, None
    close = difflib.get_close_matches(wanted_content.strip().upper(), list(by_upper.keys()), n=1, cutoff=0)
    matched = by_upper[close[0]] if close else next(iter(by_upper.values()))
    note = (f"pole '{field_label}': '{wanted_content.strip()}' nie znaleziono w słowniku -- "
           f"znane niedopasowanie, dopasowano do '{matched.get('content')}'.")
    return matched.get("id"), note


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
    """Zwraca (attrs, błąd, notatki). `notatki` to lista tekstów o
    zastosowanych dopasowaniach przybliżonych (literówki w Excelu) —
    puste, jeśli wszystko dopasowało się dokładnie."""
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
                if err:
                    vid2, note2 = find_value_id_by_alias(dict_values(), n, elem.get("name"))
                    if vid2 is None:
                        vid2, note2 = find_value_id_by_grupa_zawodowa(
                            dict_values(), n, row.get("grupa_zawodowa"), elem.get("name"))
                    if vid2 is None:
                        return None, err, notes
                    vid, note = vid2, note2
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
                                  f"'otwarte' ani 'zamknięte' — nie wiadomo jaki 'Rodzaj zadania "
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
                          f"to GUI jeszcze nie obsługuje."), notes

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


def list_sheet_names(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    return wb.sheetnames


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


def find_header_column(ws, header_row, predicate):
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=header_row, column=col).value
        if header and predicate(str(header)):
            return col
    return None


COLUMN_SPECS = {
    "usluga_pracownik":            ("Usługa pracownik", lambda h: h.strip() == "Usługa pracownik"),
    "kategoria_usluga_pracownik":  ("Kategoria usługi (pracownik)", lambda h: h.startswith("Kategoria usługi(") and h.rstrip().endswith("(pracownik)")),
    "zadanie_pracownik":           ("Zadanie pracownik", lambda h: h.strip() == "Zadanie pracownik"),
    "kategoria_zadania_pracownik": ("Kategoria zadania (pracownik)", lambda h: h.startswith("Kategoria zadania(") and h.rstrip().lower().endswith("(pracownik)")),
    "zadanie_mieszkaniec":         ("Zadanie mieszkaniec", lambda h: h.strip() == "Zadanie mieszkaniec"),
    "kategoria_zadania_mieszkaniec": ("Kategoria zadania (Mieszkaniec)", lambda h: h.startswith("Kategoria zadania(") and h.rstrip().lower().endswith("(mieszkaniec)")),
    "usluga_mieszkaniec":          ("Usługa mieszkaniec", lambda h: h.strip() == "Usługa mieszkaniec"),
    "czas_realizacji":             ("Nominalny czas realizacji", lambda h: h.strip() == "Nominalny czas realizacji"),
    "stanowisko":                  ("Stanowisko", lambda h: h.strip() == "Stanowisko"),
    "ilosc_pracownikow":           ("Ilość pracowników", lambda h: h.startswith("Ilość pracowników")),
    "priorytet":                   ("Piorytet zadania", lambda h: h.startswith("Piorytet zadania")),
    "kanal":                       ("Kanał komunikacyjny", lambda h: h.startswith("Kanał komunikacyjny")),
}


def read_tasks_from_excel(xlsx_path, sheet_name):
    """GUI-owa wersja: zwraca (lista_wierszy, błąd) zamiast przerywać program."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return None, f"Arkusz '{sheet_name}' nie istnieje w pliku. Dostępne: {wb.sheetnames}"
    ws = wb[sheet_name]
    merge_map = _build_merge_map(ws)

    cols = {}
    for key, (label, predicate) in COLUMN_SPECS.items():
        col = find_header_column(ws, 1, predicate)
        if col is None:
            return None, f"Nie znaleziono kolumny '{label}' w wierszu nagłówka arkusza '{ws.title}'."
        cols[key] = col
    # Opcjonalna kolumna -- używana tylko jako drugorzędne źródło dopasowania
    # dla pola 'Stanowisko' (patrz find_value_id_by_grupa_zawodowa), starsze
    # pliki Excela mogą jej nie mieć.
    grupa_zawodowa_col = find_header_column(ws, 1, lambda h: h.strip() == "Grupa zawodowa")

    rows = []
    seen_pracownik = set()
    for r in range(2, ws.max_row + 1):
        zadanie_pracownik = _cell_value(ws, merge_map, r, cols["zadanie_pracownik"])
        if not zadanie_pracownik or not str(zadanie_pracownik).strip():
            continue
        zadanie_pracownik = str(zadanie_pracownik).strip()
        if zadanie_pracownik.lower() in seen_pracownik:
            continue
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
            "grupa_zawodowa": (str(_cell_value(ws, merge_map, r, grupa_zawodowa_col)).strip()
                               if grupa_zawodowa_col and _cell_value(ws, merge_map, r, grupa_zawodowa_col) else None),
            "ilosc_pracownikow": get("ilosc_pracownikow"),
            "priorytet": get("priorytet"),
            "kanal": get("kanal"),
        }
        missing = [k for k in ("czas_realizacji", "stanowisko", "ilosc_pracownikow", "priorytet")
                  if not row[k]]
        if missing:
            return None, f"Wiersz {r} ('{zadanie_pracownik}'): brak wartości w polach: {', '.join(missing)}."
        if zadanie_mieszkaniec and not row["kategoria_zadania_mieszkaniec"]:
            return None, f"Wiersz {r}: '{zadanie_mieszkaniec}' nie ma podanej kategorii zadania (mieszkaniec)."
        rows.append(row)

    return rows, None


def plan_tasks(rows, existing_by_content):
    resident_planned = {}
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


# ─────────────────────────────────────────────────────────────────────────────
class App(ttkb.Window):
    def __init__(self):
        super().__init__(title=f"Syrena — Rodzaje zadań grupowych z Excela [{ENV_LABEL}]",
                          themename=THEME_NAME, size=(980, 780), resizable=(True, True))

        self.client = Client()
        self._xlsx_path = None
        self._structure_elements = []
        self._existing_by_content = {}
        self._next_display_order = 0
        self._resident_to_create = []
        self._employee_to_create = []

        self._build_ui()

    # ── BUILD UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        env_icon = "⚠ " if IS_PROD else ""
        hdr = ttkb.Frame(self, bootstyle=f"@{ACCENT_STYLE}", height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ttkb.Label(hdr, text="  Syrena — Rodzaje zadań grupowych z Excela", bootstyle=f"@{ACCENT_STYLE}",
                   font=FONT_TITLE).pack(side="left", padx=(12, 4))
        ttkb.Label(hdr, text=f"{env_icon}[{ENV_LABEL}]", bootstyle=f"@{ACCENT_STYLE}",
                   font=("Segoe UI", 12, "bold")).pack(side="left")
        self._lbl_status = ttkb.Label(hdr, text="●  Niezalogowany", bootstyle=f"@{ACCENT_STYLE}", font=FONT_SMALL)
        self._lbl_status.pack(side="right", padx=12)
        if IS_PROD:
            ttkb.Label(hdr, text="PRODUKCJA — rodzaje zadań trafią od razu do systemu",
                       bootstyle=f"@{ACCENT_STYLE}", font=FONT_TINY).pack(side="right", padx=12)

        nb = ttkb.Notebook(self, bootstyle=ACCENT_STYLE)
        nb.pack(fill="both", expand=True, padx=10, pady=8)
        self._nb = nb

        t1 = ttkb.Frame(nb); nb.add(t1, text="  1. Logowanie  ")
        t2 = ttkb.Frame(nb); nb.add(t2, text="  2. Import rodzajów zadań  ")

        self._tab_login(t1)
        self._tab_import(t2)

        lf = ttkb.Labelframe(self, text=" Log ", bootstyle="secondary")
        lf.pack(fill="x", padx=10, pady=(0, 8))
        self._log = ttkb.ScrolledText(
            lf, height=8, font=("Consolas", 9), bg="#1e1e1e", fg="#ccc",
            insertbackground="white", state="disabled", wrap="word", auto_hide=True)
        self._log.pack(fill="x", padx=6, pady=4)
        self._log.tag_config("ok", foreground="#6dbf67")
        self._log.tag_config("err", foreground="#e06c6c")
        self._log.tag_config("inf", foreground="#9cdcfe")

    # ── TAB 1: Logowanie ──────────────────────────────────────────────────────
    def _tab_login(self, p):
        outer = ttkb.Frame(p, padding=(24, 20))
        outer.pack(anchor="nw", fill="x")
        ttkb.Label(outer, text="Zaloguj się do API Syrena",
                   font=FONT_TITLE).pack(anchor="w", pady=(0, 4))
        ttkb.Label(outer, text=f"Środowisko: {ENV_LABEL}   •   Host: {AUTH_URL}",
                   bootstyle="secondary", font=FONT_TINY).pack(anchor="w", pady=(0, 14))

        card = _card(outer)
        card.pack(anchor="w")
        f = ttkb.Frame(card, padding=(20, 18))
        f.pack()

        for row, (lbl, attr, kw) in enumerate([
            ("Login:", "_e_user", {}),
            ("Hasło:", "_e_pass", {"show": "*"}),
            ("Organizacja (ID):", "_e_org", {}),
        ]):
            ttkb.Label(f, text=lbl, font=FONT_BASE).grid(row=row, column=0, sticky="w", pady=6)
            e = ttkb.Entry(f, width=34, font=FONT_BASE, bootstyle=ACCENT_STYLE, **kw)
            e.grid(row=row, column=1, padx=12, pady=6, sticky="w")
            setattr(self, attr, e)
        self._e_user.insert(0, _load_last_login())
        self._e_org.insert(0, str(ORG_ID_DEFAULT))
        self._e_pass.bind("<Return>", lambda _: self._do_login())
        self._e_user.bind("<Return>", lambda _: self._e_pass.focus_set())

        ttkb.Label(f, text="Hasło nie jest nigdzie zapisywane — tylko login.",
                   bootstyle="secondary", font=FONT_TINY).grid(row=3, column=0, columnspan=2, sticky="w", pady=(2, 12))
        bf = ttkb.Frame(f)
        bf.grid(row=4, column=0, columnspan=2, sticky="w")
        btn_login = ttkb.Button(bf, text="Zaloguj", command=self._do_login,
                                 bootstyle=ACCENT_STYLE, padding=(18, 7))
        btn_login.pack(side="left")
        self._lbl_auth = ttkb.Label(bf, text="", font=FONT_BASE)
        self._lbl_auth.pack(side="left", padx=12)

    # ── TAB 2: Import rodzajów zadań ─────────────────────────────────────────
    def _tab_import(self, p):
        outer = ttkb.Frame(p, padding=(16, 14))
        outer.pack(fill="both", expand=True)

        sec_file = ttkb.Labelframe(outer, text=" Plik Excel ", bootstyle=ACCENT_STYLE, padding=10)
        sec_file.pack(fill="x", pady=(0, 8))
        fr = ttkb.Frame(sec_file); fr.pack(fill="x")
        ttkb.Button(fr, text="📂 Wybierz plik...", command=self._choose_xlsx,
                    bootstyle="secondary", padding=(10, 5)).pack(side="left")
        self._lbl_xlsx = ttkb.Label(fr, text="— nie wybrano —", font=FONT_SMALL, bootstyle="secondary")
        self._lbl_xlsx.pack(side="left", padx=10)

        fr2 = ttkb.Frame(sec_file); fr2.pack(fill="x", pady=(8, 0))
        ttkb.Label(fr2, text="Arkusz:", font=FONT_SMALL).pack(side="left")
        self._sheet_var = tk.StringVar(value=DEFAULT_SHEET)
        self._cb_sheet = ttkb.Combobox(fr2, textvariable=self._sheet_var, state="readonly",
                                       width=20, font=FONT_SMALL, bootstyle=ACCENT_STYLE)
        self._cb_sheet.pack(side="left", padx=(6, 0))
        ttkb.Label(sec_file,
                 text="Kolejność wysyłki: najpierw zadania mieszkańca, potem zadania pracownika "
                      "(automatycznie połączone z odpowiednim zadaniem mieszkańca przy tworzeniu).",
                 bootstyle="secondary", font=FONT_TINY, wraplength=900, justify="left").pack(anchor="w", pady=(8, 0))

        sec_go = ttkb.Labelframe(outer, text=" Podgląd i wysyłka ", bootstyle=ACCENT_STYLE, padding=10)
        sec_go.pack(fill="both", expand=True)
        gob = ttkb.Frame(sec_go); gob.pack(fill="x")
        ttkb.Button(gob, text="① 🔍 Pokaż podgląd", command=self._show_preview,
                    bootstyle="secondary", padding=(14, 6)).pack(side="left")
        send_style = "danger" if IS_PROD else "success"
        self._btn_send = ttkb.Button(gob, text="②  ▶  Wyślij rodzaje zadań", command=self._do_send,
                  bootstyle=send_style, padding=(14, 6), state="disabled")
        self._btn_send.pack(side="left", padx=8)
        self._progress = ttkb.Progressbar(gob, length=240, mode="determinate", bootstyle=f"{ACCENT_STYLE}-striped")
        self._progress.pack(side="left", padx=12)
        self._lbl_progress = ttkb.Label(gob, text="", font=FONT_SMALL)
        self._lbl_progress.pack(side="left")
        ttkb.Label(sec_go, text="Kolejność: najpierw \"Pokaż podgląd\" (nic nie wysyła) — dopiero wtedy odblokuje się \"Wyślij rodzaje zadań\".",
                 bootstyle="secondary", font=FONT_TINY).pack(anchor="w", pady=(4, 0))

        self._preview_text = ttkb.ScrolledText(
            sec_go, height=14, font=("Consolas", 9), bg="#f7f7f7", fg="#222",
            state="disabled", wrap="word", auto_hide=True)
        self._preview_text.pack(fill="both", expand=True, pady=(8, 0))

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _log_msg(self, msg, tag="inf"):
        from datetime import datetime as _dt
        ts = _dt.now().strftime("%H:%M:%S")
        self._log.configure(state="normal")
        self._log.insert("end", f"[{ts}] {msg}\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")
        logging.info(msg)

    def _bg(self, fn): return threading.Thread(target=fn, daemon=True).start()

    def _set_preview_text(self, text):
        self._preview_text.configure(state="normal")
        self._preview_text.delete("1.0", "end")
        self._preview_text.insert("end", text)
        self._preview_text.configure(state="disabled")

    # ── LOGOWANIE ─────────────────────────────────────────────────────────────
    def _do_login(self):
        user, pw = self._e_user.get().strip(), self._e_pass.get()
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Logowanie", "Organizacja musi być liczbą (ID)."); return
        if not user or not pw:
            messagebox.showwarning("Logowanie", "Podaj login i hasło"); return
        self._lbl_auth.config(text="Logowanie...", bootstyle="secondary")

        def _t():
            ok = self.client.login(user, pw)
            if not ok:
                self.after(0, lambda: self._lbl_auth.config(text="Błąd logowania", bootstyle="danger"))
                self.after(0, lambda: self._log_msg("Błąd logowania — sprawdź login/hasło", "err"))
                return
            _save_last_login(user)
            if not self.client.select_organization(org_id):
                self.after(0, lambda: self._lbl_auth.config(
                    text=f"Zalogowano, ale nie udało się przełączyć na org {org_id}", bootstyle="danger"))
                self.after(0, lambda: self._log_msg(f"Nie udało się przełączyć na organizationId={org_id}", "err"))
                return
            self.after(0, lambda: self._lbl_auth.config(text="✓ Zalogowano", bootstyle="success"))
            self.after(0, lambda: self._lbl_status.config(text=f"●  Zalogowano (org {org_id})"))
            self.after(0, lambda: self._log_msg(f"Zalogowano jako {user}, organizacja {org_id}", "ok"))
            self.after(0, lambda: self._nb.select(1))
        self._bg(_t)

    # ── PLIK EXCEL ────────────────────────────────────────────────────────────
    def _choose_xlsx(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik Excel", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Wszystkie pliki", "*.*")])
        if not path:
            return
        self._xlsx_path = path
        self._lbl_xlsx.config(text=os.path.basename(path), bootstyle="success")
        try:
            sheets = list_sheet_names(path)
        except Exception as e:
            messagebox.showerror("Excel", f"Nie udało się otworzyć pliku:\n{e}")
            self._log_msg(f"Błąd otwierania {path}: {e}", "err")
            return
        self._cb_sheet.configure(values=sheets)
        if DEFAULT_SHEET in sheets:
            self._sheet_var.set(DEFAULT_SHEET)
        elif sheets:
            self._sheet_var.set(sheets[0])
        self._log_msg(f"Wybrano plik: {path} (arkusze: {', '.join(sheets)})", "inf")

    # ── PODGLĄD ───────────────────────────────────────────────────────────────
    def _show_preview(self):
        if not self._xlsx_path:
            messagebox.showwarning("Podgląd", "Najpierw wybierz plik Excel."); return
        sheet = self._sheet_var.get().strip()
        if not sheet:
            messagebox.showwarning("Podgląd", "Wybierz arkusz."); return

        self._btn_send.config(state="disabled")
        self._set_preview_text("Wczytywanie Excela i słowników...")

        def _t():
            rows, err = read_tasks_from_excel(self._xlsx_path, sheet)
            if err:
                self.after(0, lambda: messagebox.showerror("Excel", err))
                self.after(0, lambda: self._set_preview_text(f"Błąd: {err}"))
                self.after(0, lambda: self._log_msg(f"Błąd odczytu Excela: {err}", "err"))
                return
            if not rows:
                self.after(0, lambda: self._set_preview_text("Brak zadań do wprowadzenia w tym arkuszu."))
                return

            structure = fetch_structure_elements(self.client, ZADANIA_DICTIONARY_ID)
            if not structure:
                self.after(0, lambda: self._set_preview_text(
                    f"Błąd: nie udało się pobrać struktury słownika {ZADANIA_DICTIONARY_ID}."))
                return

            existing = fetch_dictionary_values(self.client, ZADANIA_DICTIONARY_ID)
            existing_by_content = {(v.get("content") or "").strip().lower(): v.get("id") for v in existing}

            resident_to_create, employee_to_create, skipped = plan_tasks(rows, existing_by_content)

            # walidacja "na sucho" -- sprawdź, czy da się zbudować atrybuty (bez wysyłki),
            # żeby pokazać ewentualne błędy w podglądzie zamiast dopiero przy wysyłce.
            dict_value_cache = {}
            errors = []
            fuzzy_notes = []
            for row in resident_to_create:
                _, e, notes = resolve_task_value_attributes(self.client, structure, row, "mieszkaniec", dict_value_cache)
                if e:
                    errors.append(f"{row['zadanie_mieszkaniec']}: {e}")
                fuzzy_notes.extend(notes)
            for row in employee_to_create:
                fake_subordinate = 1 if row["zadanie_mieszkaniec"] else None
                _, e, notes = resolve_task_value_attributes(self.client, structure, row, "pracownik", dict_value_cache,
                                                             subordinate_id=fake_subordinate)
                if e:
                    errors.append(f"{row['zadanie_pracownik']}: {e}")
                fuzzy_notes.extend(notes)

            self.after(0, lambda: self._render_preview(
                structure, existing_by_content, len(existing),
                resident_to_create, employee_to_create, skipped, errors, fuzzy_notes))
        self._bg(_t)

    def _render_preview(self, structure, existing_by_content, existing_count,
                        resident_to_create, employee_to_create, skipped, errors, fuzzy_notes=None):
        self._structure_elements = structure
        self._existing_by_content = existing_by_content
        self._next_display_order = existing_count
        self._resident_to_create = resident_to_create
        self._employee_to_create = employee_to_create

        lines = []
        lines.append(f"Zadania mieszkańca do utworzenia: {len(resident_to_create)}")
        for row in resident_to_create:
            lines.append(f"   [wiersz {row['row']}] {row['zadanie_mieszkaniec']}  "
                         f"(kategoria: {row['kategoria_zadania_mieszkaniec']}, usługa: {row['usluga_mieszkaniec']})")
        lines.append("")
        lines.append(f"Zadania pracownika do utworzenia: {len(employee_to_create)}")
        for row in employee_to_create:
            link = f"  -> podrzędne: {row['zadanie_mieszkaniec']}" if row["zadanie_mieszkaniec"] else ""
            lines.append(f"   [wiersz {row['row']}] {row['zadanie_pracownik']}  "
                         f"(kategoria: {row['kategoria_zadania_pracownik']}, usługa: {row['usluga_pracownik']}){link}")
        if skipped:
            lines.append("")
            lines.append(f"⚠ Pomijam (już istnieją): {', '.join(skipped)}")
        fuzzy_notes = sorted(set(fuzzy_notes or []))
        if fuzzy_notes:
            lines.append("")
            lines.append("⚠ Dopasowania przybliżone (literówka w Excelu?) — SPRAWDŹ zanim wyślesz:")
            for n in fuzzy_notes:
                lines.append(f"   {n}")
        if errors:
            lines.append("")
            lines.append("✗ Błędy (te wiersze NIE zostaną utworzone):")
            for e in errors:
                lines.append(f"   {e}")
        lines.append("")
        lines.append("─" * 70)
        lines.append(f"RAZEM: {len(resident_to_create) + len(employee_to_create)} rodzajów zadań do utworzenia")
        self._set_preview_text("\n".join(lines))
        self._log_msg(f"Podgląd: {len(resident_to_create)} zadań mieszkańca, {len(employee_to_create)} "
                      f"zadań pracownika, {len(skipped)} pominiętych, {len(fuzzy_notes)} dopasowań "
                      f"przybliżonych, {len(errors)} błędów", "inf")

        total = len(resident_to_create) + len(employee_to_create)
        # Błędne wiersze i tak są pomijane pojedynczo w _do_send (patrz log
        # "✗ ... — <błąd>"), więc kilka błędnych wierszy (np. nierozpoznany
        # czas trwania) nie powinno blokować wysyłki reszty poprawnych zadań.
        self._btn_send.config(state="normal" if total else "disabled")

    # ── WYSYŁKA ───────────────────────────────────────────────────────────────
    def _do_send(self):
        total = len(self._resident_to_create) + len(self._employee_to_create)
        if not total:
            messagebox.showwarning("Wysyłka", "Najpierw kliknij 'Pokaż podgląd'.")
            return
        warn_prefix = "⚠ PRODUKCJA — rodzaje zadań trafią od razu do systemu.\n\n" if IS_PROD else ""
        if not messagebox.askyesno("Potwierdzenie wysyłki",
            f"{warn_prefix}Dodać {total} rodzajów zadań do systemu "
            f"({len(self._resident_to_create)} mieszkańca + {len(self._employee_to_create)} pracownika)?\n"
            f"Tej operacji nie można cofnąć hurtowo."):
            return

        resident_to_create = list(self._resident_to_create)
        employee_to_create = list(self._employee_to_create)
        structure = self._structure_elements
        display_order = self._next_display_order
        resident_id_by_content = dict(self._existing_by_content)

        self._btn_send.config(state="disabled")
        self._progress.configure(maximum=total, value=0)
        self._lbl_progress.config(text="Wysyłanie...")

        def _t():
            nonlocal display_order
            ok = fail = 0
            step = 0
            dict_value_cache = {}

            for row in resident_to_create:
                step += 1
                attrs, err, notes = resolve_task_value_attributes(
                    self.client, structure, row, "mieszkaniec", dict_value_cache)
                name = row["zadanie_mieszkaniec"]
                for note in notes:
                    self.after(0, lambda note=note: self._log_msg(f"  ⚠ {note}", "inf"))
                if err:
                    fail += 1
                    self.after(0, lambda name=name, err=err: self._log_msg(f"✗ {name} — {err}", "err"))
                else:
                    new_id, err2 = post_task_kind(self.client, name, display_order, attrs)
                    display_order += 1
                    if new_id is None:
                        fail += 1
                        self.after(0, lambda name=name, err2=err2: self._log_msg(f"✗ {name} — {err2}", "err"))
                    else:
                        ok += 1
                        resident_id_by_content[name.lower()] = new_id
                        self.after(0, lambda name=name: self._log_msg(f"✓ {name} — utworzono", "ok"))
                self.after(0, lambda step=step, ok=ok, fail=fail: (
                    self._progress.configure(value=step),
                    self._lbl_progress.config(text=f"{step}/{total}  ✓ {ok}  ✗ {fail}")))

            for row in employee_to_create:
                step += 1
                name = row["zadanie_pracownik"]
                subordinate_id = None
                if row["zadanie_mieszkaniec"]:
                    subordinate_id = resident_id_by_content.get(row["zadanie_mieszkaniec"].lower())
                    if subordinate_id is None:
                        fail += 1
                        self.after(0, lambda name=name, dep=row["zadanie_mieszkaniec"]: self._log_msg(
                            f"✗ {name} — brak utworzonego zadania mieszkańca '{dep}' (nie powiodło się wcześniej?)", "err"))
                        self.after(0, lambda step=step, ok=ok, fail=fail: (
                            self._progress.configure(value=step),
                            self._lbl_progress.config(text=f"{step}/{total}  ✓ {ok}  ✗ {fail}")))
                        continue
                attrs, err, notes = resolve_task_value_attributes(
                    self.client, structure, row, "pracownik", dict_value_cache, subordinate_id=subordinate_id)
                for note in notes:
                    self.after(0, lambda note=note: self._log_msg(f"  ⚠ {note}", "inf"))
                if err:
                    fail += 1
                    self.after(0, lambda name=name, err=err: self._log_msg(f"✗ {name} — {err}", "err"))
                else:
                    new_id, err2 = post_task_kind(self.client, name, display_order, attrs)
                    display_order += 1
                    if new_id is None:
                        fail += 1
                        self.after(0, lambda name=name, err2=err2: self._log_msg(f"✗ {name} — {err2}", "err"))
                    else:
                        ok += 1
                        self.after(0, lambda name=name: self._log_msg(f"✓ {name} — utworzono", "ok"))
                self.after(0, lambda step=step, ok=ok, fail=fail: (
                    self._progress.configure(value=step),
                    self._lbl_progress.config(text=f"{step}/{total}  ✓ {ok}  ✗ {fail}")))

            summary = f"Zakończono: {ok}/{total} utworzono, {fail} błędów"
            self.after(0, lambda: self._log_msg(f"— {summary} —", "ok" if fail == 0 else "err"))
            self.after(0, lambda: messagebox.showinfo("Gotowe", summary))
            self._resident_to_create = []
            self._employee_to_create = []
        self._bg(_t)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception as e:
        logging.error(f"Nieobsłużony wyjątek przy starcie okna: {e}\n{traceback.format_exc()}")
        _fatal_startup_error("Błąd uruchomienia okna", str(e), exc=e)
