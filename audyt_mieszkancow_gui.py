#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Syrena — Audyt i eksport: mieszkańcy / pracownicy / zadania (GUI)
Domyślnie config_testowy.json obok skryptu; inny plik jako argument CLI
(np. `python audyt_mieszkancow_gui.py config_produkcja.json` dla produkcji).
Używa tego samego pliku config co task_assign_gui.py (auth_url,
employee_api_url, organization_id, env_label, shift_url) — jeśli brakuje w
nim beneficiary_url/shift_url, są wyliczane z auth_url (:5010 -> :5020 /
:5070).

Logika (ten sam mechanizm co audyt_mieszkancow.py, w wersji CLI):
  Mieszkańcy:  GET :5020/api/beneficiary/by-organization-id/paged?statusList=1
  Audyt (mieszkaniec/pracownik): GET :5000/api/audit/<beneficiary|employee>/{id}/paged?page=&pageSize=&changeFrom=...

Zakładki "Pracownicy" i "Zadania" (dodane 2026-09-15, na podstawie ruchu
przechwyconego w przeglądarce -- patrz notatki przy poszczególnych funkcjach
fetch_* poniżej co jest POTWIERDZONE, a co jest założeniem do zweryfikowania):
  Pracownicy (aktywni):  GET :5000/api/employee/by-organization-id/paged
    -- NIEPOTWIERDZONE wprost (analogia do .../deleted/... i do zadań, patrz niżej)
  Pracownicy (usunięci): GET :5000/api/employee/deleted/by-organization-id/paged
  Zadania (aktywne):     GET :5070/api/task/by-organization-id
    -- w przechwyconym ruchu widziane TYLKO z wąskim search=<id>&searchFields=id
       (odpowiedź to goła lista, nie {results,...}) -- stronicowanie dla pełnej
       listy NIE jest potwierdzone, patrz ostrzeżenie w fetch_active_tasks.
  Zadania (usunięte):    GET :5070/api/task/deleted/by-organization-id/paged

Zakres dat: "Od" idzie do API jako changeFrom (zmniejsza ilość pobieranych
danych). "Do" NIE jest wspieranym parametrem API (nie ma potwierdzonego
changeTo) — filtrowane lokalnie po pobraniu, po polu changeTime.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading, requests, json, os, sys, time, traceback, logging, calendar
from datetime import datetime, date as _date, timedelta

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("pip install openpyxl"); sys.exit(1)

# ── Konfiguracja ─────────────────────────────────────────────────────────────
SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
CONFIG_NAME     = sys.argv[1] if len(sys.argv) > 1 else "config_testowy.json"
CONFIG_PATH     = os.path.join(SCRIPT_DIR, CONFIG_NAME)
CRASH_LOG_PATH  = os.path.join(SCRIPT_DIR, "audyt_gui_crash.log")
LAST_LOGIN_PATH = os.path.join(SCRIPT_DIR, "audyt_gui_last_login.json")


def _fatal_startup_error(title, message, exc=None):
    """Okno się otwiera i od razu zamyka, gdy wyjątek poleci PRZED mainloop()
    (np. przy dwukliku bez konsoli). Zapisuje traceback do pliku obok
    skryptu i pokazuje okienko zamiast cichego zamknięcia."""
    detail = f"{message}\n\n{traceback.format_exc()}" if exc else message
    try:
        with open(CRASH_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {title}\n{detail}\n")
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
    LOG_PATH = os.path.join(SCRIPT_DIR, CFG.get("log_file", "audyt_gui.log"))
    logging.basicConfig(filename=LOG_PATH, level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s",
                        encoding="utf-8")

    AUTH_URL        = CFG["auth_url"]
    EMP_URL         = CFG.get("employee_api_url", AUTH_URL.replace(":5010", ":5000"))
    BENEFICIARY_URL = CFG.get("beneficiary_url", AUTH_URL.replace(":5010", ":5020"))
    TASK_URL        = CFG.get("shift_url", AUTH_URL.replace(":5010", ":5070"))
    ORG_ID_DEFAULT  = CFG.get("organization_id", 1)
    ENV_LABEL       = CFG.get("env_label", "TEST")
except KeyError as e:
    _fatal_startup_error(
        f"Brak klucza w {CONFIG_NAME}",
        f"W pliku konfiguracyjnym brakuje wymaganego klucza: {e}\n"
        f"Wymagane: auth_url.", exc=e)


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


# ── Kolory i typografia (ten sam schemat co task_assign_gui.py) ──────────────
IS_PROD = ENV_LABEL.strip().upper() in ("PROD", "PRODUKCJA", "PRODUCTION")
if ENV_LABEL == "TEST":
    ACCENT = "#e8734a"
elif IS_PROD:
    ACCENT = "#c0392b"
else:
    ACCENT = "#2c5f8a"

ACCENT_DARK  = "#1f2937"
BG           = "#fff8f0" if ENV_LABEL == "TEST" else ("#fdf2f1" if IS_PROD else "#f4f6f8")
WHITE        = "#ffffff"
CARD_BORDER  = "#e2e5ea"
TXT_MUTED    = "#6b7280"
SUCCESS_CLR  = "#2a7a3a"
DANGER_CLR   = "#c0392b"
NEUTRAL_BTN  = "#4a7da8"

FONT_TITLE   = ("Segoe UI", 12, "bold")
FONT_SECTION = ("Segoe UI", 10, "bold")
FONT_BASE    = ("Segoe UI", 10)
FONT_SMALL   = ("Segoe UI", 9)
FONT_TINY    = ("Segoe UI", 8)


def _lighten(hexcolor, amount=0.15):
    hexcolor = hexcolor.lstrip("#")
    r, g, b = int(hexcolor[0:2], 16), int(hexcolor[2:4], 16), int(hexcolor[4:6], 16)
    r = min(255, int(r + (255 - r) * amount))
    g = min(255, int(g + (255 - g) * amount))
    b = min(255, int(b + (255 - b) * amount))
    return f"#{r:02x}{g:02x}{b:02x}"


def _add_hover(widget, base_color, hover_amount=0.15):
    hover_color = _lighten(base_color, hover_amount)
    widget.bind("<Enter>", lambda _e: widget.configure(bg=hover_color))
    widget.bind("<Leave>", lambda _e: widget.configure(bg=base_color))


def _card(parent, bg):
    return tk.Frame(parent, bg=bg, highlightbackground=CARD_BORDER,
                     highlightthickness=1, bd=0)


# ── Backend (ten sam mechanizm co audyt_mieszkancow.py) ──────────────────────
TIMEOUT             = 20
AUDIT_TIMEOUT        = 60
AUDIT_TIMEOUT_RETRY  = 120
PAGE_SIZE_DEFAULT    = 100
MAX_PAGES_DEFAULT    = 50
CHANGE_TYPE_LABELS   = {1: "Dodano", 2: "Zmieniono"}


class Client:
    def __init__(self):
        self.servers = {"auth": AUTH_URL, "employee": EMP_URL, "beneficiary": BENEFICIARY_URL, "task": TASK_URL}
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

    def _get(self, url, params=None, _retry=True, timeout=None):
        timeout = timeout or TIMEOUT
        logging.debug(f"GET {url} params={params} timeout={timeout}")
        try:
            r = self.s.get(url, params=params, timeout=timeout)
        except requests.exceptions.Timeout as e:
            logging.error(f"Timeout przy GET {url} (limit={timeout}s): {e}")
            return None, f"TIMEOUT:{e}"
        except Exception as e:
            logging.error(f"Wyjątek przy GET {url}: {e}")
            return None, str(e)
        if r.status_code == 401 and _retry and self._creds:
            logging.warning("Token wygasł (401) — ponowne logowanie")
            if self.login(*self._creds):
                if self._current_org_id is not None:
                    self.select_organization(self._current_org_id)
                return self._get(url, params=params, _retry=False, timeout=timeout)
        return r, None


def fetch_active_beneficiaries(client, org_id, status_list="1"):
    url = f"{client.servers['beneficiary']}/api/beneficiary/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id,
                  "statusList": status_list}
        r, err = client._get(url, params=params)
        if err or r is None:
            logging.error(f"Błąd pobierania mieszkańców org={org_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logging.error(f"Mieszkańcy org={org_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Mieszkańcy org={org_id}: odpowiedź nie jest JSON-em: {e}")
            break
        items = data.get("results") or data.get("items") or data.get("data") or []
        all_items.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    logging.info(f"Organizacja {org_id}: {len(all_items)} aktywnych mieszkańców (statusList={status_list})")
    return all_items


def fetch_beneficiary_audit(client, beneficiary_id, date_from=None,
                             page_size=PAGE_SIZE_DEFAULT, max_pages=MAX_PAGES_DEFAULT):
    url = f"{client.servers['employee']}/api/audit/beneficiary/{beneficiary_id}/paged"
    page, all_events = 1, []
    while page <= max_pages:
        params = {"page": page, "pageSize": page_size}
        if date_from:
            params["changeFrom"] = f"{date_from}T00:00:00.000Z"
        r, err = client._get(url, params=params, timeout=AUDIT_TIMEOUT)
        if err and err.startswith("TIMEOUT:"):
            logging.warning(f"Audyt mieszkańca {beneficiary_id} str.{page}: timeout po {AUDIT_TIMEOUT}s "
                            f"— ponawiam z limitem {AUDIT_TIMEOUT_RETRY}s...")
            r, err = client._get(url, params=params, timeout=AUDIT_TIMEOUT_RETRY)
        if err or r is None:
            logging.error(f"Audyt mieszkańca {beneficiary_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logging.error(f"Audyt mieszkańca {beneficiary_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Audyt mieszkańca {beneficiary_id}: odpowiedź nie jest JSON-em: {e}")
            break
        items = data.get("results") or data.get("items") or data.get("data") or []
        all_events.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    else:
        logging.warning(f"Audyt mieszkańca {beneficiary_id}: osiągnięto limit {max_pages} stron.")
    if not all_events:
        logging.warning(f"Audyt mieszkańca {beneficiary_id}: brak zdarzeń audytu")
    return all_events


# ── Pracownicy ────────────────────────────────────────────────────────────────
def fetch_active_employees(client, org_id):
    """POTWIERDZONE tylko pośrednio: przechwycony ruch pokazał
    .../employee/deleted/by-organization-id/paged (patrz fetch_deleted_employees)
    -- ten (bez 'deleted') zakłada analogiczny adres przez analogię do pary
    /api/task/by-organization-id + /api/task/deleted/by-organization-id/paged,
    która JEST potwierdzona w tym samym ruchu. Jeśli się myli, HTTP != 200
    zostanie zalogowany, nie ukryty."""
    url = f"{client.servers['employee']}/api/employee/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id}
        r, err = client._get(url, params=params)
        if err or r is None:
            logging.error(f"Błąd pobierania pracowników org={org_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logging.error(f"Pracownicy org={org_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Pracownicy org={org_id}: odpowiedź nie jest JSON-em: {e}")
            break
        items = data.get("results") if isinstance(data, dict) else data
        items = items or []
        all_items.extend(items)
        total_pages = data.get("totalNumberOfPages") if isinstance(data, dict) else None
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    logging.info(f"Organizacja {org_id}: {len(all_items)} aktywnych pracowników")
    return all_items


def fetch_deleted_employees(client, org_id):
    """POTWIERDZONE przechwyconym ruchem: GET .../employee/deleted/by-organization-id/paged"""
    url = f"{client.servers['employee']}/api/employee/deleted/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id,
                  "orderBy": "id", "ascending": "true"}
        r, err = client._get(url, params=params)
        if err or r is None:
            logging.error(f"Błąd pobierania usuniętych pracowników org={org_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logging.error(f"Usunięci pracownicy org={org_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Usunięci pracownicy org={org_id}: odpowiedź nie jest JSON-em: {e}")
            break
        items = data.get("results") or []
        all_items.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    logging.info(f"Organizacja {org_id}: {len(all_items)} usuniętych/nieaktywnych pracowników")
    return all_items


def fetch_employee_audit(client, employee_id, date_from=None,
                         page_size=PAGE_SIZE_DEFAULT, max_pages=MAX_PAGES_DEFAULT):
    """POTWIERDZONE przechwyconym ruchem: GET :5000/api/audit/employee/{id}/paged
    -- dokładny odpowiednik fetch_beneficiary_audit."""
    url = f"{client.servers['employee']}/api/audit/employee/{employee_id}/paged"
    page, all_events = 1, []
    while page <= max_pages:
        params = {"page": page, "pageSize": page_size}
        if date_from:
            params["changeFrom"] = f"{date_from}T00:00:00.000Z"
        r, err = client._get(url, params=params, timeout=AUDIT_TIMEOUT)
        if err and err.startswith("TIMEOUT:"):
            logging.warning(f"Audyt pracownika {employee_id} str.{page}: timeout po {AUDIT_TIMEOUT}s "
                            f"— ponawiam z limitem {AUDIT_TIMEOUT_RETRY}s...")
            r, err = client._get(url, params=params, timeout=AUDIT_TIMEOUT_RETRY)
        if err or r is None:
            logging.error(f"Audyt pracownika {employee_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logging.error(f"Audyt pracownika {employee_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Audyt pracownika {employee_id}: odpowiedź nie jest JSON-em: {e}")
            break
        items = data.get("results") or data.get("items") or data.get("data") or []
        all_events.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    else:
        logging.warning(f"Audyt pracownika {employee_id}: osiągnięto limit {max_pages} stron.")
    if not all_events:
        logging.warning(f"Audyt pracownika {employee_id}: brak zdarzeń audytu")
    return all_events


def _employee_label(e):
    return f"{e.get('surname', '')} {e.get('firstName', '')}".strip() or f"#{e.get('id')}"


# ── Zadania ───────────────────────────────────────────────────────────────────
def fetch_active_tasks(client, org_id):
    """UWAGA -- NIE W PEŁNI POTWIERDZONE: w przechwyconym ruchu ten endpoint
    widziany był tylko z wąskim filtrem search=<id>&searchFields=id (jeden
    wynik), a odpowiedź to GOŁA LISTA JSON, nie {results, totalNumberOfPages}
    jak w innych 'paged' endpointach tego skryptu. Stronicowanie po
    page/pageSize dla PEŁNEJ listy zadań organizacji NIE jest potwierdzone.
    Próbujemy tych parametrów; jeśli odpowiedź to goła lista o długości
    dokładnie pageSize (może być więcej, których nie widzimy), logujemy
    WYRAŹNE ostrzeżenie zamiast cicho zwracać niepełne dane."""
    url = f"{client.servers['task']}/api/task/by-organization-id"
    page, page_size, all_items, warned = 1, 200, [], False
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id}
        r, err = client._get(url, params=params)
        if err or r is None:
            logging.error(f"Błąd pobierania zadań org={org_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logging.error(f"Zadania org={org_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Zadania org={org_id}: odpowiedź nie jest JSON-em: {e}")
            break
        if isinstance(data, list):
            if not warned:
                logging.warning(
                    "fetch_active_tasks: /api/task/by-organization-id zwrócił gołą listę (nie "
                    "{results,...}) -- stronicowanie page/pageSize dla PEŁNEJ listy zadań NIE jest "
                    "potwierdzone przechwyconym ruchem. Jeśli liczba zadań w eksporcie wygląda za "
                    "niska, zweryfikuj ręcznie w przeglądarce (Network -> /api/task/by-organization-id "
                    "bez filtra 'search').")
                warned = True
            all_items.extend(data)
            if len(data) < page_size:
                break
            page += 1
            continue
        items = data.get("results") or data.get("items") or data.get("data") or []
        all_items.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    logging.info(f"Organizacja {org_id}: {len(all_items)} aktywnych zadań")
    return all_items


def fetch_deleted_tasks(client, org_id):
    """POTWIERDZONE przechwyconym ruchem: GET .../task/deleted/by-organization-id/paged"""
    url = f"{client.servers['task']}/api/task/deleted/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id,
                  "orderBy": "id", "ascending": "true"}
        r, err = client._get(url, params=params)
        if err or r is None:
            logging.error(f"Błąd pobierania usuniętych zadań org={org_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logging.error(f"Usunięte zadania org={org_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Usunięte zadania org={org_id}: odpowiedź nie jest JSON-em: {e}")
            break
        items = data.get("results") or []
        all_items.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    logging.info(f"Organizacja {org_id}: {len(all_items)} usuniętych zadań")
    return all_items


def fetch_all_beneficiaries_any_status(client, org_id):
    """UWAGA -- ZAŁOŻENIE do zweryfikowania: pominięcie parametru statusList ma
    (zgodnie z typowym REST-owym wzorcem) zwrócić mieszkańców w KAŻDYM statusie,
    nie tylko aktywnych. Używane WYŁĄCZNIE do ustalenia w zakładce 'Zadania',
    czy mieszkaniec przypisany do zadania jest aktywny -- NIE do właściwego
    audytu mieszkańców (tam nadal fetch_active_beneficiaries ze statusList=1,
    bez zmian). Jeśli to założenie jest błędne, lista będzie identyczna z
    aktywnymi i filtr 'pomiń nieaktywnych' po prostu nic nie odfiltruje --
    nie zepsuje to eksportu, tylko nie zawęzi go tak jak oczekiwano."""
    url = f"{client.servers['beneficiary']}/api/beneficiary/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id}
        r, err = client._get(url, params=params)
        if err or r is None or r.status_code != 200:
            logging.error(f"Błąd pobierania wszystkich mieszkańców org={org_id} str.{page}: "
                          f"{err or (r and r.status_code)}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Wszyscy mieszkańcy org={org_id}: odpowiedź nie jest JSON-em: {e}")
            break
        items = data.get("results") or data.get("items") or data.get("data") or []
        all_items.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    return all_items


def _task_employee_names(t, employee_lookup=None):
    names, seen = [], set()
    top_name = (t.get("executingEmployeeName") or "").strip()
    if top_name:
        names.append(top_name); seen.add(top_name)
    for ex in t.get("taskExecutors") or []:
        n = f"{ex.get('employeeSurname', '')} {ex.get('employeeFirstName', '')}".strip()
        eid = ex.get("employeeId")
        if not n and employee_lookup is not None and eid in employee_lookup:
            n = employee_lookup[eid]
        if n and n not in seen:
            seen.add(n); names.append(n)
    return ", ".join(names)


def _task_employee_ids(t):
    ids = set()
    if t.get("executingEmployeeId"):
        ids.add(t["executingEmployeeId"])
    for ex in t.get("taskExecutors") or []:
        if ex.get("employeeId"):
            ids.add(ex["employeeId"])
    return ids


def flatten_audit_tree(events, resident_id, resident_name, org_id):
    rows = []

    def walk(node, path, top_id, inherited_time, inherited_user, inherited_ip):
        name = node.get("name") or ""
        segment = name if name else f"[#{node.get('id')}]"
        new_path = path + [segment]
        change_time = node.get("changeTime") or inherited_time
        user_name = node.get("userName") or inherited_user
        ip_address = node.get("ipAddress") or inherited_ip
        change_type = node.get("changeType")

        for pc in node.get("propertyChanges", []) or []:
            rows.append({
                "organizationId": org_id,
                "residentId": resident_id,
                "residentName": resident_name,
                "auditEventId": top_id,
                "changeTime": change_time,
                "changeType": CHANGE_TYPE_LABELS.get(change_type, f"typ {change_type}"),
                "userName": user_name,
                "ipAddress": ip_address,
                "path": " > ".join(new_path),
                "propertyName": pc.get("propertyName", ""),
                "oldValue": pc.get("oldValue", ""),
                "newValue": pc.get("newValue", ""),
            })

        for child in node.get("entities", []) or []:
            walk(child, new_path, top_id, change_time, user_name, ip_address)

    for ev in events:
        walk(ev, [], ev.get("id"), ev.get("changeTime") or "", ev.get("userName") or "", ev.get("ipAddress") or "")

    return rows


def filter_rows_by_date_to(rows, date_to):
    """'Do' nie jest wspierane przez API — filtr lokalny po changeTime (prefiks YYYY-MM-DD)."""
    if not date_to:
        return rows
    return [r for r in rows if not r.get("changeTime") or r["changeTime"][:10] <= date_to]


def _resident_label(r):
    return f"{r.get('surname', '')} {r.get('firstName', '')}".strip() or f"#{r.get('id')}"


def save_excel(out_path, all_rows, residents_summary, subject_label="Mieszkaniec",
              subject_id_label="ID mieszkańca", title_prefix="Audyt zmian mieszkańców",
              sheet2_title="Podsumowanie mieszkańców"):
    wb = openpyxl.Workbook()

    def fill(c): return PatternFill("solid", fgColor=c)
    def thin():
        s = Side(style="thin", color="D1D5DB")
        return Border(left=s, right=s, top=s, bottom=s)

    ws = wb.active
    ws.title = "Zmiany"
    cols = ["organizationId", "residentId", "residentName", "auditEventId", "changeTime",
            "changeType", "userName", "ipAddress", "path", "propertyName", "oldValue", "newValue"]
    labels = ["OrgID", subject_id_label, subject_label, "ID zdarzenia", "Data zmiany",
              "Typ", "Użytkownik", "IP", "Ścieżka", "Pole", "Stara wartość", "Nowa wartość"]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    c = ws["A1"]
    c.value = f"{title_prefix}  |  {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  {len(all_rows)} zmian"
    c.font = Font(name="Segoe UI", bold=True, size=12, color="FFFFFF")
    c.fill = fill("1A3A6B")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    widths = [8, 12, 22, 12, 20, 12, 18, 14, 45, 30, 30, 30]
    for col, (lbl, w) in enumerate(zip(labels, widths), 1):
        cell = ws.cell(row=2, column=col, value=lbl)
        cell.font = Font(name="Segoe UI", bold=True, size=10, color="FFFFFF")
        cell.fill = fill("1A3A6B")
        cell.border = thin()
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col)].width = w

    CHANGE_FILL = "F5B7B1"
    for i, row in enumerate(all_rows):
        r = 3 + i
        is_change = row.get("changeType") == "Zmieniono"
        old_v = str(row.get("oldValue") or "").strip()
        new_v = str(row.get("newValue") or "").strip()
        is_deletion = is_change and old_v and not new_v
        bg = CHANGE_FILL if is_change else ("EBF3FF" if i % 2 == 0 else "FFFFFF")
        for col, key in enumerate(cols, 1):
            v = row.get(key, "")
            cell = ws.cell(row=r, column=col, value=str(v) if v is not None else "")
            cell.fill = fill(bg)
            cell.border = thin()
            strike = is_deletion and key == "oldValue"
            cell.font = Font(name="Segoe UI", size=9, strike=strike)
            cell.alignment = Alignment(wrap_text=(key in ("path", "oldValue", "newValue")), vertical="top")

    ws.freeze_panes = "A3"
    if all_rows:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{2 + len(all_rows)}"

    ws2 = wb.create_sheet(sheet2_title)
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    c2 = ws2["A1"]
    c2.value = f"{sheet2_title}  |  {len(residents_summary)}"
    c2.font = Font(name="Segoe UI", bold=True, size=12, color="FFFFFF")
    c2.fill = fill("1A3A6B")
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    headers2 = [subject_id_label, subject_label, "Zdarzeń audytu", "Zmian pól", "Ostatnia zmiana"]
    widths2 = [12, 24, 16, 12, 20]
    for col, (lbl, w) in enumerate(zip(headers2, widths2), 1):
        cell = ws2.cell(row=2, column=col, value=lbl)
        cell.font = Font(name="Segoe UI", bold=True, size=10, color="FFFFFF")
        cell.fill = fill("1A3A6B")
        cell.border = thin()
        cell.alignment = Alignment(horizontal="center")
        ws2.column_dimensions[get_column_letter(col)].width = w

    for i, s in enumerate(sorted(residents_summary, key=lambda x: x["residentName"])):
        r = 3 + i
        bg = "EBF3FF" if i % 2 == 0 else "FFFFFF"
        vals = [s["residentId"], s["residentName"], s["eventCount"], s["changeCount"], s["lastChange"]]
        for col, v in enumerate(vals, 1):
            cell = ws2.cell(row=r, column=col, value=v)
            cell.fill = fill(bg)
            cell.border = thin()
            cell.font = Font(name="Segoe UI", size=9)
    ws2.freeze_panes = "A3"
    if residents_summary:
        ws2.auto_filter.ref = f"A2:E{2 + len(residents_summary)}"

    wb.save(out_path)


def save_tasks_excel(out_path, task_rows, resident_summary):
    wb = openpyxl.Workbook()

    def fill(c): return PatternFill("solid", fgColor=c)
    def thin():
        s = Side(style="thin", color="D1D5DB")
        return Border(left=s, right=s, top=s, bottom=s)

    active_txt = {True: "tak", False: "nie", None: "?"}

    ws = wb.active
    ws.title = "Zadania"
    cols = ["organizationId", "residentId", "residentName", "residentActive", "taskId", "taskName",
            "taskCategoryName", "employees", "taskStatus", "isDeleted", "plannedStartTime",
            "plannedFinishTime", "normativeTime"]
    labels = ["OrgID", "ID mieszkańca", "Mieszkaniec", "Mieszkaniec aktywny", "ID zadania", "Zadanie",
              "Kategoria", "Pracownicy", "Status (kod)", "Usunięte", "Planowany start",
              "Planowane zakończenie", "Czas normatywny (min)"]
    widths = [8, 12, 22, 14, 12, 42, 18, 30, 12, 10, 20, 20, 16]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    c = ws["A1"]
    c.value = f"Zadania  |  {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  {len(task_rows)} zadań"
    c.font = Font(name="Segoe UI", bold=True, size=12, color="FFFFFF")
    c.fill = fill("1A3A6B")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    for col, (lbl, w) in enumerate(zip(labels, widths), 1):
        cell = ws.cell(row=2, column=col, value=lbl)
        cell.font = Font(name="Segoe UI", bold=True, size=10, color="FFFFFF")
        cell.fill = fill("1A3A6B")
        cell.border = thin()
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col)].width = w

    DELETED_FILL = "F5B7B1"
    for i, row in enumerate(task_rows):
        r = 3 + i
        bg = DELETED_FILL if row.get("isDeleted") else ("EBF3FF" if i % 2 == 0 else "FFFFFF")
        for col, key in enumerate(cols, 1):
            v = row.get(key, "")
            if key == "residentActive":
                v = active_txt.get(v, v)
            elif key == "isDeleted":
                v = "tak" if v else "nie"
            cell = ws.cell(row=r, column=col, value=str(v) if v is not None else "")
            cell.fill = fill(bg)
            cell.border = thin()
            cell.font = Font(name="Segoe UI", size=9)
            cell.alignment = Alignment(wrap_text=(key in ("taskName", "employees")), vertical="top")

    ws.freeze_panes = "A3"
    if task_rows:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{2 + len(task_rows)}"

    ws2 = wb.create_sheet("Podsumowanie wg mieszkańca")
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    c2 = ws2["A1"]
    c2.value = f"Podsumowanie wg mieszkańca  |  {len(resident_summary)} mieszkańców"
    c2.font = Font(name="Segoe UI", bold=True, size=12, color="FFFFFF")
    c2.fill = fill("1A3A6B")
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    headers2 = ["ID mieszkańca", "Mieszkaniec", "Aktywny", "Liczba zadań"]
    widths2 = [12, 24, 10, 14]
    for col, (lbl, w) in enumerate(zip(headers2, widths2), 1):
        cell = ws2.cell(row=2, column=col, value=lbl)
        cell.font = Font(name="Segoe UI", bold=True, size=10, color="FFFFFF")
        cell.fill = fill("1A3A6B")
        cell.border = thin()
        cell.alignment = Alignment(horizontal="center")
        ws2.column_dimensions[get_column_letter(col)].width = w

    for i, s in enumerate(sorted(resident_summary, key=lambda x: x["residentName"])):
        r = 3 + i
        bg = "EBF3FF" if i % 2 == 0 else "FFFFFF"
        vals = [s["residentId"], s["residentName"], active_txt.get(s["residentActive"], s["residentActive"]),
                s["taskCount"]]
        for col, v in enumerate(vals, 1):
            cell = ws2.cell(row=r, column=col, value=v)
            cell.fill = fill(bg)
            cell.border = thin()
            cell.font = Font(name="Segoe UI", size=9)
    ws2.freeze_panes = "A3"
    if resident_summary:
        ws2.auto_filter.ref = f"A2:D{2 + len(resident_summary)}"

    wb.save(out_path)


class _CheckboxList:
    """Treeview z checkboxami (kolumny sel/id/nazwa) + stan zaznaczenia +
    filtr tekstowy -- ta sama logika była wcześniej powielona tylko dla
    mieszkańców; teraz współdzielona między zakładkami Mieszkańcy,
    Pracownicy (audyt) i Zadania (filtr po pracowniku)."""
    def __init__(self, tree, filter_entry, sel_label, label_fn):
        self.tree = tree
        self.filter_entry = filter_entry
        self.sel_label = sel_label
        self.label_fn = label_fn
        self.items = []
        self.all_rows = []
        self.checked_value_ids = set()   # ID rekordu (stabilne -- PRZETRWA filtrowanie/odświeżenie)
        self.tree.bind("<Button-1>", self._on_click)
        if filter_entry is not None:
            filter_entry.bind("<KeyRelease>", lambda _e: self.apply_filter())

    def set_items(self, items):
        self.items = sorted(items, key=lambda it: self.label_fn(it).lower())
        self.checked_value_ids = set()
        self.refresh(self.items)

    def refresh(self, items):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.all_rows = []
        for it in items:
            checked = it.get("id") in self.checked_value_ids
            iid = self.tree.insert("", "end", values=("☑" if checked else "☐", it.get("id", ""), self.label_fn(it)),
                                   tags=("checked",) if checked else ("unchecked",))
            self.all_rows.append((iid, it))
        self._update_label()

    def apply_filter(self):
        q = self.filter_entry.get().strip().lower() if self.filter_entry else ""
        filtered = [it for it in self.items if q in self.label_fn(it).lower()] if q else self.items
        self.refresh(filtered)

    def _on_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid or self.tree.identify_region(event.x, event.y) != "cell":
            return
        it = next((it for row_iid, it in self.all_rows if row_iid == iid), None)
        if it is None:
            return
        vid = it.get("id")
        vals = list(self.tree.item(iid, "values"))
        if vid in self.checked_value_ids:
            self.checked_value_ids.discard(vid); vals[0] = "☐"
            self.tree.item(iid, values=vals, tags=("unchecked",))
        else:
            self.checked_value_ids.add(vid); vals[0] = "☑"
            self.tree.item(iid, values=vals, tags=("checked",))
        self._update_label()

    def sel_all(self):
        for iid, it in self.all_rows:
            self.checked_value_ids.add(it.get("id"))
            vals = list(self.tree.item(iid, "values")); vals[0] = "☑"
            self.tree.item(iid, values=vals, tags=("checked",))
        self._update_label()

    def desel_all(self):
        for iid, it in self.all_rows:
            self.checked_value_ids.discard(it.get("id"))
            vals = list(self.tree.item(iid, "values")); vals[0] = "☐"
            self.tree.item(iid, values=vals, tags=("unchecked",))
        self._update_label()

    def _update_label(self):
        if self.sel_label is not None:
            self.sel_label.config(text=f"{len(self.checked_value_ids)} zaznaczonych")

    def selected(self):
        by_id = {it.get("id"): it for it in self.items}
        return [by_id[vid] for vid in self.checked_value_ids if vid in by_id]


# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Syrena — Audyt i eksport: mieszkańcy / pracownicy / zadania [{ENV_LABEL}]")
        self.geometry("1050x820")
        self.resizable(True, True)
        self.configure(bg=BG)

        self.client = Client()
        self._res_list = None          # _CheckboxList mieszkańców (zakładka 2)
        self._emp_list = None          # _CheckboxList pracowników do audytu (zakładka 3)
        self._task_emp_list = None     # _CheckboxList pracowników - filtr zadań (zakładka 4)

        self._build_ui()

    # ── BUILD UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", font=FONT_BASE, padding=(16, 8))
        style.map("TNotebook.Tab", background=[("selected", WHITE)],
                  foreground=[("selected", ACCENT_DARK)])

        env_icon = "⚠ " if IS_PROD else ""
        hdr = tk.Frame(self, bg=ACCENT, height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="  Syrena — Audyt i eksport",
                 bg=ACCENT, fg="white", font=FONT_TITLE).pack(side="left", padx=(12, 4))
        tk.Label(hdr, text=f"{env_icon}[{ENV_LABEL}]",
                 bg=ACCENT, fg="white" if not IS_PROD else "#ffe9e6",
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        self._lbl_status = tk.Label(hdr, text="●  Niezalogowany",
                                    bg=ACCENT, fg="#ffddcc", font=FONT_SMALL)
        self._lbl_status.pack(side="right", padx=12)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=8)
        self._nb = nb

        t1 = tk.Frame(nb, bg=WHITE); nb.add(t1, text="  1. Logowanie  ")
        t2 = tk.Frame(nb, bg=WHITE); nb.add(t2, text="  2. Mieszkańcy  ")
        t3 = tk.Frame(nb, bg=WHITE); nb.add(t3, text="  3. Pracownicy  ")
        t4 = tk.Frame(nb, bg=WHITE); nb.add(t4, text="  4. Zadania  ")

        self._tab_login(t1)
        self._tab_residents(t2)
        self._tab_employees(t3)
        self._tab_tasks(t4)

        lf = tk.LabelFrame(self, text=" Log ", bg=BG, fg=ACCENT_DARK, font=FONT_SECTION)
        lf.pack(fill="x", padx=10, pady=(0, 8))
        self._log = scrolledtext.ScrolledText(
            lf, height=6, font=("Consolas", 9),
            bg="#1e1e1e", fg="#ccc", insertbackground="white",
            state="disabled", wrap="word")
        self._log.pack(fill="x", padx=6, pady=4)
        self._log.tag_config("ok",  foreground="#6dbf67")
        self._log.tag_config("err", foreground="#e06c6c")
        self._log.tag_config("inf", foreground="#9cdcfe")

    # ── TAB 1: Logowanie ──────────────────────────────────────────────────────
    def _tab_login(self, p):
        outer = tk.Frame(p, bg=p["bg"], padx=24, pady=20)
        outer.pack(anchor="nw", fill="x")

        tk.Label(outer, text="Zaloguj się do API Syrena",
                 bg=p["bg"], fg=ACCENT_DARK, font=FONT_TITLE).pack(anchor="w", pady=(0, 4))
        tk.Label(outer, text=f"Środowisko: {ENV_LABEL}   •   Host: {AUTH_URL}",
                 bg=p["bg"], fg=TXT_MUTED, font=FONT_TINY).pack(anchor="w", pady=(0, 14))

        card = _card(outer, WHITE)
        card.pack(anchor="w")
        f = tk.Frame(card, bg=WHITE, padx=20, pady=18)
        f.pack()

        for row, (lbl, attr, kw) in enumerate([
            ("Login:", "_e_user", {}),
            ("Hasło:", "_e_pass", {"show": "*"}),
            ("Organizacja (ID):", "_e_org", {}),
        ]):
            tk.Label(f, text=lbl, bg=WHITE, font=FONT_BASE).grid(row=row, column=0, sticky="w", pady=6)
            e = tk.Entry(f, width=34, font=FONT_BASE, relief="solid",
                         highlightthickness=1, highlightbackground=CARD_BORDER,
                         highlightcolor=ACCENT, bd=1, **kw)
            e.grid(row=row, column=1, padx=12, pady=6, sticky="w")
            setattr(self, attr, e)
        self._e_user.insert(0, _load_last_login())
        self._e_org.insert(0, str(ORG_ID_DEFAULT))
        self._e_pass.bind("<Return>", lambda _: self._do_login())
        self._e_user.bind("<Return>", lambda _: self._e_pass.focus_set())

        tk.Label(f, text="Hasło nie jest nigdzie zapisywane — tylko login.",
                 bg=WHITE, fg=TXT_MUTED, font=FONT_TINY).grid(
                     row=3, column=0, columnspan=2, sticky="w", pady=(2, 12))
        bf = tk.Frame(f, bg=WHITE)
        bf.grid(row=4, column=0, columnspan=2, sticky="w")
        btn_login = tk.Button(bf, text="Zaloguj", command=self._do_login,
                  bg=ACCENT, fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), activebackground=_lighten(ACCENT),
                  activeforeground="white", padx=18, pady=7, cursor="hand2", bd=0)
        btn_login.pack(side="left")
        _add_hover(btn_login, ACCENT)
        self._lbl_auth = tk.Label(bf, text="", bg=WHITE, font=FONT_BASE)
        self._lbl_auth.pack(side="left", padx=12)

    # ── Wspólny budulec: lista z checkboxami (mieszkańcy/pracownicy) ─────────────
    def _build_checklist_ui(self, parent, section_title, load_btn_text, load_cmd,
                            name_col_title, label_fn, height=10):
        sec = tk.LabelFrame(parent, text=f" {section_title} ",
                            bg=parent["bg"], font=FONT_SECTION, fg=ACCENT_DARK, padx=10, pady=8)
        sec.pack(fill="both", expand=True, padx=12, pady=(10, 4))

        row0 = tk.Frame(sec, bg=parent["bg"])
        row0.pack(fill="x", pady=(0, 4))
        btn = tk.Button(row0, text=load_btn_text, command=load_cmd,
                        bg=NEUTRAL_BTN, fg="white", relief="flat",
                        font=FONT_SMALL, padx=10, pady=4, bd=0, cursor="hand2")
        btn.pack(side="left")
        _add_hover(btn, NEUTRAL_BTN)
        lbl_status = tk.Label(row0, text="", bg=parent["bg"], font=FONT_SMALL, fg="#555")
        lbl_status.pack(side="left", padx=8)

        flt_row = tk.Frame(sec, bg=parent["bg"])
        flt_row.pack(fill="x", pady=(2, 4))
        tk.Label(flt_row, text="Filtr (nazwisko/imię):", bg=parent["bg"], font=FONT_SMALL).pack(side="left")
        e_filter = tk.Entry(flt_row, width=40, font=FONT_SMALL)
        e_filter.pack(side="left", padx=6)

        tf = tk.Frame(sec, bg=parent["bg"])
        tf.pack(fill="both", expand=True, pady=4)
        tree = ttk.Treeview(tf, columns=("sel", "id", "name"), show="headings",
                            selectmode="none", height=height)
        tree.heading("sel", text="✓"); tree.heading("id", text="ID"); tree.heading("name", text=name_col_title)
        tree.column("sel", width=28, stretch=False, anchor="center")
        tree.column("id", width=80, stretch=False, anchor="center")
        tree.column("name", width=320, stretch=True)
        tree.tag_configure("checked", background="#e8f5e9")
        tree.tag_configure("unchecked", background="#ffffff")
        vsb = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tf.rowconfigure(0, weight=1); tf.columnconfigure(0, weight=1)

        bot = tk.Frame(sec, bg=parent["bg"])
        bot.pack(fill="x", pady=(2, 0))
        lbl_sel = tk.Label(bot, text="0 zaznaczonych", bg=parent["bg"], font=FONT_SMALL, fg="#444")
        lbl_sel.pack(side="right")

        checklist = _CheckboxList(tree, e_filter, lbl_sel, label_fn)
        tk.Button(bot, text="Zaznacz wszystkich", command=checklist.sel_all,
                  relief="flat", font=FONT_SMALL, cursor="hand2").pack(side="left", padx=2)
        tk.Button(bot, text="Odznacz wszystkich", command=checklist.desel_all,
                  relief="flat", font=FONT_SMALL, cursor="hand2").pack(side="left", padx=2)
        tk.Button(flt_row, text="Wyczyść", relief="flat", font=FONT_SMALL, cursor="hand2",
                  command=lambda: (e_filter.delete(0, "end"), checklist.apply_filter())).pack(side="left")

        return checklist, lbl_status

    def _build_date_range_ui(self, parent, section_title="2. Zakres dat"):
        sec_date = tk.LabelFrame(parent, text=f" {section_title} ",
                                 bg=parent["bg"], font=FONT_SECTION, fg=ACCENT_DARK, padx=10, pady=8)
        sec_date.pack(fill="x", padx=12, pady=4)

        dr = tk.Frame(sec_date, bg=parent["bg"])
        dr.pack(fill="x")
        tk.Label(dr, text="Od:", bg=parent["bg"], font=FONT_SMALL).pack(side="left")
        e_from = tk.Entry(dr, width=12, font=FONT_SMALL)
        e_from.pack(side="left", padx=(4, 14))
        tk.Label(dr, text="Do:", bg=parent["bg"], font=FONT_SMALL).pack(side="left")
        e_to = tk.Entry(dr, width=12, font=FONT_SMALL)
        e_to.insert(0, _date.today().strftime("%Y-%m-%d"))
        e_to.pack(side="left", padx=(4, 14))

        for lbl, days in [("Ostatni miesiąc", 30), ("Ostatnie 3 mies.", 90),
                          ("Ostatni rok", 365), ("Cała historia", None)]:
            tk.Button(dr, text=lbl, relief="flat", font=FONT_SMALL, cursor="hand2",
                      command=lambda d=days: self._set_range(d, e_from, e_to)).pack(side="left", padx=2)

        tk.Label(sec_date,
                 text="Format: YYYY-MM-DD.  'Od' idzie do API (mniej danych do pobrania).  "
                      "'Do' NIE jest wspierane przez API — filtrowane lokalnie po pobraniu.",
                 bg=parent["bg"], fg=TXT_MUTED, font=FONT_TINY).pack(anchor="w", pady=(4, 0))
        return e_from, e_to

    def _build_go_button_ui(self, parent, text, command):
        sec_go = tk.Frame(parent, bg=parent["bg"], padx=12, pady=10)
        sec_go.pack(fill="x")
        btn_go = tk.Button(sec_go, text=text, command=command,
                  bg=SUCCESS_CLR if not IS_PROD else DANGER_CLR, fg="white", relief="flat",
                  font=("Segoe UI", 11, "bold"), padx=18, pady=8, bd=0, cursor="hand2")
        btn_go.pack(side="left")
        _add_hover(btn_go, SUCCESS_CLR if not IS_PROD else DANGER_CLR)
        progress = ttk.Progressbar(sec_go, length=300, mode="determinate")
        progress.pack(side="left", padx=16)
        lbl_progress = tk.Label(sec_go, text="", bg=parent["bg"], font=FONT_SMALL, fg="#444")
        lbl_progress.pack(side="left")
        return progress, lbl_progress

    # ── TAB 2: Mieszkańcy ─────────────────────────────────────────────────────
    def _tab_residents(self, p):
        self._res_list, self._lbl_res_status = self._build_checklist_ui(
            p, "1. Aktywni mieszkańcy", "↓ Pobierz aktywnych mieszkańców",
            self._load_residents, "Mieszkaniec", _resident_label)
        self._e_from, self._e_to = self._build_date_range_ui(p)
        self._progress, self._lbl_progress = self._build_go_button_ui(
            p, "▶  Pobierz historię zmian i zapisz Excel", self._do_fetch_and_save)

    # ── TAB 3: Pracownicy (audyt zmian) ──────────────────────────────────────
    def _tab_employees(self, p):
        self._emp_list, self._lbl_emp_status = self._build_checklist_ui(
            p, "1. Aktywni pracownicy", "↓ Pobierz aktywnych pracowników",
            self._load_employees, "Pracownik", _employee_label)
        self._v_incl_deleted_emp = tk.BooleanVar(value=False)
        tk.Checkbutton(p, text="Uwzględnij usuniętych/nieaktywnych pracowników (odśwież listę po zmianie)",
                       variable=self._v_incl_deleted_emp, bg=p["bg"],
                       font=FONT_SMALL).pack(anchor="w", padx=12)
        self._e_emp_from, self._e_emp_to = self._build_date_range_ui(p)
        self._emp_progress, self._lbl_emp_progress = self._build_go_button_ui(
            p, "▶  Pobierz historię zmian i zapisz Excel", self._do_fetch_and_save_employees)

    # ── TAB 4: Zadania ────────────────────────────────────────────────────────
    def _tab_tasks(self, p):
        self._task_emp_list, self._lbl_task_emp_status = self._build_checklist_ui(
            p, "1. Pracownicy (filtr — puste = wszyscy)", "↓ Pobierz aktywnych pracowników",
            self._load_task_employees, "Pracownik", _employee_label, height=7)
        self._v_incl_deleted_task_emp = tk.BooleanVar(value=False)
        tk.Checkbutton(p, text="Uwzględnij usuniętych/nieaktywnych pracowników w liście filtra (odśwież po zmianie)",
                       variable=self._v_incl_deleted_task_emp, bg=p["bg"],
                       font=FONT_SMALL).pack(anchor="w", padx=12)

        sec_opt = tk.LabelFrame(p, text=" 2. Opcje ",
                                bg=p["bg"], font=FONT_SECTION, fg=ACCENT_DARK, padx=10, pady=8)
        sec_opt.pack(fill="x", padx=12, pady=4)
        self._v_incl_deleted_tasks = tk.BooleanVar(value=False)
        tk.Checkbutton(sec_opt, text="Uwzględnij usunięte zadania",
                       variable=self._v_incl_deleted_tasks, bg=p["bg"],
                       font=FONT_SMALL).pack(anchor="w")
        self._v_skip_inactive_residents = tk.BooleanVar(value=False)
        tk.Checkbutton(sec_opt, text="Pomiń zadania mieszkańców nieaktywnych",
                       variable=self._v_skip_inactive_residents, bg=p["bg"],
                       font=FONT_SMALL).pack(anchor="w")
        tk.Label(sec_opt,
                 text="UWAGA: stronicowanie pełnej listy aktywnych zadań organizacji nie jest "
                      "w 100% potwierdzone (patrz log przy pobieraniu) — jeśli liczba zadań "
                      "wygląda za niska, zgłoś to i zweryfikujemy razem.",
                 bg=p["bg"], fg=TXT_MUTED, font=FONT_TINY, wraplength=760, justify="left").pack(anchor="w", pady=(4, 0))

        self._task_progress, self._lbl_task_progress = self._build_go_button_ui(
            p, "▶  Pobierz zadania i zapisz Excel", self._do_fetch_and_save_tasks)

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _log_msg(self, msg, tag="inf"):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.configure(state="normal")
        self._log.insert("end", f"[{ts}] {msg}\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")
        logging.info(msg)

    def _bg(self, fn): return threading.Thread(target=fn, daemon=True).start()

    def _set_range(self, days, e_from, e_to):
        e_to.delete(0, "end"); e_to.insert(0, _date.today().strftime("%Y-%m-%d"))
        e_from.delete(0, "end")
        if days is not None:
            e_from.insert(0, (_date.today() - timedelta(days=days)).strftime("%Y-%m-%d"))

    # ── LOGOWANIE ─────────────────────────────────────────────────────────────
    def _do_login(self):
        user, pw = self._e_user.get().strip(), self._e_pass.get()
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Logowanie", "Organizacja musi być liczbą (ID)."); return
        if not user or not pw:
            messagebox.showwarning("Logowanie", "Podaj login i hasło"); return
        self._lbl_auth.config(text="Logowanie...", fg="#888")

        def _t():
            ok = self.client.login(user, pw)
            if not ok:
                self.after(0, lambda: self._lbl_auth.config(text="Błąd logowania", fg=DANGER_CLR))
                self.after(0, lambda: self._log_msg("Błąd logowania — sprawdź login/hasło", "err"))
                return
            _save_last_login(user)
            if not self.client.select_organization(org_id):
                self.after(0, lambda: self._lbl_auth.config(
                    text=f"Zalogowano, ale nie udało się przełączyć na org {org_id}", fg=DANGER_CLR))
                self.after(0, lambda: self._log_msg(f"Nie udało się przełączyć na organizationId={org_id}", "err"))
                return
            self.after(0, lambda: self._lbl_auth.config(text="✓ Zalogowano", fg=SUCCESS_CLR))
            self.after(0, lambda: self._lbl_status.config(text=f"●  Zalogowano (org {org_id})", fg="#aaffaa"))
            self.after(0, lambda: self._log_msg(f"Zalogowano jako {user}, organizacja {org_id}", "ok"))
            self.after(0, lambda: self._nb.select(1))
        self._bg(_t)

    # ── MIESZKAŃCY ────────────────────────────────────────────────────────────
    def _load_residents(self):
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Mieszkańcy", "Organizacja musi być liczbą (ID)."); return
        self._lbl_res_status.config(text="Pobieranie...")

        def _t():
            residents = fetch_active_beneficiaries(self.client, org_id)
            self.after(0, lambda: self._res_list.set_items(residents))
            self.after(0, lambda: self._lbl_res_status.config(text=f"Pobrano {len(residents)} aktywnych"))
            self.after(0, lambda: self._log_msg(f"Pobrano {len(residents)} aktywnych mieszkańców", "ok"))
        self._bg(_t)

    # ── PRACOWNICY (zakładka 3 — audyt) ──────────────────────────────────────
    def _load_employees(self):
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Pracownicy", "Organizacja musi być liczbą (ID)."); return
        incl_deleted = self._v_incl_deleted_emp.get()
        self._lbl_emp_status.config(text="Pobieranie...")

        def _t():
            employees = fetch_active_employees(self.client, org_id)
            if incl_deleted:
                employees = employees + fetch_deleted_employees(self.client, org_id)
            self.after(0, lambda: self._emp_list.set_items(employees))
            self.after(0, lambda: self._lbl_emp_status.config(text=f"Pobrano {len(employees)}"))
            self.after(0, lambda: self._log_msg(f"Pobrano {len(employees)} pracowników "
                                                f"({'aktywni + usunięci' if incl_deleted else 'aktywni'})", "ok"))
        self._bg(_t)

    # ── PRACOWNICY (zakładka 4 — filtr do zadań) ─────────────────────────────
    def _load_task_employees(self):
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Pracownicy", "Organizacja musi być liczbą (ID)."); return
        incl_deleted = self._v_incl_deleted_task_emp.get()
        self._lbl_task_emp_status.config(text="Pobieranie...")

        def _t():
            employees = fetch_active_employees(self.client, org_id)
            if incl_deleted:
                employees = employees + fetch_deleted_employees(self.client, org_id)
            self.after(0, lambda: self._task_emp_list.set_items(employees))
            self.after(0, lambda: self._lbl_task_emp_status.config(text=f"Pobrano {len(employees)}"))
            self.after(0, lambda: self._log_msg(f"Pobrano {len(employees)} pracowników do filtra zadań "
                                                f"({'aktywni + usunięci' if incl_deleted else 'aktywni'})", "ok"))
        self._bg(_t)

    # ── POBIERZ I ZAPISZ: mieszkańcy ─────────────────────────────────────────
    def _do_fetch_and_save(self):
        selected = self._res_list.selected() if self._res_list else []
        if not selected:
            messagebox.showwarning("Audyt", "Zaznacz przynajmniej jednego mieszkańca")
            return
        date_from = self._e_from.get().strip() or None
        date_to = self._e_to.get().strip() or None
        org_id_str = self._e_org.get().strip()

        out_path = filedialog.asksaveasfilename(
            title="Zapisz jako", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="audyt_mieszkancow.xlsx")
        if not out_path:
            return

        self._progress.configure(maximum=len(selected), value=0)
        self._lbl_progress.config(text="Pobieranie...")

        def _t():
            all_rows, residents_summary = [], []
            for i, res in enumerate(selected, 1):
                rid = res.get("id")
                rname = _resident_label(res)
                self.after(0, lambda i=i, rname=rname: self._lbl_progress.config(
                    text=f"{i}/{len(selected)}  {rname}"))
                events = fetch_beneficiary_audit(self.client, rid, date_from=date_from)
                rows = flatten_audit_tree(events, rid, rname, org_id_str)
                rows = filter_rows_by_date_to(rows, date_to)
                all_rows.extend(rows)
                last_change = max((ev.get("changeTime") or "" for ev in events), default="")
                residents_summary.append({"residentId": rid, "residentName": rname,
                                          "eventCount": len(events), "changeCount": len(rows),
                                          "lastChange": last_change})
                self.after(0, lambda i=i: self._progress.configure(value=i))

            try:
                save_excel(out_path, all_rows, residents_summary)
                msg = f"Zapisano {out_path}  ({len(all_rows)} zmian, {len(selected)} mieszkańców)"
                self.after(0, lambda: self._log_msg(msg, "ok"))
                self.after(0, lambda: messagebox.showinfo("Gotowe", msg))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._log_msg(f"Błąd zapisu Excela: {err_msg}", "err"))
                self.after(0, lambda: messagebox.showerror("Błąd zapisu", err_msg))
            self.after(0, lambda: self._lbl_progress.config(text="Gotowe"))
        self._bg(_t)

    # ── POBIERZ I ZAPISZ: pracownicy ─────────────────────────────────────────
    def _do_fetch_and_save_employees(self):
        selected = self._emp_list.selected() if self._emp_list else []
        if not selected:
            messagebox.showwarning("Audyt", "Zaznacz przynajmniej jednego pracownika")
            return
        date_from = self._e_emp_from.get().strip() or None
        date_to = self._e_emp_to.get().strip() or None
        org_id_str = self._e_org.get().strip()

        out_path = filedialog.asksaveasfilename(
            title="Zapisz jako", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="audyt_pracownikow.xlsx")
        if not out_path:
            return

        self._emp_progress.configure(maximum=len(selected), value=0)
        self._lbl_emp_progress.config(text="Pobieranie...")

        def _t():
            all_rows, employees_summary = [], []
            for i, emp in enumerate(selected, 1):
                eid = emp.get("id")
                ename = _employee_label(emp)
                self.after(0, lambda i=i, ename=ename: self._lbl_emp_progress.config(
                    text=f"{i}/{len(selected)}  {ename}"))
                events = fetch_employee_audit(self.client, eid, date_from=date_from)
                rows = flatten_audit_tree(events, eid, ename, org_id_str)
                rows = filter_rows_by_date_to(rows, date_to)
                all_rows.extend(rows)
                last_change = max((ev.get("changeTime") or "" for ev in events), default="")
                employees_summary.append({"residentId": eid, "residentName": ename,
                                          "eventCount": len(events), "changeCount": len(rows),
                                          "lastChange": last_change})
                self.after(0, lambda i=i: self._emp_progress.configure(value=i))

            try:
                save_excel(out_path, all_rows, employees_summary, subject_label="Pracownik",
                          subject_id_label="ID pracownika", title_prefix="Audyt zmian pracowników",
                          sheet2_title="Podsumowanie pracowników")
                msg = f"Zapisano {out_path}  ({len(all_rows)} zmian, {len(selected)} pracowników)"
                self.after(0, lambda: self._log_msg(msg, "ok"))
                self.after(0, lambda: messagebox.showinfo("Gotowe", msg))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._log_msg(f"Błąd zapisu Excela: {err_msg}", "err"))
                self.after(0, lambda: messagebox.showerror("Błąd zapisu", err_msg))
            self.after(0, lambda: self._lbl_emp_progress.config(text="Gotowe"))
        self._bg(_t)

    # ── POBIERZ I ZAPISZ: zadania ─────────────────────────────────────────────
    def _do_fetch_and_save_tasks(self):
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Zadania", "Organizacja musi być liczbą (ID)."); return

        selected_employees = self._task_emp_list.selected() if self._task_emp_list else []
        employee_id_filter = {e.get("id") for e in selected_employees} or None
        employee_lookup = {e.get("id"): _employee_label(e) for e in selected_employees}
        include_deleted = self._v_incl_deleted_tasks.get()
        skip_inactive = self._v_skip_inactive_residents.get()

        out_path = filedialog.asksaveasfilename(
            title="Zapisz jako", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="zadania.xlsx")
        if not out_path:
            return

        self._task_progress.configure(mode="indeterminate")
        self._task_progress.start(12)
        self._lbl_task_progress.config(text="Pobieranie...")

        def _t():
            self.after(0, lambda: self._lbl_task_progress.config(text="Pobieranie zadań aktywnych..."))
            tasks = [dict(t, isDeleted=False) for t in fetch_active_tasks(self.client, org_id)]
            if include_deleted:
                self.after(0, lambda: self._lbl_task_progress.config(text="Pobieranie zadań usuniętych..."))
                tasks.extend(dict(t, isDeleted=True) for t in fetch_deleted_tasks(self.client, org_id))

            resident_active_map = {}
            if skip_inactive:
                self.after(0, lambda: self._lbl_task_progress.config(text="Ustalanie aktywności mieszkańców..."))
                active_ids = {b.get("id") for b in fetch_active_beneficiaries(self.client, org_id)}
                all_residents = fetch_all_beneficiaries_any_status(self.client, org_id)
                for b in all_residents:
                    resident_active_map[b.get("id")] = b.get("id") in active_ids
                for rid in active_ids:
                    resident_active_map.setdefault(rid, True)

            org_id_str = self._e_org.get().strip()
            task_rows, by_resident = [], {}
            for t in tasks:
                if employee_id_filter is not None and not (_task_employee_ids(t) & employee_id_filter):
                    continue
                rid = t.get("visitorId")
                rname = (t.get("visitorName") or "").strip() or "— (bez przypisanego mieszkańca)"
                ractive = resident_active_map.get(rid) if rid is not None else None
                if skip_inactive and rid is not None and ractive is False:
                    continue
                task_rows.append({
                    "organizationId": org_id_str,
                    "residentId": rid, "residentName": rname, "residentActive": ractive,
                    "taskId": t.get("id"), "taskName": t.get("name") or t.get("taskKindName") or "",
                    "taskCategoryName": t.get("taskCategoryName") or "",
                    "employees": _task_employee_names(t, employee_lookup),
                    "taskStatus": t.get("taskStatus"), "isDeleted": t.get("isDeleted", False),
                    "plannedStartTime": t.get("plannedStartTime") or "",
                    "plannedFinishTime": t.get("plannedFinishTime") or "",
                    "normativeTime": t.get("normativeTime"),
                })
                key = rid if rid is not None else rname
                if key not in by_resident:
                    by_resident[key] = {"residentId": rid, "residentName": rname,
                                        "residentActive": ractive, "taskCount": 0}
                by_resident[key]["taskCount"] += 1

            try:
                save_tasks_excel(out_path, task_rows, list(by_resident.values()))
                msg = f"Zapisano {out_path}  ({len(task_rows)} zadań, {len(by_resident)} mieszkańców)"
                self.after(0, lambda: self._log_msg(msg, "ok"))
                self.after(0, lambda: messagebox.showinfo("Gotowe", msg))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._log_msg(f"Błąd zapisu Excela: {err_msg}", "err"))
                self.after(0, lambda: messagebox.showerror("Błąd zapisu", err_msg))
            self.after(0, self._task_progress.stop)
            self.after(0, lambda: self._task_progress.configure(mode="determinate"))
            self.after(0, lambda: self._lbl_task_progress.config(text="Gotowe"))
        self._bg(_t)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception as e:
        logging.error(f"Nieobsłużony wyjątek przy starcie okna: {e}\n{traceback.format_exc()}")
        _fatal_startup_error("Błąd uruchomienia okna", str(e), exc=e)
