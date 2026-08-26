#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Syrena — Dodawanie zadań ad hoc (GUI)
Domyślnie config_testowy.json obok skryptu; inny plik jako argument CLI
(np. `python zadania_ad_hoc_gui.py config_produkcja.json` dla produkcji).
Ten sam mechanizm co zadania_ad_hoc.py (wersja CLI) — patrz tam po
szczegóły przepływu API i kształtu payloadu POST :5070/api/task.

BEZPIECZEŃSTWO: to skrypt PISZĄCY do systemu. "Pokaż podgląd" nic nie
wysyła. Wysyłka wymaga jawnego potwierdzenia w oknie dialogowym z
liczbą zadań (i dodatkowym ostrzeżeniem, gdy środowisko to PRODUKCJA).
"""

import tkinter as tk
from tkinter import messagebox
import threading, requests, json, os, sys, time, traceback, logging, calendar
from datetime import datetime, date as _date, timedelta

try:
    import ttkbootstrap as ttkb
    HAS_TTKB = True
except ImportError:
    HAS_TTKB = False

try:
    from zoneinfo import ZoneInfo
    _PL_TZ = ZoneInfo("Europe/Warsaw")
    def _dt_local(year, month, day, hour, minute=0):
        return datetime(year, month, day, hour, minute, tzinfo=_PL_TZ)
except Exception:
    from datetime import timezone as _timezone
    def _dt_local(year, month, day, hour, minute=0):
        offset = 2 if 3 < month < 10 or (month == 3 and day >= 25) or (month == 10 and day < 25) else 1
        return datetime(year, month, day, hour, minute, tzinfo=_timezone(timedelta(hours=offset)))

def _to_utc_iso(local_dt):
    from datetime import timezone as _timezone
    return local_dt.astimezone(_timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


# ── Konfiguracja ─────────────────────────────────────────────────────────────
SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
CONFIG_NAME     = sys.argv[1] if len(sys.argv) > 1 else "config_testowy.json"
CONFIG_PATH     = os.path.join(SCRIPT_DIR, CONFIG_NAME)
CRASH_LOG_PATH  = os.path.join(SCRIPT_DIR, "zadania_ad_hoc_gui_crash.log")
LAST_LOGIN_PATH = os.path.join(SCRIPT_DIR, "zadania_ad_hoc_gui_last_login.json")


def _fatal_startup_error(title, message, exc=None):
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


if not HAS_TTKB:
    _fatal_startup_error(
        "Brak biblioteki ttkbootstrap",
        "To GUI wymaga pakietu 'ttkbootstrap'.\nZainstaluj go poleceniem:\n\n"
        "    pip install ttkbootstrap")


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
    LOG_PATH = os.path.join(SCRIPT_DIR, CFG.get("log_file", "zadania_ad_hoc_gui.log"))
    logging.basicConfig(filename=LOG_PATH, level=logging.DEBUG,
                        format="%(asctime)s %(levelname)s %(message)s",
                        encoding="utf-8")

    AUTH_URL        = CFG["auth_url"]
    EMP_URL         = CFG.get("employee_api_url", AUTH_URL.replace(":5010", ":5000"))
    BENEFICIARY_URL = CFG.get("beneficiary_url", AUTH_URL.replace(":5010", ":5020"))
    SHIFT_URL       = CFG.get("shift_url", AUTH_URL.replace(":5010", ":5070"))
    ORG_ID_DEFAULT  = CFG.get("organization_id", 1)
    ENV_LABEL       = CFG.get("env_label", "TEST")
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
# Kolor środowiska wyrażony jako słowo-klucz bootstyle (zamiast ręcznie
# liczonych heksów) — "united" ma pomarańczowy jako primary i czerwony jako
# danger, więc dobrze pasuje do wcześniejszego schematu TEST=pomarańczowy,
# PRODUKCJA=czerwony tego repo.
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
    """Ramka z obramowaniem (odpowiednik dawnej ręcznej 'karty')."""
    return ttkb.Labelframe(parent, text="", bootstyle="secondary")


# ── Backend (ten sam mechanizm co zadania_ad_hoc.py) ─────────────────────────
TIMEOUT = 20
DURATION_ATTRIBUTE_KIND = 11
DEFAULT_SERVICE_NAME = "zadania ad hoc"
DEFAULT_PRIORITY = 2
DEFAULT_STATUS = 3
TASK_STATUS_LABELS = {1: "Nowe", 2: "W trakcie", 3: "Do wykonania",
                       4: "Wykonane", 5: "Anulowane", 6: "Zawieszone"}
# Mapowanie liczba->nazwa niepotwierdzone w 100% (przeniesione z innego skryptu w tym
# repo). Jeśli brakuje statusu (np. "Do przypisania"), pole Status w GUI jest edytowalne
# -- można wpisać numer bezpośrednio, zamiast być ograniczonym do tej listy.
STATUS_LABEL_TO_CODE = {v.lower(): k for k, v in TASK_STATUS_LABELS.items()}
DOW_LABELS = ["Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Nie"]


class Client:
    def __init__(self):
        self.servers = {"auth": AUTH_URL, "employee": EMP_URL,
                        "beneficiary": BENEFICIARY_URL, "shift": SHIFT_URL}
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


def fetch_active_beneficiaries(client, org_id, status_list="1"):
    url = f"{client.servers['beneficiary']}/api/beneficiary/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id, "statusList": status_list}
        r, err = client._get(url, params=params)
        if err or r is None or r.status_code != 200:
            logging.error(f"Błąd pobierania mieszkańców org={org_id} str.{page}: {err or (r and r.status_code)}")
            break
        data = r.json()
        items = data.get("results") or data.get("items") or data.get("data") or []
        all_items.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    logging.info(f"Organizacja {org_id}: {len(all_items)} aktywnych mieszkańców")
    return all_items


def _resident_label(r):
    return f"{r.get('surname', '')} {r.get('firstName', '')}".strip() or f"#{r.get('id')}"


def fetch_visitor_services(client, org_id, max_pages=50):
    url = f"{client.servers['shift']}/api/service/by-visitor-organization-id/paged"
    all_items, page, page_size = [], 1, 100
    search_fields = "name,serviceCategoryName,planRegistrationNumber,planStatus,visitorSurname,visitorFirstName,executionTime,description"
    while page <= max_pages:
        params = {"page": page, "pageSize": page_size, "orderBy": "planRegistrationNumber",
                  "ascending": "false", "searchFields": search_fields,
                  "planStatus": 1, "visitorType": 3}
        r, err = client._get(url, params=params)
        if err or r is None or r.status_code != 200:
            logging.error(f"Błąd pobierania usług str.{page}: {err or (r and r.status_code)}")
            break
        data = r.json()
        items = data.get("results") or data.get("items") or data.get("data") or []
        all_items.extend(items)
        total_pages = data.get("totalNumberOfPages")
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < page_size:
            break
        page += 1
    else:
        logging.warning(f"Usługi mieszkańców: osiągnięto limit {max_pages} stron.")
    logging.info(f"Pobrano {len(all_items)} usług mieszkańców")
    return all_items


def find_adhoc_service_id(services, resident_id, service_name=DEFAULT_SERVICE_NAME):
    for s in services:
        if s.get("visitorId") == resident_id and service_name in (s.get("serviceKindName") or "").lower():
            return s.get("id")
    return None


def fetch_all_task_kinds(client):
    r, err = client._get(f"{client.servers['auth']}/api/dictionary-value/by-dictionary-kind/53",
                         params={"withAttributes": "true"})
    if err or r is None or r.status_code != 200:
        logging.error(f"Błąd pobierania rodzajów zadań: {err or (r and r.status_code)}")
        return []
    data = r.json()
    items = data if isinstance(data, list) else (data.get("items") or data.get("data") or [])
    items.sort(key=lambda x: (x.get("content") or "").lower())
    logging.info(f"Słownik 53: {len(items)} rodzajów zadań")
    return items


def extract_default_duration(task_kind):
    for attr in task_kind.get("valueAttributes", []) or []:
        if attr.get("attributeKind") == DURATION_ATTRIBUTE_KIND:
            v = attr.get("value")
            if isinstance(v, int) and not isinstance(v, bool):
                return v
    return None


def fetch_rooms(client, org_id):
    r, err = client._get(f"{client.servers['employee']}/api/room/by-organization-id",
                         params={"organizationId": org_id})
    if err or r is None or r.status_code != 200:
        logging.error(f"Błąd pobierania sal: {err or (r and r.status_code)}")
        return []
    data = r.json()
    items = data if isinstance(data, list) else (data.get("items") or data.get("data") or [])
    logging.info(f"Pobrano {len(items)} sal/pomieszczeń")
    return items


# Kandydackie nazwy pól z ID pokoju w odpowiedzi beneficiary-residence — endpoint
# jest potwierdzony (używany w api_completeness_report.py), ale dokładny kształt
# odpowiedzi NIE jest potwierdzony, stąd przeszukiwanie kilku wariantów.
ROOM_FIELD_CANDIDATES = ("roomId", "room_id", "executionRoomId", "residenceRoomId", "currentRoomId")


def fetch_resident_room_id(client, beneficiary_id):
    """Best-effort: GET beneficiary-residence/by-beneficiary-id/{id}, szuka pola
    wyglądającego na ID pokoju. Zwraca (room_id, nazwa_pola) albo (None, None)."""
    url = f"{client.servers['beneficiary']}/api/beneficiary-residence/by-beneficiary-id/{beneficiary_id}"
    r, err = client._get(url)
    if err or r is None or r.status_code != 200:
        logging.warning(f"Mieszkaniec {beneficiary_id}: błąd pobierania zamieszkania: {err or (r and r.status_code)}")
        return None, None
    try:
        data = r.json()
    except Exception as e:
        logging.warning(f"Mieszkaniec {beneficiary_id}: odpowiedź beneficiary-residence nie jest JSON-em: {e}")
        return None, None
    if isinstance(data, list):
        data = data[0] if data else {}
    if not isinstance(data, dict):
        return None, None
    for key in ROOM_FIELD_CANDIDATES:
        v = data.get(key)
        if isinstance(v, int) and not isinstance(v, bool):
            logging.info(f"Mieszkaniec {beneficiary_id}: pokój zamieszkania = {v} (pole '{key}')")
            return v, key
    for key, v in data.items():
        if "room" in key.lower() and isinstance(v, int) and not isinstance(v, bool):
            logging.info(f"Mieszkaniec {beneficiary_id}: pokój zamieszkania = {v} (pole '{key}', dopasowanie przybliżone)")
            return v, key
    logging.warning(f"Mieszkaniec {beneficiary_id}: nie znaleziono pola z ID pokoju w beneficiary-residence. "
                    f"Dostępne pola: {list(data.keys())}")
    return None, None


def generate_dates(date_from, date_to, every_n_days=1, weekdays=None):
    d0 = datetime.strptime(date_from, "%Y-%m-%d").date()
    d1 = datetime.strptime(date_to, "%Y-%m-%d").date()
    dates, cur, i = [], d0, 0
    while cur <= d1:
        if weekdays is not None:
            if cur.weekday() in weekdays:
                dates.append(cur)
        elif i % every_n_days == 0:
            dates.append(cur)
        cur += timedelta(days=1)
        i += 1
    return dates


def build_task_payload(task_kind, service_id, room_id, occ_date, time_of_day,
                        duration_minutes, priority=DEFAULT_PRIORITY, status=DEFAULT_STATUS):
    hh, mm = [int(x) for x in time_of_day.split(":")]
    start_local = _dt_local(occ_date.year, occ_date.month, occ_date.day, hh, mm)
    finish_local = start_local + timedelta(minutes=duration_minutes)
    return {
        "id": 0, "rowVersion": 0, "isDeleted": False,
        "name": task_kind.get("content", ""),
        "description": None, "executingEmployeeId": None,
        "executionRoomId": room_id, "finishTime": None, "isEvaluated": False,
        "normativeTime": duration_minutes,
        "plannedFinishTime": _to_utc_iso(finish_local),
        "plannedStartTime": _to_utc_iso(start_local),
        "serviceId": service_id, "specialSkills": [], "startTime": None,
        "taskKindId": task_kind.get("id"), "taskPriority": priority, "taskStatus": status,
    }


def post_task(client, payload):
    r, err = client._post(f"{client.servers['shift']}/api/task", json_body=payload)
    if err:
        return False, err
    if r.status_code not in (200, 201, 204):
        try:
            detail = r.text[:300]
        except Exception:
            detail = ""
        return False, f"HTTP {r.status_code}: {detail}"
    return True, None


# ─────────────────────────────────────────────────────────────────────────────
class App(ttkb.Window):
    def __init__(self):
        super().__init__(title=f"Syrena — Zadania ad hoc [{ENV_LABEL}]",
                          themename=THEME_NAME, size=(1050, 820), resizable=(True, True))

        self.client = Client()
        self._residents = []
        self._resident_rows = []
        self._checked_ids = set()
        self._task_kinds = []
        self._selected_kind = None
        self._rooms = []
        self._selected_room = None
        self._pending_tasks = []   # [(resident, date, payload), ...] z ostatniego podglądu

        self._build_ui()

    # ── BUILD UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        env_icon = "⚠ " if IS_PROD else ""
        hdr = ttkb.Frame(self, bootstyle=f"@{ACCENT_STYLE}", height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ttkb.Label(hdr, text="  Syrena — Dodawanie zadań ad hoc", bootstyle=f"@{ACCENT_STYLE}",
                   font=FONT_TITLE).pack(side="left", padx=(12, 4))
        ttkb.Label(hdr, text=f"{env_icon}[{ENV_LABEL}]", bootstyle=f"@{ACCENT_STYLE}",
                   font=("Segoe UI", 12, "bold")).pack(side="left")
        self._lbl_status = ttkb.Label(hdr, text="●  Niezalogowany", bootstyle=f"@{ACCENT_STYLE}", font=FONT_SMALL)
        self._lbl_status.pack(side="right", padx=12)
        if IS_PROD:
            ttkb.Label(hdr, text="PRODUKCJA — zadania trafią od razu do systemu",
                       bootstyle=f"@{ACCENT_STYLE}", font=FONT_TINY).pack(side="right", padx=12)

        nb = ttkb.Notebook(self, bootstyle=ACCENT_STYLE)
        nb.pack(fill="both", expand=True, padx=10, pady=8)
        self._nb = nb

        t1 = ttkb.Frame(nb); nb.add(t1, text="  1. Logowanie  ")
        t2 = ttkb.Frame(nb); nb.add(t2, text="  2. Mieszkańcy, zadanie, sala  ")
        t3 = ttkb.Frame(nb); nb.add(t3, text="  3. Harmonogram i wysyłka  ")

        self._tab_login(t1)
        self._tab_selection(t2)
        self._tab_schedule(t3)

        lf = ttkb.Labelframe(self, text=" Log ", bootstyle="secondary")
        lf.pack(fill="x", padx=10, pady=(0, 8))
        self._log = ttkb.ScrolledText(
            lf, height=7, font=("Consolas", 9), bg="#1e1e1e", fg="#ccc",
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

    # ── TAB 2: Wybór ──────────────────────────────────────────────────────────
    def _tab_selection(self, p):
        cols_frame = ttkb.Frame(p)
        cols_frame.pack(fill="both", expand=True, padx=12, pady=10)
        cols_frame.columnconfigure(0, weight=1)
        cols_frame.columnconfigure(1, weight=1)
        cols_frame.rowconfigure(0, weight=1)

        # ── Mieszkańcy ──
        sec_res = ttkb.Labelframe(cols_frame, text=" Mieszkańcy (wielokrotny wybór) ",
                                  bootstyle=ACCENT_STYLE, padding=8)
        sec_res.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        row0 = ttkb.Frame(sec_res)
        row0.pack(fill="x", pady=(0, 4))
        btn_res = ttkb.Button(row0, text="↓ Pobierz mieszkańców", command=self._load_residents,
                  bootstyle="secondary", padding=(8, 3))
        btn_res.pack(side="left")
        self._lbl_res_status = ttkb.Label(row0, text="", font=FONT_TINY, bootstyle="secondary")
        self._lbl_res_status.pack(side="left", padx=6)

        self._e_res_filter = ttkb.Entry(sec_res, font=FONT_SMALL)
        self._e_res_filter.pack(fill="x", pady=(0, 4))
        self._e_res_filter.bind("<KeyRelease>", lambda _: self._apply_res_filter())

        tf = ttkb.Frame(sec_res)
        tf.pack(fill="both", expand=True)
        self._res_tree = ttkb.Treeview(tf, columns=("sel", "id", "name"), show="headings",
                                       selectmode="none", height=10, bootstyle=ACCENT_STYLE)
        self._res_tree.heading("sel", text="✓"); self._res_tree.heading("id", text="ID")
        self._res_tree.heading("name", text="Mieszkaniec")
        self._res_tree.column("sel", width=26, stretch=False, anchor="center")
        self._res_tree.column("id", width=60, stretch=False, anchor="center")
        self._res_tree.column("name", width=180, stretch=True)
        self._res_tree.bind("<Button-1>", self._on_res_click)
        self._res_tree.tag_configure("checked", background="#e8f5e9")
        vsb = ttkb.Scrollbar(tf, orient="vertical", command=self._res_tree.yview, bootstyle=f"{ACCENT_STYLE}-round")
        self._res_tree.configure(yscrollcommand=vsb.set)
        self._res_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tf.rowconfigure(0, weight=1); tf.columnconfigure(0, weight=1)

        bot = ttkb.Frame(sec_res)
        bot.pack(fill="x", pady=(4, 0))
        ttkb.Button(bot, text="Wszyscy", command=self._sel_all_res,
                    bootstyle="link").pack(side="left", padx=1)
        ttkb.Button(bot, text="Nikt", command=self._desel_all_res,
                    bootstyle="link").pack(side="left", padx=1)
        self._lbl_res_sel = ttkb.Label(bot, text="0 zaznaczonych", font=FONT_TINY, bootstyle="secondary")
        self._lbl_res_sel.pack(side="right")

        # ── Rodzaj zadania + Sala (prawa kolumna, jedno pod drugim) ──
        right_col = ttkb.Frame(cols_frame)
        right_col.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        right_col.rowconfigure(0, weight=1)
        right_col.rowconfigure(1, weight=1)

        sec_kind = ttkb.Labelframe(right_col, text=" Rodzaj zadania (słownik 53) — jeden ",
                                   bootstyle=ACCENT_STYLE, padding=8)
        sec_kind.pack(fill="both", expand=True, pady=(0, 6))
        row1 = ttkb.Frame(sec_kind)
        row1.pack(fill="x", pady=(0, 4))
        btn_kind = ttkb.Button(row1, text="↓ Pobierz rodzaje zadań", command=self._load_kinds,
                  bootstyle="secondary", padding=(8, 3))
        btn_kind.pack(side="left")
        self._lbl_kind_status = ttkb.Label(row1, text="", font=FONT_TINY, bootstyle="secondary")
        self._lbl_kind_status.pack(side="left", padx=6)
        self._e_kind_filter = ttkb.Entry(sec_kind, font=FONT_SMALL)
        self._e_kind_filter.pack(fill="x", pady=(0, 4))
        self._e_kind_filter.bind("<KeyRelease>", lambda _: self._apply_kind_filter())
        self._lb_kinds = ttkb.Listbox(sec_kind, font=FONT_SMALL, height=6, exportselection=False)
        self._lb_kinds.pack(fill="both", expand=True)
        self._lb_kinds.bind("<<ListboxSelect>>", self._on_kind_select)
        self._lbl_kind_sel = ttkb.Label(sec_kind, text="— nie wybrano —", font=FONT_TINY,
                                        bootstyle="secondary", wraplength=380, justify="left")
        self._lbl_kind_sel.pack(anchor="w", pady=(4, 0))

        sec_room = ttkb.Labelframe(right_col, text=" Sala/pomieszczenie ",
                                   bootstyle=ACCENT_STYLE, padding=8)
        sec_room.pack(fill="both", expand=True, pady=(6, 0))

        self._room_mode = tk.StringVar(value="fixed")
        ttkb.Radiobutton(sec_room, text="Wybierz z listy (jedna sala dla wszystkich):",
                       variable=self._room_mode, value="fixed", bootstyle=ACCENT_STYLE,
                       command=self._on_room_mode_change).pack(anchor="w")
        row2 = ttkb.Frame(sec_room)
        row2.pack(fill="x", pady=(0, 4))
        btn_room = ttkb.Button(row2, text="↓ Pobierz sale", command=self._load_rooms,
                  bootstyle="secondary", padding=(8, 3))
        btn_room.pack(side="left")
        self._lbl_room_status = ttkb.Label(row2, text="", font=FONT_TINY, bootstyle="secondary")
        self._lbl_room_status.pack(side="left", padx=6)
        self._e_room_filter = ttkb.Entry(sec_room, font=FONT_SMALL)
        self._e_room_filter.pack(fill="x", pady=(0, 4))
        self._e_room_filter.bind("<KeyRelease>", lambda _: self._apply_room_filter())
        self._lb_rooms = ttkb.Listbox(sec_room, font=FONT_SMALL, height=4, exportselection=False)
        self._lb_rooms.pack(fill="both", expand=True)
        self._lb_rooms.bind("<<ListboxSelect>>", self._on_room_select)
        self._lbl_room_sel = ttkb.Label(sec_room, text="— nie wybrano —", font=FONT_TINY, bootstyle="secondary")
        self._lbl_room_sel.pack(anchor="w", pady=(2, 6))

        ttkb.Radiobutton(sec_room, text="Pokój zamieszkania mieszkańca (automatycznie, osobno dla każdego)",
                       variable=self._room_mode, value="own", bootstyle=ACCENT_STYLE,
                       command=self._on_room_mode_change).pack(anchor="w")
        ttkb.Label(sec_room, text="   best-effort — sprawdź wynik w podglądzie przed wysyłką",
                 bootstyle="secondary", font=FONT_TINY).pack(anchor="w")
        ttkb.Radiobutton(sec_room, text="Brak sali (zadanie nie jest realizowane w pomieszczeniu)",
                       variable=self._room_mode, value="none", bootstyle=ACCENT_STYLE,
                       command=self._on_room_mode_change).pack(anchor="w", pady=(0, 4))

    # ── TAB 3: Harmonogram i wysyłka ──────────────────────────────────────────
    def _tab_schedule(self, p):
        outer = ttkb.Frame(p, padding=(16, 14))
        outer.pack(fill="both", expand=True)

        sec_date = ttkb.Labelframe(outer, text=" Zakres dat i godzina ",
                                   bootstyle=ACCENT_STYLE, padding=10)
        sec_date.pack(fill="x", pady=(0, 8))
        dr = ttkb.Frame(sec_date)
        dr.pack(fill="x")
        ttkb.Label(dr, text="Od:", font=FONT_SMALL).pack(side="left")
        self._e_from = ttkb.DateEntry(dr, width=11, date_format=r"%Y-%m-%d",
                                      bootstyle=ACCENT_STYLE, first_weekday=0)
        self._e_from.pack(side="left", padx=(4, 14))
        ttkb.Label(dr, text="Do:", font=FONT_SMALL).pack(side="left")
        self._e_to = ttkb.DateEntry(dr, width=11, date_format=r"%Y-%m-%d",
                                    bootstyle=ACCENT_STYLE, first_weekday=0,
                                    value=_date.today() + timedelta(days=30))
        self._e_to.pack(side="left", padx=(4, 14))
        ttkb.Label(dr, text="Godzina:", font=FONT_SMALL).pack(side="left")
        self._sp_hour = ttkb.Combobox(dr, values=[f"{h:02d}" for h in range(24)],
                                      state="readonly", width=3, font=FONT_SMALL,
                                      bootstyle=ACCENT_STYLE)
        self._sp_hour.set("10")
        self._sp_hour.pack(side="left", padx=(4, 0))
        ttkb.Label(dr, text=":", font=FONT_SMALL).pack(side="left")
        self._sp_minute = ttkb.Combobox(dr, values=[f"{m:02d}" for m in range(0, 60, 5)],
                                        state="readonly", width=3, font=FONT_SMALL,
                                        bootstyle=ACCENT_STYLE)
        self._sp_minute.set("00")
        self._sp_minute.pack(side="left", padx=(0, 4))

        sec_freq = ttkb.Labelframe(outer, text=" Częstotliwość — wybierz JEDEN z dwóch trybów ",
                                   bootstyle=ACCENT_STYLE, padding=10)
        sec_freq.pack(fill="x", pady=(0, 8))
        self._freq_mode = tk.StringVar(value="every_n")
        fr1 = ttkb.Frame(sec_freq); fr1.pack(fill="x", anchor="w", pady=2)
        ttkb.Radiobutton(fr1, text="Co", variable=self._freq_mode, value="every_n",
                       bootstyle=ACCENT_STYLE, command=self._on_freq_mode_change).pack(side="left")
        self._e_every_n = ttkb.Entry(fr1, width=4, font=FONT_SMALL)
        self._e_every_n.insert(0, "1")
        self._e_every_n.pack(side="left", padx=4)
        ttkb.Label(fr1, text="dni (1 = codziennie)", font=FONT_SMALL).pack(side="left")

        fr2 = ttkb.Frame(sec_freq); fr2.pack(fill="x", anchor="w", pady=(8, 2))
        self._rb_weekdays = ttkb.Radiobutton(fr2, text="Tylko wybrane dni tygodnia:", variable=self._freq_mode,
                       value="weekdays", bootstyle=ACCENT_STYLE, command=self._on_freq_mode_change)
        self._rb_weekdays.pack(side="left")
        self._dow_vars = {}
        self._dow_checks = []
        for i, lbl in enumerate(DOW_LABELS):
            v = tk.BooleanVar(value=(i < 5))
            self._dow_vars[i] = v
            cb = ttkb.Checkbutton(fr2, text=lbl, variable=v, bootstyle=f"{ACCENT_STYLE}-toolbutton")
            cb.pack(side="left", padx=2)
            self._dow_checks.append(cb)
        self._on_freq_mode_change()

        sec_params = ttkb.Labelframe(outer, text=" Parametry zadania ",
                                     bootstyle=ACCENT_STYLE, padding=10)
        sec_params.pack(fill="x", pady=(0, 8))
        pr = ttkb.Frame(sec_params); pr.pack(fill="x")
        ttkb.Label(pr, text="Czas trwania (min):", font=FONT_SMALL).pack(side="left")
        self._e_duration = ttkb.Entry(pr, width=6, font=FONT_SMALL)
        self._e_duration.insert(0, "15")
        self._e_duration.pack(side="left", padx=(4, 14))
        ttkb.Label(pr, text="Priorytet:", font=FONT_SMALL).pack(side="left")
        self._e_priority = ttkb.Entry(pr, width=4, font=FONT_SMALL)
        self._e_priority.insert(0, str(DEFAULT_PRIORITY))
        self._e_priority.pack(side="left", padx=(4, 14))
        ttkb.Label(pr, text="Status:", font=FONT_SMALL).pack(side="left")
        self._status_var = tk.StringVar(value=TASK_STATUS_LABELS[DEFAULT_STATUS])
        status_menu = ttkb.Combobox(pr, textvariable=self._status_var, state="normal",
                                    values=list(TASK_STATUS_LABELS.values()), width=16, font=FONT_SMALL,
                                    bootstyle=ACCENT_STYLE)
        status_menu.pack(side="left", padx=4)
        ttkb.Label(pr, text="Bezpiecznik max zadań:", font=FONT_SMALL).pack(side="left", padx=(14, 0))
        self._e_max_tasks = ttkb.Entry(pr, width=6, font=FONT_SMALL)
        self._e_max_tasks.insert(0, "300")
        self._e_max_tasks.pack(side="left", padx=4)
        ttkb.Label(sec_params,
                 text="Priorytet i status: numery na liście nie są w 100% potwierdzone. "
                      "Jeśli brakuje statusu na liście (np. 'Do przypisania'), wpisz jego numer ręcznie w pole Status.",
                 bootstyle="secondary", font=FONT_TINY, wraplength=900, justify="left").pack(anchor="w", pady=(6, 0))

        sec_go = ttkb.Labelframe(outer, text=" Podgląd i wysyłka ", bootstyle=ACCENT_STYLE, padding=10)
        sec_go.pack(fill="both", expand=True)
        gob = ttkb.Frame(sec_go); gob.pack(fill="x")
        btn_preview = ttkb.Button(gob, text="① 🔍 Pokaż podgląd", command=self._show_preview,
                  bootstyle="secondary", padding=(14, 6))
        btn_preview.pack(side="left")
        send_style = "danger" if IS_PROD else "success"
        self._btn_send = ttkb.Button(gob, text="②  ▶  Wyślij zadania", command=self._do_send,
                  bootstyle=send_style, padding=(14, 6), state="disabled")
        self._btn_send.pack(side="left", padx=8)
        self._progress = ttkb.Progressbar(gob, length=240, mode="determinate", bootstyle=f"{ACCENT_STYLE}-striped")
        self._progress.pack(side="left", padx=12)
        self._lbl_progress = ttkb.Label(gob, text="", font=FONT_SMALL)
        self._lbl_progress.pack(side="left")
        ttkb.Label(sec_go, text="Kolejność: najpierw \"Pokaż podgląd\" (nic nie wysyła) — dopiero wtedy odblokuje się \"Wyślij zadania\".",
                 bootstyle="secondary", font=FONT_TINY).pack(anchor="w", pady=(4, 0))

        self._preview_text = ttkb.ScrolledText(
            sec_go, height=9, font=("Consolas", 9), bg="#f7f7f7", fg="#222",
            state="disabled", wrap="word", auto_hide=True)
        self._preview_text.pack(fill="both", expand=True, pady=(8, 0))

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _log_msg(self, msg, tag="inf"):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.configure(state="normal")
        self._log.insert("end", f"[{ts}] {msg}\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")
        logging.info(msg)

    def _bg(self, fn): return threading.Thread(target=fn, daemon=True).start()

    def _on_freq_mode_change(self):
        """Wyszarza tryb częstotliwości, który akurat nie jest wybrany — żeby było
        jasne, który zestaw pól faktycznie coś robi."""
        if self._freq_mode.get() == "weekdays":
            self._e_every_n.configure(state="disabled")
            for cb in self._dow_checks:
                cb.configure(state="normal")
        else:
            self._e_every_n.configure(state="normal")
            for cb in self._dow_checks:
                cb.configure(state="disabled")

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

    # ── MIESZKAŃCY ────────────────────────────────────────────────────────────
    def _load_residents(self):
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Mieszkańcy", "Organizacja musi być liczbą (ID)."); return
        self._lbl_res_status.config(text="Pobieranie...")

        def _t():
            residents = fetch_active_beneficiaries(self.client, org_id)
            residents = sorted(residents, key=lambda r: _resident_label(r).lower())
            self.after(0, lambda: self._render_residents(residents))
            self.after(0, lambda: self._lbl_res_status.config(text=f"{len(residents)} aktywnych"))
            self.after(0, lambda: self._log_msg(f"Pobrano {len(residents)} aktywnych mieszkańców", "ok"))
        self._bg(_t)

    def _render_residents(self, residents):
        self._residents = residents
        self._checked_ids = set()
        self._refresh_res_tree(residents)

    def _refresh_res_tree(self, residents):
        for iid in self._res_tree.get_children():
            self._res_tree.delete(iid)
        self._resident_rows = []
        for r in residents:
            iid = self._res_tree.insert("", "end", values=("☐", r.get("id", ""), _resident_label(r)))
            self._resident_rows.append((iid, r))
        self._update_res_sel()

    def _apply_res_filter(self):
        q = self._e_res_filter.get().strip().lower()
        filtered = [r for r in self._residents if q in _resident_label(r).lower()] if q else self._residents
        self._refresh_res_tree(filtered)

    def _on_res_click(self, event):
        iid = self._res_tree.identify_row(event.y)
        if not iid or self._res_tree.identify_region(event.x, event.y) != "cell":
            return
        vals = list(self._res_tree.item(iid, "values"))
        if iid in self._checked_ids:
            self._checked_ids.discard(iid); vals[0] = "☐"
            self._res_tree.item(iid, values=vals, tags=())
        else:
            self._checked_ids.add(iid); vals[0] = "☑"
            self._res_tree.item(iid, values=vals, tags=("checked",))
        self._update_res_sel()

    def _sel_all_res(self):
        for iid, _ in self._resident_rows:
            self._checked_ids.add(iid)
            vals = list(self._res_tree.item(iid, "values")); vals[0] = "☑"
            self._res_tree.item(iid, values=vals, tags=("checked",))
        self._update_res_sel()

    def _desel_all_res(self):
        for iid, _ in self._resident_rows:
            self._checked_ids.discard(iid)
            vals = list(self._res_tree.item(iid, "values")); vals[0] = "☐"
            self._res_tree.item(iid, values=vals, tags=())
        self._update_res_sel()

    def _update_res_sel(self):
        self._lbl_res_sel.config(text=f"{len(self._checked_ids)} zaznaczonych")

    def _get_selected_residents(self):
        iid_map = {iid: r for iid, r in self._resident_rows}
        return [iid_map[iid] for iid in self._checked_ids if iid in iid_map]

    # ── RODZAJ ZADANIA ────────────────────────────────────────────────────────
    def _load_kinds(self):
        self._lbl_kind_status.config(text="Pobieranie...")

        def _t():
            kinds = fetch_all_task_kinds(self.client)
            self.after(0, lambda: self._render_kinds(kinds))
            self.after(0, lambda: self._lbl_kind_status.config(text=f"{len(kinds)} rodzajów"))
            self.after(0, lambda: self._log_msg(f"Pobrano {len(kinds)} rodzajów zadań", "ok"))
        self._bg(_t)

    def _render_kinds(self, kinds):
        self._task_kinds = kinds
        self._refresh_kinds_list(kinds)

    def _refresh_kinds_list(self, kinds):
        self._kinds_display = kinds
        self._lb_kinds.delete(0, "end")
        for k in kinds:
            self._lb_kinds.insert("end", f"[{k['id']}] {k.get('content','')}")

    def _apply_kind_filter(self):
        q = self._e_kind_filter.get().strip().lower()
        filtered = [k for k in self._task_kinds if q in (k.get("content") or "").lower()] if q else self._task_kinds
        self._refresh_kinds_list(filtered)

    def _on_kind_select(self, _):
        sel = self._lb_kinds.curselection()
        if not sel:
            return
        self._selected_kind = self._kinds_display[sel[0]]
        auto = extract_default_duration(self._selected_kind)
        if auto is not None:
            self._e_duration.delete(0, "end")
            self._e_duration.insert(0, str(auto))
        self._lbl_kind_sel.config(
            text=f"Wybrano: {self._selected_kind.get('content','')}"
                 + (f"  (sugerowany czas: {auto} min)" if auto is not None else ""),
            bootstyle="success")

    # ── SALA ──────────────────────────────────────────────────────────────────
    def _load_rooms(self):
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Sala", "Organizacja musi być liczbą (ID)."); return
        self._lbl_room_status.config(text="Pobieranie...")

        def _t():
            rooms = fetch_rooms(self.client, org_id)
            self.after(0, lambda: self._render_rooms(rooms))
            self.after(0, lambda: self._lbl_room_status.config(text=f"{len(rooms)} sal"))
            self.after(0, lambda: self._log_msg(f"Pobrano {len(rooms)} sal", "ok"))
        self._bg(_t)

    def _render_rooms(self, rooms):
        self._rooms = rooms
        self._refresh_rooms_list(rooms)

    def _refresh_rooms_list(self, rooms):
        self._rooms_display = rooms
        self._lb_rooms.delete(0, "end")
        for r in rooms:
            self._lb_rooms.insert("end", f"[{r.get('id')}] {r.get('name', r.get('content',''))}")

    def _apply_room_filter(self):
        q = self._e_room_filter.get().strip().lower()
        filtered = [r for r in self._rooms if q in str(r.get("name", r.get("content",""))).lower()] if q else self._rooms
        self._refresh_rooms_list(filtered)

    def _on_room_select(self, _):
        sel = self._lb_rooms.curselection()
        if not sel:
            return
        self._selected_room = self._rooms_display[sel[0]]
        self._lbl_room_sel.config(
            text=f"Wybrano: [{self._selected_room.get('id')}] {self._selected_room.get('name', self._selected_room.get('content',''))}",
            bootstyle="success")

    def _on_room_mode_change(self):
        mode = self._room_mode.get()
        if mode == "fixed":
            self._lb_rooms.configure(state="normal")
            self._e_room_filter.configure(state="normal")
        else:
            self._lb_rooms.configure(state="disabled")
            self._e_room_filter.configure(state="disabled")
            if mode == "none":
                self._lbl_room_sel.config(text="Wybrano: brak sali (zadanie nie w pomieszczeniu)", bootstyle="success")
            else:
                self._lbl_room_sel.config(text="Wybrano: pokój zamieszkania (automatycznie per mieszkaniec)", bootstyle="success")

    # ── HARMONOGRAM: PODGLĄD ─────────────────────────────────────────────────
    def _collect_schedule_params(self):
        errors = []
        residents = self._get_selected_residents()
        if not residents:
            errors.append("Nie zaznaczono żadnego mieszkańca (zakładka 2).")
        if not self._selected_kind:
            errors.append("Nie wybrano rodzaju zadania (zakładka 2).")
        if self._room_mode.get() == "fixed" and not self._selected_room:
            errors.append("Wybierz salę z listy, albo zmień tryb na 'pokój zamieszkania' / 'brak sali' (zakładka 2).")
        date_from = self._e_from.entry.get().strip()
        date_to = self._e_to.entry.get().strip()
        time_of_day = f"{self._sp_hour.get().strip().zfill(2)}:{self._sp_minute.get().strip().zfill(2)}"
        try:
            datetime.strptime(date_from, "%Y-%m-%d"); datetime.strptime(date_to, "%Y-%m-%d")
        except ValueError:
            errors.append("Zakres dat musi być w formacie YYYY-MM-DD.")
        if not re_time_ok(time_of_day):
            errors.append("Godzina musi być w formacie HH:MM.")
        try:
            duration = int(self._e_duration.get().strip())
        except ValueError:
            errors.append("Czas trwania musi być liczbą całkowitą (minuty).")
            duration = None
        try:
            priority = int(self._e_priority.get().strip())
        except ValueError:
            errors.append("Priorytet musi być liczbą całkowitą.")
            priority = DEFAULT_PRIORITY
        try:
            max_tasks = int(self._e_max_tasks.get().strip())
        except ValueError:
            errors.append("Bezpiecznik musi być liczbą całkowitą.")
            max_tasks = 300
        status_raw = self._status_var.get().strip()
        status = STATUS_LABEL_TO_CODE.get(status_raw.lower())
        if status is None:
            try:
                status = int(status_raw)
            except ValueError:
                errors.append(f"Nieznany status '{status_raw}' — wpisz nazwę z listy albo numer statusu wprost.")
                status = DEFAULT_STATUS

        weekdays = None
        every_n = 1
        if self._freq_mode.get() == "weekdays":
            weekdays = {i for i, v in self._dow_vars.items() if v.get()}
            if not weekdays:
                errors.append("Wybierz przynajmniej jeden dzień tygodnia.")
        else:
            try:
                every_n = max(1, int(self._e_every_n.get().strip()))
            except ValueError:
                errors.append("Częstotliwość 'co N dni' musi być liczbą całkowitą.")

        if errors:
            return None, errors
        return {
            "residents": residents, "date_from": date_from, "date_to": date_to,
            "time_of_day": time_of_day, "duration": duration, "priority": priority,
            "status": status, "max_tasks": max_tasks, "weekdays": weekdays, "every_n": every_n,
            "room_mode": self._room_mode.get(),
        }, []

    def _show_preview(self):
        params, errors = self._collect_schedule_params()
        if errors:
            messagebox.showwarning("Podgląd", "\n".join(errors))
            return
        self._btn_send.config(state="disabled")
        self._set_preview_text("Pobieranie usług mieszkańców...")

        def _t():
            services = fetch_visitor_services(self.client, int(self._e_org.get().strip()))
            occ_dates = generate_dates(params["date_from"], params["date_to"],
                                       every_n_days=params["every_n"], weekdays=params["weekdays"])
            pending = []
            missing = []
            resolved_rooms = {}   # rname -> (room_id, pole) tylko dla trybu "own"
            for res in params["residents"]:
                rid, rname = res.get("id"), _resident_label(res)
                service_id = find_adhoc_service_id(services, rid)
                if service_id is None:
                    missing.append(rname)
                    continue
                if params["room_mode"] == "own":
                    room_id, field = fetch_resident_room_id(self.client, rid)
                    resolved_rooms[rname] = (room_id, field)
                    if room_id is None:
                        missing.append(f"{rname} (nie znaleziono pokoju zamieszkania)")
                        continue
                elif params["room_mode"] == "none":
                    room_id = None
                else:
                    room_id = self._selected_room.get("id")
                for d in occ_dates:
                    payload = build_task_payload(self._selected_kind, service_id, room_id,
                                                 d, params["time_of_day"], params["duration"],
                                                 priority=params["priority"], status=params["status"])
                    pending.append((res, d, payload))
            self.after(0, lambda: self._render_preview(pending, missing, occ_dates, params, resolved_rooms))
        self._bg(_t)

    def _render_preview(self, pending, missing, occ_dates, params, resolved_rooms=None):
        resolved_rooms = resolved_rooms or {}
        self._pending_tasks = pending
        lines = []
        lines.append(f"Rodzaj zadania: {self._selected_kind.get('content','')}")
        if params["room_mode"] == "own":
            room_desc = "pokój zamieszkania mieszkańca (automatycznie, patrz niżej)"
        elif params["room_mode"] == "none":
            room_desc = "brak (zadanie nie w pomieszczeniu)"
        else:
            room_id = self._selected_room.get('id')
            room_desc = f"[{room_id}] {self._selected_room.get('name', self._selected_room.get('content',''))}"
        lines.append(f"Sala: {room_desc}")
        lines.append(f"Okres: {params['date_from']} .. {params['date_to']}  |  godzina {params['time_of_day']}  |  "
                     f"{len(occ_dates)} dat  |  czas trwania {params['duration']} min  |  "
                     f"priorytet {params['priority']}  |  status {params['status']} "
                     f"({TASK_STATUS_LABELS.get(params['status'], 'nieznany — wpisany ręcznie')})")
        lines.append("─" * 70)
        by_res = {}
        for res, d, _ in pending:
            by_res.setdefault(_resident_label(res), []).append(d)
        for name, dates in sorted(by_res.items()):
            room_note = ""
            if name in resolved_rooms:
                rid, field = resolved_rooms[name]
                room_note = f"  [pokój: {rid}, pole '{field}']"
            lines.append(f"  {name}: {len(dates)} zadań ({dates[0]} .. {dates[-1]}){room_note}")
        if missing:
            lines.append("")
            lines.append(f"⚠ Pominięci: {', '.join(missing)}")
        lines.append("─" * 70)
        lines.append(f"RAZEM: {len(pending)} zadań do utworzenia")
        self._set_preview_text("\n".join(lines))
        self._log_msg(f"Podgląd: {len(pending)} zadań ({len(missing)} mieszkańców pominiętych — brak usługi)", "inf")

        if not pending:
            self._btn_send.config(state="disabled")
            return
        try:
            max_tasks = int(self._e_max_tasks.get().strip())
        except ValueError:
            max_tasks = 300
        if len(pending) > max_tasks:
            self._log_msg(f"{len(pending)} zadań przekracza bezpiecznik max={max_tasks} — wysyłka zablokowana.", "err")
            self._btn_send.config(state="disabled")
        else:
            self._btn_send.config(state="normal")

    # ── WYSYŁKA ───────────────────────────────────────────────────────────────
    def _do_send(self):
        if not self._pending_tasks:
            messagebox.showwarning("Wysyłka", "Najpierw kliknij 'Pokaż podgląd'.")
            return
        n = len(self._pending_tasks)
        warn_prefix = "⚠ PRODUKCJA — zadania trafią od razu do systemu.\n\n" if IS_PROD else ""
        if not messagebox.askyesno("Potwierdzenie wysyłki",
            f"{warn_prefix}Wysłać {n} zadań do systemu?\nTej operacji nie można cofnąć hurtowo."):
            return

        tasks = list(self._pending_tasks)
        self._btn_send.config(state="disabled")
        self._progress.configure(maximum=len(tasks), value=0)
        self._lbl_progress.config(text="Wysyłanie...")

        def _t():
            ok = fail = 0
            for i, (res, d, payload) in enumerate(tasks, 1):
                success, err = post_task(self.client, payload)
                rname = _resident_label(res)
                if success:
                    ok += 1
                    self.after(0, lambda rname=rname, d=d: self._log_msg(f"✓ {rname} {d} — utworzono", "ok"))
                else:
                    fail += 1
                    self.after(0, lambda rname=rname, d=d, err=err: self._log_msg(f"✗ {rname} {d} — {err}", "err"))
                self.after(0, lambda i=i, ok=ok, fail=fail: (
                    self._progress.configure(value=i),
                    self._lbl_progress.config(text=f"{i}/{len(tasks)}  ✓ {ok}  ✗ {fail}")))
                time.sleep(0.3)
            summary = f"Zakończono: {ok}/{len(tasks)} utworzono, {fail} błędów"
            self.after(0, lambda: self._log_msg(f"— {summary} —", "ok" if fail == 0 else "err"))
            self.after(0, lambda: messagebox.showinfo("Gotowe", summary))
            self._pending_tasks = []
        self._bg(_t)


def re_time_ok(s):
    import re as _re
    return bool(_re.fullmatch(r"(?:[01]?\d|2[0-3]):[0-5]\d", s or ""))


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception as e:
        logging.error(f"Nieobsłużony wyjątek przy starcie okna: {e}\n{traceback.format_exc()}")
        _fatal_startup_error("Błąd uruchomienia okna", str(e), exc=e)
