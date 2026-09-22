#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Syrena — Sprawdzenie użycia rodzaju zadania (GUI)
Domyślnie config_testowy.json obok skryptu; inny plik jako argument CLI
(np. `python sprawdz_uzycie_rodzaju_zadania_gui.py config_produkcja.json`
dla produkcji). Używa tego samego pliku config co audyt_mieszkancow_gui.py
(auth_url, employee_api_url, organization_id, env_label, shift_url).

To jest GUI-wersja sprawdz_uzycie_rodzaju_zadania.py (ten sam mechanizm
sprawdzania -- patrz komentarz u góry tamtego pliku dla pełnego opisu
dwóch sygnałów dopasowania po ID/po nazwie i dlaczego pole ID w
odpowiedzi API jest wykrywane dynamicznie, nie zakładane na sztywno).

Sprawdza, czy dany wpis w słowniku 53 ("Rodzaj zadania") jest gdziekolwiek
używany przez konkretne zadania (:5070/api/task, aktywne i usunięte) --
przydatne PRZED usunięciem wpisu ze słownika (patrz usun_ze_slownika_gui.py).
Można podać kilka organizacji naraz -- skrypt NIE zgaduje "całego systemu"
automatycznie, bo nie ma potwierdzonego endpointu do wylistowania
wszystkich organizacji.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading, requests, json, os, sys, re, traceback, logging
from datetime import datetime

# ── Konfiguracja ─────────────────────────────────────────────────────────────
SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
CONFIG_NAME     = sys.argv[1] if len(sys.argv) > 1 else "config_testowy.json"
CONFIG_PATH     = os.path.join(SCRIPT_DIR, CONFIG_NAME)
CRASH_LOG_PATH  = os.path.join(SCRIPT_DIR, "sprawdz_uzycie_rodzaju_zadania_gui_crash.log")
LAST_LOGIN_PATH = os.path.join(SCRIPT_DIR, "sprawdz_uzycie_rodzaju_zadania_gui_last_login.json")


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
    LOG_PATH = os.path.join(SCRIPT_DIR, CFG.get("log_file", "sprawdz_uzycie_rodzaju_zadania_gui.log"))
    logging.basicConfig(filename=LOG_PATH, level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s",
                        encoding="utf-8")

    AUTH_URL       = CFG["auth_url"]
    EMP_URL        = CFG.get("employee_api_url", AUTH_URL.replace(":5010", ":5000"))
    SHIFT_URL      = CFG.get("shift_url", AUTH_URL.replace(":5010", ":5070"))
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


# ── Kolory i typografia (ten sam schemat co inne GUI w tym repo) ─────────────
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
WARN_CLR     = "#b8860b"

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
    return tk.Frame(parent, bg=bg, highlightbackground=CARD_BORDER, highlightthickness=1, bd=0)


# ── Backend (ta sama logika co sprawdz_uzycie_rodzaju_zadania.py) ───────────
TIMEOUT = 20
# /api/task/by-organization-id zwraca (jak widać w praktyce) całą listę zadań
# organizacji naraz, bez działającego stronicowania -- dla organizacji z dużą
# liczbą zadań bywa wolne i potrafi przekroczyć TIMEOUT=20s. Ten sam wzorzec
# "spróbuj dłużej po timeout" co AUDIT_TIMEOUT/AUDIT_TIMEOUT_RETRY w
# audyt_mieszkancow_gui.py.
TASK_TIMEOUT = 60
TASK_TIMEOUT_RETRY = 180
ZADANIA_DICTIONARY_ID = 53
_TASK_KIND_ID_FIELD_RE = re.compile(r"(?i)^taskkind.*id$")


class Client:
    def __init__(self):
        self.servers = {"auth": AUTH_URL, "employee": EMP_URL, "shift": SHIFT_URL}
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
        if r.status_code == 401 and _retry and self._relogin():
            return self._get(url, params=params, _retry=False)
        return r, None


def fetch_dictionary_values(client, dictionary_id):
    r, err = client._get(f"{client.servers['employee']}/api/dictionary-value/by-dictionary-id/{dictionary_id}")
    if err or r is None or r.status_code != 200:
        logging.error(f"Błąd pobierania wartości słownika {dictionary_id}: {err or (r and r.status_code)}")
        return []
    data = r.json()
    return data if isinstance(data, list) else (data.get("items") or data.get("data") or [])


def fetch_active_tasks(client, org_id):
    """UWAGA -- NIE W PEŁNI POTWIERDZONE (patrz sprawdz_uzycie_rodzaju_zadania.py
    i audyt_mieszkancow_gui.py::fetch_active_tasks): stronicowanie dla PEŁNEJ
    listy zadań organizacji nie jest potwierdzone przechwyconym ruchem."""
    url = f"{client.servers['shift']}/api/task/by-organization-id"
    page, page_size, all_items, warned = 1, 200, [], False
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id}
        r, err = client._get(url, params=params, timeout=TASK_TIMEOUT)
        if err and err.startswith("TIMEOUT:"):
            logging.warning(f"Zadania org={org_id} str.{page}: timeout po {TASK_TIMEOUT}s "
                            f"— ponawiam z limitem {TASK_TIMEOUT_RETRY}s...")
            r, err = client._get(url, params=params, timeout=TASK_TIMEOUT_RETRY)
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
                    "fetch_active_tasks: /api/task/by-organization-id zwrócił gołą listę -- "
                    "stronicowanie dla PEŁNEJ listy zadań NIE jest potwierdzone.")
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
    url = f"{client.servers['shift']}/api/task/deleted/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id,
                  "orderBy": "id", "ascending": "true"}
        r, err = client._get(url, params=params, timeout=TASK_TIMEOUT)
        if err and err.startswith("TIMEOUT:"):
            logging.warning(f"Usunięte zadania org={org_id} str.{page}: timeout po {TASK_TIMEOUT}s "
                            f"— ponawiam z limitem {TASK_TIMEOUT_RETRY}s...")
            r, err = client._get(url, params=params, timeout=TASK_TIMEOUT_RETRY)
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


def detect_task_kind_id_field(tasks):
    for t in tasks:
        if not isinstance(t, dict):
            continue
        for k in t.keys():
            if _TASK_KIND_ID_FIELD_RE.match(k):
                return k
    return None


def _task_employee_names(t):
    names, seen = [], set()
    top_name = (t.get("executingEmployeeName") or "").strip()
    if top_name:
        names.append(top_name); seen.add(top_name)
    for ex in t.get("taskExecutors") or []:
        n = f"{ex.get('employeeSurname', '')} {ex.get('employeeFirstName', '')}".strip()
        if n and n not in seen:
            seen.add(n); names.append(n)
    return ", ".join(names)


def match_tasks(tasks, task_kind_id, kind_content, id_field):
    kind_content_norm = (kind_content or "").strip().lower()
    hits = []
    for t in tasks:
        matched_by = []
        if id_field and t.get(id_field) == task_kind_id:
            matched_by.append("id")
        if kind_content_norm:
            name = (t.get("name") or t.get("taskKindName") or "").strip().lower()
            if name and name == kind_content_norm:
                matched_by.append("nazwa")
        if matched_by:
            hits.append((t, matched_by))
    return hits


def check_task_kind_usage(client, org_ids, task_kind_id, report_cb):
    """report_cb(line, tag) -- wypisuje jedną linię raportu do GUI (tag: inf/ok/err/warn)."""
    total_hits = 0
    any_reliable_check = False

    for org_id in org_ids:
        report_cb(f"— Organizacja {org_id} —", "inf")
        if not client.select_organization(org_id):
            report_cb(f"  ✗ Nie udało się przełączyć na organizationId={org_id} — pomijam.", "err")
            continue

        dict_values = fetch_dictionary_values(client, ZADANIA_DICTIONARY_ID)
        kind_entry = next((v for v in dict_values if v.get("id") == task_kind_id), None)
        kind_content = kind_entry.get("content") if kind_entry else None
        if kind_content:
            report_cb(f"  Wpis w słowniku 53: '{kind_content}'", "inf")
        else:
            report_cb(f"  ⚠ W słowniku 53 tej organizacji NIE MA wpisu o ID={task_kind_id} "
                     f"— dopasowanie po nazwie pominięte dla tej organizacji.", "warn")

        report_cb("  Pobieranie zadań... (organizacje z dużą liczbą zadań mogą potrwać nawet kilka minut)", "inf")
        active = fetch_active_tasks(client, org_id)
        deleted = fetch_deleted_tasks(client, org_id)
        all_tasks = active + deleted
        report_cb(f"  Pobrano: {len(active)} aktywnych zadań, {len(deleted)} usuniętych zadań.", "inf")

        id_field = detect_task_kind_id_field(all_tasks)
        if id_field:
            report_cb(f"  Wykryto pole ID rodzaju zadania: '{id_field}' — dopasowanie po ID aktywne.", "inf")
        else:
            report_cb("  ⚠ Brak pola pasującego do wzorca 'taskKind...Id' w zadaniach — "
                      "dopasowanie WYŁĄCZNIE po nazwie (mniej pewne).", "warn")

        if not id_field and not kind_content:
            report_cb(f"  ✗ Brak wiarygodnego sygnału dla tej organizacji — pomijam, wynik NIEROZSTRZYGNIĘTY.", "err")
            continue
        any_reliable_check = True

        hits_active = match_tasks(active, task_kind_id, kind_content, id_field)
        hits_deleted = match_tasks(deleted, task_kind_id, kind_content, id_field)

        if not hits_active and not hits_deleted:
            report_cb(f"  ✓ Nie znaleziono żadnego zadania używającego ID={task_kind_id}.", "ok")
            continue

        total_hits += len(hits_active) + len(hits_deleted)
        for label, hits in (("AKTYWNE", hits_active), ("USUNIĘTE", hits_deleted)):
            if not hits:
                continue
            report_cb(f"  ✗ Znaleziono {len(hits)} zadań ({label}) używających ID={task_kind_id}:", "err")
            for t, matched_by in hits:
                name = t.get("name") or t.get("taskKindName") or "(brak nazwy)"
                resident = t.get("visitorName") or (f"visitorId={t['visitorId']}" if t.get("visitorId") else "—")
                employees = _task_employee_names(t) or "—"
                status = t.get("taskStatus")
                signal = "+".join(matched_by)
                report_cb(f"     [id={t.get('id')}] '{name}' | mieszkaniec: {resident} | "
                         f"pracownicy: {employees} | status: {status} | dopasowanie: {signal}", "err")

    return total_hits, any_reliable_check


# ── GUI ───────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Syrena — Sprawdzenie użycia rodzaju zadania [{ENV_LABEL}]")
        self.geometry("880x620")
        self.resizable(True, True)
        self.configure(bg=BG)

        self.client = Client()
        self._build_ui()

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
        tk.Label(hdr, text="  Syrena — Sprawdzenie użycia rodzaju zadania",
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
        t2 = tk.Frame(nb, bg=WHITE); nb.add(t2, text="  2. Sprawdzenie  ")

        self._tab_login(t1)
        self._tab_check(t2)

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
        ]):
            tk.Label(f, text=lbl, bg=WHITE, font=FONT_BASE).grid(row=row, column=0, sticky="w", pady=6)
            e = tk.Entry(f, width=34, font=FONT_BASE, relief="solid",
                         highlightthickness=1, highlightbackground=CARD_BORDER,
                         highlightcolor=ACCENT, bd=1, **kw)
            e.grid(row=row, column=1, padx=12, pady=6, sticky="w")
            setattr(self, attr, e)
        self._e_user.insert(0, _load_last_login())
        self._e_pass.bind("<Return>", lambda _: self._do_login())
        self._e_user.bind("<Return>", lambda _: self._e_pass.focus_set())

        tk.Label(f, text="Hasło nie jest nigdzie zapisywane — tylko login.",
                 bg=WHITE, fg=TXT_MUTED, font=FONT_TINY).grid(
                     row=2, column=0, columnspan=2, sticky="w", pady=(2, 12))
        bf = tk.Frame(f, bg=WHITE)
        bf.grid(row=3, column=0, columnspan=2, sticky="w")
        btn_login = tk.Button(bf, text="Zaloguj", command=self._do_login,
                  bg=ACCENT, fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), activebackground=_lighten(ACCENT),
                  activeforeground="white", padx=18, pady=7, cursor="hand2", bd=0)
        btn_login.pack(side="left")
        _add_hover(btn_login, ACCENT)
        self._lbl_auth = tk.Label(bf, text="", bg=WHITE, font=FONT_BASE)
        self._lbl_auth.pack(side="left", padx=12)

    # ── TAB 2: Sprawdzenie ────────────────────────────────────────────────────
    def _tab_check(self, p):
        sec = tk.LabelFrame(p, text=" 1. Parametry sprawdzenia ",
                            bg=p["bg"], font=FONT_SECTION, fg=ACCENT_DARK, padx=10, pady=8)
        sec.pack(fill="x", padx=12, pady=(10, 4))

        row1 = tk.Frame(sec, bg=p["bg"])
        row1.pack(fill="x", pady=(0, 6))
        tk.Label(row1, text="Rodzaj zadania (ID ze słownika 53):",
                 bg=p["bg"], font=FONT_SMALL).pack(side="left")
        self._e_kind_id = tk.Entry(row1, width=12, font=FONT_SMALL)
        self._e_kind_id.pack(side="left", padx=6)

        row2 = tk.Frame(sec, bg=p["bg"])
        row2.pack(fill="x")
        tk.Label(row2, text="Organizacje (ID, oddzielone spacją):",
                 bg=p["bg"], font=FONT_SMALL).pack(side="left")
        self._e_org_ids = tk.Entry(row2, width=30, font=FONT_SMALL)
        self._e_org_ids.pack(side="left", padx=6)
        self._e_org_ids.insert(0, str(ORG_ID_DEFAULT))

        tk.Label(sec,
                 text="Sprawdza zadania AKTYWNE i USUNIĘTE we wszystkich podanych organizacjach. "
                      "Skrypt NIE zgaduje 'całego systemu' -- podaj wszystkie organizacje, które chcesz sprawdzić.",
                 bg=p["bg"], fg=TXT_MUTED, font=FONT_TINY, wraplength=820, justify="left").pack(
                     anchor="w", pady=(6, 0))

        sec_go = tk.Frame(p, bg=p["bg"], padx=12, pady=12)
        sec_go.pack(fill="x")
        btn_go = tk.Button(sec_go, text="▶  Sprawdź użycie", command=self._do_check,
                  bg=SUCCESS_CLR if not IS_PROD else DANGER_CLR, fg="white", relief="flat",
                  font=("Segoe UI", 11, "bold"), padx=18, pady=8, bd=0, cursor="hand2")
        btn_go.pack(side="left")
        _add_hover(btn_go, SUCCESS_CLR if not IS_PROD else DANGER_CLR)
        self._progress = ttk.Progressbar(sec_go, length=220, mode="indeterminate")
        self._progress.pack(side="left", padx=16)

        self._lbl_result = tk.Label(p, text="", bg=p["bg"], font=("Segoe UI", 11, "bold"))
        self._lbl_result.pack(anchor="w", padx=12, pady=(0, 4))

        lf = tk.LabelFrame(p, text=" Raport ", bg=p["bg"], fg=ACCENT_DARK, font=FONT_SECTION)
        lf.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self._report = scrolledtext.ScrolledText(
            lf, font=("Consolas", 9),
            bg="#1e1e1e", fg="#ccc", insertbackground="white",
            state="disabled", wrap="word")
        self._report.pack(fill="both", expand=True, padx=6, pady=6)
        self._report.tag_config("ok",   foreground="#6dbf67")
        self._report.tag_config("err",  foreground="#e06c6c")
        self._report.tag_config("warn", foreground="#e0b86c")
        self._report.tag_config("inf",  foreground="#9cdcfe")

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _report_msg(self, msg, tag="inf"):
        self._report.configure(state="normal")
        self._report.insert("end", msg + "\n", tag)
        self._report.see("end")
        self._report.configure(state="disabled")
        logging.info(msg)

    def _bg(self, fn): return threading.Thread(target=fn, daemon=True).start()

    # ── LOGOWANIE ─────────────────────────────────────────────────────────────
    def _do_login(self):
        user, pw = self._e_user.get().strip(), self._e_pass.get()
        if not user or not pw:
            messagebox.showwarning("Logowanie", "Podaj login i hasło"); return
        self._lbl_auth.config(text="Logowanie...", fg="#888")

        def _t():
            ok = self.client.login(user, pw)
            if not ok:
                self.after(0, lambda: self._lbl_auth.config(text="Błąd logowania", fg=DANGER_CLR))
                return
            _save_last_login(user)
            self.after(0, lambda: self._lbl_auth.config(text="✓ Zalogowano", fg=SUCCESS_CLR))
            self.after(0, lambda: self._lbl_status.config(text="●  Zalogowano", fg="#aaffaa"))
            self.after(0, lambda: self._nb.select(1))
        self._bg(_t)

    # ── SPRAWDŹ UŻYCIE ────────────────────────────────────────────────────────
    def _do_check(self):
        if "Authorization" not in self.client.s.headers:
            messagebox.showwarning("Sprawdzenie", "Zaloguj się najpierw (zakładka 1)."); return
        try:
            task_kind_id = int(self._e_kind_id.get().strip())
        except ValueError:
            messagebox.showwarning("Sprawdzenie", "ID rodzaju zadania musi być liczbą."); return
        raw_orgs = self._e_org_ids.get().strip()
        try:
            org_ids = [int(x) for x in raw_orgs.replace(",", " ").split()]
        except ValueError:
            messagebox.showwarning("Sprawdzenie", "Organizacje muszą być liczbami oddzielonymi spacją."); return
        if not org_ids:
            messagebox.showwarning("Sprawdzenie", "Podaj przynajmniej jedną organizację."); return

        self._report.configure(state="normal"); self._report.delete("1.0", "end"); self._report.configure(state="disabled")
        self._lbl_result.config(text="Sprawdzanie...", fg=ACCENT_DARK)
        self._progress.start(12)

        def _t():
            try:
                total_hits, any_reliable = check_task_kind_usage(
                    self.client, org_ids, task_kind_id,
                    report_cb=lambda msg, tag: self.after(0, lambda: self._report_msg(msg, tag)))
            except Exception as e:
                self.after(0, lambda: self._report_msg(f"Nieoczekiwany błąd: {e}", "err"))
                self.after(0, lambda: self._progress.stop())
                self.after(0, lambda: self._lbl_result.config(text="✗ Błąd sprawdzania", fg=DANGER_CLR))
                return

            self.after(0, lambda: self._progress.stop())
            if not any_reliable:
                self.after(0, lambda: self._lbl_result.config(
                    text="⚠ NIEROZSTRZYGNIĘTE — brak wiarygodnego sygnału (patrz raport)", fg=WARN_CLR))
            elif total_hits == 0:
                self.after(0, lambda: self._lbl_result.config(
                    text=f"✓ NIE UŻYWANY w żadnej z podanych organizacji ({org_ids})", fg=SUCCESS_CLR))
            else:
                self.after(0, lambda: self._lbl_result.config(
                    text=f"✗ UŻYWANY — łącznie {total_hits} zadań", fg=DANGER_CLR))
        self._bg(_t)


def main():
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        _fatal_startup_error("Nieoczekiwany błąd startu", str(e), exc=e)


if __name__ == "__main__":
    main()
