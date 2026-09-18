#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Syrena — Eksport spotkań do Excela (GUI)
Domyślnie config_testowy.json obok skryptu; inny plik jako argument CLI
(np. `python eksport_spotkan_gui.py config_produkcja.json` dla produkcji).
Używa tego samego pliku config co audyt_mieszkancow_gui.py (auth_url,
employee_api_url, organization_id, env_label) — jeśli brakuje w nim
beneficiary_url, jest wyliczany z auth_url (:5010 -> :5020).

Endpointy (POTWIERDZONE przechwyconym ruchem przeglądarki, 2026-09-17):
  Lista spotkań:    GET :5020/api/meeting/by-organization-id/paged
                     ?page=&pageSize=&orderBy=&ascending=&search=&searchFields=
                     (pola do search: meetingTime,kindName,placeName,topic,
                      purpose,summary,urlLink)
  Szczegóły:         GET :5020/api/meeting/{id}
                     -> placeId, kindId, topic, purpose, summary, urlLink(?),
                        leaders (id-y), members (id-y), subjects (id-y),
                        plannedSubjects (id-y)
  Słownik Miejsce:          GET :5000/api/dictionary-value/by-dictionary-kind/72?withAttributes=true
  Słownik Rodzaj spotkania: GET :5000/api/dictionary-value/by-dictionary-kind/94?withAttributes=true
  Uczestnicy (pracownicy):  GET :5010/api/visitor/by-organization-id?typeList=4
  Uczestnicy (mieszkańcy):  GET :5010/api/visitor/by-organization-id?typeList=3

UWAGA -- ZAŁOŻENIE do zweryfikowania: w przechwyconym przykładzie szczegółów
spotkania pola "leaders"/"members"/"subjects"/"plannedSubjects" to listy ID
uczestników, ale nie było jasne z samego przykładu, które pole to "kto
prowadził", a które "kto planowo/faktycznie uczestniczył". Zamiast zgadywać
która kolumna to co, ten skrypt bierze WSZYSTKIE cztery pola razem i dzieli
uczestników na "Pracownicy" / "Mieszkańcy" wyłącznie po tym, w którym z dwóch
słowników /api/visitor (typeList=4 czy typeList=3) dane ID faktycznie
występuje -- to nie wymaga zgadywania znaczenia poszczególnych pól.
"Wnioski" zmapowane na pole "summary" (opisowe pole wypełniane po spotkaniu),
"Link do spotkania" na "urlLink" -- oba do zweryfikowania na żywych danych,
gdzie faktycznie są wypełnione (w przechwyconym przykładzie oba puste).
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading, requests, json, os, sys, traceback, logging
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
CRASH_LOG_PATH  = os.path.join(SCRIPT_DIR, "eksport_spotkan_gui_crash.log")
LAST_LOGIN_PATH = os.path.join(SCRIPT_DIR, "eksport_spotkan_gui_last_login.json")


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
    LOG_PATH = os.path.join(SCRIPT_DIR, CFG.get("log_file", "eksport_spotkan_gui.log"))
    logging.basicConfig(filename=LOG_PATH, level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s",
                        encoding="utf-8")

    AUTH_URL        = CFG["auth_url"]
    EMP_URL         = CFG.get("employee_api_url", AUTH_URL.replace(":5010", ":5000"))
    BENEFICIARY_URL = CFG.get("beneficiary_url", AUTH_URL.replace(":5010", ":5020"))
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


# ── Kolory i typografia (ten sam schemat co audyt_mieszkancow_gui.py) ────────
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
    return tk.Frame(parent, bg=bg, highlightbackground=CARD_BORDER, highlightthickness=1, bd=0)


# ── Backend ───────────────────────────────────────────────────────────────────
TIMEOUT           = 20
PAGE_SIZE_DEFAULT = 100
SEARCH_FIELDS     = "meetingTime,kindName,placeName,topic,purpose,summary,urlLink"


class Client:
    def __init__(self):
        self.servers = {"auth": AUTH_URL, "employee": EMP_URL, "beneficiary": BENEFICIARY_URL}
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


def fetch_meetings(client, org_id, search=None):
    """POTWIERDZONE przechwyconym ruchem: GET .../meeting/by-organization-id/paged"""
    url = f"{client.servers['beneficiary']}/api/meeting/by-organization-id/paged"
    page, page_size, all_items = 1, PAGE_SIZE_DEFAULT, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id,
                  "orderBy": "meetingTime", "ascending": "true"}
        if search:
            params["search"] = search
            params["searchFields"] = SEARCH_FIELDS
        r, err = client._get(url, params=params)
        if err or r is None:
            logging.error(f"Błąd pobierania spotkań org={org_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logging.error(f"Spotkania org={org_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Spotkania org={org_id}: odpowiedź nie jest JSON-em: {e}")
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
    logging.info(f"Organizacja {org_id}: {len(all_items)} spotkań")
    return all_items


def fetch_meeting_detail(client, meeting_id):
    """POTWIERDZONE przechwyconym ruchem: GET .../meeting/{id}"""
    r, err = client._get(f"{client.servers['beneficiary']}/api/meeting/{meeting_id}")
    if err or r is None:
        logging.error(f"Błąd pobierania szczegółów spotkania {meeting_id}: {err}")
        return {}
    if r.status_code != 200:
        logging.error(f"Szczegóły spotkania {meeting_id}: HTTP {r.status_code}")
        return {}
    try:
        return r.json() or {}
    except Exception as e:
        logging.error(f"Szczegóły spotkania {meeting_id}: odpowiedź nie jest JSON-em: {e}")
        return {}


def fetch_dictionary_by_kind(client, kind_id):
    """POTWIERDZONE przechwyconym ruchem: GET .../dictionary-value/by-dictionary-kind/{kind}"""
    r, err = client._get(f"{client.servers['employee']}/api/dictionary-value/by-dictionary-kind/{kind_id}",
                         params={"withAttributes": "true"})
    if err or r is None or r.status_code != 200:
        logging.error(f"Błąd pobierania słownika (kind={kind_id}): {err or (r and r.status_code)}")
        return []
    data = r.json()
    return data if isinstance(data, list) else (data.get("items") or data.get("data") or [])


def fetch_visitors(client, org_id, type_list):
    """POTWIERDZONE przechwyconym ruchem: GET .../visitor/by-organization-id?typeList=<3|4>
    (3 = mieszkańcy, 4 = pracownicy -- ustalone z przykładowych pól odpowiedzi:
    typeList=3 ma 'pesel'/'status', typeList=4 ma 'occupationName').
    UWAGA: przechwycony request NIE MA parametru organizationId (tylko
    typeList) -- serwer najwyraźniej ustala organizację z samego tokena.
    org_id tu przyjmowane tylko dla spójności sygnatury z resztą modułu,
    NIE wysyłane."""
    r, err = client._get(f"{client.servers['auth']}/api/visitor/by-organization-id",
                         params={"typeList": type_list})
    if err or r is None or r.status_code != 200:
        logging.error(f"Błąd pobierania visitor typeList={type_list}: {err or (r and r.status_code)}")
        return []
    data = r.json()
    return data if isinstance(data, list) else (data.get("items") or data.get("data") or [])


def _visitor_label(v):
    return f"{v.get('surname', '')} {v.get('firstName', '')}".strip() or f"#{v.get('id')}"


def fetch_all_beneficiaries(client, org_id):
    """Dodatkowe źródło ID/nazwisk mieszkańców -- ZAŁOŻENIE do zweryfikowania:
    użytkownik potwierdził, że pola uczestników spotkania (leaders/members/
    subjects/plannedSubjects) zawierają ID mieszkańców, ale nie wiadomo czy
    to ta sama przestrzeń ID co /api/visitor (typeList=3) czy ta z
    /api/beneficiary (używana gdzie indziej w tym repo, np.
    audyt_mieszkancow_gui.py). Zamiast zgadywać którą, pobieramy WSZYSTKICH
    mieszkańców (bez filtra statusu, żeby złapać też nieaktywnych/wypisanych,
    do których spotkanie może się jeszcze odnosić) i łączymy z listą z
    /api/visitor przy dopasowywaniu uczestników."""
    url = f"{client.servers['beneficiary']}/api/beneficiary/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id}
        r, err = client._get(url, params=params)
        if err or r is None or r.status_code != 200:
            logging.error(f"Błąd pobierania wszystkich mieszkańców (beneficiary) org={org_id} str.{page}: "
                          f"{err or (r and r.status_code)}")
            break
        try:
            data = r.json()
        except Exception as e:
            logging.error(f"Wszyscy mieszkańcy (beneficiary) org={org_id}: odpowiedź nie jest JSON-em: {e}")
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


def build_meeting_rows(client, org_id, meetings, progress_cb=None):
    places = {v.get("id"): v.get("content") for v in fetch_dictionary_by_kind(client, 72)}
    kinds = {v.get("id"): v.get("content") for v in fetch_dictionary_by_kind(client, 94)}
    employees = fetch_visitors(client, org_id, 4)
    residents = fetch_visitors(client, org_id, 3) + fetch_all_beneficiaries(client, org_id)
    employee_ids = {v.get("id") for v in employees}
    resident_ids = {v.get("id") for v in residents}
    name_by_id = {v.get("id"): _visitor_label(v) for v in employees + residents}

    rows = []
    for i, m in enumerate(meetings, 1):
        if progress_cb:
            progress_cb(i, len(meetings), m.get("topic") or "")
        detail = fetch_meeting_detail(client, m.get("id")) or {}
        place_id = detail.get("placeId")
        miejsce = places.get(place_id) or m.get("placeName") or ""
        kind_id = detail.get("kindId", m.get("kindId"))
        rodzaj = kinds.get(kind_id) or m.get("kindName") or ""

        participant_ids = set()
        for field in ("leaders", "members", "subjects", "plannedSubjects"):
            participant_ids |= set(detail.get(field) or [])
        pracownicy = sorted(name_by_id.get(pid, f"#{pid}") for pid in participant_ids if pid in employee_ids)
        mieszkancy = sorted(name_by_id.get(pid, f"#{pid}") for pid in participant_ids if pid in resident_ids)

        rows.append({
            "organizationId": org_id,
            "meetingId": m.get("id"),
            "Data spotkania": detail.get("meetingTime") or m.get("meetingTime") or "",
            "Rodzaj spotkania": rodzaj,
            "Miejsce": miejsce,
            "Temat": detail.get("topic") or m.get("topic") or "",
            "Cel": detail.get("purpose") or m.get("purpose") or "",
            "Wnioski": detail.get("summary") or "",
            "Link do spotkania": detail.get("urlLink") or detail.get("link") or "",
            "Pracownicy (uczestnicy)": ", ".join(pracownicy),
            "Mieszkańcy (uczestnicy)": ", ".join(mieszkancy),
            # POTWIERDZONE przez użytkownika 2026-09-18: pole 'qualifications'
            # na obiekcie spotkania to "Obserwacje" (widoczne w zakładce
            # "Wnioski i obserwacje" w UI, razem z 'summary' -> Wnioski).
            "Obserwacje": detail.get("qualifications") or "",
        })
    return rows


def filter_rows_by_date(rows, date_from, date_to):
    def in_range(r):
        t = (r.get("Data spotkania") or "")[:10]
        if not t:
            return True
        if date_from and t < date_from:
            return False
        if date_to and t > date_to:
            return False
        return True
    return [r for r in rows if in_range(r)]


def save_excel(out_path, rows):
    wb = openpyxl.Workbook()

    def fill(c): return PatternFill("solid", fgColor=c)
    def thin():
        s = Side(style="thin", color="D1D5DB")
        return Border(left=s, right=s, top=s, bottom=s)

    ws = wb.active
    ws.title = "Spotkania"
    cols = ["Data spotkania", "Rodzaj spotkania", "Miejsce", "Temat", "Cel",
            "Link do spotkania", "Pracownicy (uczestnicy)", "Mieszkańcy (uczestnicy)", "Wnioski", "Obserwacje"]
    widths = [18, 22, 20, 34, 26, 26, 34, 34, 34, 40]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    c = ws["A1"]
    c.value = f"Spotkania  |  {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  {len(rows)} spotkań"
    c.font = Font(name="Segoe UI", bold=True, size=12, color="FFFFFF")
    c.fill = fill("1A3A6B")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    for col, (lbl, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(row=2, column=col, value=lbl)
        cell.font = Font(name="Segoe UI", bold=True, size=10, color="FFFFFF")
        cell.fill = fill("1A3A6B")
        cell.border = thin()
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col)].width = w

    for i, row in enumerate(rows):
        r = 3 + i
        bg = "EBF3FF" if i % 2 == 0 else "FFFFFF"
        for col, key in enumerate(cols, 1):
            v = row.get(key, "")
            cell = ws.cell(row=r, column=col, value=str(v) if v is not None else "")
            cell.fill = fill(bg)
            cell.border = thin()
            cell.font = Font(name="Segoe UI", size=9)
            cell.alignment = Alignment(wrap_text=(key in ("Temat", "Cel", "Wnioski",
                                                          "Pracownicy (uczestnicy)",
                                                          "Mieszkańcy (uczestnicy)",
                                                          "Obserwacje")), vertical="top")

    ws.freeze_panes = "A3"
    if rows:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{2 + len(rows)}"

    wb.save(out_path)


# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Syrena — Eksport spotkań [{ENV_LABEL}]")
        self.geometry("900x600")
        self.resizable(True, True)
        self.configure(bg=BG)

        self.client = Client()
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
        tk.Label(hdr, text="  Syrena — Eksport spotkań",
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
        t2 = tk.Frame(nb, bg=WHITE); nb.add(t2, text="  2. Spotkania  ")

        self._tab_login(t1)
        self._tab_meetings(t2)

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

    # ── TAB 2: Spotkania ──────────────────────────────────────────────────────
    def _tab_meetings(self, p):
        sec_filter = tk.LabelFrame(p, text=" 1. Filtr (opcjonalnie) ",
                                   bg=p["bg"], font=FONT_SECTION, fg=ACCENT_DARK, padx=10, pady=8)
        sec_filter.pack(fill="x", padx=12, pady=(10, 4))

        row_s = tk.Frame(sec_filter, bg=p["bg"])
        row_s.pack(fill="x", pady=(0, 4))
        tk.Label(row_s, text="Szukaj (temat/cel/miejsce/wnioski...):",
                 bg=p["bg"], font=FONT_SMALL).pack(side="left")
        self._e_search = tk.Entry(row_s, width=40, font=FONT_SMALL)
        self._e_search.pack(side="left", padx=6)

        dr = tk.Frame(sec_filter, bg=p["bg"])
        dr.pack(fill="x", pady=(4, 0))
        tk.Label(dr, text="Od:", bg=p["bg"], font=FONT_SMALL).pack(side="left")
        self._e_from = tk.Entry(dr, width=12, font=FONT_SMALL)
        self._e_from.pack(side="left", padx=(4, 14))
        tk.Label(dr, text="Do:", bg=p["bg"], font=FONT_SMALL).pack(side="left")
        self._e_to = tk.Entry(dr, width=12, font=FONT_SMALL)
        self._e_to.pack(side="left", padx=(4, 14))
        for lbl, days in [("Ostatni miesiąc", 30), ("Ostatnie 3 mies.", 90),
                          ("Ostatni rok", 365), ("Wyczyść", None)]:
            tk.Button(dr, text=lbl, relief="flat", font=FONT_SMALL, cursor="hand2",
                      command=lambda d=days: self._set_range(d)).pack(side="left", padx=2)
        tk.Label(sec_filter,
                 text="Format dat: YYYY-MM-DD. Filtrowane lokalnie po dacie spotkania, po pobraniu "
                      "wszystkich pasujących wyszukiwaniu (jeśli podane).",
                 bg=p["bg"], fg=TXT_MUTED, font=FONT_TINY).pack(anchor="w", pady=(6, 0))

        sec_go = tk.Frame(p, bg=p["bg"], padx=12, pady=14)
        sec_go.pack(fill="x")
        btn_go = tk.Button(sec_go, text="▶  Pobierz spotkania i zapisz Excel",
                  command=self._do_fetch_and_save,
                  bg=SUCCESS_CLR if not IS_PROD else DANGER_CLR, fg="white", relief="flat",
                  font=("Segoe UI", 11, "bold"), padx=18, pady=8, bd=0, cursor="hand2")
        btn_go.pack(side="left")
        _add_hover(btn_go, SUCCESS_CLR if not IS_PROD else DANGER_CLR)
        self._progress = ttk.Progressbar(sec_go, length=300, mode="determinate")
        self._progress.pack(side="left", padx=16)
        self._lbl_progress = tk.Label(sec_go, text="", bg=p["bg"], font=FONT_SMALL, fg="#444")
        self._lbl_progress.pack(side="left")

        tk.Label(p,
                 text="UWAGA: 'Wnioski' czytane z pola 'summary', 'Link do spotkania' z 'urlLink' -- "
                      "oba były puste w przechwyconym przykładzie, więc mapowanie nie jest w 100% "
                      "sprawdzone na rzeczywistych danych. Uczestnicy dzieleni na Pracownicy/Mieszkańcy "
                      "przez sprawdzenie, w którym z dwóch słowników /api/visitor dane ID występuje.",
                 bg=p["bg"], fg=TXT_MUTED, font=FONT_TINY, wraplength=820, justify="left").pack(
                     anchor="w", padx=12, pady=(0, 8))

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _log_msg(self, msg, tag="inf"):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.configure(state="normal")
        self._log.insert("end", f"[{ts}] {msg}\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")
        logging.info(msg)

    def _bg(self, fn): return threading.Thread(target=fn, daemon=True).start()

    def _set_range(self, days):
        self._e_from.delete(0, "end"); self._e_to.delete(0, "end")
        if days is not None:
            self._e_to.insert(0, _date.today().strftime("%Y-%m-%d"))
            self._e_from.insert(0, (_date.today() - timedelta(days=days)).strftime("%Y-%m-%d"))

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

    # ── POBIERZ I ZAPISZ ──────────────────────────────────────────────────────
    def _do_fetch_and_save(self):
        try:
            org_id = int(self._e_org.get().strip())
        except ValueError:
            messagebox.showwarning("Spotkania", "Organizacja musi być liczbą (ID)."); return
        search = self._e_search.get().strip() or None
        date_from = self._e_from.get().strip() or None
        date_to = self._e_to.get().strip() or None

        out_path = filedialog.asksaveasfilename(
            title="Zapisz jako", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="spotkania.xlsx")
        if not out_path:
            return

        self._progress.configure(mode="indeterminate")
        self._progress.start(12)
        self._lbl_progress.config(text="Pobieranie listy spotkań...")

        def _t():
            meetings = fetch_meetings(self.client, org_id, search=search)
            self.after(0, lambda: self._log_msg(f"Pobrano {len(meetings)} spotkań (lista)", "ok"))
            self.after(0, lambda: self._progress.configure(mode="determinate", maximum=max(len(meetings), 1), value=0))

            def progress_cb(i, total, topic):
                self.after(0, lambda: (
                    self._progress.configure(value=i),
                    self._lbl_progress.config(text=f"{i}/{total}  {topic}")))

            rows = build_meeting_rows(self.client, org_id, meetings, progress_cb=progress_cb)
            rows = filter_rows_by_date(rows, date_from, date_to)

            try:
                save_excel(out_path, rows)
                msg = f"Zapisano {out_path}  ({len(rows)} spotkań)"
                self.after(0, lambda: self._log_msg(msg, "ok"))
                self.after(0, lambda: messagebox.showinfo("Gotowe", msg))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._log_msg(f"Błąd zapisu Excela: {err_msg}", "err"))
                self.after(0, lambda: messagebox.showerror("Błąd zapisu", err_msg))
            self.after(0, lambda: self._lbl_progress.config(text="Gotowe"))
        self._bg(_t)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception as e:
        logging.error(f"Nieobsłużony wyjątek przy starcie okna: {e}\n{traceback.format_exc()}")
        _fatal_startup_error("Błąd uruchomienia okna", str(e), exc=e)
