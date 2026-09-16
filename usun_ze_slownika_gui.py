#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Syrena — Usuwanie pozycji ze słownika (GUI)
Domyślnie config_testowy.json obok skryptu; inny plik jako argument CLI
(np. `python usun_ze_slownika_gui.py config_produkcja.json` dla
produkcji).

Do naprawiania pomyłek zrobionych przez dodaj_uslugi_pracownikow_gui.py
(słownik 51 — Usługi) i dodaj_rodzaje_zadan_gui.py (słownik 53 — Rodzaje
zadań) — ale działa na dowolnym słowniku (podajesz jego ID).

Przepływ (dokładnie taki jak w przechwyconym ruchu przeglądarki,
usuwanie pozycji słownika):
  1. GET    :5000/api/dictionary-value/by-dictionary-id/<id>
         -> lista pozycji do wyboru (id, treść, kolejność)
  2. DELETE :5000/api/dictionary-value/<id_pozycji>
         -> usuwa jedną pozycję (bez treści w body)

BEZPIECZEŃSTWO: to skrypt PISZĄCY (kasujący!) dane w systemie. Usuwanie
pozycji słownika jest NIEODWRACALNE i może zepsuć inne wpisy, które się
do niej odwołują (np. zadanie pracownika wskazujące usuniętą pozycję
jako "Rodzaj zadania podrzędnego") — serwer może odrzucić takie
usunięcie, ale to sprawdza dopiero przy próbie. "Pokaż podgląd" nic nie
usuwa. Usuwanie wymaga jawnego potwierdzenia w oknie dialogowym z pełną
listą pozycji (nie tylko liczbą), i dodatkowego ostrzeżenia, gdy
środowisko to PRODUKCJA. Nic nie jest zaznaczone domyślnie.
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import threading, requests, json, os, sys, traceback, logging

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
CRASH_LOG_PATH  = os.path.join(SCRIPT_DIR, "usun_ze_slownika_gui_crash.log")
LAST_LOGIN_PATH = os.path.join(SCRIPT_DIR, "usun_ze_slownika_gui_last_login.json")


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
    LOG_PATH = os.path.join(SCRIPT_DIR, CFG.get("log_file", "usun_ze_slownika_gui.log"))
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

DICTIONARY_PRESETS = [
    ("51 — Usługi", 51),
    ("53 — Rodzaje zadań", 53),
]


def _card(parent):
    return ttkb.Labelframe(parent, text="", bootstyle="secondary")


# ── Backend ────────────────────────────────────────────────────────────────
TIMEOUT = 20


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

    def _delete(self, url, _retry=True):
        logging.debug(f"DELETE {url}")
        try:
            r = self.s.delete(url, timeout=TIMEOUT)
        except Exception as e:
            logging.error(f"Wyjątek przy DELETE {url}: {e}")
            return None, str(e)
        logging.debug(f"-> HTTP {r.status_code} {url}: {r.text[:500]}")
        if r.status_code == 401 and _retry and self._relogin():
            return self._delete(url, _retry=False)
        return r, None


def fetch_dictionary_values(client, dictionary_id):
    r, err = client._get(f"{client.servers['employee']}/api/dictionary-value/by-dictionary-id/{dictionary_id}")
    if err or r is None or r.status_code != 200:
        logging.error(f"Błąd pobierania wartości słownika {dictionary_id}: {err or (r and r.status_code)}")
        return None, (err or f"HTTP {r.status_code if r else '?'}")
    data = r.json()
    items = data if isinstance(data, list) else (data.get("items") or data.get("data") or [])
    logging.info(f"Słownik {dictionary_id}: {len(items)} pozycji")
    return items, None


def delete_dictionary_value(client, value_id):
    r, err = client._delete(f"{client.servers['employee']}/api/dictionary-value/{value_id}")
    if err:
        return False, err
    if r.status_code not in (200, 201, 204):
        try:
            detail = r.text[:300]
        except Exception:
            detail = ""
        return False, f"HTTP {r.status_code}: {detail}"
    return True, None


# ── Zaznaczanie z Excela ─────────────────────────────────────────────────────
# Odtwarza dokładnie te teksty (kolumny "Zadanie pracownik" / "Zadanie
# mieszkaniec"), które dodaj_rodzaje_zadan_gui.py wysyła jako 'content'
# rodzaju zadania (słownik 53) -- żeby móc zaznaczyć w tym narzędziu to, co
# dany plik Excela by utworzył, i to skasować.
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


def find_header_column_optional(ws, header_row, predicate):
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=header_row, column=col).value
        if header and predicate(str(header)):
            return col
    return None


def extract_task_names_from_excel(xlsx_path):
    """Zbiera wszystkie wartości z kolumn 'Zadanie pracownik' i 'Zadanie
    mieszkaniec' ze WSZYSTKICH arkuszy pliku -- pomija po cichu arkusze,
    które nie mają żadnej z tych kolumn (np. inny typ pliku). Zwraca zbiór
    tekstów znormalizowanych (strip + lower) do porównania z 'content' w
    słowniku."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    names = set()
    for ws in wb.worksheets:
        col_p = find_header_column_optional(ws, 1, lambda h: h.strip() == "Zadanie pracownik")
        col_m = find_header_column_optional(ws, 1, lambda h: h.strip() == "Zadanie mieszkaniec")
        if col_p is None and col_m is None:
            continue
        merge_map = _build_merge_map(ws)
        for r in range(2, ws.max_row + 1):
            for col in (col_p, col_m):
                if col is None:
                    continue
                v = _cell_value(ws, merge_map, r, col)
                if v is not None and str(v).strip():
                    names.add(str(v).strip().lower())
    return names


# ─────────────────────────────────────────────────────────────────────────────
class App(ttkb.Window):
    def __init__(self):
        super().__init__(title=f"Syrena — Usuwanie pozycji ze słownika [{ENV_LABEL}]",
                          themename=THEME_NAME, size=(900, 760), resizable=(True, True))

        self.client = Client()
        self._items = []
        self._item_rows = []
        self._checked_value_ids = set()   # ID pozycji słownika (stabilne -- PRZETRWA sortowanie/filtrowanie)
        self._sort_col = "content"
        self._sort_reverse = False

        self._build_ui()

    # ── BUILD UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        env_icon = "⚠ " if IS_PROD else ""
        hdr = ttkb.Frame(self, bootstyle=f"@{ACCENT_STYLE}", height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ttkb.Label(hdr, text="  Syrena — Usuwanie pozycji ze słownika", bootstyle=f"@{ACCENT_STYLE}",
                   font=FONT_TITLE).pack(side="left", padx=(12, 4))
        ttkb.Label(hdr, text=f"{env_icon}[{ENV_LABEL}]", bootstyle=f"@{ACCENT_STYLE}",
                   font=("Segoe UI", 12, "bold")).pack(side="left")
        self._lbl_status = ttkb.Label(hdr, text="●  Niezalogowany", bootstyle=f"@{ACCENT_STYLE}", font=FONT_SMALL)
        self._lbl_status.pack(side="right", padx=12)
        if IS_PROD:
            ttkb.Label(hdr, text="PRODUKCJA — usunięcie jest natychmiastowe i nieodwracalne",
                       bootstyle=f"@{ACCENT_STYLE}", font=FONT_TINY).pack(side="right", padx=12)

        nb = ttkb.Notebook(self, bootstyle=ACCENT_STYLE)
        nb.pack(fill="both", expand=True, padx=10, pady=8)
        self._nb = nb

        t1 = ttkb.Frame(nb); nb.add(t1, text="  1. Logowanie  ")
        t2 = ttkb.Frame(nb); nb.add(t2, text="  2. Usuwanie  ")

        self._tab_login(t1)
        self._tab_delete(t2)

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

    # ── TAB 2: Usuwanie ───────────────────────────────────────────────────────
    def _tab_delete(self, p):
        outer = ttkb.Frame(p, padding=(16, 14))
        outer.pack(fill="both", expand=True)

        sec_dict = ttkb.Labelframe(outer, text=" Słownik ", bootstyle=ACCENT_STYLE, padding=10)
        sec_dict.pack(fill="x", pady=(0, 8))
        fr = ttkb.Frame(sec_dict); fr.pack(fill="x")
        ttkb.Label(fr, text="ID słownika:", font=FONT_SMALL).pack(side="left")
        self._dict_var = tk.StringVar(value=DICTIONARY_PRESETS[0][0])
        self._cb_dict = ttkb.Combobox(fr, textvariable=self._dict_var, width=20, font=FONT_SMALL,
                                      bootstyle=ACCENT_STYLE,
                                      values=[label for label, _ in DICTIONARY_PRESETS])
        self._cb_dict.pack(side="left", padx=(6, 6))
        self._cb_dict.bind("<<ComboboxSelected>>", self._on_preset_pick)
        self._e_dict_id = ttkb.Entry(fr, width=8, font=FONT_SMALL)
        self._e_dict_id.insert(0, str(DICTIONARY_PRESETS[0][1]))
        self._e_dict_id.pack(side="left")
        btn_load = ttkb.Button(fr, text="↓ Wczytaj pozycje", command=self._load_items,
                               bootstyle="secondary", padding=(10, 5))
        btn_load.pack(side="left", padx=(10, 0))
        self._lbl_dict_status = ttkb.Label(fr, text="", font=FONT_TINY, bootstyle="secondary")
        self._lbl_dict_status.pack(side="left", padx=6)

        sec_items = ttkb.Labelframe(outer, text=" Pozycje (wielokrotny wybór) ", bootstyle=ACCENT_STYLE, padding=10)
        sec_items.pack(fill="both", expand=True, pady=(0, 8))
        self._e_filter = ttkb.Entry(sec_items, font=FONT_SMALL)
        self._e_filter.pack(fill="x", pady=(0, 4))
        self._e_filter.bind("<KeyRelease>", lambda _: self._apply_filter())

        tf = ttkb.Frame(sec_items)
        tf.pack(fill="both", expand=True)
        self._tree = ttkb.Treeview(tf, columns=("sel", "id", "content", "order"), show="headings",
                                   selectmode="none", height=14, bootstyle=ACCENT_STYLE)
        self._tree.heading("sel", text="✓")
        self._tree.heading("id", text="ID", command=lambda: self._sort_by("id"))
        self._tree.heading("content", text="Treść", command=lambda: self._sort_by("content"))
        self._tree.heading("order", text="Kolejność", command=lambda: self._sort_by("order"))
        self._tree.column("sel", width=26, stretch=False, anchor="center")
        self._tree.column("id", width=70, stretch=False, anchor="center")
        self._tree.column("content", width=500, stretch=True)
        self._tree.column("order", width=80, stretch=False, anchor="center")
        self._tree.bind("<Button-1>", self._on_row_click)
        self._tree.tag_configure("checked", background="#fde7e7")
        vsb = ttkb.Scrollbar(tf, orient="vertical", command=self._tree.yview, bootstyle=f"{ACCENT_STYLE}-round")
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tf.rowconfigure(0, weight=1); tf.columnconfigure(0, weight=1)

        bot = ttkb.Frame(sec_items); bot.pack(fill="x", pady=(4, 0))
        ttkb.Button(bot, text="Wszystkie", command=self._sel_all, bootstyle="link").pack(side="left", padx=1)
        ttkb.Button(bot, text="Żadne", command=self._desel_all, bootstyle="link").pack(side="left", padx=1)
        ttkb.Button(bot, text="📂 Zaznacz z Excela...", command=self._sel_from_excel,
                    bootstyle="secondary", padding=(8, 3)).pack(side="left", padx=(10, 1))
        self._lbl_sel = ttkb.Label(bot, text="0 zaznaczonych", font=FONT_TINY, bootstyle="secondary")
        self._lbl_sel.pack(side="right")

        sec_go = ttkb.Labelframe(outer, text=" Usuwanie ", bootstyle=ACCENT_STYLE, padding=10)
        sec_go.pack(fill="x")
        gob = ttkb.Frame(sec_go); gob.pack(fill="x")
        self._btn_delete = ttkb.Button(gob, text="🗑  Usuń zaznaczone", command=self._do_delete,
                  bootstyle="danger", padding=(14, 6), state="disabled")
        self._btn_delete.pack(side="left")
        self._progress = ttkb.Progressbar(gob, length=240, mode="determinate", bootstyle=f"{ACCENT_STYLE}-striped")
        self._progress.pack(side="left", padx=12)
        self._lbl_progress = ttkb.Label(gob, text="", font=FONT_SMALL)
        self._lbl_progress.pack(side="left")
        ttkb.Label(sec_go, text="Usunięcie jest NATYCHMIASTOWE i NIEODWRACALNE. Sprawdź listę w oknie "
                                "potwierdzenia zanim klikniesz [tak].",
                 bootstyle="danger", font=FONT_TINY).pack(anchor="w", pady=(6, 0))

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

    def _on_preset_pick(self, _event):
        label = self._dict_var.get()
        for lbl, dict_id in DICTIONARY_PRESETS:
            if lbl == label:
                self._e_dict_id.delete(0, "end")
                self._e_dict_id.insert(0, str(dict_id))
                return

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

    # ── POZYCJE SŁOWNIKA ──────────────────────────────────────────────────────
    def _load_items(self):
        try:
            dict_id = int(self._e_dict_id.get().strip())
        except ValueError:
            messagebox.showwarning("Słownik", "ID słownika musi być liczbą."); return
        self._lbl_dict_status.config(text="Pobieranie...")
        self._btn_delete.config(state="disabled")

        def _t():
            items, err = fetch_dictionary_values(self.client, dict_id)
            if err:
                self.after(0, lambda: self._lbl_dict_status.config(text="Błąd"))
                self.after(0, lambda: messagebox.showerror("Słownik", f"Nie udało się pobrać słownika {dict_id}:\n{err}"))
                self.after(0, lambda: self._log_msg(f"Błąd pobierania słownika {dict_id}: {err}", "err"))
                return
            items = sorted(items, key=lambda v: (v.get("content") or "").lower())
            self.after(0, lambda: self._render_items(items))
            self.after(0, lambda: self._lbl_dict_status.config(text=f"{len(items)} pozycji"))
            self.after(0, lambda: self._log_msg(f"Słownik {dict_id}: pobrano {len(items)} pozycji", "ok"))
        self._bg(_t)

    def _render_items(self, items):
        self._items = items
        self._checked_value_ids = set()
        self._refresh_tree(self._filtered_sorted_items())
        self._update_sort_headers()

    def _refresh_tree(self, items):
        for iid in self._tree.get_children():
            self._tree.delete(iid)
        self._item_rows = []
        for v in items:
            checked = v.get("id") in self._checked_value_ids
            iid = self._tree.insert("", "end", values=("☑" if checked else "☐", v.get("id", ""),
                                                        v.get("content", ""), v.get("displayOrder", "")),
                                    tags=("checked",) if checked else ())
            self._item_rows.append((iid, v))
        self._update_sel_label()

    def _filtered_sorted_items(self):
        q = self._e_filter.get().strip().lower()
        items = [v for v in self._items if q in (v.get("content") or "").lower()] if q else list(self._items)

        def key(v):
            if self._sort_col == "id":
                return v.get("id") or 0
            if self._sort_col == "order":
                return v.get("displayOrder") or 0
            return (v.get("content") or "").lower()

        items.sort(key=key, reverse=self._sort_reverse)
        return items

    def _apply_filter(self):
        self._refresh_tree(self._filtered_sorted_items())

    def _sort_by(self, col):
        """Kliknięcie nagłówka kolumny -- sortuje po niej (np. 'Kolejność',
        żeby zebrać razem ostatnio dodane pozycje po ich numerze sortowania
        i łatwo zaznaczyć tylko je do usunięcia). Ponowne kliknięcie tej
        samej kolumny odwraca kierunek."""
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col = col
            self._sort_reverse = False
        self._refresh_tree(self._filtered_sorted_items())
        self._update_sort_headers()

    def _update_sort_headers(self):
        labels = {"id": "ID", "content": "Treść", "order": "Kolejność"}
        for col, label in labels.items():
            arrow = ""
            if self._sort_col == col:
                arrow = " ▼" if self._sort_reverse else " ▲"
            self._tree.heading(col, text=label + arrow)

    def _on_row_click(self, event):
        iid = self._tree.identify_row(event.y)
        if not iid or self._tree.identify_region(event.x, event.y) != "cell":
            return
        v = next((v for row_iid, v in self._item_rows if row_iid == iid), None)
        if v is None:
            return
        vid = v.get("id")
        vals = list(self._tree.item(iid, "values"))
        if vid in self._checked_value_ids:
            self._checked_value_ids.discard(vid); vals[0] = "☐"
            self._tree.item(iid, values=vals, tags=())
        else:
            self._checked_value_ids.add(vid); vals[0] = "☑"
            self._tree.item(iid, values=vals, tags=("checked",))
        self._update_sel_label()

    def _sel_all(self):
        for iid, v in self._item_rows:
            self._checked_value_ids.add(v.get("id"))
            vals = list(self._tree.item(iid, "values")); vals[0] = "☑"
            self._tree.item(iid, values=vals, tags=("checked",))
        self._update_sel_label()

    def _desel_all(self):
        for iid, v in self._item_rows:
            self._checked_value_ids.discard(v.get("id"))
            vals = list(self._tree.item(iid, "values")); vals[0] = "☐"
            self._tree.item(iid, values=vals, tags=())
        self._update_sel_label()

    def _update_sel_label(self):
        n = len(self._checked_value_ids)
        self._lbl_sel.config(text=f"{n} zaznaczonych")
        self._btn_delete.config(state="normal" if n else "disabled")

    def _sel_from_excel(self):
        """Wczytuje plik Excela w formacie dodaj_rodzaje_zadan_gui.py
        (kolumny 'Zadanie pracownik' / 'Zadanie mieszkaniec') i zaznacza w
        aktualnie wczytanym słowniku te pozycje, których treść dokładnie
        odpowiada nazwom z tego pliku -- żeby móc skasować tylko to, co ten
        konkretny Excel by utworzył."""
        if not HAS_OPENPYXL:
            messagebox.showerror("Zaznacz z Excela",
                "Ta funkcja wymaga pakietu 'openpyxl'.\nZainstaluj go poleceniem:\n\n    pip install openpyxl")
            return
        if not self._item_rows:
            messagebox.showwarning("Zaznacz z Excela", "Najpierw wczytaj pozycje słownika."); return
        path = filedialog.askopenfilename(
            title="Wybierz plik Excel", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Wszystkie pliki", "*.*")])
        if not path:
            return
        self._log_msg(f"Wczytywanie nazw zadań z: {path}", "inf")

        def _t():
            try:
                names = extract_task_names_from_excel(path)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Zaznacz z Excela", f"Nie udało się odczytać pliku:\n{e}"))
                self.after(0, lambda: self._log_msg(f"Błąd odczytu {path}: {e}", "err"))
                return
            self.after(0, lambda: self._apply_excel_selection(names, path))
        self._bg(_t)

    def _apply_excel_selection(self, names, path):
        if not names:
            messagebox.showwarning("Zaznacz z Excela",
                "Nie znaleziono w tym pliku kolumny 'Zadanie pracownik' ani 'Zadanie mieszkaniec' "
                "(albo są puste) -- to nie jest plik w formacie dodaj_rodzaje_zadan_gui.py?")
            return
        matched = 0
        for v in self._items:   # cały wczytany słownik, nie tylko aktualnie widoczny (filtr/sortowanie)
            content = (v.get("content") or "").strip().lower()
            if content in names:
                matched += 1
                self._checked_value_ids.add(v.get("id"))
        self._refresh_tree(self._filtered_sorted_items())
        unmatched = len(names) - matched
        self._log_msg(
            f"Excel {os.path.basename(path)}: {len(names)} nazw zadań, zaznaczono {matched} pasujących "
            f"pozycji w aktualnie wczytanym słowniku"
            + (f" ({unmatched} nazw z Excela nie znaleziono w tym słowniku — może są w innym słowniku "
               f"albo już usunięte)" if unmatched > 0 else ""), "ok")

    def _get_selected_items(self):
        by_id = {v.get("id"): v for v in self._items}
        return [by_id[vid] for vid in self._checked_value_ids if vid in by_id]

    # ── USUWANIE ──────────────────────────────────────────────────────────────
    def _do_delete(self):
        selected = self._get_selected_items()
        if not selected:
            messagebox.showwarning("Usuwanie", "Nie zaznaczono żadnej pozycji."); return
        # Przy dużym zaznaczeniu (np. "Zaznacz z Excela" na kilkuset pozycjach)
        # wypisanie WSZYSTKICH w jednym messageboxie robi z niego nieczytelne/
        # zbyt duże okno -- przyciski [tak]/[nie] mogą wypadać poza ekran, co
        # wygląda jak "przycisk nic nie robi". Ograniczamy podgląd, licznik
        # zawsze pokazuje pełną liczbę.
        PREVIEW_LIMIT = 30
        preview = selected[:PREVIEW_LIMIT]
        listing = "\n".join(f"  [{v.get('id')}] {v.get('content','')}" for v in preview)
        if len(selected) > PREVIEW_LIMIT:
            listing += f"\n  ... i jeszcze {len(selected) - PREVIEW_LIMIT} innych"
        warn_prefix = "⚠ PRODUKCJA — usunięcie jest natychmiastowe.\n\n" if IS_PROD else ""
        if not messagebox.askyesno("Potwierdzenie usunięcia",
            f"{warn_prefix}Usunąć {len(selected)} pozycji ze słownika?\n"
            f"Tej operacji NIE MOŻNA cofnąć.\n\n{listing}"):
            return

        items = list(selected)
        self._btn_delete.config(state="disabled")
        self._progress.configure(maximum=len(items), value=0)
        self._lbl_progress.config(text="Usuwanie...")

        def _t():
            ok = fail = 0
            for i, v in enumerate(items, 1):
                vid, content = v.get("id"), v.get("content", "")
                success, err = delete_dictionary_value(self.client, vid)
                if success:
                    ok += 1
                    self.after(0, lambda content=content, vid=vid: self._log_msg(f"✓ [{vid}] {content} — usunięto", "ok"))
                else:
                    fail += 1
                    self.after(0, lambda content=content, vid=vid, err=err: self._log_msg(f"✗ [{vid}] {content} — {err}", "err"))
                self.after(0, lambda i=i, ok=ok, fail=fail: (
                    self._progress.configure(value=i),
                    self._lbl_progress.config(text=f"{i}/{len(items)}  ✓ {ok}  ✗ {fail}")))
            summary = f"Zakończono: {ok}/{len(items)} usunięto, {fail} błędów"
            self.after(0, lambda: self._log_msg(f"— {summary} —", "ok" if fail == 0 else "err"))
            self.after(0, lambda: messagebox.showinfo("Gotowe", summary))
            self.after(0, self._load_items)
        self._bg(_t)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception as e:
        logging.error(f"Nieobsłużony wyjątek przy starcie okna: {e}\n{traceback.format_exc()}")
        _fatal_startup_error("Błąd uruchomienia okna", str(e), exc=e)
