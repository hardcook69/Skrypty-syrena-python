#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Wyciąga z systemu SYRENA wszystkie AKTYWNE (uruchomione, planStatus=1) plany
razem z ich usługami, zadaniami i ISTNIEJĄCYMI WYZWALACZAMI (godziny/dni,
tak jak w zakładce "3 Harmonogramy" w harmonogramy_gui.py), i zapisuje jeden
plik Excel z dwoma arkuszami:

  jeden arkusz NA KAŻDE PIĘTRO (np. "V piętro"), w układzie Głównej Matrycy:
      jeden WIERSZ = jeden mieszkaniec (Mieszkaniec, Pokój), jedna KOLUMNA =
      jedna kombinacja (nazwa zadania, wzorzec dnia) - np. "Toalety
      indywidualne (poniedziałek)" i "(wtorek)" to dwie różne kolumny, bo mają
      inne godziny. W komórce: zakres godzin np. "13:00-13:30" (koniec liczony
      z czasu trwania zadania). Kolumny posortowane malejąco wg częstotliwości
      (codzienne na początku, miesięczne/rzadsze na końcu), w obrębie tej samej
      częstotliwości chronologicznie wg godziny w ciągu dnia.
  "Import"    — pusta matryca do wypełnienia (ten sam układ co
                matryca_import_harmonogramow.xlsx) — żeby dodać nowe
                usługi/zadania/harmonogramy odpowiedniej osobie.

Mieszkaniec i pokój — POTWIERDZONE zapytaniem z przechwytu w przeglądarce:
  GET :5020/api/beneficiary/by-organization-id/paged?...&statusList=1
  zwraca wprost {firstName, surname, roomNumber, roomName, storeyName, ...}
  dla każdego aktywnego mieszkańca. Pobierane RAZ na starcie, potem łączone
  z planami po plan["visitorId"] == beneficiary["id"] (zgodność sprawdzana
  empirycznie w tym przebiegu — patrz podsumowanie na końcu w konsoli:
  ile planów dopasowano do rejestru mieszkańców).
  Uwaga: część rekordów ma puste firstName/surname w samym systemie
  (potwierdzone w przechwycie) — to nie błąd skryptu.

DIAGNOSTYKA: przy pierwszych 2 planach zapisuje surowy JSON (plan + pierwsza
usługa + pierwsze zadanie) do debug_pierwszy_rekord.json.

Tylko odczyt (same GET) — bezpieczne do uruchamiania bez ograniczeń.

Użycie:
  python wyciagnij_aktywne_plany.py                    # config_testowy.json
  python wyciagnij_aktywne_plany.py config_produkcja.json
"""
import sys
import time
import getpass
import logging
import os
import json
from datetime import datetime

try:
    import requests
except ImportError:
    print("pip install requests"); sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("pip install openpyxl"); sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_NAME = sys.argv[1] if len(sys.argv) > 1 else "config_testowy.json"
CONFIG_PATH = os.path.join(SCRIPT_DIR, CONFIG_NAME)
DEBUG_JSON_PATH = os.path.join(SCRIPT_DIR, "debug_pierwszy_rekord.json")
TIMEOUT = 20

TASK_KIND_ID_CANDIDATES = ("taskKindId", "taskDefinitionKindId", "kindId")

logger = logging.getLogger("wyciagnij_aktywne_plany")

DOW_EN_TO_PL_FULL = {"MON": "poniedziałek", "TUE": "wtorek", "WED": "środa", "THU": "czwartek",
                      "FRI": "piątek", "SAT": "sobota", "SUN": "niedziela"}
MON_EN_TO_PL = {"JAN": "Sty", "FEB": "Luty", "MAR": "Marz", "APR": "Kwi", "MAY": "Maj", "JUN": "Cze",
                "JUL": "Lip", "AUG": "Sie", "SEP": "Wrz", "OCT": "Paź", "NOV": "Lis", "DEC": "Gru"}


def _parse_times(hr, mn, duration_minutes):
    """Zwraca (lista_zakresów_HH:MM-HH:MM, minuty_pierwszego_wystąpienia_od_północy).

    Pole godziny (i/lub minuty) w CRON może być listą oddzielaną przecinkami -
    POTWIERDZONE w natywnym UI SYRENA: wyzwalacz "3 razy dziennie" pokazuje się
    tam jako "O 07:25, 16:25 i 19:25" - czyli godzina="7,16,19", minuta="25"
    (iloczyn kartezjański obu list daje wszystkie pary godzina:minuta)."""
    try:
        hours = [int(h) for h in str(hr).split(",")]
        minutes = [int(m) for m in str(mn).split(",")]
    except (ValueError, TypeError):
        return [], 0
    duration = duration_minutes if isinstance(duration_minutes, (int, float)) else 0
    starts = sorted(h * 60 + m for h in hours for m in minutes)
    ranges = []
    for start_total in starts:
        end_total = (start_total + int(duration)) % (24 * 60)
        ranges.append(f"{start_total // 60:02d}:{start_total % 60:02d}-{end_total // 60:02d}:{end_total % 60:02d}")
    return ranges, (starts[0] if starts else 0)


def trigger_columns(task_name, cron, duration_minutes):
    """Rozbija JEDEN wyzwalacz na listę kolumn siatki: (col_key, col_label,
    freq_score, sort_minutes, wartość_komórki).

    Kolumna = unikalna kombinacja (nazwa zadania, wzorzec dnia) - np. "Toalety
    indywidualne (poniedziałek)" i "(wtorek)" to DWIE różne kolumny, bo mają
    różne godziny u różnych mieszkańców (tak jak w ręcznie prowadzonej Głównej
    Matrycy). freq_score = przybliżona liczba wystąpień/miesiąc, do sortowania
    malejąco (najczęstsze pierwsze, miesięczne/rzadsze na końcu) - patrz main().
    Uwzględnia wyzwalacze z kilkoma porami dziennie (np. "3 razy dziennie")."""
    p = (cron or "").split()
    if len(p) < 6:
        return []
    _, mn, hr, dom, mon, dow = p[:6]
    time_ranges, sort_minutes = _parse_times(hr, mn, duration_minutes)
    tr = "; ".join(time_ranges)
    occurrences = max(1, len(time_ranges))

    def day_col(day_label, freq_score, suffix=""):
        key = f"{task_name}|{day_label}"
        label = f"{task_name} ({day_label}{suffix})" if day_label else task_name
        return (key, label, freq_score * occurrences, sort_minutes, tr)

    def monthly_col(freq_score, text):
        return (f"{task_name}|miesięczne", f"{task_name} (miesięczne)", freq_score * occurrences, 999999, text)

    if dow not in ("*", "?"):
        if "#" in dow:
            parts = dow.split(",")
            if len(parts) > 1:
                d = parts[0].split("#")[0]
                return [day_col(DOW_EN_TO_PL_FULL.get(d, d), 2.17, ", co 2 tyg.")]
            d, n = dow.split("#")
            return [monthly_col(1, f"{n}. {DOW_EN_TO_PL_FULL.get(d, d)} mies.  {tr}")]
        if "," in dow:
            days = [d for d in dow.split(",") if d in DOW_EN_TO_PL_FULL]
            return [day_col(DOW_EN_TO_PL_FULL.get(d, d), len(days) * 4.33) for d in days]
        return [day_col(DOW_EN_TO_PL_FULL.get(dow, dow), 4.33)]

    if dom not in ("?", "*"):
        if "/" in dom:
            s, n = dom.split("/")
            try:
                freq = 30 / int(n) if int(n) else 1
            except ValueError:
                freq = 1
            return [monthly_col(freq, f"co {n} dni od {s}.  {tr}")]
        if mon not in ("*", "?"):
            return [monthly_col(0.1, f"{dom} {MON_EN_TO_PL.get(mon, mon)}  {tr}")]
        return [monthly_col(1, f"{dom}. dzień mies.  {tr}")]

    return [day_col("", 30)]


def load_cfg():
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            cfg = json.load(f)
    except FileNotFoundError:
        print(f"Nie znaleziono {CONFIG_PATH}."); sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"Błędny format {CONFIG_NAME}: {e}"); sys.exit(1)
    cfg.setdefault("organization_id", 1)
    cfg.setdefault("delay_between_requests", 0.15)
    cfg.setdefault("beneficiary_url", cfg["auth_url"].replace(":5010", ":5020"))
    cfg.setdefault("log_file", "wyciagnij_aktywne_plany.log")
    return cfg


class Client:
    def __init__(self, cfg):
        self.cfg = cfg
        self.s = requests.Session()
        self.s.headers.update({"Content-Type": "application/json", "Accept": "application/json"})
        self._creds = None
        self._current_org_id = None

    def login(self, username, password):
        for key in ("login", "userName", "username"):
            try:
                r = self.s.post(f"{self.cfg['auth_url']}/api/authentication/token",
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
            r = self.s.post(f"{self.cfg['auth_url']}/api/authentication/token",
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

    def _get(self, base_url, path, params=None, _retry=True):
        url = f"{base_url}{path}"
        logger.debug(f"GET {url} params={params}")
        try:
            r = self.s.get(url, params=params, timeout=TIMEOUT)
        except Exception as e:
            logger.error(f"Wyjątek przy GET {url}: {e}")
            return None, str(e)
        if r.status_code == 401 and _retry and self._relogin():
            return self._get(base_url, path, params=params, _retry=False)
        if r.status_code != 200:
            return None, f"HTTP {r.status_code}: {r.text[:300]}"
        try:
            return r.json(), None
        except Exception as e:
            return None, f"Odpowiedź nie jest JSON-em: {e}"


def fetch_active_plans(client, org_id):
    all_plans, page = [], 1
    while True:
        d, e = client._get(client.cfg["shift_url"], "/api/plan/by-visitor-organization-id/paged",
                            {"page": page, "pageSize": 200, "orderBy": "id", "ascending": "true",
                             "planStatus": 1, "visitorType": 3})
        if e:
            logger.error(f"Błąd pobierania planów str.{page}: {e}")
            break
        items = d.get("results", []) if isinstance(d, dict) else []
        all_plans.extend(items)
        total_pages = d.get("totalNumberOfPages") if isinstance(d, dict) else None
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < 200:
            break
        page += 1
    return all_plans


def fetch_services(client, plan_id):
    d, e = client._get(client.cfg["shift_url"], f"/api/service-definition/by-plan-id/{plan_id}")
    if e:
        logger.warning(f"Usługi planu {plan_id}: {e}")
        return []
    return d if isinstance(d, list) else []


def fetch_tasks(client, service_id):
    d, e = client._get(client.cfg["shift_url"], f"/api/task-definition/by-service-definition-id/{service_id}")
    if e:
        logger.warning(f"Zadania usługi {service_id}: {e}")
        return []
    return d if isinstance(d, list) else []


def fetch_triggers(client, task_def_id):
    d, e = client._get(client.cfg["shift_url"], f"/api/schedule-trigger/by-task-definition-id/{task_def_id}")
    if e:
        logger.warning(f"Wyzwalacze zadania {task_def_id}: {e}")
        return []
    return d if isinstance(d, list) else []


def fetch_beneficiary_roster(client, org_id):
    """GET :5020/api/beneficiary/by-organization-id/paged?...&statusList=1 —
    potwierdzone przechwytem w przeglądarce: zwraca wprost firstName, surname,
    roomNumber, roomName, storeyName dla każdego aktywnego mieszkańca.
    Zwraca dict {beneficiaryId: {surname, firstname, room, storey}}."""
    roster, page = {}, 1
    while True:
        d, e = client._get(client.cfg["beneficiary_url"], "/api/beneficiary/by-organization-id/paged",
                            {"page": page, "pageSize": 200, "orderBy": "pesel",
                             "ascending": "false", "statusList": "1"})
        if e:
            logger.error(f"Błąd pobierania rejestru mieszkańców str.{page}: {e}")
            break
        items = d.get("results", []) if isinstance(d, dict) else []
        for b in items:
            surname = (b.get("surname") or "").strip()
            firstname = (b.get("firstName") or "").strip()
            name = f"{surname} {firstname}".strip() or f"(brak nazwiska, id={b.get('id')})"
            roster[b["id"]] = {
                "name": name,
                "room": b.get("roomNumber") or "",
                "roomName": b.get("roomName") or "",
                "storey": b.get("storeyName") or "",
            }
        total_pages = d.get("totalNumberOfPages") if isinstance(d, dict) else None
        if total_pages is not None:
            if page >= total_pages:
                break
        elif len(items) < 200:
            break
        page += 1
    return roster


def extract_task_kind_id(task):
    for key in TASK_KIND_ID_CANDIDATES:
        v = task.get(key)
        if v not in (None, ""):
            return v, key
    return "", None


def build_import_sheet(wb):
    """Ten sam układ co matryca_import_harmonogramow.xlsx - pusta, gotowa do
    wypełnienia (wartości kopiowane z arkusza ObecnyStan)."""
    headers = [
        "GrupaPlanu", "PeselMieszkanca", "PlanBazowyId", "NowaNazwaPlanu",
        "ServiceKindId", "TaskKindId", "CzasNormatywny(min)", "Priorytet(1-4)",
        "WyzwalaczNazwa", "Tryb",
        "Godzina(0-23)", "Minuta(0-59)", "DniTygodnia", "DzienMiesiaca(1-31)",
        "Miesiac", "KtoreWystapienie(1-5)", "RokOd", "RokDo",
        "UruchomPlanPoImporcie(TAK/NIE)",
    ]
    widths = [11, 16, 12, 22, 13, 11, 18, 14, 20, 22, 13, 12, 26, 18, 12, 19, 8, 8, 26]
    ws = wb.create_sheet("Import")
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1B5E20")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 42
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def _safe_sheet_name(name, used):
    base = (name or "Nieznane piętro").strip()[:31] or "Nieznane piętro"
    candidate = base
    n = 2
    while candidate in used:
        suffix = f" ({n})"
        candidate = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def build_storey_sheets(wb, storey_columns, storey_residents):
    """Jeden arkusz na piętro, w układzie Głównej Matrycy: jeden wiersz = jeden
    mieszkaniec, jedna kolumna = jedna kombinacja (zadanie, dzień). Kolumny
    posortowane malejąco wg częstotliwości (najczęstsze - np. codzienne - na
    początku, miesięczne/rzadsze na końcu), w obrębie tej samej częstotliwości
    chronologicznie wg godziny w ciągu dnia."""
    used_names = set()
    for storey in sorted(storey_columns.keys(), key=lambda s: s.lower()):
        cols = storey_columns[storey]
        col_keys = sorted(cols.keys(), key=lambda k: (-cols[k]["freq_score"], cols[k]["sort_minutes"]))
        headers = ["Mieszkaniec", "Pokoj"] + [cols[k]["label"] for k in col_keys]

        ws = wb.create_sheet(_safe_sheet_name(storey, used_names))
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1B5E20")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[1].height = 30

        residents = sorted(storey_residents[storey].keys(), key=lambda rk: rk[0].lower())
        for resident, room in residents:
            cells = storey_residents[storey][(resident, room)]
            ws.append([resident, room] + ["; ".join(cells.get(k, [])) for k in col_keys])

        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 8
        for i in range(len(col_keys)):
            ws.column_dimensions[get_column_letter(i + 3)].width = 16
        ws.freeze_panes = "C2"


def main():
    logging.basicConfig(level=logging.DEBUG, encoding="utf-8",
                         format="%(asctime)s %(levelname)s %(message)s")
    cfg = load_cfg()
    fh = logging.FileHandler(os.path.join(SCRIPT_DIR, cfg["log_file"]), encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logging.getLogger().addHandler(fh)

    print(f"Środowisko: {cfg.get('env_label', '?')}  |  {cfg['auth_url']}")
    user = input("Login: ").strip()
    password = getpass.getpass("Hasło: ")

    client = Client(cfg)
    if not client.login(user, password):
        print("Błąd logowania."); sys.exit(1)
    org_id = cfg.get("organization_id", 1)
    if not client.select_organization(org_id):
        print(f"Nie udało się przełączyć na organizationId={org_id}."); sys.exit(1)
    print("Zalogowano.")

    delay = cfg.get("delay_between_requests", 0.15)

    print("Pobieranie rejestru mieszkańców (nazwisko, pokój)...")
    roster = fetch_beneficiary_roster(client, org_id)
    print(f"Rejestr: {len(roster)} aktywnych mieszkańców.")

    print("Pobieranie aktywnych planów...")
    plans = fetch_active_plans(client, org_id)
    print(f"Znaleziono {len(plans)} aktywnych planów.")

    debug_records = []
    storey_columns = {}    # storey -> {col_key: {"label":..., "freq_score":..., "sort_minutes":...}}
    storey_residents = {}  # storey -> {(resident, room): {col_key: [wartości]}}
    taskkind_source_seen = set()
    matched, unmatched = 0, 0
    total_tasks = 0

    for i, plan in enumerate(plans, 1):
        plan_id = plan.get("id")
        visitor_id = plan.get("visitorId")
        person = roster.get(visitor_id)
        if person:
            matched += 1
            resident, room, storey = person["name"], person["room"], person["storey"] or "(nieznane piętro)"
        else:
            unmatched += 1
            surname = (plan.get("visitorSurname") or "").strip()
            firstname = (plan.get("visitorFirstName") or "").strip()
            resident = f"{surname} {firstname}".strip() or f"(brak dopasowania, visitorId={visitor_id})"
            room, storey = "", "(nieznane piętro)"
        print(f"  [{i}/{len(plans)}] Plan {plan_id} — {resident}")

        services = fetch_services(client, plan_id)
        time.sleep(delay)

        for si, svc in enumerate(services):
            tasks = fetch_tasks(client, svc["id"])
            time.sleep(delay)
            if len(debug_records) < 2 and si == 0:
                debug_records.append({"plan": plan, "service": svc, "task": tasks[0] if tasks else None})
            for t in tasks:
                triggers = fetch_triggers(client, t["id"])
                time.sleep(delay)
                tkid, tk_src = extract_task_kind_id(t)
                taskkind_source_seen.add(tk_src)
                duration = t.get("normativeTime")
                task_name = t.get("taskKindName", "") or svc.get("serviceKindName", "")
                if not triggers:
                    continue  # brak harmonogramu = nic do pokazania w tej siatce godzin

                total_tasks += 1
                cols = storey_columns.setdefault(storey, {})
                cells = storey_residents.setdefault(storey, {}).setdefault((resident, room), {})
                for tr in triggers:
                    for key, label, freq_score, sort_minutes, value in trigger_columns(
                            task_name, tr.get("cron", ""), duration):
                        existing = cols.get(key)
                        if existing is None or freq_score > existing["freq_score"]:
                            cols[key] = {"label": label, "freq_score": freq_score, "sort_minutes": sort_minutes}
                        cells.setdefault(key, []).append(value)

    with open(DEBUG_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(debug_records, f, ensure_ascii=False, indent=2, default=str)
    print(f"Zrzut surowego JSON (do diagnostyki pól TaskKindId): {DEBUG_JSON_PATH}")
    print(f"Źródła TaskKindId użyte w tym przebiegu: {taskkind_source_seen}")
    print(f"Dopasowanie planów do rejestru mieszkańców: {matched} dopasowanych, {unmatched} niedopasowanych"
          + (" (visitorId != id z rejestru dla części planów)" if unmatched else ""))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_storey_sheets(wb, storey_columns, storey_residents)
    build_import_sheet(wb)

    out_name = f"aktywne_plany_slownik_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    out_path = os.path.join(SCRIPT_DIR, out_name)
    wb.save(out_path)
    print(f"\nZapisano {total_tasks} zaplanowanych zadań w {len(storey_columns)} arkuszach (piętrach) do: {out_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nPrzerwano.")
    except Exception as e:
        logging.error(f"Nieobsłużony wyjątek: {e}", exc_info=True)
        print(f"Błąd: {e}")
        sys.exit(1)
