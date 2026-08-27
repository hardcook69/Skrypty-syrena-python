#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Wyciąga z systemu SYRENA wszystkie AKTYWNE (uruchomione, planStatus=1) plany
razem z ich usługami, zadaniami i ISTNIEJĄCYMI WYZWALACZAMI (godziny/dni,
tak jak w zakładce "3 Harmonogramy" w harmonogramy_gui.py), i zapisuje jeden
plik Excel z dwoma arkuszami:

  jeden arkusz NA KAŻDE PIĘTRO (np. "V piętro") — siatka:
      Mieszkaniec | Nazwa zadania | Pon | Wt | Śr | Czw | Pt | Sob | Ndz | Miesięczne
      Godziny w formacie zakresu np. "13:00-13:30" (koniec liczony z czasu trwania
      zadania). Harmonogramy nie-tygodniowe (N-ty dzień miesiąca, konkretna data,
      co N dni) trafiają do osobnej kolumny "Miesięczne", bo nie pasują do siatki
      dni tygodnia.
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

DOW_GRID = ["Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Ndz"]
DOW_CRON_TO_GRID = {"MON": 0, "TUE": 1, "WED": 2, "THU": 3, "FRI": 4, "SAT": 5, "SUN": 6}
DOW_EN_TO_PL_FULL = {"MON": "poniedziałek", "TUE": "wtorek", "WED": "środa", "THU": "czwartek",
                      "FRI": "piątek", "SAT": "sobota", "SUN": "niedziela"}
MON_EN_TO_PL = {"JAN": "Sty", "FEB": "Luty", "MAR": "Marz", "APR": "Kwi", "MAY": "Maj", "JUN": "Cze",
                "JUL": "Lip", "AUG": "Sie", "SEP": "Wrz", "OCT": "Paź", "NOV": "Lis", "DEC": "Gru"}


def _time_range(hr, mn, duration_minutes):
    """np. 13:00-13:30 - koniec liczony z czasu trwania zadania (normativeTime)."""
    try:
        start_total = int(hr) * 60 + int(mn)
    except (ValueError, TypeError):
        return ""
    duration = duration_minutes if isinstance(duration_minutes, (int, float)) else 0
    end_total = (start_total + int(duration)) % (24 * 60)
    return f"{start_total // 60:02d}:{start_total % 60:02d}-{end_total // 60:02d}:{end_total % 60:02d}"


def classify_trigger(cron, duration_minutes):
    """Rozkłada CRON (wygenerowany przez TriggerBuilder w harmonogramy_gui.py) na
    dni tygodnia z zakresem godzin, ALBO tekst dla osobnej kolumny "Miesięczne"
    (N-ty dzień miesiąca / N-ty dzień tyg. mies. / konkretna data / co N dni) -
    te tryby nie są "co tydzień w X", więc nie pasują do siatki dni tygodnia.
    Zwraca (indeksy_dni: list[int] 0=Pon..6=Ndz, tekst_godzin, tekst_miesięczny|None)."""
    p = (cron or "").split()
    if len(p) < 6:
        return [], "", None
    _, mn, hr, dom, mon, dow = p[:6]
    tr = _time_range(hr, mn, duration_minutes)

    if dow not in ("*", "?"):
        if "#" in dow:
            parts = dow.split(",")
            if len(parts) > 1:
                # co dwa tygodnie - nadal konkretny dzień tygodnia, tylko nie co tydzień
                d = parts[0].split("#")[0]
                idx = DOW_CRON_TO_GRID.get(d)
                return ([idx] if idx is not None else []), f"{tr} (co 2 tyg.)", None
            d, n = dow.split("#")
            return [], "", f"{n}. {DOW_EN_TO_PL_FULL.get(d, d)} mies.  {tr}"
        if "," in dow:
            idxs = [DOW_CRON_TO_GRID[d] for d in dow.split(",") if d in DOW_CRON_TO_GRID]
            return idxs, tr, None
        idx = DOW_CRON_TO_GRID.get(dow)
        return ([idx] if idx is not None else []), tr, None

    if dom not in ("?", "*"):
        if "/" in dom:
            s, n = dom.split("/")
            return [], "", f"co {n} dni od {s}.  {tr}"
        if mon not in ("*", "?"):
            return [], "", f"{dom} {MON_EN_TO_PL.get(mon, mon)}  {tr}"
        return [], "", f"{dom}. dzień mies.  {tr}"

    return list(range(7)), tr, None


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


def build_storey_sheets(wb, storey_rows):
    """Jeden arkusz na piętro: Mieszkaniec | Nazwa zadania | Pon..Ndz | Miesięczne -
    tak jak w Szablon_zada_.xlsx, podzielone na moduły/piętra jak w Głównej Matrycy."""
    headers = ["Mieszkaniec", "Nazwa zadania"] + DOW_GRID + ["Miesięczne"]
    widths = [22, 30] + [16] * 7 + [26]
    used_names = set()
    for storey in sorted(storey_rows.keys(), key=lambda s: s.lower()):
        rows = sorted(storey_rows[storey], key=lambda r: (r["Mieszkaniec"].lower(), r["NazwaZadania"].lower()))
        ws = wb.create_sheet(_safe_sheet_name(storey, used_names))
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1B5E20")
        for r in rows:
            ws.append([r["Mieszkaniec"], r["NazwaZadania"]] + [r["days"][i] for i in range(7)] + [r["Miesieczne"]])
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"


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
    storey_rows = {}
    taskkind_source_seen = set()
    matched, unmatched = 0, 0
    total_tasks = 0

    for i, plan in enumerate(plans, 1):
        plan_id = plan.get("id")
        visitor_id = plan.get("visitorId")
        person = roster.get(visitor_id)
        if person:
            matched += 1
            resident, storey = person["name"], person["storey"] or "(nieznane piętro)"
        else:
            unmatched += 1
            surname = (plan.get("visitorSurname") or "").strip()
            firstname = (plan.get("visitorFirstName") or "").strip()
            resident = f"{surname} {firstname}".strip() or f"(brak dopasowania, visitorId={visitor_id})"
            storey = "(nieznane piętro)"
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

                days = {d: [] for d in range(7)}
                monthly = []
                for tr in triggers:
                    idxs, day_time, monthly_text = classify_trigger(tr.get("cron", ""), duration)
                    for idx in idxs:
                        if day_time:
                            days[idx].append(day_time)
                    if monthly_text:
                        monthly.append(monthly_text)
                if not triggers:
                    monthly = ["(brak harmonogramu)"]

                total_tasks += 1
                storey_rows.setdefault(storey, []).append({
                    "Mieszkaniec": resident,
                    "NazwaZadania": t.get("taskKindName", "") or svc.get("serviceKindName", ""),
                    "days": {d: "; ".join(v) for d, v in days.items()},
                    "Miesieczne": "; ".join(monthly),
                })

    with open(DEBUG_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(debug_records, f, ensure_ascii=False, indent=2, default=str)
    print(f"Zrzut surowego JSON (do diagnostyki pól TaskKindId): {DEBUG_JSON_PATH}")
    print(f"Źródła TaskKindId użyte w tym przebiegu: {taskkind_source_seen}")
    print(f"Dopasowanie planów do rejestru mieszkańców: {matched} dopasowanych, {unmatched} niedopasowanych"
          + (" (visitorId != id z rejestru dla części planów)" if unmatched else ""))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_storey_sheets(wb, storey_rows)
    build_import_sheet(wb)

    out_name = f"aktywne_plany_slownik_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    out_path = os.path.join(SCRIPT_DIR, out_name)
    wb.save(out_path)
    print(f"\nZapisano {total_tasks} zadań w {len(storey_rows)} arkuszach (piętrach) do: {out_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nPrzerwano.")
    except Exception as e:
        logging.error(f"Nieobsłużony wyjątek: {e}", exc_info=True)
        print(f"Błąd: {e}")
        sys.exit(1)
