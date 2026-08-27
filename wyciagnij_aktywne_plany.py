#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Wyciąga z systemu SYRENA wszystkie AKTYWNE (uruchomione, planStatus=1) plany
razem z ich usługami i zadaniami, i zapisuje jako słownik referencyjny do
Excela — pogrupowany po mieszkańcu (nazwisko + imię), żeby przy wypełnianiu
matrycy importu harmonogramów nie trzeba było zgadywać ID (PlanBazowyId,
ServiceKindId, TaskKindId), tylko znaleźć właściwy wiersz po nazwisku.

Pokój mieszkańca jest dołączany jako pomocnicza kolumna do odróżnienia dwóch
osób o tym samym imieniu i nazwisku — best-effort, dwa źródła po kolei:
  1. pole w samym obiekcie planu, jeśli API je tam umieszcza (szukane po
     nazwie klucza zawierającej "room"/"pokoj" — nazwa pola niepotwierdzona)
  2. GET :5020/api/beneficiary-residence/by-beneficiary-id/{visitorId} —
     zakłada, że visitorId (używany przez /api/plan, /api/service) to ten
     sam numer co beneficiaryId (używany przez usługi/obserwacje). TO
     ZAŁOŻENIE NIE JEST POTWIERDZONE — jeśli się nie sprawdzi, kolumna
     Pokoj zostanie po prostu pusta dla tego wiersza (log ostrzega o tym raz).

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
from datetime import datetime
import json

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
TIMEOUT = 20
ROOM_FIELD_CANDIDATES = ("roomId", "room_id", "roomName", "executionRoomId",
                          "residenceRoomId", "currentRoomId")

logger = logging.getLogger("wyciagnij_aktywne_plany")


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
    d, e = client._get(client.cfg["shift_url"], "/api/plan/by-visitor-organization-id/paged",
                        {"page": 1, "pageSize": 500, "orderBy": "id", "ascending": "true",
                         "planStatus": 1, "visitorType": 3})
    if e:
        logger.error(f"Błąd pobierania planów: {e}")
        return []
    return d.get("results", []) if isinstance(d, dict) else []


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


_room_cache = {}
_room_fallback_warned = False


def find_room_best_effort(client, plan):
    """Zwraca (pokoj_string, zrodlo) albo (None, None). Best-effort, patrz
    docstring modułu — kolejność źródeł i ostrzeżenie o niepewności."""
    global _room_fallback_warned
    for key in ROOM_FIELD_CANDIDATES:
        v = plan.get(key)
        if v not in (None, ""):
            return str(v), f"plan.{key}"

    visitor_id = plan.get("visitorId")
    if visitor_id is None:
        return None, None
    if visitor_id in _room_cache:
        return _room_cache[visitor_id]

    d, e = client._get(client.cfg["beneficiary_url"],
                        f"/api/beneficiary-residence/by-beneficiary-id/{visitor_id}")
    if e or not d:
        if not _room_fallback_warned:
            logger.warning("beneficiary-residence/by-beneficiary-id z visitorId nie zwraca danych — "
                            "prawdopodobnie visitorId != beneficiaryId. Kolumna Pokoj zostanie pusta.")
            _room_fallback_warned = True
        _room_cache[visitor_id] = (None, None)
        return None, None

    rec = d[0] if isinstance(d, list) and d else d if isinstance(d, dict) else {}
    for key, val in rec.items():
        if "room" in key.lower() and val not in (None, ""):
            result = (str(val), f"beneficiary-residence.{key}")
            _room_cache[visitor_id] = result
            return result
    _room_cache[visitor_id] = (None, None)
    return None, None


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

    print("Pobieranie aktywnych planów...")
    plans = fetch_active_plans(client, org_id)
    print(f"Znaleziono {len(plans)} aktywnych planów.")

    rows = []
    for i, plan in enumerate(plans, 1):
        plan_id = plan.get("id")
        surname = (plan.get("visitorSurname") or "").strip()
        firstname = (plan.get("visitorFirstName") or "").strip()
        resident = f"{surname} {firstname}".strip() or f"(brak nazwiska, visitorId={plan.get('visitorId')})"
        room, room_src = find_room_best_effort(client, plan)
        print(f"  [{i}/{len(plans)}] Plan {plan_id} — {resident}")

        services = fetch_services(client, plan_id)
        time.sleep(delay)
        if not services:
            rows.append({
                "Mieszkaniec": resident, "Pokoj": room or "",
                "PlanId": plan_id, "NazwaPlanu": plan.get("name", ""),
                "WaznyOd": (plan.get("validFrom") or "")[:10], "WaznyDo": (plan.get("validTo") or "")[:10],
                "ServiceKindId": "", "NazwaUslugi": "(brak usług)",
                "TaskKindId": "", "NazwaZadania": "", "CzasNormatywny": "",
            })
            continue

        for svc in services:
            tasks = fetch_tasks(client, svc["id"])
            time.sleep(delay)
            if not tasks:
                rows.append({
                    "Mieszkaniec": resident, "Pokoj": room or "",
                    "PlanId": plan_id, "NazwaPlanu": plan.get("name", ""),
                    "WaznyOd": (plan.get("validFrom") or "")[:10], "WaznyDo": (plan.get("validTo") or "")[:10],
                    "ServiceKindId": svc.get("serviceKindId", ""),
                    "NazwaUslugi": svc.get("serviceKindName", ""),
                    "TaskKindId": "", "NazwaZadania": "(brak zadań)", "CzasNormatywny": "",
                })
                continue
            for t in tasks:
                rows.append({
                    "Mieszkaniec": resident, "Pokoj": room or "",
                    "PlanId": plan_id, "NazwaPlanu": plan.get("name", ""),
                    "WaznyOd": (plan.get("validFrom") or "")[:10], "WaznyDo": (plan.get("validTo") or "")[:10],
                    "ServiceKindId": svc.get("serviceKindId", ""),
                    "NazwaUslugi": svc.get("serviceKindName", ""),
                    "TaskKindId": t.get("taskKindId", ""),
                    "NazwaZadania": t.get("taskKindName", ""),
                    "CzasNormatywny": t.get("normativeTime", ""),
                })

    rows.sort(key=lambda r: (r["Mieszkaniec"].lower(), r["PlanId"] or 0,
                              r["ServiceKindId"] if isinstance(r["ServiceKindId"], int) else 0,
                              r["TaskKindId"] if isinstance(r["TaskKindId"], int) else 0))

    headers = ["Mieszkaniec", "Pokoj", "PlanId", "NazwaPlanu", "WaznyOd", "WaznyDo",
               "ServiceKindId", "NazwaUslugi", "TaskKindId", "NazwaZadania", "CzasNormatywny"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AktywnePlany"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1B5E20")
    for r in rows:
        ws.append([r[h] for h in headers])
    widths = [24, 10, 9, 26, 12, 12, 13, 26, 11, 26, 15]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    out_name = f"aktywne_plany_slownik_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    out_path = os.path.join(SCRIPT_DIR, out_name)
    wb.save(out_path)
    print(f"\nZapisano {len(rows)} wierszy do: {out_path}")
    if _room_fallback_warned:
        print("UWAGA: kolumna 'Pokoj' jest pusta dla części/wszystkich wierszy — "
              "zobacz log, prawdopodobnie visitorId != beneficiaryId (niepotwierdzone założenie).")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nPrzerwano.")
    except Exception as e:
        logging.error(f"Nieobsłużony wyjątek: {e}", exc_info=True)
        print(f"Błąd: {e}")
        sys.exit(1)
