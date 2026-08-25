#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dodawanie zadań mieszkańcom pod usługę "Zadania ad hoc" — na wybrany
okres, częstotliwość i porę dnia. Jeden rodzaj zadania na uruchomienie
(API przyjmuje jedno zadanie na request), ale dla wielu mieszkańców
i wielu dat naraz.

Przepływ (dokładnie taki jak w przechwyconym ruchu przeglądarki):
  1. GET  :5020/api/beneficiary/by-organization-id/paged?statusList=1
         -> lista aktywnych mieszkańców do wyboru
  2. GET  :5070/api/service/by-visitor-organization-id/paged
         ?planStatus=1&visitorType=3
         -> usługi mieszkańców; szukamy tej z serviceKindName="Zadania ad hoc"
           dla każdego wybranego mieszkańca (po visitorId) -> serviceId
  3. GET  :5010/api/dictionary-value/by-dictionary-kind/53?withAttributes=true
         -> lista rodzajów zadań (czynności, słownik 53) do wyboru
           -- JEDEN rodzaj na uruchomienie skryptu
  4. GET  :5000/api/room/by-organization-id
         -> lista sal/pomieszczeń -- JEDNO pomieszczenie na uruchomienie
           (jeśli różni mieszkańcy mają różne sale, uruchom skrypt osobno)
  5. Generuje listę dat wg zakresu + częstotliwości, dla każdej pary
     (mieszkaniec, data) buduje i POST-uje:

     POST :5070/api/task
       {id:0, rowVersion:0, isDeleted:false, name:<z dict 53>,
        description:null, executingEmployeeId:null, executionRoomId:<wybrane>,
        finishTime:null, isEvaluated:false, normativeTime:<z attrybutu 11 albo podane>,
        plannedFinishTime:<ISO UTC>, plannedStartTime:<ISO UTC>,
        serviceId:<z kroku 2>, specialSkills:[], startTime:null,
        taskKindId:<z kroku 3>, taskPriority:<domyślnie 2>, taskStatus:<domyślnie 3>}

BEZPIECZEŃSTWO: ten skrypt PISZE do systemu. Domyślnie działa w trybie
--dry-run (tylko pokazuje co by zrobił, nic nie wysyła). Żeby faktycznie
wysłać zadania: --execute, i mimo to trzeba potwierdzić [tak] w konsoli
tuż przed wysyłką.

Uruchomienie:
  python zadania_ad_hoc.py --login user@example.com --org-id 1
      # dry-run, wszystko interaktywnie (mieszkańcy, zadanie, sala, harmonogram)
  python zadania_ad_hoc.py --org-id 1 --execute
      # jw., ale faktycznie wysyła po potwierdzeniu
  python zadania_ad_hoc.py --org-id 1 --resident-ids 1554,1600 \\
      --task-kind-id 30722 --room-id 520 \\
      --date-from 2026-09-01 --date-to 2026-09-30 --time 10:00 --execute
"""
import sys
import re
import time
import argparse
import getpass
import logging
from datetime import datetime, date as _date, timedelta

try:
    import requests
except ImportError:
    print("pip install requests"); sys.exit(1)

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


ENVS = {
    "prod": "apedps01.bzmw.gov.pl",
    "test": "ttapedps01.bzmw.gov.pl",
}

TIMEOUT = 20
GROUPING_ATTRIBUTE_KIND = 39
DURATION_ATTRIBUTE_KIND = 11   # "plain int", 100% niepewne, ale zgodne z próbką: 10,60,3,1,20,30 minut
DEFAULT_SERVICE_NAME = "zadania ad hoc"
DEFAULT_PRIORITY = 2
DEFAULT_STATUS = 3   # "Do wykonania" wg mapowania w task_assign_gui.py
TASK_STATUS_LABELS = {1: "Nowe", 2: "W trakcie", 3: "Do wykonania",
                       4: "Wykonane", 5: "Anulowane", 6: "Zawieszone"}

DOW_PL_MAP = {"pon": 0, "wt": 1, "sr": 2, "śr": 2, "czw": 3, "pt": 4, "sob": 5, "nie": 6}

logger = logging.getLogger("zadania_ad_hoc")


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


# ── Mieszkańcy ────────────────────────────────────────────────────────────────
def fetch_active_beneficiaries(client, org_id, status_list="1"):
    url = f"{client.servers['beneficiary']}/api/beneficiary/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id, "statusList": status_list}
        r, err = client._get(url, params=params)
        if err or r is None or r.status_code != 200:
            logger.error(f"Błąd pobierania mieszkańców org={org_id} str.{page}: {err or (r and r.status_code)}")
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
    logger.info(f"Organizacja {org_id}: {len(all_items)} aktywnych mieszkańców")
    return all_items


def _resident_label(r):
    return f"{r.get('surname', '')} {r.get('firstName', '')}".strip() or f"#{r.get('id')}"


def _parse_index_spec(spec, n):
    result = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                a, b = int(a), int(b)
            except ValueError:
                return None
            result.extend(range(a, b + 1))
        else:
            try:
                result.append(int(part))
            except ValueError:
                return None
    seen, out = set(), []
    for i in result:
        if 1 <= i <= n and i not in seen:
            seen.add(i)
            out.append(i)
    return out


def select_from_list_interactive(items, label_fn, prompt_title, multi=True):
    """Uniwersalny interaktywny picker: numery/zakresy, 'all' (jeśli multi),
    albo tekst do filtrowania. Zwraca listę wybranych (multi) albo jeden
    element (not multi) albo None jeśli anulowano."""
    pool = items
    while True:
        print(f"\n{prompt_title} ({len(pool)}):")
        for i, it in enumerate(pool, 1):
            print(f"  {i:>3}. {label_fn(it)}")
        hint = "numery/zakresy np. '1,3,5-8'" + ("  |  'all' = wszystkie" if multi else "")
        print(f"\nWpisz: {hint}  |  tekst = filtruj  |  puste = anuluj")
        choice = input("> ").strip()
        if not choice:
            return [] if multi else None
        if multi and choice.lower() == "all":
            return pool
        if re.fullmatch(r"[\d,\-\s]+", choice):
            indices = _parse_index_spec(choice, len(pool))
            if indices:
                selected = [pool[i - 1] for i in indices]
                return selected if multi else selected[0]
            print("Nie rozpoznano numerów — spróbuj ponownie.")
            continue
        q = choice.lower()
        filtered = [it for it in pool if q in label_fn(it).lower()]
        if not filtered:
            print(f"Brak dopasowań dla '{choice}'.")
            continue
        pool = filtered


# ── Usługa "Zadania ad hoc" ───────────────────────────────────────────────────
def fetch_visitor_services(client, org_id, max_pages=50):
    """GET :5070/api/service/by-visitor-organization-id/paged — pobiera WSZYSTKIE
    strony raz na cały przebieg skryptu (potem filtrujemy lokalnie per mieszkaniec,
    żeby nie odpytywać API osobno dla każdego)."""
    url = f"{client.servers['shift']}/api/service/by-visitor-organization-id/paged"
    all_items, page, page_size = [], 1, 100
    search_fields = "name,serviceCategoryName,planRegistrationNumber,planStatus,visitorSurname,visitorFirstName,executionTime,description"
    while page <= max_pages:
        params = {"page": page, "pageSize": page_size, "orderBy": "planRegistrationNumber",
                  "ascending": "false", "searchFields": search_fields,
                  "planStatus": 1, "visitorType": 3}
        r, err = client._get(url, params=params)
        if err or r is None or r.status_code != 200:
            logger.error(f"Błąd pobierania usług str.{page}: {err or (r and r.status_code)}")
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
        logger.warning(f"Usługi mieszkańców: osiągnięto limit {max_pages} stron — "
                        f"część usług mogła nie zostać pobrana.")
    logger.info(f"Pobrano {len(all_items)} usług mieszkańców (planStatus=1, visitorType=3)")
    return all_items


def find_adhoc_service_id(services, resident_id, service_name=DEFAULT_SERVICE_NAME):
    for s in services:
        if s.get("visitorId") == resident_id and service_name in (s.get("serviceKindName") or "").lower():
            return s.get("id")
    return None


# ── Rodzaje zadań (słownik 53) ────────────────────────────────────────────────
def fetch_all_task_kinds(client):
    r, err = client._get(f"{client.servers['auth']}/api/dictionary-value/by-dictionary-kind/53",
                         params={"withAttributes": "true"})
    if err or r is None or r.status_code != 200:
        logger.error(f"Błąd pobierania rodzajów zadań (słownik 53): {err or (r and r.status_code)}")
        return []
    data = r.json()
    items = data if isinstance(data, list) else (data.get("items") or data.get("data") or [])
    items.sort(key=lambda x: (x.get("content") or "").lower())
    logger.info(f"Słownik 53: {len(items)} rodzajów zadań")
    return items


def extract_default_duration(task_kind):
    """Próbuje wyciągnąć domyślny czas trwania (minuty) z attributeKind=11.
    Niepewne mapowanie (patrz docstring modułu) — zawsze do potwierdzenia
    przez użytkownika, nigdy nie używane bez pytania."""
    for attr in task_kind.get("valueAttributes", []) or []:
        if attr.get("attributeKind") == DURATION_ATTRIBUTE_KIND:
            v = attr.get("value")
            if isinstance(v, int) and not isinstance(v, bool):
                return v
    return None


# ── Sale/pomieszczenia ────────────────────────────────────────────────────────
def fetch_rooms(client, org_id):
    r, err = client._get(f"{client.servers['employee']}/api/room/by-organization-id",
                         params={"organizationId": org_id})
    if err or r is None or r.status_code != 200:
        logger.error(f"Błąd pobierania sal: {err or (r and r.status_code)}")
        return []
    data = r.json()
    items = data if isinstance(data, list) else (data.get("items") or data.get("data") or [])
    logger.info(f"Pobrano {len(items)} sal/pomieszczeń")
    return items


def _room_label(r):
    return f"{r.get('name', r.get('content', ''))}"


# ── Harmonogram ────────────────────────────────────────────────────────────────
def parse_weekdays(spec):
    """'pon,śr,pt' -> {0,2,4}. Zwraca None jeśli spec puste."""
    if not spec:
        return None
    out = set()
    for part in spec.split(","):
        part = part.strip().lower()
        if part in DOW_PL_MAP:
            out.add(DOW_PL_MAP[part])
    return out or None


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


# ── Budowa payloadu i wysyłka ──────────────────────────────────────────────────
def build_task_payload(task_kind, service_id, room_id, occ_date, time_of_day,
                        duration_minutes, priority=DEFAULT_PRIORITY, status=DEFAULT_STATUS):
    hh, mm = [int(x) for x in time_of_day.split(":")]
    start_local = _dt_local(occ_date.year, occ_date.month, occ_date.day, hh, mm)
    finish_local = start_local + timedelta(minutes=duration_minutes)
    return {
        "id": 0,
        "rowVersion": 0,
        "isDeleted": False,
        "name": task_kind.get("content", ""),
        "description": None,
        "executingEmployeeId": None,
        "executionRoomId": room_id,
        "finishTime": None,
        "isEvaluated": False,
        "normativeTime": duration_minutes,
        "plannedFinishTime": _to_utc_iso(finish_local),
        "plannedStartTime": _to_utc_iso(start_local),
        "serviceId": service_id,
        "specialSkills": [],
        "startTime": None,
        "taskKindId": task_kind.get("id"),
        "taskPriority": priority,
        "taskStatus": status,
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
def parse_args():
    p = argparse.ArgumentParser(description="Dodawanie zadań mieszkańcom pod usługę 'Zadania ad hoc'")
    p.add_argument("--login")
    p.add_argument("--password")
    p.add_argument("--env", choices=["prod", "test"], default="test")
    p.add_argument("--org-id", type=int, default=1)
    p.add_argument("--resident-ids", help="Nie pytaj interaktywnie o mieszkańców — ID po przecinku.")
    p.add_argument("--task-kind-id", type=int, help="Nie pytaj interaktywnie o rodzaj zadania — ID ze słownika 53.")
    p.add_argument("--room-id", type=int, help="Nie pytaj interaktywnie o salę — ID pomieszczenia.")
    p.add_argument("--service-name", default=DEFAULT_SERVICE_NAME,
                    help=f"Fragment nazwy usługi do wyszukania (domyślnie '{DEFAULT_SERVICE_NAME}').")
    p.add_argument("--date-from")
    p.add_argument("--date-to")
    p.add_argument("--time", help="Godzina zadania HH:MM (czas lokalny PL).")
    p.add_argument("--every-n-days", type=int, default=1, help="Co ile dni (1=codziennie). Ignorowane gdy podano --weekdays.")
    p.add_argument("--weekdays", help="Tylko te dni tygodnia, np. 'pon,sr,pt' (nadpisuje --every-n-days).")
    p.add_argument("--duration-minutes", type=int, help="Długość zadania w minutach (domyślnie z atrybutu rodzaju zadania, jeśli dostępny).")
    p.add_argument("--priority", type=int, default=DEFAULT_PRIORITY)
    p.add_argument("--status", type=int, default=DEFAULT_STATUS,
                    choices=list(TASK_STATUS_LABELS), help="1=Nowe 2=W trakcie 3=Do wykonania 4=Wykonane 5=Anulowane 6=Zawieszone")
    p.add_argument("--max-tasks", type=int, default=300, help="Bezpiecznik: maksymalna liczba zadań do utworzenia w jednym uruchomieniu.")
    p.add_argument("--sleep", type=float, default=0.3)
    p.add_argument("--execute", action="store_true", help="Bez tego flaga: tylko podgląd (dry-run), nic nie wysyła.")
    p.add_argument("--log-file", default="zadania_ad_hoc.log")
    return p.parse_args()


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
    print("  Dodawanie zadań ad hoc — API SYRENA")
    print(f"  Środowisko: {args.env} ({host})  |  organizationId={args.org_id}")
    if not args.execute:
        print("  TRYB: DRY-RUN (nic nie zostanie wysłane — dodaj --execute)")
    print("=" * 60)

    u = args.login or input("Login (email): ").strip()
    pw = args.password or getpass.getpass(f"Hasło dla {u}: ")
    if not client.login(u, pw):
        print("✗ Logowanie nieudane"); sys.exit(1)
    if not client.select_organization(args.org_id):
        print(f"✗ Nie udało się przełączyć na organizationId={args.org_id}"); sys.exit(1)

    # ── Mieszkańcy ──
    residents = fetch_active_beneficiaries(client, args.org_id)
    if not residents:
        print("Brak aktywnych mieszkańców — kończę."); sys.exit(0)
    if args.resident_ids:
        wanted = {int(x.strip()) for x in args.resident_ids.split(",") if x.strip()}
        selected_residents = [r for r in residents if r.get("id") in wanted]
        missing = wanted - {r.get("id") for r in selected_residents}
        if missing:
            print(f"⚠ Nie znaleziono wśród aktywnych mieszkańców ID: {sorted(missing)}")
    else:
        residents_sorted = sorted(residents, key=lambda r: _resident_label(r).lower())
        selected_residents = select_from_list_interactive(residents_sorted, _resident_label, "Mieszkańcy", multi=True)
    if not selected_residents:
        print("Nie wybrano żadnego mieszkańca — kończę."); sys.exit(0)
    print(f"→ wybrano {len(selected_residents)} mieszkańców")

    # ── Rodzaj zadania ──
    if args.task_kind_id:
        all_kinds = fetch_all_task_kinds(client)
        task_kind = next((k for k in all_kinds if k.get("id") == args.task_kind_id), None)
        if not task_kind:
            print(f"✗ Nie znaleziono rodzaju zadania o ID={args.task_kind_id}"); sys.exit(1)
    else:
        all_kinds = fetch_all_task_kinds(client)
        if not all_kinds:
            print("Brak rodzajów zadań — kończę."); sys.exit(0)
        task_kind = select_from_list_interactive(
            all_kinds, lambda k: f"[{k['id']}] {k.get('content','')}", "Rodzaj zadania (słownik 53)", multi=False)
        if not task_kind:
            print("Nie wybrano rodzaju zadania — kończę."); sys.exit(0)
    print(f"→ rodzaj zadania: [{task_kind['id']}] {task_kind.get('content','')}")

    # ── Sala ──
    if args.room_id:
        room_id = args.room_id
    else:
        rooms = fetch_rooms(client, args.org_id)
        if not rooms:
            print("Brak sal — podaj ręcznie --room-id."); sys.exit(1)
        room = select_from_list_interactive(rooms, _room_label, "Sala/pomieszczenie (jedna dla całego uruchomienia)", multi=False)
        if not room:
            print("Nie wybrano sali — kończę."); sys.exit(0)
        room_id = room.get("id")
    print(f"→ sala: ID={room_id}")

    # ── Harmonogram ──
    date_from = args.date_from or input("Data od (YYYY-MM-DD): ").strip()
    date_to = args.date_to or input("Data do (YYYY-MM-DD): ").strip()
    time_of_day = args.time or input("Godzina zadania (HH:MM): ").strip()
    weekdays = parse_weekdays(args.weekdays) if args.weekdays else None
    if weekdays:
        print(f"→ tylko dni tygodnia: {sorted(weekdays)} (0=pon..6=nie)")
    else:
        print(f"→ co {args.every_n_days} dni")

    duration = args.duration_minutes
    if duration is None:
        auto = extract_default_duration(task_kind)
        if auto is not None:
            ans = input(f"Długość zadania w minutach [domyślna z rodzaju zadania: {auto}]: ").strip()
            duration = int(ans) if ans else auto
        else:
            ans = input("Długość zadania w minutach: ").strip()
            duration = int(ans) if ans else 15

    occ_dates = generate_dates(date_from, date_to, every_n_days=args.every_n_days, weekdays=weekdays)
    if not occ_dates:
        print("Brak dat w wybranym zakresie — kończę."); sys.exit(0)
    print(f"→ {len(occ_dates)} dat: {occ_dates[0]} .. {occ_dates[-1]}")

    # ── Usługi "Zadania ad hoc" ──
    services = fetch_visitor_services(client, args.org_id)
    tasks_to_create = []
    for res in selected_residents:
        rid, rname = res.get("id"), _resident_label(res)
        service_id = find_adhoc_service_id(services, rid, args.service_name)
        if service_id is None:
            logger.warning(f"Mieszkaniec {rname} (id={rid}): brak usługi '{args.service_name}' — pomijam")
            continue
        for d in occ_dates:
            payload = build_task_payload(task_kind, service_id, room_id, d, time_of_day,
                                         duration, priority=args.priority, status=args.status)
            tasks_to_create.append((rname, rid, d, payload))

    if not tasks_to_create:
        print("Brak zadań do utworzenia (żaden mieszkaniec nie ma usługi 'Zadania ad hoc'?) — kończę.")
        sys.exit(0)

    print(f"\n{'='*60}\nPODGLĄD: {len(tasks_to_create)} zadań do utworzenia")
    print(f"  Rodzaj zadania: {task_kind.get('content','')}")
    print(f"  Priorytet={args.priority}  Status={args.status} ({TASK_STATUS_LABELS.get(args.status,'?')})  Czas trwania={duration} min")
    print(f"{'='*60}")
    by_resident = {}
    for rname, rid, d, _ in tasks_to_create:
        by_resident.setdefault((rid, rname), []).append(d)
    for (rid, rname), dates in by_resident.items():
        print(f"  {rname} (id={rid}): {len(dates)} zadań, {dates[0]} .. {dates[-1]}")

    if len(tasks_to_create) > args.max_tasks:
        print(f"\n✗ {len(tasks_to_create)} zadań przekracza limit bezpieczeństwa --max-tasks={args.max_tasks}. "
              f"Zmniejsz zakres/liczbę mieszkańców albo podnieś limit świadomie.")
        sys.exit(1)

    if not args.execute:
        print("\nTryb dry-run — nic nie wysłano. Dodaj --execute, żeby faktycznie utworzyć zadania.")
        sys.exit(0)

    confirm = input(f"\nPotwierdź wysłanie {len(tasks_to_create)} zadań do PRODUKCJI/systemu [tak/nie]: ").strip().lower()
    if confirm not in ("tak", "t", "yes", "y"):
        print("Anulowano — nic nie wysłano."); sys.exit(0)

    ok = fail = 0
    for i, (rname, rid, d, payload) in enumerate(tasks_to_create, 1):
        success, err = post_task(client, payload)
        if success:
            ok += 1
            logger.info(f"[{i}/{len(tasks_to_create)}] ✓ {rname} {d} — utworzono")
        else:
            fail += 1
            logger.error(f"[{i}/{len(tasks_to_create)}] ✗ {rname} {d} — {err}")
        if args.sleep:
            time.sleep(args.sleep)

    print(f"\n{'='*60}\nZakończono: {ok}/{len(tasks_to_create)} utworzono, {fail} błędów")
    print(f"Log: {args.log_file}")


if __name__ == "__main__":
    main()
