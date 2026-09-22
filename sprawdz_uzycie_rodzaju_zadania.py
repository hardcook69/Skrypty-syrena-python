#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sprawdzenie, czy dany rodzaj zadania (wpis w słowniku 53 "Rodzaj zadania",
identyfikowany po ID) jest gdziekolwiek UŻYWANY przez konkretne zadania
(instancje, :5070/api/task) w systemie -- przydatne PRZED usunięciem wpisu
ze słownika (patrz usun_ze_slownika_gui.py), żeby nie usunąć czegoś, co
nadal jest przypisane do zadań mieszkańców/pracowników.

Sprawdza WSZYSTKIE organizacje podane w --org-id (może być kilka). Skrypt
NIE zgaduje "całego systemu" automatycznie -- nie ma w tym repo
potwierdzonego endpointu do wylistowania wszystkich organizacji (patrz
inne skrypty), więc organizacje trzeba podać jawnie. Jeśli nie podasz
żadnej, użyty zostanie tylko organizationId=1 (domyślny, jak w innych
skryptach) -- SPRAWDŹ, czy to na pewno wszystkie organizacje, których
dotyczy pytanie.

Mechanizm dopasowania -- DWA NIEZALEŻNE sygnały, bo pole z ID rodzaju
zadania w odpowiedzi GET /api/task NIE jest potwierdzone przechwyconym
ruchem (patrz ostrzeżenie w fetch_active_tasks, ten sam problem co w
audyt_mieszkancow_gui.py):
  1. Po ID: jeśli w pobranych zadaniach w ogóle występuje pole pasujące do
     wzorca "taskKind...Id" (wykrywane DYNAMICZNIE z pierwszego zadania,
     NIE hardkodowane w ciemno) -- dokładne porównanie liczbowe.
  2. Po treści: nazwa zadania (task.name / task.taskKindName) dokładnie
     równa treści wpisu słownika 53 o podanym ID (pobranej z
     /api/dictionary-value/by-dictionary-id/53 dla wybranej organizacji).
Dla każdego trafienia skrypt pokazuje, KTÓRY sygnał zadziałał (id/nazwa/
oba) -- trafienia oparte WYŁĄCZNIE o nazwę oznacz jako mniej pewne
(mogłyby dać fałszywy wynik, gdyby dwa różne wpisy słownika miały
identyczną treść) i tak też są opisane w raporcie.

Sprawdza zarówno zadania AKTYWNE jak i USUNIĘTE (soft-delete) -- jedne i
drugie nadal odwołują się do rodzaju zadania w bazie, więc obie kategorie
są istotne przy ocenie, czy wpis słownika można bezpiecznie skasować.

Uruchomienie:
  python sprawdz_uzycie_rodzaju_zadania.py --env test --org-id 1 --task-kind-id 30700
  python sprawdz_uzycie_rodzaju_zadania.py --env prod --org-id 1 2 3 --task-kind-id 30700
"""
import sys
import re
import argparse
import getpass
import logging

try:
    import requests
except ImportError:
    print("pip install requests"); sys.exit(1)


ENVS = {
    "prod": "apedps01.bzmw.gov.pl",
    "test": "ttapedps01.bzmw.gov.pl",
}

TIMEOUT = 20
ZADANIA_DICTIONARY_ID = 53

logger = logging.getLogger("sprawdz_uzycie_rodzaju_zadania")


class Client:
    def __init__(self, host):
        self.servers = {
            "auth": f"http://{host}:5010",
            "employee": f"http://{host}:5000",
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


def fetch_dictionary_values(client, dictionary_id):
    r, err = client._get(f"{client.servers['employee']}/api/dictionary-value/by-dictionary-id/{dictionary_id}")
    if err or r is None or r.status_code != 200:
        logger.error(f"Błąd pobierania wartości słownika {dictionary_id}: {err or (r and r.status_code)}")
        return []
    data = r.json()
    return data if isinstance(data, list) else (data.get("items") or data.get("data") or [])


def fetch_active_tasks(client, org_id):
    """UWAGA -- NIE W PEŁNI POTWIERDZONE (ten sam zastrzeżony endpoint co w
    audyt_mieszkancow_gui.py::fetch_active_tasks): stronicowanie page/
    pageSize dla PEŁNEJ listy zadań organizacji nie jest potwierdzone
    przechwyconym ruchem. Jeśli liczba zadań wygląda podejrzanie nisko,
    zweryfikuj ręcznie w przeglądarce (Network -> /api/task/by-organization-id
    bez filtra 'search')."""
    url = f"{client.servers['shift']}/api/task/by-organization-id"
    page, page_size, all_items, warned = 1, 200, [], False
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id}
        r, err = client._get(url, params=params)
        if err or r is None:
            logger.error(f"Błąd pobierania zadań org={org_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logger.error(f"Zadania org={org_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logger.error(f"Zadania org={org_id}: odpowiedź nie jest JSON-em: {e}")
            break
        if isinstance(data, list):
            if not warned:
                logger.warning(
                    "fetch_active_tasks: /api/task/by-organization-id zwrócił gołą listę (nie "
                    "{results,...}) -- stronicowanie page/pageSize dla PEŁNEJ listy zadań NIE jest "
                    "potwierdzone przechwyconym ruchem. Jeśli liczba zadań w raporcie wygląda za "
                    "niska, zweryfikuj ręcznie w przeglądarce.")
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
    logger.info(f"Organizacja {org_id}: {len(all_items)} aktywnych zadań")
    return all_items


def fetch_deleted_tasks(client, org_id):
    """POTWIERDZONE przechwyconym ruchem: GET .../task/deleted/by-organization-id/paged"""
    url = f"{client.servers['shift']}/api/task/deleted/by-organization-id/paged"
    page, page_size, all_items = 1, 200, []
    while True:
        params = {"page": page, "pageSize": page_size, "organizationId": org_id,
                  "orderBy": "id", "ascending": "true"}
        r, err = client._get(url, params=params)
        if err or r is None:
            logger.error(f"Błąd pobierania usuniętych zadań org={org_id} str.{page}: {err}")
            break
        if r.status_code != 200:
            logger.error(f"Usunięte zadania org={org_id} str.{page}: HTTP {r.status_code}")
            break
        try:
            data = r.json()
        except Exception as e:
            logger.error(f"Usunięte zadania org={org_id}: odpowiedź nie jest JSON-em: {e}")
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
    logger.info(f"Organizacja {org_id}: {len(all_items)} usuniętych zadań")
    return all_items


_TASK_KIND_ID_FIELD_RE = re.compile(r"(?i)^taskkind.*id$")


def detect_task_kind_id_field(tasks):
    """Szuka w pierwszym niepustym zadaniu klucza pasującego do wzorca
    'taskKind...Id' (np. taskKindId) -- pole to NIE jest potwierdzone w
    przechwyconym ruchu, więc wykrywane jest DYNAMICZNIE z prawdziwej
    odpowiedzi API zamiast zakładane na sztywno. Zwraca nazwę pola albo
    None, jeśli żadne zadanie go nie ma (wtedy używane jest tylko
    dopasowanie po nazwie, patrz match_tasks)."""
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
    """Zwraca listę (task, matched_by) dla zadań, które trafiają w
    task_kind_id -- matched_by to lista z 'id' i/lub 'nazwa', pokazująca
    który sygnał zadziałał (patrz opis mechanizmu na górze pliku)."""
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


def parse_args():
    p = argparse.ArgumentParser(
        description="Sprawdza, czy dany rodzaj zadania (słownik 53) jest używany przez jakieś zadania w systemie.")
    p.add_argument("--login")
    p.add_argument("--password")
    p.add_argument("--env", choices=["prod", "test"], default="test")
    p.add_argument("--org-id", type=int, nargs="+", default=[1],
                    help="Jedna lub więcej organizacji do sprawdzenia (domyślnie: 1).")
    p.add_argument("--task-kind-id", type=int, required=True,
                    help="ID wpisu w słowniku 53 (Rodzaj zadania), którego użycie sprawdzamy.")
    p.add_argument("--log-file", default="sprawdz_uzycie_rodzaju_zadania.log")
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
    print("  Sprawdzenie użycia rodzaju zadania (słownik 53) — API SYRENA")
    print(f"  Środowisko: {args.env} ({host})")
    print(f"  Rodzaj zadania ID: {args.task_kind_id}")
    print(f"  Organizacje: {args.org_id}")
    print("=" * 60)

    u = args.login or input("\nLogin (email): ").strip()
    pw = args.password or getpass.getpass(f"Hasło dla {u}: ")
    if not client.login(u, pw):
        print("✗ Logowanie nieudane"); sys.exit(1)

    total_hits = 0
    any_reliable_check = False

    for org_id in args.org_id:
        print(f"\n{'-'*60}\nOrganizacja {org_id}")
        if not client.select_organization(org_id):
            print(f"  ✗ Nie udało się przełączyć na organizationId={org_id} — pomijam.")
            continue

        dict_values = fetch_dictionary_values(client, ZADANIA_DICTIONARY_ID)
        kind_entry = next((v for v in dict_values if v.get("id") == args.task_kind_id), None)
        kind_content = kind_entry.get("content") if kind_entry else None
        if kind_content:
            print(f"  Wpis w słowniku 53: '{kind_content}'")
        else:
            print(f"  ⚠ W słowniku 53 tej organizacji NIE MA wpisu o ID={args.task_kind_id} "
                  f"— dopasowanie po nazwie będzie pominięte dla tej organizacji.")

        active = fetch_active_tasks(client, org_id)
        deleted = fetch_deleted_tasks(client, org_id)
        all_tasks = active + deleted
        print(f"  Pobrano: {len(active)} aktywnych zadań, {len(deleted)} usuniętych zadań.")

        id_field = detect_task_kind_id_field(all_tasks)
        if id_field:
            print(f"  Wykryto pole ID rodzaju zadania w odpowiedzi API: '{id_field}' — dopasowanie po ID aktywne.")
        else:
            print(f"  ⚠ Żadne zadanie nie ma pola pasującego do wzorca 'taskKind...Id' — "
                  f"dopasowanie WYŁĄCZNIE po nazwie (mniej pewne).")

        if not id_field and not kind_content:
            print(f"  ✗ Brak jakiegokolwiek wiarygodnego sygnału dla tej organizacji "
                  f"(brak pola ID w zadaniach I brak wpisu w słowniku) — pomijam, wynik NIEROZSTRZYGNIĘTY.")
            continue
        any_reliable_check = True

        hits_active = match_tasks(active, args.task_kind_id, kind_content, id_field)
        hits_deleted = match_tasks(deleted, args.task_kind_id, kind_content, id_field)

        if not hits_active and not hits_deleted:
            print(f"  ✓ Nie znaleziono żadnego zadania używającego ID={args.task_kind_id}.")
            continue

        total_hits += len(hits_active) + len(hits_deleted)
        for label, hits in (("AKTYWNE", hits_active), ("USUNIĘTE", hits_deleted)):
            if not hits:
                continue
            print(f"  ✗ Znaleziono {len(hits)} zadań ({label}) używających ID={args.task_kind_id}:")
            for t, matched_by in hits:
                name = t.get("name") or t.get("taskKindName") or "(brak nazwy)"
                resident = t.get("visitorName") or (f"visitorId={t['visitorId']}" if t.get("visitorId") else "—")
                employees = _task_employee_names(t) or "—"
                status = t.get("taskStatus")
                signal = "+".join(matched_by)
                print(f"     [id={t.get('id')}] '{name}' | mieszkaniec: {resident} | "
                     f"pracownicy: {employees} | status: {status} | dopasowanie: {signal}")

    print(f"\n{'='*60}")
    if not any_reliable_check:
        print(f"NIEROZSTRZYGNIĘTE: dla żadnej z podanych organizacji nie udało się uzyskać "
             f"wiarygodnego sygnału (patrz ostrzeżenia wyżej). NIE traktuj tego jako 'nieużywane'.")
    elif total_hits == 0:
        print(f"WYNIK: rodzaj zadania ID={args.task_kind_id} NIE jest używany w żadnej z "
             f"sprawdzonych organizacji ({args.org_id}).")
    else:
        print(f"WYNIK: rodzaj zadania ID={args.task_kind_id} JEST używany — łącznie {total_hits} "
             f"zadań w sprawdzonych organizacjach ({args.org_id}).")
    print(f"Log: {args.log_file}")


if __name__ == "__main__":
    main()
