// ==UserScript==
// @name         Obserwacje – szybkie dodawanie z raportu dziennego
// @namespace    https://apedps01.bzmw.gov.pl/
// @version      1.4
// @updateURL    https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/ObserwacjeSkrot.user.js
// @downloadURL  https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/ObserwacjeSkrot.user.js
// @description  Dodawanie obserwacji mieszkańcom bezpośrednio z okna "Edycja raportu: Dzienny" - zapisuje się na serwerze (POST /api/observation), widoczne dla każdego kto ma zainstalowany ten sam skrypt.
// @match        *://apedps01.bzmw.gov.pl/*
// @match        *://ttapedps01.bzmw.gov.pl/*
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const EMPLOYEE_PORT = 5000;
    const BENEFICIARY_PORT = 5020;
    const OBSERVATION_PORT = 5020;

    // ---------- 1. Podsłuch tokena Bearer (appka dokleja go do kazdego XHR) ----------
    let capturedToken = null;

    const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
        if (typeof name === 'string' && name.toLowerCase() === 'authorization' && /^Bearer /.test(value)) {
            capturedToken = value;
        }
        return origSetHeader.apply(this, arguments);
    };

    // ---------- 2. Podsłuch odpowiedzi z reportBody (lista mieszkańców wymagających monitorowania) ----------
    let currentResidentEntries = []; // [{beneficiaryIds:[...], guardianshipReasonId}]
    const residentNames = {}; // beneficiaryId -> "Nazwisko Imię"
    let currentEmployeeId = null;

    const origOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function (method, url) {
        this._obsUrl = url;
        return origOpen.apply(this, arguments);
    };

    const origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function (body) {
        this.addEventListener('load', function () {
            try {
                if (this._obsUrl && this._obsUrl.indexOf('/api/daily-shift-report') !== -1 && this.responseText) {
                    const data = JSON.parse(this.responseText);
                    if (data && typeof data.reportBody === 'string') {
                        const rb = JSON.parse(data.reportBody);
                        if (Array.isArray(rb.requiredMonitoringBeneficiaryListItem)) {
                            currentResidentEntries = rb.requiredMonitoringBeneficiaryListItem;
                            resolveResidentNames();
                            refreshSelectOptions();
                            fetchAlertHistory();
                        }
                    }
                }
            } catch (e) {
                // nie nasze dane / nie JSON - ignorujemy
            }
        });
        return origSend.apply(this, arguments);
    };

    function apiBase(port) {
        return 'http://' + location.hostname + ':' + port;
    }

    function ensureEmployeeId() {
        if (currentEmployeeId || !capturedToken) return;
        fetch(apiBase(EMPLOYEE_PORT) + '/api/employee/by-user-id', {
            headers: { Authorization: capturedToken }
        })
            .then((r) => r.json())
            .then((data) => {
                if (data && typeof data.value !== 'undefined') {
                    currentEmployeeId = data.value;
                }
            })
            .catch((e) => console.warn('[Obserwacje] nie udało się pobrać employeeId:', e));
    }

    function resolveResidentNames() {
        if (!capturedToken) return;
        const ids = new Set();
        currentResidentEntries.forEach((item) => (item.beneficiaryIds || []).forEach((id) => ids.add(id)));
        ids.forEach((id) => {
            if (residentNames[id]) return;
            fetch(apiBase(BENEFICIARY_PORT) + '/api/beneficiary/by-organization-id?search=' + id + '&searchFields=id', {
                headers: { Authorization: capturedToken }
            })
                .then((r) => r.json())
                .then((arr) => {
                    if (Array.isArray(arr) && arr[0]) {
                        const b = arr[0];
                        residentNames[id] = (b.surname + ' ' + b.firstName).trim() || ('ID ' + id);
                        refreshSelectOptions();
                        fetchAlertHistory();
                    }
                })
                .catch(() => {});
        });
    }

    // ---------- 3a. Panel DODAWANIA (przywrócony stary układ: lista + pole + wnioski + przycisk) ----------
    const ALERT_PREFIX = 'Alert(raport dzień/noc): ';
    let selectEl = null;
    let addStatusEl = null;
    let addEditableSpan = null;

    function refreshSelectOptions() {
        if (!selectEl) return;
        const prevValue = selectEl.value;
        selectEl.innerHTML = '';
        currentResidentEntries.forEach((entry) => {
            (entry.beneficiaryIds || []).forEach((id) => {
                const opt = document.createElement('option');
                opt.value = id;
                opt.textContent = (residentNames[id] || ('ID ' + id)) + ' (przyczyna nadzoru: ' + entry.guardianshipReasonId + ')';
                selectEl.appendChild(opt);
            });
        });
        if (prevValue) selectEl.value = prevValue;
    }

    function createObservationField() {
        const div = document.createElement('div');
        div.contentEditable = 'true';
        div.style.cssText = 'min-height:40px;border:1px solid #ccc;padding:4px;background:#fff;font-size:13px;box-sizing:border-box;margin-bottom:6px;word-break:break-word;overflow-wrap:anywhere;';

        const prefixSpan = document.createElement('span');
        prefixSpan.textContent = ALERT_PREFIX;
        prefixSpan.style.color = '#999';
        prefixSpan.contentEditable = 'false';
        div.appendChild(prefixSpan);

        const editableSpan = document.createElement('span');
        editableSpan.style.color = '#000';
        div.appendChild(editableSpan);

        return { div, editableSpan };
    }

    function buildAddPanel() {
        const wrap = document.createElement('div');
        wrap.style.cssText = 'margin:10px 0;padding:10px;border:1px solid #ccc;border-radius:6px;background:#fafafa;font-size:13px;';

        const title = document.createElement('div');
        title.textContent = '📝 Dodaj obserwację (Tampermonkey)';
        title.style.cssText = 'font-weight:bold;margin-bottom:6px;';
        wrap.appendChild(title);

        selectEl = document.createElement('select');
        selectEl.style.cssText = 'width:100%;margin-bottom:6px;padding:4px;';
        wrap.appendChild(selectEl);

        const { div: obsField, editableSpan } = createObservationField();
        addEditableSpan = editableSpan;
        wrap.appendChild(obsField);

        const conclusionsEl = document.createElement('textarea');
        conclusionsEl.placeholder = 'Wnioski';
        conclusionsEl.rows = 2;
        conclusionsEl.style.cssText = 'width:100%;margin-bottom:6px;padding:4px;box-sizing:border-box;';
        wrap.appendChild(conclusionsEl);

        const btn = document.createElement('button');
        btn.textContent = 'Zapisz obserwację';
        btn.type = 'button';
        btn.style.cssText = 'padding:6px 14px;cursor:pointer;';
        wrap.appendChild(btn);

        addStatusEl = document.createElement('div');
        addStatusEl.style.cssText = 'margin-top:6px;';
        wrap.appendChild(addStatusEl);

        btn.addEventListener('click', () => {
            const beneficiaryId = Number(selectEl.value);
            if (!beneficiaryId) {
                addStatusEl.textContent = 'Wybierz mieszkańca.';
                addStatusEl.style.color = 'red';
                return;
            }
            if (!capturedToken) {
                addStatusEl.textContent = 'Brak przechwyconego tokenu - wykonaj dowolną akcję w aplikacji i spróbuj ponownie.';
                addStatusEl.style.color = 'red';
                return;
            }
            if (!currentEmployeeId) {
                addStatusEl.textContent = 'Trwa pobieranie danych pracownika, spróbuj za chwilę.';
                addStatusEl.style.color = 'red';
                ensureEmployeeId();
                return;
            }

            const payload = {
                id: 0,
                rowVersion: 0,
                isDeleted: false,
                beneficiaryId: beneficiaryId,
                employeeId: currentEmployeeId,
                observationConclusions: conclusionsEl.value,
                observationContent: ALERT_PREFIX + editableSpan.textContent.trim(),
                observationTime: new Date().toISOString()
            };

            btn.disabled = true;
            addStatusEl.textContent = 'Zapisywanie...';
            addStatusEl.style.color = 'black';

            fetch(apiBase(OBSERVATION_PORT) + '/api/observation', {
                method: 'POST',
                headers: {
                    Authorization: capturedToken,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            })
                .then((r) => {
                    if (!r.ok) {
                        return r.text().then((t) => {
                            throw new Error('HTTP ' + r.status + (t ? ': ' + t : ''));
                        });
                    }
                    return r.json().catch(() => null);
                })
                .then(() => {
                    addStatusEl.textContent = '✔ Zapisano.';
                    addStatusEl.style.color = 'green';
                    editableSpan.textContent = '';
                    conclusionsEl.value = '';
                    fetchAlertHistory();
                })
                .catch((e) => {
                    addStatusEl.textContent = '✗ Błąd zapisu: ' + e.message;
                    addStatusEl.style.color = 'red';
                    console.error('[Obserwacje] błąd zapisu:', e);
                })
                .finally(() => {
                    btn.disabled = false;
                });
        });

        refreshSelectOptions();
        return wrap;
    }

    // ---------- 3b. Tabela PODGLĄDU alertów (GET /api/observation/.../paged, filtrowane po frazie ALERT_PREFIX) ----------
    let historyBodyEl = null;
    let historyStatusEl = null;

    function fetchAlertHistory() {
        if (!capturedToken || !historyBodyEl) return;
        historyStatusEl.textContent = 'Wczytywanie...';
        const url = apiBase(OBSERVATION_PORT) +
            '/api/observation/by-organization-id/paged?page=1&pageSize=200&orderBy=observationTime&ascending=false&onlyActive=true';

        fetch(url, { headers: { Authorization: capturedToken } })
            .then((r) => r.json())
            .then((data) => {
                if (!data || !Array.isArray(data.results)) {
                    historyStatusEl.textContent = '';
                    return;
                }
                const relevantIds = new Set();
                currentResidentEntries.forEach((e) => (e.beneficiaryIds || []).forEach((id) => relevantIds.add(id)));

                const filtered = data.results.filter(
                    (o) =>
                        typeof o.observationContent === 'string' &&
                        o.observationContent.indexOf(ALERT_PREFIX) !== -1 &&
                        relevantIds.has(o.beneficiaryId)
                );

                historyBodyEl.innerHTML = '';
                if (filtered.length === 0) {
                    historyStatusEl.textContent = 'Brak alertów dla mieszkańców z tego raportu (przeszukano ostatnie ' + data.results.length + ' obserwacji).';
                    return;
                }
                historyStatusEl.textContent = '';
                filtered.forEach((o) => {
                    const tr = document.createElement('tr');
                    [
                        residentNames[o.beneficiaryId] || ('ID ' + o.beneficiaryId),
                        new Date(o.observationTime).toLocaleString('pl-PL'),
                        o.observationContent,
                        o.observationConclusions
                    ].forEach((text) => {
                        const td = document.createElement('td');
                        td.style.cssText = 'padding:4px;border:1px solid #ddd;vertical-align:top;word-break:break-word;overflow-wrap:anywhere;max-width:0;';
                        td.textContent = text;
                        tr.appendChild(td);
                    });
                    historyBodyEl.appendChild(tr);
                });
            })
            .catch((e) => {
                historyStatusEl.textContent = 'Błąd pobierania: ' + e.message;
                console.error('[Obserwacje] błąd pobierania historii alertów:', e);
            });
    }

    function buildHistoryTable() {
        const wrap = document.createElement('div');
        wrap.style.cssText = 'margin:10px 0;padding:10px;border:1px solid #ccc;border-radius:6px;background:#fafafa;font-size:13px;';

        const titleRow = document.createElement('div');
        titleRow.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;';
        const title = document.createElement('div');
        title.textContent = '📋 Alerty (raport dzień/noc) dla mieszkańców z tego raportu';
        title.style.cssText = 'font-weight:bold;';
        titleRow.appendChild(title);
        const refreshBtn = document.createElement('button');
        refreshBtn.textContent = 'Odśwież';
        refreshBtn.type = 'button';
        refreshBtn.style.cssText = 'padding:2px 8px;cursor:pointer;';
        refreshBtn.addEventListener('click', fetchAlertHistory);
        titleRow.appendChild(refreshBtn);
        wrap.appendChild(titleRow);

        historyStatusEl = document.createElement('div');
        historyStatusEl.style.cssText = 'margin-bottom:6px;color:#666;';
        wrap.appendChild(historyStatusEl);

        const table = document.createElement('table');
        table.style.cssText = 'width:100%;max-width:100%;table-layout:fixed;border-collapse:collapse;';
        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        ['Mieszkaniec', 'Data', 'Treść', 'Wnioski'].forEach((h) => {
            const th = document.createElement('th');
            th.textContent = h;
            th.style.cssText = 'padding:4px;border:1px solid #ddd;background:#eee;text-align:left;word-break:break-word;overflow-wrap:anywhere;';
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);
        table.appendChild(thead);

        historyBodyEl = document.createElement('tbody');
        table.appendChild(historyBodyEl);
        wrap.appendChild(table);

        return wrap;
    }

    // ---------- 3c. Wstrzykiwanie obu bloków w okno "Edycja raportu: Dzienny" ----------
    // Renderowane w Shadow DOM z doczepionym Water.css (CDN) - izoluje style w obie strony
    // (nic z frameworku nie wycieka na resztę strony SYRENA, nic ze SYRENY nie miesza się do
    // frameworku), przy zachowaniu dziedziczenia fontu/koloru z otaczającej aplikacji (Shadow
    // DOM blokuje reguły CSS, ale nie dziedziczone wartości obliczone). Brak taga "body" tutaj
    // celowo - to widget wstrzykiwany W ramach istniejącej strony, nie osobne okno.
    const WATER_CSS_URL = 'https://cdn.jsdelivr.net/npm/water.css@2/out/water.css';

    function tryInjectPanel() {
        if (document.getElementById('obs-skrot-panel')) return;

        const headings = document.querySelectorAll('*');
        for (const el of headings) {
            if (el.children.length === 0 && el.textContent && el.textContent.trim() === 'Mieszkańcy wymagający monitorowania') {
                const section = el.closest('div');
                if (!section || section.dataset.obsPanelInjected) continue;

                const host = document.createElement('div');
                host.id = 'obs-skrot-panel';
                const shadow = host.attachShadow({ mode: 'open' });

                const link = document.createElement('link');
                link.rel = 'stylesheet';
                link.href = WATER_CSS_URL;
                shadow.appendChild(link);

                const fixupStyle = document.createElement('style');
                fixupStyle.textContent = 'input,select,textarea{border:1px solid #999;}';
                shadow.appendChild(fixupStyle);

                shadow.appendChild(buildHistoryTable());
                shadow.appendChild(buildAddPanel());

                section.dataset.obsPanelInjected = '1';
                section.parentElement
                    ? section.parentElement.insertBefore(host, section.nextSibling)
                    : section.appendChild(host);

                ensureEmployeeId();
                fetchAlertHistory();
                break;
            }
        }
    }

    const observer = new MutationObserver(() => {
        tryInjectPanel();
    });

    function start() {
        observer.observe(document.body, { childList: true, subtree: true });
        tryInjectPanel();
    }

    if (document.body) {
        start();
    } else {
        document.addEventListener('DOMContentLoaded', start);
    }
})();
