// ==UserScript==
// @name         Obserwacje – szybkie dodawanie z raportu dziennego
// @namespace    https://apedps01.bzmw.gov.pl/
// @version      1.1
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
                            refreshTableRows();
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
                        refreshTableRows();
                    }
                })
                .catch(() => {});
        });
    }

    // ---------- 3. Tabela wstrzykiwana w okno "Edycja raportu: Dzienny" ----------
    const ALERT_PREFIX = 'Alert(raport dzień/noc): ';
    let tableBodyEl = null;
    const rowRefs = {}; // beneficiaryId -> {nameCell, editableSpan, conclusionsEl, statusEl}

    function createObservationField() {
        const div = document.createElement('div');
        div.contentEditable = 'true';
        div.style.cssText = 'min-height:40px;border:1px solid #ccc;padding:4px;background:#fff;font-size:13px;';

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

    function buildRow(beneficiaryId, guardianshipReasonId) {
        const tr = document.createElement('tr');

        const nameTd = document.createElement('td');
        nameTd.style.cssText = 'padding:4px;border:1px solid #ddd;vertical-align:top;';
        nameTd.textContent = residentNames[beneficiaryId] || ('ID ' + beneficiaryId);
        tr.appendChild(nameTd);

        const historyTd = document.createElement('td');
        historyTd.style.cssText = 'padding:4px;border:1px solid #ddd;vertical-align:top;color:#999;font-size:12px;';
        historyTd.textContent = 'brak podglądu (czeka na dane z GET /api/observation)';
        tr.appendChild(historyTd);

        const contentTd = document.createElement('td');
        contentTd.style.cssText = 'padding:4px;border:1px solid #ddd;vertical-align:top;';
        const { div: obsField, editableSpan } = createObservationField();
        contentTd.appendChild(obsField);
        tr.appendChild(contentTd);

        const conclusionsTd = document.createElement('td');
        conclusionsTd.style.cssText = 'padding:4px;border:1px solid #ddd;vertical-align:top;';
        const conclusionsEl = document.createElement('textarea');
        conclusionsEl.placeholder = 'Wnioski';
        conclusionsEl.rows = 2;
        conclusionsEl.style.cssText = 'width:100%;box-sizing:border-box;font-size:13px;';
        conclusionsTd.appendChild(conclusionsEl);
        tr.appendChild(conclusionsTd);

        const actionTd = document.createElement('td');
        actionTd.style.cssText = 'padding:4px;border:1px solid #ddd;vertical-align:top;';
        const btn = document.createElement('button');
        btn.textContent = 'Zapisz';
        btn.type = 'button';
        btn.style.cssText = 'padding:4px 10px;cursor:pointer;';
        const statusEl = document.createElement('div');
        statusEl.style.cssText = 'margin-top:4px;font-size:12px;';
        actionTd.appendChild(btn);
        actionTd.appendChild(statusEl);
        tr.appendChild(actionTd);

        btn.addEventListener('click', () => {
            if (!capturedToken) {
                statusEl.textContent = 'Brak tokenu - wykonaj akcję w aplikacji.';
                statusEl.style.color = 'red';
                return;
            }
            if (!currentEmployeeId) {
                statusEl.textContent = 'Czekam na dane pracownika...';
                statusEl.style.color = 'red';
                ensureEmployeeId();
                return;
            }

            const observationContent = ALERT_PREFIX + editableSpan.textContent.trim();

            const payload = {
                id: 0,
                rowVersion: 0,
                isDeleted: false,
                beneficiaryId: beneficiaryId,
                employeeId: currentEmployeeId,
                observationConclusions: conclusionsEl.value,
                observationContent: observationContent,
                observationTime: new Date().toISOString()
            };

            btn.disabled = true;
            statusEl.textContent = 'Zapisywanie...';
            statusEl.style.color = 'black';

            fetch(apiBase(OBSERVATION_PORT) + '/api/observation', {
                method: 'POST',
                headers: {
                    Authorization: capturedToken,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            })
                .then((r) => {
                    if (!r.ok) throw new Error('HTTP ' + r.status);
                    return r.json().catch(() => null);
                })
                .then(() => {
                    statusEl.textContent = '✔ Zapisano.';
                    statusEl.style.color = 'green';
                    editableSpan.textContent = '';
                    conclusionsEl.value = '';
                })
                .catch((e) => {
                    statusEl.textContent = '✗ Błąd: ' + e.message + ' (sprawdź konsolę - możliwa blokada CORS)';
                    statusEl.style.color = 'red';
                    console.error('[Obserwacje] błąd zapisu:', e);
                })
                .finally(() => {
                    btn.disabled = false;
                });
        });

        rowRefs[beneficiaryId] = { nameTd, historyTd };
        return tr;
    }

    function refreshTableRows() {
        if (!tableBodyEl) return;
        tableBodyEl.innerHTML = '';
        currentResidentEntries.forEach((entry) => {
            (entry.beneficiaryIds || []).forEach((id) => {
                if (rowRefs[id]) {
                    rowRefs[id].nameTd.textContent = residentNames[id] || ('ID ' + id);
                    tableBodyEl.appendChild(rowRefs[id].nameTd.closest('tr'));
                } else {
                    tableBodyEl.appendChild(buildRow(id, entry.guardianshipReasonId));
                }
            });
        });
    }

    function buildTable() {
        const wrap = document.createElement('div');
        wrap.style.cssText = 'margin:10px 0;padding:10px;border:1px solid #ccc;border-radius:6px;background:#fafafa;';

        const title = document.createElement('div');
        title.textContent = '📝 Obserwacje monitorowanych mieszkańców (Tampermonkey)';
        title.style.cssText = 'font-weight:bold;margin-bottom:6px;font-size:13px;';
        wrap.appendChild(title);

        const table = document.createElement('table');
        table.style.cssText = 'width:100%;border-collapse:collapse;font-size:13px;';

        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        ['Mieszkaniec', 'Ostatnie obserwacje', 'Nowa obserwacja', 'Wnioski', ''].forEach((h) => {
            const th = document.createElement('th');
            th.textContent = h;
            th.style.cssText = 'padding:4px;border:1px solid #ddd;background:#eee;text-align:left;';
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);
        table.appendChild(thead);

        tableBodyEl = document.createElement('tbody');
        table.appendChild(tableBodyEl);
        wrap.appendChild(table);

        refreshTableRows();
        return wrap;
    }

    function tryInjectPanel() {
        if (document.getElementById('obs-skrot-panel')) return;

        const headings = document.querySelectorAll('*');
        for (const el of headings) {
            if (el.children.length === 0 && el.textContent && el.textContent.trim() === 'Mieszkańcy wymagający monitorowania') {
                const section = el.closest('div');
                if (!section || section.dataset.obsPanelInjected) continue;
                const panel = buildTable();
                panel.id = 'obs-skrot-panel';
                section.dataset.obsPanelInjected = '1';
                section.parentElement
                    ? section.parentElement.insertBefore(panel, section.nextSibling)
                    : section.appendChild(panel);
                ensureEmployeeId();
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
