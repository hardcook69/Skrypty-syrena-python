// ==UserScript==
// @name         Zadania ad hoc – dodawanie z przeglądarki
// @namespace    https://apedps01.bzmw.gov.pl/
// @version      1.4
// @updateURL    https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/ZadaniaAdHocSkrot.user.js
// @downloadURL  https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/ZadaniaAdHocSkrot.user.js
// @description  Odpowiednik zadania_ad_hoc_gui.py w przeglądarce - wielu mieszkańców, zakres dat, powtarzalność, podgląd i wysyłka - zapisuje się na serwerze (POST /api/task), widoczne dla każdego kto ma zainstalowany ten sam skrypt.
// @match        *://apedps01.bzmw.gov.pl/*
// @match        *://ttapedps01.bzmw.gov.pl/*
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const AUTH_PORT = 5010;
    const EMPLOYEE_PORT = 5000;
    const BENEFICIARY_PORT = 5020;
    const SHIFT_PORT = 5070;
    const ORG_ID = 1;
    const DURATION_ATTRIBUTE_KIND = 11;
    const DEFAULT_PRIORITY = 2;
    const DEFAULT_STATUS = 3;
    const TASK_STATUS_LABELS = { 1: 'Nowe', 2: 'W trakcie', 3: 'Do wykonania', 4: 'Wykonane', 5: 'Anulowane', 6: 'Zawieszone' };
    const PRIORITY_LABELS = { 1: 'Niski', 2: 'Średni', 3: 'Wysoki', 4: 'Najwyższy' };
    const STATUS_LABEL_TO_CODE = Object.fromEntries(Object.entries(TASK_STATUS_LABELS).map(([k, v]) => [v.toLowerCase(), Number(k)]));
    const DOW_LABELS = ['Pon', 'Wt', 'Śr', 'Czw', 'Pt', 'Sob', 'Nie'];
    const ROOM_FIELD_CANDIDATES = ['roomId', 'room_id', 'executionRoomId', 'residenceRoomId', 'currentRoomId'];

    // ---------- Podsłuch tokena Bearer ----------
    let capturedToken = null;
    const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
        if (typeof name === 'string' && name.toLowerCase() === 'authorization' && /^Bearer /.test(value)) {
            capturedToken = value;
        }
        return origSetHeader.apply(this, arguments);
    };

    function apiBase(port) {
        return 'http://' + location.hostname + ':' + port;
    }

    function authFetch(url) {
        return fetch(url, { headers: { Authorization: capturedToken } }).then((r) => {
            if (!r.ok) throw new Error('HTTP ' + r.status);
            return r.json();
        });
    }

    // ---------- DOM helper ----------
    function el(tag, attrs, ...children) {
        const e = document.createElement(tag);
        Object.entries(attrs || {}).forEach(([k, v]) => {
            if (k === 'style') e.style.cssText = v;
            else if (k.startsWith('on') && typeof v === 'function') e.addEventListener(k.slice(2), v);
            else if (k === 'text') e.textContent = v;
            else e.setAttribute(k, v);
        });
        children.flat().forEach((c) => {
            if (c === null || c === undefined) return;
            e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
        });
        return e;
    }

    function residentLabel(r) {
        return (((r.surname || '') + ' ' + (r.firstName || '')).trim()) || ('#' + r.id);
    }

    // ---------- Fetche danych (te same endpointy co zadania_ad_hoc_gui.py) ----------
    function fetchBeneficiaries() {
        const all = [];
        function page(p) {
            const qs = new URLSearchParams({ page: p, pageSize: 200, organizationId: ORG_ID, statusList: '1' });
            return authFetch(apiBase(BENEFICIARY_PORT) + '/api/beneficiary/by-organization-id/paged?' + qs).then((data) => {
                const items = data.results || data.items || data.data || [];
                all.push(...items);
                const totalPages = data.totalNumberOfPages;
                if (totalPages != null ? p < totalPages : items.length >= 200) {
                    return page(p + 1);
                }
                return all;
            });
        }
        return page(1);
    }

    function fetchTaskKinds() {
        return authFetch(apiBase(AUTH_PORT) + '/api/dictionary-value/by-dictionary-kind/53?withAttributes=true').then((data) => {
            const items = Array.isArray(data) ? data : (data.items || data.data || []);
            items.sort((a, b) => (a.content || '').toLowerCase().localeCompare((b.content || '').toLowerCase()));
            return items;
        });
    }

    function extractDefaultDuration(taskKind) {
        for (const attr of taskKind.valueAttributes || []) {
            if (attr.attributeKind === DURATION_ATTRIBUTE_KIND && typeof attr.value === 'number') return attr.value;
        }
        return null;
    }

    function fetchRooms() {
        return authFetch(apiBase(EMPLOYEE_PORT) + '/api/room/by-organization-id?' + new URLSearchParams({ organizationId: ORG_ID })).then((data) =>
            Array.isArray(data) ? data : (data.items || data.data || [])
        );
    }

    function fetchVisitorServices() {
        const all = [];
        const searchFields = 'name,serviceCategoryName,planRegistrationNumber,planStatus,visitorSurname,visitorFirstName,executionTime,description';
        function page(p) {
            const qs = new URLSearchParams({
                page: p, pageSize: 100, orderBy: 'planRegistrationNumber', ascending: 'false',
                searchFields, planStatus: 1, visitorType: 3
            });
            return authFetch(apiBase(SHIFT_PORT) + '/api/service/by-visitor-organization-id/paged?' + qs).then((data) => {
                const items = data.results || data.items || data.data || [];
                all.push(...items);
                const totalPages = data.totalNumberOfPages;
                if ((totalPages != null ? p < totalPages : items.length >= 100) && p < 50) {
                    return page(p + 1);
                }
                return all;
            });
        }
        return page(1);
    }

    function findAdhocServiceId(services, residentId) {
        const svc = services.find((s) => s.visitorId === residentId && (s.serviceKindName || '').toLowerCase().includes('zadania ad hoc'));
        return svc ? svc.id : null;
    }

    function fetchResidentRoomId(beneficiaryId) {
        return authFetch(apiBase(BENEFICIARY_PORT) + '/api/beneficiary-residence/by-beneficiary-id/' + beneficiaryId)
            .then((data) => {
                const d = Array.isArray(data) ? (data[0] || {}) : (data || {});
                for (const key of ROOM_FIELD_CANDIDATES) {
                    if (typeof d[key] === 'number') return { roomId: d[key], field: key };
                }
                for (const key of Object.keys(d)) {
                    if (key.toLowerCase().includes('room') && typeof d[key] === 'number') return { roomId: d[key], field: key };
                }
                return { roomId: null, field: null };
            })
            .catch(() => ({ roomId: null, field: null }));
    }

    // ---------- Generowanie dat / payload (odpowiednik generate_dates / build_task_payload) ----------
    function generateDates(dateFrom, dateTo, everyNDays, weekdaysSet) {
        const [fy, fm, fd] = dateFrom.split('-').map(Number);
        const [ty, tm, td] = dateTo.split('-').map(Number);
        let cur = new Date(fy, fm - 1, fd);
        const end = new Date(ty, tm - 1, td);
        const dates = [];
        let i = 0;
        while (cur <= end) {
            if (weekdaysSet) {
                const monFirst = (cur.getDay() + 6) % 7;
                if (weekdaysSet.has(monFirst)) dates.push(new Date(cur));
            } else if (i % everyNDays === 0) {
                dates.push(new Date(cur));
            }
            cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1);
            i++;
        }
        return dates;
    }

    function buildTaskPayload(taskKind, serviceId, roomId, occDate, hh, mm, durationMinutes, priority, status) {
        const start = new Date(occDate.getFullYear(), occDate.getMonth(), occDate.getDate(), hh, mm, 0, 0);
        const finish = new Date(start.getTime() + durationMinutes * 60000);
        return {
            id: 0, rowVersion: 0, isDeleted: false,
            name: taskKind.content || '',
            description: null, executingEmployeeId: null,
            executionRoomId: roomId, finishTime: null, isEvaluated: false,
            normativeTime: durationMinutes,
            plannedFinishTime: finish.toISOString(),
            plannedStartTime: start.toISOString(),
            serviceId: serviceId, specialSkills: [], startTime: null,
            taskKindId: taskKind.id, taskPriority: priority, taskStatus: status
        };
    }

    function postTask(payload) {
        return fetch(apiBase(SHIFT_PORT) + '/api/task', {
            method: 'POST',
            headers: { Authorization: capturedToken, 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        }).then((r) => {
            if (!r.ok) return r.text().then((t) => { throw new Error('HTTP ' + r.status + (t ? ': ' + t : '')); });
            return r.json().catch(() => null);
        });
    }

    // ---------- Stan ----------
    const state = {
        residents: [], selectedResidentIds: new Set(),
        taskKinds: [], selectedKind: null,
        rooms: [], selectedRoom: null, roomMode: 'fixed',
        pendingTasks: []
    };

    // ---------- Budowa modala ----------
    let modalEls = {};

    function logMsg(msg, kind) {
        const ts = new Date().toLocaleTimeString('pl-PL');
        const color = kind === 'ok' ? '#3c763d' : kind === 'err' ? '#a94442' : '#31708f';
        modalEls.log.appendChild(el('div', { style: 'color:' + color + ';' }, '[' + ts + '] ' + msg));
        modalEls.log.scrollTop = modalEls.log.scrollHeight;
    }

    function buildResidentSection(preselectName) {
        const filterInput = el('input', { type: 'text', placeholder: 'Filtruj...', style: 'width:100%;margin-bottom:4px;padding:4px;box-sizing:border-box;' });
        const listDiv = el('div', { style: 'max-height:220px;overflow-y:auto;border:1px solid #ccc;padding:4px;' });
        const countLbl = el('span', { style: 'font-size:12px;color:#666;margin-left:8px;' }, '0 zaznaczonych');
        const preselectLbl = el('div', { style: 'font-size:12px;color:#b45309;margin-bottom:4px;' }, '');
        const expandToggle = el('a', { href: '#', style: 'font-size:12px;display:none;' }, 'Zmień wybór / pokaż pełną listę');

        function renderList(filterText) {
            listDiv.innerHTML = '';
            const q = (filterText || '').toLowerCase();
            state.residents
                .filter((r) => !q || residentLabel(r).toLowerCase().includes(q))
                .forEach((r) => {
                    const cb = el('input', { type: 'checkbox' });
                    cb.checked = state.selectedResidentIds.has(r.id);
                    cb.addEventListener('change', () => {
                        if (cb.checked) state.selectedResidentIds.add(r.id);
                        else state.selectedResidentIds.delete(r.id);
                        countLbl.textContent = state.selectedResidentIds.size + ' zaznaczonych';
                    });
                    listDiv.appendChild(el('label', { style: 'display:block;padding:2px 0;font-size:13px;' }, cb, ' ' + residentLabel(r)));
                });
        }
        filterInput.addEventListener('input', () => renderList(filterInput.value));

        const statusLbl = el('span', { style: 'font-size:12px;color:#666;margin-left:8px;' }, '');
        const loadBtn = el('button', {
            type: 'button', style: 'padding:4px 10px;cursor:pointer;',
            onclick: () => {
                statusLbl.textContent = 'Pobieranie...';
                fetchBeneficiaries().then((residents) => {
                    residents.sort((a, b) => residentLabel(a).toLowerCase().localeCompare(residentLabel(b).toLowerCase()));
                    state.residents = residents;
                    statusLbl.textContent = residents.length + ' aktywnych';
                    if (preselectName) {
                        const match = residents.find((r) => residentLabel(r).toLowerCase() === preselectName.toLowerCase());
                        if (match) {
                            state.selectedResidentIds.add(match.id);
                            preselectLbl.textContent = '✓ Zaznaczono automatycznie: ' + residentLabel(match);
                            expandableWrap.style.display = 'none';
                            expandToggle.style.display = 'inline';
                        } else {
                            preselectLbl.textContent = '⚠ Nie znaleziono dokładnego dopasowania dla „' + preselectName + '” — zaznacz ręcznie.';
                        }
                    }
                    renderList(filterInput.value);
                    countLbl.textContent = state.selectedResidentIds.size + ' zaznaczonych';
                    logMsg('Pobrano ' + residents.length + ' aktywnych mieszkańców', 'ok');
                }).catch((e) => { statusLbl.textContent = 'Błąd'; logMsg('Błąd pobierania mieszkańców: ' + e.message, 'err'); });
            }
        }, '↓ Pobierz mieszkańców');
        if (preselectName) {
            preselectLbl.textContent = 'Wczytywanie mieszkańca „' + preselectName + '”...';
            setTimeout(() => loadBtn.click(), 0);
        }

        const allBtn = el('button', { type: 'button', style: 'padding:2px 8px;cursor:pointer;margin-right:4px;', onclick: () => {
            state.residents.forEach((r) => state.selectedResidentIds.add(r.id));
            renderList(filterInput.value); countLbl.textContent = state.selectedResidentIds.size + ' zaznaczonych';
        } }, 'Wszyscy');
        const noneBtn = el('button', { type: 'button', style: 'padding:2px 8px;cursor:pointer;', onclick: () => {
            state.selectedResidentIds.clear();
            renderList(filterInput.value); countLbl.textContent = '0 zaznaczonych';
        } }, 'Nikt');

        expandToggle.addEventListener('click', (e) => {
            e.preventDefault();
            expandableWrap.style.display = '';
            expandToggle.style.display = 'none';
        });

        var expandableWrap = el('div', {},
            el('div', { style: 'margin-bottom:4px;' }, loadBtn, statusLbl),
            filterInput, listDiv,
            el('div', { style: 'margin-top:4px;' }, allBtn, noneBtn, countLbl)
        );

        return el('fieldset', { style: 'border:1px solid #ccc;padding:8px;margin-bottom:8px;' },
            el('legend', {}, 'Mieszkańcy (wielokrotny wybór)'),
            preselectLbl, expandToggle,
            expandableWrap
        );
    }

    function buildKindSection() {
        const filterInput = el('input', { type: 'text', placeholder: 'Filtruj...', style: 'width:100%;margin-bottom:4px;padding:4px;box-sizing:border-box;' });
        const listDiv = el('div', { style: 'max-height:140px;overflow-y:auto;border:1px solid #ccc;padding:4px;' });
        const selLbl = el('div', { style: 'font-size:12px;color:#3c763d;margin-top:4px;' }, '— nie wybrano —');
        const statusLbl = el('span', { style: 'font-size:12px;color:#666;margin-left:8px;' }, '');

        function renderList(filterText) {
            listDiv.innerHTML = '';
            const q = (filterText || '').toLowerCase();
            state.taskKinds
                .filter((k) => !q || (k.content || '').toLowerCase().includes(q))
                .forEach((k) => {
                    listDiv.appendChild(el('div', {
                        style: 'padding:3px 4px;cursor:pointer;font-size:13px;' + (state.selectedKind && state.selectedKind.id === k.id ? 'background:#e8f5e9;' : ''),
                        onclick: () => {
                            state.selectedKind = k;
                            const auto = extractDefaultDuration(k);
                            if (auto !== null && modalEls.duration) modalEls.duration.value = auto;
                            selLbl.textContent = 'Wybrano: ' + (k.content || '') + (auto !== null ? '  (sugerowany czas: ' + auto + ' min)' : '');
                            renderList(filterInput.value);
                        }
                    }, '[' + k.id + '] ' + (k.content || '')));
                });
        }
        filterInput.addEventListener('input', () => renderList(filterInput.value));

        const loadBtn = el('button', {
            type: 'button', style: 'padding:4px 10px;cursor:pointer;',
            onclick: () => {
                statusLbl.textContent = 'Pobieranie...';
                fetchTaskKinds().then((kinds) => {
                    state.taskKinds = kinds;
                    statusLbl.textContent = kinds.length + ' rodzajów';
                    renderList(filterInput.value);
                    logMsg('Pobrano ' + kinds.length + ' rodzajów zadań', 'ok');
                }).catch((e) => { statusLbl.textContent = 'Błąd'; logMsg('Błąd pobierania rodzajów zadań: ' + e.message, 'err'); });
            }
        }, '↓ Pobierz rodzaje zadań (słownik 53)');

        return el('fieldset', { style: 'border:1px solid #ccc;padding:8px;margin-bottom:8px;' },
            el('legend', {}, 'Rodzaj zadania — jeden'),
            el('div', { style: 'margin-bottom:4px;' }, loadBtn, statusLbl),
            filterInput, listDiv, selLbl
        );
    }

    function buildRoomSection() {
        const filterInput = el('input', { type: 'text', placeholder: 'Filtruj...', style: 'width:100%;margin-bottom:4px;padding:4px;box-sizing:border-box;' });
        const listDiv = el('div', { style: 'max-height:100px;overflow-y:auto;border:1px solid #ccc;padding:4px;' });
        const selLbl = el('div', { style: 'font-size:12px;color:#3c763d;margin-top:4px;' }, '— nie wybrano —');
        const statusLbl = el('span', { style: 'font-size:12px;color:#666;margin-left:8px;' }, '');

        function renderList(filterText) {
            listDiv.innerHTML = '';
            const q = (filterText || '').toLowerCase();
            state.rooms
                .filter((r) => !q || String(r.name || r.content || '').toLowerCase().includes(q))
                .forEach((r) => {
                    listDiv.appendChild(el('div', {
                        style: 'padding:3px 4px;cursor:pointer;font-size:13px;' + (state.selectedRoom && state.selectedRoom.id === r.id ? 'background:#e8f5e9;' : ''),
                        onclick: () => {
                            state.selectedRoom = r;
                            selLbl.textContent = 'Wybrano: [' + r.id + '] ' + (r.name || r.content || '');
                            renderList(filterInput.value);
                        }
                    }, '[' + r.id + '] ' + (r.name || r.content || '')));
                });
        }
        filterInput.addEventListener('input', () => renderList(filterInput.value));

        const loadBtn = el('button', {
            type: 'button', style: 'padding:4px 10px;cursor:pointer;',
            onclick: () => {
                statusLbl.textContent = 'Pobieranie...';
                fetchRooms().then((rooms) => {
                    state.rooms = rooms;
                    statusLbl.textContent = rooms.length + ' sal';
                    renderList(filterInput.value);
                    logMsg('Pobrano ' + rooms.length + ' sal', 'ok');
                }).catch((e) => { statusLbl.textContent = 'Błąd'; logMsg('Błąd pobierania sal: ' + e.message, 'err'); });
            }
        }, '↓ Pobierz sale');

        const radioName = 'obs-room-mode';
        function radio(value, label) {
            const r = el('input', { type: 'radio', name: radioName, value });
            r.checked = state.roomMode === value;
            r.addEventListener('change', () => {
                state.roomMode = value;
                filterInput.disabled = value !== 'fixed';
                listDiv.style.opacity = value !== 'fixed' ? '0.5' : '1';
                if (value === 'none') selLbl.textContent = 'Wybrano: brak sali';
                else if (value === 'own') selLbl.textContent = 'Wybrano: pokój zamieszkania (automatycznie per mieszkaniec)';
                else selLbl.textContent = state.selectedRoom ? 'Wybrano: [' + state.selectedRoom.id + '] ' + (state.selectedRoom.name || '') : '— nie wybrano —';
            });
            return el('label', { style: 'display:block;font-size:13px;padding:2px 0;' }, r, ' ' + label);
        }

        return el('fieldset', { style: 'border:1px solid #ccc;padding:8px;margin-bottom:8px;' },
            el('legend', {}, 'Sala/pomieszczenie'),
            radio('fixed', 'Wybierz z listy (jedna sala dla wszystkich):'),
            el('div', { style: 'margin-bottom:4px;' }, loadBtn, statusLbl),
            filterInput, listDiv, selLbl,
            radio('own', 'Pokój zamieszkania mieszkańca (automatycznie, best-effort)'),
            radio('none', 'Brak sali')
        );
    }

    function buildScheduleSection() {
        const today = new Date();
        const in30 = new Date(today.getTime() + 30 * 86400000);
        const fmt = (d) => d.toISOString().slice(0, 10);

        const fromInput = el('input', { type: 'date', value: fmt(today), style: 'padding:3px;' });
        const toInput = el('input', { type: 'date', value: fmt(in30), style: 'padding:3px;' });
        const hourSel = el('select', { style: 'padding:3px;' });
        for (let h = 0; h < 24; h++) hourSel.appendChild(el('option', { value: String(h).padStart(2, '0') }, String(h).padStart(2, '0')));
        hourSel.value = '10';
        const minuteSel = el('select', { style: 'padding:3px;' });
        for (let m = 0; m < 60; m += 5) minuteSel.appendChild(el('option', { value: String(m).padStart(2, '0') }, String(m).padStart(2, '0')));
        minuteSel.value = '00';

        modalEls.fromInput = fromInput; modalEls.toInput = toInput;
        modalEls.hourSel = hourSel; modalEls.minuteSel = minuteSel;

        const dateRow = el('div', { style: 'display:flex;align-items:center;gap:6px;flex-wrap:wrap;' },
            'Od:', fromInput, 'Do:', toInput, 'Godzina:', hourSel, ':', minuteSel);

        const everyRadio = el('input', { type: 'radio', name: 'obs-freq-mode', checked: 'checked' });
        const weekdaysRadio = el('input', { type: 'radio', name: 'obs-freq-mode' });
        const everyNInput = el('input', { type: 'number', min: '1', value: '1', style: 'width:50px;padding:3px;' });
        const dowChecks = {};
        const dowRow = el('div', { style: 'display:flex;gap:4px;margin-left:8px;' });
        function styleDowBtn(btn) {
            btn.style.cssText = 'padding:4px 9px;font-size:12px;border-radius:4px;border:1px solid #1b5e20;'
                + (btn.checked ? 'background:#1b5e20;color:#fff;cursor:pointer;' : 'background:#fff;color:#1b5e20;cursor:pointer;')
                + (btn.disabled ? 'opacity:0.35;cursor:default;' : '');
        }
        DOW_LABELS.forEach((label, i) => {
            const btn = el('button', { type: 'button' }, label);
            btn.checked = i < 5;
            btn.disabled = true;
            btn.addEventListener('click', () => {
                if (btn.disabled) return;
                btn.checked = !btn.checked;
                styleDowBtn(btn);
            });
            styleDowBtn(btn);
            dowChecks[i] = btn;
            dowRow.appendChild(btn);
        });
        function updateFreqEnabled() {
            const weekdaysMode = weekdaysRadio.checked;
            everyNInput.disabled = weekdaysMode;
            Object.values(dowChecks).forEach((btn) => { btn.disabled = !weekdaysMode; styleDowBtn(btn); });
        }
        everyRadio.addEventListener('change', updateFreqEnabled);
        weekdaysRadio.addEventListener('change', updateFreqEnabled);

        modalEls.everyRadio = everyRadio; modalEls.weekdaysRadio = weekdaysRadio;
        modalEls.everyNInput = everyNInput; modalEls.dowChecks = dowChecks;

        const freqSection = el('fieldset', { style: 'border:1px solid #ccc;padding:8px;margin-bottom:8px;' },
            el('legend', {}, 'Częstotliwość — jeden z dwóch trybów'),
            el('label', { style: 'display:block;' }, everyRadio, ' Co ', everyNInput, ' dni (1 = codziennie)'),
            el('label', { style: 'display:block;margin-top:6px;' }, weekdaysRadio, ' Tylko wybrane dni tygodnia:'),
            dowRow
        );

        const durationInput = el('input', { type: 'number', value: '15', style: 'width:60px;padding:3px;' });
        const priorityInput = el('select', { style: 'padding:3px;' });
        Object.entries(PRIORITY_LABELS).forEach(([code, label]) => priorityInput.appendChild(el('option', { value: code }, label)));
        priorityInput.value = String(DEFAULT_PRIORITY);
        const statusSelect = el('select', { style: 'padding:3px;' });
        Object.entries(TASK_STATUS_LABELS).forEach(([code, label]) => statusSelect.appendChild(el('option', { value: code }, label)));
        statusSelect.value = String(DEFAULT_STATUS);
        const statusCustomInput = el('input', { type: 'text', placeholder: 'lub numer statusu ręcznie', style: 'width:150px;padding:3px;' });
        const maxTasksInput = el('input', { type: 'number', value: '300', style: 'width:70px;padding:3px;' });

        modalEls.duration = durationInput; modalEls.priority = priorityInput;
        modalEls.statusSelect = statusSelect; modalEls.statusCustom = statusCustomInput;
        modalEls.maxTasks = maxTasksInput;

        const paramsSection = el('fieldset', { style: 'border:1px solid #ccc;padding:8px;margin-bottom:8px;' },
            el('legend', {}, 'Parametry zadania'),
            el('div', { style: 'display:flex;align-items:center;gap:6px;flex-wrap:wrap;' },
                'Czas trwania (min):', durationInput,
                'Priorytet:', priorityInput,
                'Status:', statusSelect, statusCustomInput,
                'Bezpiecznik max zadań:', maxTasksInput
            ),
            el('div', { style: 'font-size:11px;color:#666;margin-top:4px;' }, 'Jeśli status nie pasuje do listy (np. "Do przypisania"), wpisz jego numer w pole obok.')
        );

        return el('div', {}, dateRow, el('div', { style: 'height:8px;' }), freqSection, paramsSection);
    }

    // ---------- Podgląd i wysyłka ----------
    function collectParams() {
        const errors = [];
        const residentIds = Array.from(state.selectedResidentIds);
        const residents = state.residents.filter((r) => state.selectedResidentIds.has(r.id));
        if (residents.length === 0) errors.push('Nie zaznaczono żadnego mieszkańca.');
        if (!state.selectedKind) errors.push('Nie wybrano rodzaju zadania.');
        if (state.roomMode === 'fixed' && !state.selectedRoom) errors.push('Wybierz salę z listy albo zmień tryb.');

        const dateFrom = modalEls.fromInput.value;
        const dateTo = modalEls.toInput.value;
        if (!dateFrom || !dateTo) errors.push('Uzupełnij zakres dat.');

        const duration = parseInt(modalEls.duration.value, 10);
        if (!Number.isFinite(duration)) errors.push('Czas trwania musi być liczbą.');
        const priority = parseInt(modalEls.priority.value, 10);
        if (!Number.isFinite(priority)) errors.push('Priorytet musi być liczbą.');
        const maxTasks = parseInt(modalEls.maxTasks.value, 10) || 300;

        let status = null;
        const customStatus = modalEls.statusCustom.value.trim();
        if (customStatus) {
            status = parseInt(customStatus, 10);
            if (!Number.isFinite(status)) errors.push('Ręcznie wpisany status musi być liczbą.');
        } else {
            status = parseInt(modalEls.statusSelect.value, 10);
        }

        let weekdaysSet = null;
        let everyN = 1;
        if (modalEls.weekdaysRadio.checked) {
            weekdaysSet = new Set(Object.entries(modalEls.dowChecks).filter(([, cb]) => cb.checked).map(([i]) => Number(i)));
            if (weekdaysSet.size === 0) errors.push('Wybierz przynajmniej jeden dzień tygodnia.');
        } else {
            everyN = Math.max(1, parseInt(modalEls.everyNInput.value, 10) || 1);
        }

        if (errors.length) return { errors };
        return {
            residents, dateFrom, dateTo,
            hour: parseInt(modalEls.hourSel.value, 10), minute: parseInt(modalEls.minuteSel.value, 10),
            duration, priority, status, maxTasks, weekdaysSet, everyN
        };
    }

    function showPreview() {
        const params = collectParams();
        if (params.errors) { alert(params.errors.join('\n')); return; }
        modalEls.sendBtn.disabled = true;
        modalEls.preview.textContent = 'Pobieranie usług mieszkańców...';

        fetchVisitorServices().then((services) => {
            const occDates = generateDates(params.dateFrom, params.dateTo, params.everyN, params.weekdaysSet);
            const pending = [];
            const missing = [];
            const roomPromises = [];

            params.residents.forEach((res) => {
                const serviceId = findAdhocServiceId(services, res.id);
                if (serviceId === null) { missing.push(residentLabel(res) + ' (brak usługi "zadania ad hoc")'); return; }

                if (state.roomMode === 'own') {
                    roomPromises.push(
                        fetchResidentRoomId(res.id).then(({ roomId, field }) => {
                            if (roomId === null) { missing.push(residentLabel(res) + ' (nie znaleziono pokoju zamieszkania)'); return; }
                            occDates.forEach((d) => pending.push({
                                resident: res, date: d, roomNote: 'pokój: ' + roomId + ", pole '" + field + "'",
                                payload: buildTaskPayload(state.selectedKind, serviceId, roomId, d, params.hour, params.minute, params.duration, params.priority, params.status)
                            }));
                        })
                    );
                } else {
                    const roomId = state.roomMode === 'none' ? null : state.selectedRoom.id;
                    occDates.forEach((d) => pending.push({
                        resident: res, date: d, roomNote: null,
                        payload: buildTaskPayload(state.selectedKind, serviceId, roomId, d, params.hour, params.minute, params.duration, params.priority, params.status)
                    }));
                }
            });

            Promise.all(roomPromises).then(() => {
                state.pendingTasks = pending;
                renderPreview(pending, missing, occDates, params);
            });
        }).catch((e) => {
            modalEls.preview.textContent = 'Błąd: ' + e.message;
            logMsg('Błąd podglądu: ' + e.message, 'err');
        });
    }

    function renderPreview(pending, missing, occDates, params) {
        const lines = [];
        lines.push('Rodzaj zadania: ' + (state.selectedKind.content || ''));
        const roomDesc = state.roomMode === 'own' ? 'pokój zamieszkania mieszkańca (automatycznie)'
            : state.roomMode === 'none' ? 'brak'
            : '[' + state.selectedRoom.id + '] ' + (state.selectedRoom.name || '');
        lines.push('Sala: ' + roomDesc);
        lines.push('Okres: ' + params.dateFrom + ' .. ' + params.dateTo + '  |  godzina ' +
            String(params.hour).padStart(2, '0') + ':' + String(params.minute).padStart(2, '0') +
            '  |  ' + occDates.length + ' dat  |  czas trwania ' + params.duration + ' min  |  priorytet ' +
            params.priority + '  |  status ' + params.status + ' (' + (TASK_STATUS_LABELS[params.status] || 'nieznany') + ')');
        lines.push('─'.repeat(70));

        const byRes = {};
        pending.forEach((p) => {
            const name = residentLabel(p.resident);
            (byRes[name] = byRes[name] || []).push(p);
        });
        Object.keys(byRes).sort().forEach((name) => {
            const items = byRes[name];
            const note = items[0].roomNote ? '  [' + items[0].roomNote + ']' : '';
            lines.push('  ' + name + ': ' + items.length + ' zadań' + note);
        });
        if (missing.length) { lines.push(''); lines.push('⚠ Pominięci: ' + missing.join(', ')); }
        lines.push('─'.repeat(70));
        lines.push('RAZEM: ' + pending.length + ' zadań do utworzenia');
        modalEls.preview.textContent = lines.join('\n');
        logMsg('Podgląd: ' + pending.length + ' zadań (' + missing.length + ' mieszkańców pominiętych)', 'ok');

        modalEls.sendBtn.disabled = pending.length === 0 || pending.length > params.maxTasks;
        if (pending.length > params.maxTasks) logMsg(pending.length + ' zadań przekracza bezpiecznik max=' + params.maxTasks + ' — wysyłka zablokowana.', 'err');
    }

    function doSend() {
        if (!state.pendingTasks.length) { alert("Najpierw kliknij 'Pokaż podgląd'."); return; }
        if (!capturedToken) { alert('Brak przechwyconego tokenu - wykonaj dowolną akcję w aplikacji i spróbuj ponownie.'); return; }
        const n = state.pendingTasks.length;
        if (!confirm('Wysłać ' + n + ' zadań do systemu?\nTej operacji nie można cofnąć hurtowo.')) return;

        const tasks = state.pendingTasks.slice();
        modalEls.sendBtn.disabled = true;
        let ok = 0, fail = 0;

        function next(i) {
            if (i >= tasks.length) {
                const summary = 'Zakończono: ' + ok + '/' + tasks.length + ' utworzono, ' + fail + ' błędów';
                logMsg('— ' + summary + ' —', fail === 0 ? 'ok' : 'err');
                alert(summary);
                state.pendingTasks = [];
                return;
            }
            const t = tasks[i];
            const rname = residentLabel(t.resident);
            const dstr = t.date.toISOString().slice(0, 10);
            postTask(t.payload).then(() => {
                ok++; logMsg('✓ ' + rname + ' ' + dstr + ' — utworzono', 'ok');
            }).catch((e) => {
                fail++; logMsg('✗ ' + rname + ' ' + dstr + ' — ' + e.message, 'err');
            }).finally(() => {
                modalEls.progressLbl.textContent = (i + 1) + '/' + tasks.length + '  ✓ ' + ok + '  ✗ ' + fail;
                setTimeout(() => next(i + 1), 300);
            });
        }
        next(0);
    }

    // ---------- Modal ogólny ----------
    // Renderowany w Shadow DOM z doczepionym Water.css (CDN) - klasyczny "classless"
    // framework style'uje gołe znaczniki (body/button/input/table...), więc bez izolacji
    // zepsułby wygląd całej aplikacji SYRENA. Shadow root nie przepuszcza stylów w żadną
    // stronę. Element-wrapper nazwany "body" (poza realnym dokumentem, ale z tagName BODY)
    // to zwykły trik żeby selektor `body {...}` z frameworku miał się do czego przyczepić.
    const WATER_CSS_URL = 'https://cdn.jsdelivr.net/npm/water.css@2/out/water.css';

    function buildModal(preselectName) {
        if (preselectName) {
            state.selectedResidentIds.clear();
        }
        const host = document.createElement('div');
        const shadow = host.attachShadow({ mode: 'open' });
        shadow.appendChild(el('link', { rel: 'stylesheet', href: WATER_CSS_URL }));
        // Water.css zakłada, że `body` to zwykła strona z tekstem (nadaje mu max-width +
        // margin:auto do wyśrodkowania kolumny treści) - to koliduje z użyciem go tutaj jako
        // pełnoekranowego, flexboxowego tła okna. Nadpisujemy to jawnie, plus wzmacniamy
        // zbyt subtelny domyślny obrys pól formularza.
        shadow.appendChild(el('style', {}, 'body{margin:0;max-width:none;width:100vw;height:100vh;}'
            + 'input,select,textarea{border:1px solid #999;}'));

        const backdrop = el('body', { style: 'position:fixed;inset:0;background:rgba(0,0,0,0.4);z-index:99998;display:flex;align-items:center;justify-content:center;margin:0;' });
        const panel = el('div', { style: 'background:#fff;width:900px;max-width:95vw;max-height:92vh;overflow-y:auto;border-radius:8px;padding:16px;font-family:Segoe UI,Arial,sans-serif;font-size:13px;' });

        const closeBtn = el('button', { type: 'button', style: 'float:right;padding:4px 10px;cursor:pointer;', onclick: () => host.remove() }, 'Zamknij ✕');
        const titleText = preselectName ? '🗓️ Zadania ad hoc dla: ' + preselectName + ' (Tampermonkey)' : '🗓️ Zadania ad hoc (Tampermonkey)';
        panel.appendChild(el('div', {}, el('h2', { style: 'margin:0 0 8px 0;font-size:16px;display:inline-block;' }, titleText), closeBtn));

        panel.appendChild(buildResidentSection(preselectName));
        panel.appendChild(buildKindSection());
        panel.appendChild(buildRoomSection());
        panel.appendChild(buildScheduleSection());

        const preview = el('pre', { style: 'background:#f7f7f7;border:1px solid #ccc;padding:8px;min-height:120px;max-height:220px;overflow:auto;white-space:pre-wrap;font-size:12px;' }, '');
        modalEls.preview = preview;

        const previewBtn = el('button', { type: 'button', style: 'padding:6px 14px;cursor:pointer;', onclick: showPreview }, '① 🔍 Pokaż podgląd');
        const sendBtn = el('button', { type: 'button', style: 'padding:6px 14px;cursor:pointer;margin-left:8px;', disabled: 'disabled', onclick: doSend }, '② ▶ Wyślij zadania');
        modalEls.sendBtn = sendBtn;
        const progressLbl = el('span', { style: 'margin-left:10px;font-size:12px;' }, '');
        modalEls.progressLbl = progressLbl;

        const log = el('div', { style: 'margin-top:8px;background:#1e1e1e;color:#ccc;font-family:Consolas,monospace;font-size:11px;padding:6px;height:110px;overflow-y:auto;' });
        modalEls.log = log;

        panel.appendChild(el('fieldset', { style: 'border:1px solid #ccc;padding:8px;' },
            el('legend', {}, 'Podgląd i wysyłka'),
            el('div', {}, previewBtn, sendBtn, progressLbl),
            preview
        ));
        panel.appendChild(el('div', { style: 'margin-top:6px;' }, el('div', { style: 'font-weight:bold;' }, 'Log'), log));

        backdrop.appendChild(panel);
        backdrop.addEventListener('click', (e) => { if (e.target === backdrop) host.remove(); });
        shadow.appendChild(backdrop);
        document.body.appendChild(host);
    }

    // ---------- Przycisk uruchamiający (globalny, widoczny cały czas) ----------
    function addTriggerButton() {
        if (document.getElementById('adhoc-skrot-trigger')) return;
        const btn = el('button', {
            id: 'adhoc-skrot-trigger',
            type: 'button',
            style: 'position:fixed;top:70px;right:16px;z-index:99997;padding:16px 24px;border-radius:28px;'
                + 'background:#1b5e20;color:#fff;border:3px solid #fff;cursor:pointer;font-size:17px;font-weight:bold;'
                + 'box-shadow:0 3px 12px rgba(0,0,0,0.5);',
            onclick: () => buildModal()
        }, '🗓️ Zadania ad hoc');
        document.body.appendChild(btn);
    }

    // ---------- Przycisk per wiersz mieszkańca, obok natywnego "Dodaj zadanie" ----------
    // Namierza aba-button z tekstem "Dodaj zadanie" (widoczny w master-detail sekcji
    // "Zadania ad hoc" na stronie Usługi mieszkańców), wraca do wiersza planu, z którego
    // ta sekcja się rozwinęła, i czyta Nazwisko/Imię z jego kolumn (aria-colindex 6/7).
    function findResidentNameForDetailButton(btnEl) {
        const detailRow = btnEl.closest('tr.dx-master-detail-row');
        if (!detailRow) return null;
        let sib = detailRow.previousElementSibling;
        while (sib && !sib.classList.contains('dx-data-row')) sib = sib.previousElementSibling;
        if (!sib) return null;
        const surnameTd = sib.querySelector('td[aria-colindex="6"]');
        const firstNameTd = sib.querySelector('td[aria-colindex="7"]');
        if (!surnameTd) return null;
        const surname = (surnameTd.textContent || '').trim();
        const firstName = (firstNameTd ? firstNameTd.textContent || '' : '').trim();
        return (surname + ' ' + firstName).trim() || null;
    }

    function injectRowButtons() {
        document.querySelectorAll('aba-button').forEach((aba) => {
            const dxBtn = aba.querySelector('dx-button[aria-label="Dodaj zadanie"]');
            if (!dxBtn || aba.dataset.adhocSkrotDone) return;
            aba.dataset.adhocSkrotDone = '1';
            const name = findResidentNameForDetailButton(aba);
            if (!name) return;
            const shortcutBtn = el('button', {
                type: 'button',
                style: 'margin-left:8px;padding:9px 16px;border-radius:6px;background:#1b5e20;color:#fff;border:none;cursor:pointer;font-size:15px;font-weight:bold;box-shadow:0 2px 5px rgba(0,0,0,0.35);',
                title: 'Dodaj zadania ad hoc z harmonogramem (Tampermonkey) dla: ' + name,
                onclick: () => buildModal(name)
            }, '🗓️ Harmonogram (TM)');
            aba.parentElement.appendChild(shortcutBtn);
        });
    }

    function start() {
        addTriggerButton();
        new MutationObserver(injectRowButtons).observe(document.body, { childList: true, subtree: true });
        injectRowButtons();
    }

    if (document.body) start();
    else document.addEventListener('DOMContentLoaded', start);
})();
