// ==UserScript==
// @name         Eksport spotkań do Excela
// @namespace    https://apedps01.bzmw.gov.pl/
// @version      1.0
// @updateURL    https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/EksportSpotkanSkrot.user.js
// @downloadURL  https://raw.githubusercontent.com/hardcook69/Syrena-Tempermokey/main/EksportSpotkanSkrot.user.js
// @description  Przycisk widoczny TYLKO na stronie /meetings, eksportujący listę spotkań (ze szczegółami i uczestnikami podzielonymi na pracownicy/mieszkańcy) do pliku .xlsx wprost z przeglądarki - odpowiednik eksport_spotkan_gui.py, ale bez osobnego logowania (używa tokena sesji, którą masz już otwartą).
// @match        *://apedps01.bzmw.gov.pl/*
// @match        *://ttapedps01.bzmw.gov.pl/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @run-at       document-start
// @grant        none
// ==/UserScript==

// Ten sam mechanizm zbierania danych i te same, POTWIERDZONE przechwyconym
// ruchem endpointy co w eksport_spotkan_gui.py -- patrz komentarz na górze
// tamtego pliku po pełny opis i zastrzeżenia (mapowanie Wnioski<-summary
// i Link do spotkania<-urlLink niepotwierdzone na wypełnionych danych,
// uczestnicy dzieleni na Pracownicy/Mieszkańcy po przynależności ID do
// odpowiedniego słownika /api/visitor, nie po znaczeniu pól
// leaders/members/subjects/plannedSubjects).

(function () {
    'use strict';

    const AUTH_PORT = 5010;
    const EMPLOYEE_PORT = 5000;
    const BENEFICIARY_PORT = 5020;
    const ORG_ID = 1;
    const BUTTON_ID = 'eksport-spotkan-btn';

    // ---------- Podsłuch tokena Bearer (ten sam mechanizm co ZadaniaAdHocSkrot.user.js) ----------
    let capturedToken = null;
    const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
        if (typeof name === 'string' && name.toLowerCase() === 'authorization' && /^Bearer /.test(value)) {
            capturedToken = value;
        }
        return origSetHeader.apply(this, arguments);
    };

    // ---------- Podsłuch AKTUALNEGO filtra listy spotkań ----------
    // Strona sama woła /api/meeting/by-organization-id/paged z parametrami
    // filtra (search/searchFields/kindId/itd. -- dokładnych nazw dla filtra
    // "Rodzaje spotkań" nie znamy, więc zamiast zgadywać, po prostu
    // podsłuchujemy CAŁY query string ostatniego takiego requestu strony i
    // używamy go 1:1 przy eksporcie (nadpisując tylko page/pageSize, żeby
    // pobrać WSZYSTKIE strony pasujące do filtra, nie tylko bieżącą).
    let lastMeetingListParams = null;
    const origOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function (method, url) {
        try {
            if (typeof url === 'string' && url.indexOf('/api/meeting/by-organization-id/paged') !== -1) {
                const qIndex = url.indexOf('?');
                if (qIndex !== -1) lastMeetingListParams = new URLSearchParams(url.slice(qIndex + 1));
            }
        } catch (e) { /* ignoruj -- w najgorszym razie eksport pobierze bez filtra */ }
        return origOpen.apply(this, arguments);
    };

    function apiBase(port) {
        return 'http://' + location.hostname + ':' + port;
    }

    function authFetch(url) {
        return fetch(url, { headers: { Authorization: capturedToken } }).then((r) => {
            if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + url);
            return r.json();
        });
    }

    // ---------- Fetche danych ----------
    function fetchMeetings() {
        // Baza to dokładnie to, co strona sama ostatnio wysłała (patrz podsłuch
        // wyżej) -- więc dowolny aktywny filtr (search, kind itd.) jest
        // zachowany 1:1. Brak podsłuchanego requestu (np. eksport kliknięty
        // zanim strona zdążyła cokolwiek pobrać) -> pusty filtr, jak dotychczas.
        // page/pageSize zawsze nadpisujemy, żeby przejść WSZYSTKIE strony.
        const all = [];
        function page(p) {
            const qs = lastMeetingListParams ? new URLSearchParams(lastMeetingListParams) : new URLSearchParams();
            qs.set('page', p);
            qs.set('pageSize', 200);
            if (!qs.has('orderBy')) qs.set('orderBy', 'meetingTime');
            if (!qs.has('ascending')) qs.set('ascending', 'true');
            return authFetch(apiBase(BENEFICIARY_PORT) + '/api/meeting/by-organization-id/paged?' + qs).then((data) => {
                const items = data.results || [];
                all.push(...items);
                const totalPages = data.totalNumberOfPages;
                if (totalPages != null ? p < totalPages : items.length === 200) return page(p + 1);
            });
        }
        return page(1).then(() => all);
    }

    function fetchMeetingDetail(id) {
        return authFetch(apiBase(BENEFICIARY_PORT) + '/api/meeting/' + id).catch(() => ({}));
    }

    function fetchDictionaryByKind(kind) {
        return authFetch(apiBase(EMPLOYEE_PORT) + '/api/dictionary-value/by-dictionary-kind/' + kind + '?withAttributes=true').catch(() => []);
    }

    function fetchVisitors(typeList) {
        // UWAGA: w przechwyconym ruchu ten request NIE MA parametru
        // organizationId (tylko typeList) -- serwer najwyraźniej ustala
        // organizację z samego tokena. Wcześniej dokładaliśmy organizationId
        // "na wszelki wypadek"; usunięte, żeby dokładnie odzwierciedlać
        // potwierdzony ruch zamiast zgadywać.
        const qs = new URLSearchParams({ typeList });
        return authFetch(apiBase(AUTH_PORT) + '/api/visitor/by-organization-id?' + qs).catch(() => []);
    }

    function visitorLabel(v) {
        return (((v.surname || '') + ' ' + (v.firstName || '')).trim()) || ('#' + v.id);
    }

    function fetchAllBeneficiaries() {
        // Dodatkowe źródło ID/nazwisk mieszkańców -- ZAŁOŻENIE do zweryfikowania:
        // użytkownik potwierdził że pola uczestników spotkania (leaders/members/
        // subjects/plannedSubjects) zawierają ID mieszkańców, ale nie wiadomo czy
        // to ta sama przestrzeń ID co /api/visitor (typeList=3) czy ta z
        // /api/beneficiary. Pobieramy WSZYSTKICH mieszkańców (bez filtra statusu)
        // i łączymy z listą z /api/visitor zamiast zgadywać, która jest właściwa.
        const all = [];
        function page(p) {
            const qs = new URLSearchParams({ page: p, pageSize: 200, organizationId: ORG_ID });
            return authFetch(apiBase(BENEFICIARY_PORT) + '/api/beneficiary/by-organization-id/paged?' + qs).then((data) => {
                const items = data.results || [];
                all.push(...items);
                const totalPages = data.totalNumberOfPages;
                if (totalPages != null ? p < totalPages : items.length === 200) return page(p + 1);
            });
        }
        return page(1).then(() => all).catch(() => all);
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

    function setButtonState(text, disabled) {
        const btn = document.getElementById(BUTTON_ID);
        if (!btn) return;
        btn.textContent = text;
        btn.disabled = !!disabled;
        btn.style.opacity = disabled ? '0.6' : '1';
        btn.style.cursor = disabled ? 'default' : 'pointer';
    }

    function downloadXlsx(rows) {
        const cols = ['Data spotkania', 'Rodzaj spotkania', 'Miejsce', 'Temat', 'Cel', 'Wnioski',
                      'Link do spotkania', 'Pracownicy (uczestnicy)', 'Mieszkańcy (uczestnicy)'];
        const aoa = [cols].concat(rows.map((r) => cols.map((c) => r[c] || '')));
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols'] = [{ wch: 18 }, { wch: 22 }, { wch: 20 }, { wch: 34 }, { wch: 26 },
                       { wch: 34 }, { wch: 26 }, { wch: 34 }, { wch: 34 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Spotkania');
        const stamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');
        XLSX.writeFile(wb, 'spotkania_' + stamp + '.xlsx');
    }

    function runExport() {
        if (typeof XLSX === 'undefined') {
            alert('Biblioteka do zapisu .xlsx nie wczytała się (brak dostępu do CDN?). Sprawdź połączenie i odśwież stronę.');
            return;
        }
        if (!capturedToken) {
            alert('Nie złapałem jeszcze tokena logowania — odśwież stronę, poczekaj aż się w pełni załaduje, i spróbuj ponownie.');
            return;
        }
        const activeSearch = lastMeetingListParams && lastMeetingListParams.get('search');
        setButtonState(activeSearch ? `Pobieranie (filtr: "${activeSearch}")...` : 'Pobieranie listy...', true);
        Promise.all([
            fetchDictionaryByKind(72),
            fetchDictionaryByKind(94),
            fetchVisitors(4),
            fetchVisitors(3),
            fetchAllBeneficiaries(),
            fetchMeetings(),
        ]).then(([placesArr, kindsArr, employees, residentsVisitor, residentsBeneficiary, meetings]) => {
            const places = Object.fromEntries(placesArr.map((v) => [v.id, v.content]));
            const kinds = Object.fromEntries(kindsArr.map((v) => [v.id, v.content]));
            const residents = residentsVisitor.concat(residentsBeneficiary);
            const employeeIds = new Set(employees.map((v) => v.id));
            const residentIds = new Set(residents.map((v) => v.id));
            const nameById = Object.fromEntries(employees.concat(residents).map((v) => [v.id, visitorLabel(v)]));

            const rows = [];
            function next(i) {
                if (i >= meetings.length) {
                    downloadXlsx(rows);
                    setButtonState('📊 Eksport spotkań do Excela', false);
                    return;
                }
                const m = meetings[i];
                setButtonState('Pobieranie ' + (i + 1) + '/' + meetings.length + '...', true);
                fetchMeetingDetail(m.id).then((detail) => {
                    const placeId = detail.placeId;
                    const miejsce = places[placeId] || m.placeName || '';
                    const kindId = detail.kindId != null ? detail.kindId : m.kindId;
                    const rodzaj = kinds[kindId] || m.kindName || '';
                    const participantIds = new Set();
                    ['leaders', 'members', 'subjects', 'plannedSubjects'].forEach((f) => {
                        (detail[f] || []).forEach((id) => participantIds.add(id));
                    });
                    const pracownicy = [...participantIds].filter((id) => employeeIds.has(id))
                        .map((id) => nameById[id] || ('#' + id)).sort();
                    const mieszkancy = [...participantIds].filter((id) => residentIds.has(id))
                        .map((id) => nameById[id] || ('#' + id)).sort();
                    rows.push({
                        'Data spotkania': detail.meetingTime || m.meetingTime || '',
                        'Rodzaj spotkania': rodzaj,
                        'Miejsce': miejsce,
                        'Temat': detail.topic || m.topic || '',
                        'Cel': detail.purpose || m.purpose || '',
                        'Wnioski': detail.summary || '',
                        'Link do spotkania': detail.urlLink || detail.link || '',
                        'Pracownicy (uczestnicy)': pracownicy.join(', '),
                        'Mieszkańcy (uczestnicy)': mieszkancy.join(', '),
                    });
                    next(i + 1);
                }).catch(() => next(i + 1));
            }
            next(0);
        }).catch((e) => {
            alert('Błąd eksportu: ' + e.message);
            setButtonState('📊 Eksport spotkań do Excela', false);
        });
    }

    // ---------- Przycisk widoczny TYLKO na /meetings, wstawiony obok natywnych ----------
    // Kontener z prawdziwymi przyciskami ("Generuj raport", "Dodaj spotkanie") --
    // dokładna klasa potwierdzona ze strony "Rejestr spotkań" przez użytkownika.
    // Wstawiamy się jako pierwsze dziecko (najbardziej na lewo w tym rzędzie),
    // zamiast floatować nad interfejsem jak wcześniej (position:fixed zasłaniał
    // "Generuj raport").
    function findButtonContainer() {
        return document.querySelector('.col-5.offset-7.btn-container')
            || document.querySelector('.btn-container');
    }

    function injectButton() {
        if (document.getElementById(BUTTON_ID)) return true;
        const container = findButtonContainer();
        if (!container) return false;
        const btn = el('button', {
            id: BUTTON_ID,
            text: '📊 Eksport spotkań do Excela',
            style: 'display:inline-flex; align-items:center; margin-right:8px; padding:6px 12px; ' +
                   'background:#2c5f8a; color:#fff; border:none; border-radius:4px; ' +
                   'font:600 13px "Segoe UI",sans-serif; cursor:pointer; vertical-align:middle;',
            onclick: runExport,
        });
        container.insertBefore(btn, container.firstChild);
        return true;
    }

    function removeButton() {
        const btn = document.getElementById(BUTTON_ID);
        if (btn) btn.remove();
    }

    function isMeetingsPage() {
        return location.pathname.replace(/\/+$/, '').endsWith('/meetings');
    }

    function syncButtonVisibility() {
        if (isMeetingsPage()) injectButton();
        else removeButton();
    }

    // SPA (nawigacja bez przeładowania strony, kontener przycisków renderuje
    // się asynchronicznie) -- pilnujemy co 500ms zarówno zmiany adresu, jak i
    // (re)pojawienia się kontenera, żeby doczekać się aż Angular go wyrenderuje.
    let lastPath = location.pathname;
    setInterval(() => {
        if (location.pathname !== lastPath) {
            lastPath = location.pathname;
        }
        syncButtonVisibility();
    }, 500);

    document.addEventListener('DOMContentLoaded', syncButtonVisibility);
    window.addEventListener('load', syncButtonVisibility);
})();
