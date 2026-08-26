// ==UserScript==
// @name         Wyzwalacz – przyciski zamiast checkboxów + podgląd
// @namespace    https://apedps01.bzmw.gov.pl/
// @version      8.1
// @updateURL    https://raw.githubusercontent.com/hardcook69/Skrypty-syrena-python/main/Wyzwalacz.user.js
// @downloadURL  https://raw.githubusercontent.com/hardcook69/Skrypty-syrena-python/main/Wyzwalacz.user.js
// @description  Zamienia checkboxy w oknie "Wyzwalacz" (kreator CRON, wszystkie zakładki) na przyciski (dwa style dla grup Minuta/Godzina/Dzień oraz Miesiąc/Rok), usuwa domyślne zaznaczenie "0" w Minutach/Godzinach, powiększa okno, dwuklik = zaznacz tylko tę wartość.
// @match        *://apedps01.bzmw.gov.pl/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const STYLE_ID = 'wyzwalacz-button-toggle-style';
  const TAB_ORDER = ['MINUTES', 'HOURS', 'DAY', 'MONTH', 'YEAR'];
  const TAB_LABELS = {
    MINUTES: 'Minuty',
    HOURS: 'Godziny',
    DAY: 'Dzień',
    MONTH: 'Miesiąc',
    YEAR: 'Rok',
  };

  // ---------------------------------------------------------------------
  // Style: checkbox -> przycisk
  // Zamiast literalnych klas "c-and-item-*" (które dotyczą tylko Minut/Godzin),
  // używamy wzorców [class*="-item-field"] / [class*="-item-label"] / [class*="-list"],
  // bo Dzień/Miesiąc/Rok używają różnych prefiksów (np. c-and-weekday-item-field,
  // c-and-monthday-item-label), ale zawsze z tym samym sufiksem.
  // ---------------------------------------------------------------------
  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      .c-tab-content [class*="-list"] input[class*="-item-field"].form-check-input {
        position: absolute;
        opacity: 0;
        width: 0;
        height: 0;
        margin: 0;
        pointer-events: none;
      }

      .c-tab-content [class*="-list"] [class*="-item-check"] {
        display: flex;
      }

      .c-tab-content [class*="-list"] label[class*="-item-label"].form-check-label {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 100%;
        min-width: 32px;
        height: 32px;
        margin: 2px 0;
        border-radius: 6px;
        border: 1px solid #cfd8dc;
        background: #f5f7fa;
        color: #37474f;
        font-size: 13px;
        font-weight: 500;
        cursor: pointer;
        user-select: none;
        transition: background .15s ease, color .15s ease, border-color .15s ease, transform .05s ease;
      }
      .c-tab-content [class*="-list"] label[class*="-item-label"].form-check-label:hover {
        border-color: #90a4ae;
        background: #eceff1;
      }
      .c-tab-content [class*="-list"] label[class*="-item-label"].form-check-label:active {
        transform: scale(0.93);
      }

      .c-tab-content[tab-name="MINUTES"] [class*="-list"] input[class*="-item-field"]:checked + label[class*="-item-label"],
      .c-tab-content[tab-name="HOURS"]   [class*="-list"] input[class*="-item-field"]:checked + label[class*="-item-label"],
      .c-tab-content[tab-name="DAY"]     [class*="-list"] input[class*="-item-field"]:checked + label[class*="-item-label"] {
        background: #1976d2;
        border-color: #1976d2;
        color: #ffffff;
      }

      .c-tab-content[tab-name="MONTH"] [class*="-list"] input[class*="-item-field"]:checked + label[class*="-item-label"],
      .c-tab-content[tab-name="YEAR"]  [class*="-list"] input[class*="-item-field"]:checked + label[class*="-item-label"] {
        background: #7b1fa2;
        border-color: #7b1fa2;
        color: #ffffff;
      }

      .c-tab-content [class*="-list"] input[class*="-item-field"]:disabled + label[class*="-item-label"] {
        opacity: 0.55;
      }

      /* Większe okno */
      .dx-overlay-content.dx-popup-normal:has(bs5-quartz-cron) {
        width: 760px !important;
      }

      .wyzwalacz-biweekly {
        margin: 4px 0 10px;
        padding: 10px 12px;
        border: 1px solid #c5cae9;
        border-radius: 6px;
        background: #f5f7ff;
        font-size: 12.5px;
      }
      .wyzwalacz-biweekly h4 {
        margin: 0 0 8px;
        font-size: 13px;
        color: #1a237e;
      }
      .wyzwalacz-biweekly .row {
        display: flex;
        gap: 8px;
        align-items: center;
        margin-bottom: 6px;
        flex-wrap: wrap;
      }
      .wyzwalacz-biweekly label {
        font-weight: 600;
        color: #455a64;
      }
      .wyzwalacz-biweekly input[type="date"] {
        padding: 2px 4px;
        border: 1px solid #cfd8dc;
        border-radius: 4px;
        font-size: 12.5px;
      }
      .wyzwalacz-biweekly button {
        padding: 4px 10px;
        border-radius: 4px;
        border: 1px solid #1976d2;
        background: #1976d2;
        color: #fff;
        cursor: pointer;
        font-size: 12px;
      }
      .wyzwalacz-biweekly button:hover {
        background: #125ea8;
      }
      .wyzwalacz-biweekly .plan-info {
        color: #607d8b;
        font-size: 11.5px;
      }
      .wyzwalacz-biweekly .wz-results {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
      }
      .wyzwalacz-biweekly .wz-results > .plan-info {
        flex: 1 1 100%;
      }
      .wyzwalacz-biweekly .wz-groups-holder {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        flex: 1 1 100%;
      }
      .wyzwalacz-biweekly .result-group {
        flex: 1 1 220px;
        min-width: 200px;
        max-width: 280px;
        margin-top: 0;
        padding: 6px 8px;
        background: #fff;
        border: 1px solid #e0e0e0;
        border-radius: 4px;
      }
      .wyzwalacz-biweekly .result-group b {
        color: #1976d2;
      }
      .wyzwalacz-biweekly .wz-fill-btn {
        margin-top: 4px;
        padding: 3px 8px;
        border-radius: 4px;
        border: 1px solid #2e7d32;
        background: #2e7d32;
        color: #fff;
        cursor: pointer;
        font-size: 11.5px;
      }
      .wyzwalacz-biweekly .wz-fill-btn:hover {
        background: #256428;
      }
      .wyzwalacz-biweekly .wz-fill-btn:disabled {
        opacity: 0.6;
        cursor: default;
      }
    `;
    document.head.appendChild(style);
  }

  // ---------------------------------------------------------------------
  // Pomocnicze: najbliższy kontener listy ("...-list") oraz sprawdzanie klas z sufiksem
  // ---------------------------------------------------------------------
  function hasClassSuffix(el, suffix) {
    if (!el.classList) return false;
    return Array.from(el.classList).some((c) => c.endsWith(suffix));
  }

  function closestItemList(el) {
    let node = el;
    while (node && node !== document.body) {
      if (hasClassSuffix(node, '-list')) return node;
      node = node.parentElement;
    }
    return null;
  }

  // ---------------------------------------------------------------------
  // Domyślne odznaczenie "0" (Minuty / Godziny - tam gdzie lista to dosłownie .c-and-list)
  // ---------------------------------------------------------------------
  function maybeClearDefaultZero(list) {
    if (list.dataset.wyzwalaczZeroHandled === '1') return;
    list.dataset.wyzwalaczZeroHandled = '1';

    let userInteracted = false;
    const onUserInteract = () => { userInteracted = true; };
    list.addEventListener('click', onUserInteract, { capture: true });

    let attempts = 0;
    const maxAttempts = 25; // ~1s
    const stop = () => list.removeEventListener('click', onUserInteract, true);

    const tick = () => {
      if (userInteracted) { stop(); return; }
      attempts += 1;

      const zeroInput = list.querySelector('.c-and-item[item-value="0"] input.c-and-item-field');
      if (zeroInput && zeroInput.checked) {
        const otherChecked = Array.from(list.querySelectorAll('input.c-and-item-field'))
          .some((inp) => inp !== zeroInput && inp.checked);
        if (!otherChecked) {
          zeroInput.click();
        }
        stop();
        return;
      }

      if (attempts < maxAttempts) {
        setTimeout(tick, 40);
      } else {
        stop();
      }
    };

    tick();
  }

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // ---------------------------------------------------------------------
  // Pomocnik "co 2 tygodnie": wykrywa plan z URL, pobiera datę końca planu,
  // liczy terminy co 14 dni i grupuje miesiące o identycznym zestawie dni.
  // UWAGA: tylko liczy i pokazuje - niczego nie wysyła do systemu.
  // ---------------------------------------------------------------------
  const MONTHS_PL = ['Sty', 'Lut', 'Mar', 'Kwi', 'Maj', 'Cze', 'Lip', 'Sie', 'Wrz', 'Paź', 'Lis', 'Gru'];


  function getPlanIdFromUrl() {
    const m = location.pathname.match(/beneficiary-plans\/(\d+)/);
    return m ? m[1] : null;
  }

  function plDateToIso(plDate) {
    const m = (plDate || '').match(/(\d{2})\.(\d{2})\.(\d{4})/);
    if (!m) return null;
    return m[3] + '-' + m[2] + '-' + m[1];
  }

  // Próba odczytu "Data od"/"Data do" bezpośrednio z panelu właściwości planu,
  // widocznego na stronie (aba-simple-data-view > .property). Prostsze i pewniejsze
  // niż fetch do API (który wymagał nieznanych nagłówków/CORS).
  function getPlanDatesFromDOM() {
    const properties = document.querySelectorAll('.property');
    let dataOd = null;
    let dataDo = null;
    properties.forEach((prop) => {
      const label = prop.querySelector('.property-label');
      const dataSpan = prop.querySelector('.property-data span');
      if (!label || !dataSpan) return;
      const labelText = label.textContent.trim();
      if (labelText === 'Data od' && !dataOd) dataOd = dataSpan.textContent.trim();
      if (labelText === 'Data do' && !dataDo) dataDo = dataSpan.textContent.trim();
    });
    return { dataOd, dataDo };
  }

  async function fetchPlanDates(planId) {
    const url = 'http://ttapedps01.bzmw.gov.pl:5070/api/plan/view/' + planId;
    const res = await fetch(url, { credentials: 'include' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    return { validFrom: data.validFrom, validTo: data.validTo, name: data.name };
  }

  function computeBiweeklyDates(startDateStr, endDateStr) {
    const start = new Date(startDateStr + 'T00:00:00');
    const end = new Date(endDateStr + 'T23:59:59');
    const dates = [];
    const d = new Date(start);
    while (d <= end) {
      dates.push(new Date(d));
      d.setDate(d.getDate() + 14);
    }
    return dates;
  }

  function groupByMonth(dates) {
    const map = new Map();
    dates.forEach((d) => {
      const key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      if (!map.has(key)) {
        map.set(key, { label: MONTHS_PL[d.getMonth()] + ' ' + d.getFullYear(), days: [] });
      }
      map.get(key).days.push(d.getDate());
    });
    return map;
  }

  const WEEKDAYS_PL = ['Ndz.', 'Pon.', 'Wt.', 'Śr.', 'Czw.', 'Pt.', 'Sob.']; // index = JS getDay()

  // Grupuje po "który to dzień tygodnia w miesiącu" (1.-5.), zamiast po numerze dnia
  // miesiąca - dzięki temu jeden trigger może obejmować od razu kilka miesięcy,
  // które w danym cyklu wypadają na tę samą "N-tą środę" (czy inny dzień tygodnia).
  function groupByNthWeekday(dates) {
    const map = new Map(); // nth -> Set(month label)
    dates.forEach((d) => {
      const nth = Math.floor((d.getDate() - 1) / 7) + 1;
      const monthLabel = MONTHS_PL[d.getMonth()] + ' ' + d.getFullYear();
      if (!map.has(nth)) map.set(nth, new Set());
      map.get(nth).add(monthLabel);
    });

    const weekdayIndex = dates[0].getDay(); // 0=Ndz..6=Sob (JS)
    const weekdayValue = weekdayIndex + 1; // 1..7 - dopasowane do <select> w appce
    const weekdayLabel = WEEKDAYS_PL[weekdayIndex];

    return Array.from(map.entries())
      .sort((a, b) => a[0] - b[0])
      .map(([nth, monthsSet]) => ({
        nth,
        weekdayValue,
        weekdayLabel,
        months: Array.from(monthsSet),
      }));
  }

  const MONTH_ITEM_VALUES = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

  function monthLabelToNumber(label) {
    const abbr = label.split(' ')[0];
    const idx = MONTHS_PL.indexOf(abbr);
    return idx >= 0 ? MONTH_ITEM_VALUES[idx] : null;
  }

  function switchToTab(container, tabName) {
    return new Promise((resolve) => {
      const btn = Array.from(container.querySelectorAll('.c-tab')).find((b) => b.classList.contains(tabName));
      if (!btn) { resolve(null); return; }
      btn.click();
      let attempts = 40;
      const check = () => {
        const tc = container.querySelector('.c-tab-content[tab-name="' + tabName + '"]');
        if (tc) { resolve(tc); return; }
        attempts -= 1;
        if (attempts <= 0) { resolve(null); return; }
        setTimeout(check, 40);
      };
      check();
    });
  }

  // Klika i SPRAWDZA, czy zmiana faktycznie się przyjęła (odczytuje .checked po
  // odczekaniu), zamiast zakładać że się udało. Jeśli nie - próbuje ponownie.
  // Angular/appka potrafi "cofnąć" zmianę, jeśli kolejne kliknięcia lecą za szybko
  // jedno po drugim bez czasu na przetworzenie.
  function clickAndVerify(input, desiredChecked, attemptsLeft, done) {
    if (input.checked === desiredChecked) { done(true); return; }
    if (attemptsLeft <= 0) {
      console.log('[wyzwalacz] clickAndVerify GAVE UP', { id: input.id, desiredChecked, checkedNow: input.checked });
      done(false);
      return;
    }
    input.click();
    setTimeout(() => {
      console.log('[wyzwalacz] clickAndVerify attempt', {
        id: input.id,
        desiredChecked,
        attemptsLeft,
        checkedNow: input.checked,
      });
      clickAndVerify(input, desiredChecked, attemptsLeft - 1, done);
    }, 180);
  }

  function selectAndValues(tabContent, radioSelector, listSelector, values) {
    return new Promise((resolve) => {
      const radio = tabContent.querySelector(radioSelector);
      const list = tabContent.querySelector(listSelector);
      if (!radio || !list) { resolve(false); return; }
      if (!radio.checked) radio.click();

      let attempts = 30;
      const waitEnabled = () => {
        const items = Array.from(list.querySelectorAll('input[class*="-item-field"]'));
        const allEnabled = items.length > 0 && items.every((inp) => !inp.disabled);
        if (!allEnabled) {
          attempts -= 1;
          if (attempts <= 0) { resolve(false); return; }
          setTimeout(waitEnabled, 40);
          return;
        }
        runTwoPhase();
      };

      const runTwoPhase = () => {
        const targetSet = new Set(values.map(String));

        // FAZA 1: sprawdź czego brakuje spośród wartości docelowych, zaznacz brakujące,
        // POCZEKAJ, potem sprawdź OD NOWA czy naprawdę wszystkie docelowe są zaznaczone
        // (nie zakładaj że się udało tylko dlatego że kliknięcie "przeszło") - powtarzaj
        // rundy, aż będzie komplet.
        const checkMissingRound = (round) => {
          const missing = values.filter((v) => {
            const item = list.querySelector('[item-value="' + v + '"] input[class*="-item-field"]');
            return item && !item.checked;
          });

          console.log('[wyzwalacz] FAZA 1 runda zaznaczania', { round, missingCount: missing.length, missing });

          if (missing.length === 0) {
            console.log('[wyzwalacz] FAZA 1 zakończona - wszystkie docelowe zaznaczone');
            setTimeout(() => uncheckWrongRound(1), 500);
            return;
          }
          if (round > 6) {
            console.log('[wyzwalacz] FAZA 1 poddaję się - nie udało się zaznaczyć wszystkiego', { missing });
            resolve(false);
            return;
          }

          const processMissing = (queue) => {
            const v = queue.shift();
            if (v === undefined) {
              setTimeout(() => checkMissingRound(round + 1), 500);
              return;
            }
            const item = list.querySelector('[item-value="' + v + '"] input[class*="-item-field"]');
            if (!item) { processMissing(queue); return; }
            clickAndVerify(item, true, 4, () => processMissing(queue));
          };
          processMissing(missing.slice());
        };

        // FAZA 2: sprawdź na świeżo co jest zaznaczone, odznacz to, czego nie ma
        // na liście docelowej. Powtarzaj rundy, aż będzie czysto (appka potrafi
        // cofać zmiany z opóźnieniem, więc jedna runda czasem nie wystarcza).
        const uncheckWrongRound = (round) => {
          const wrong = Array.from(list.querySelectorAll('input[class*="-item-field"]'))
            .filter((inp) => {
              const itemDiv = inp.closest('[item-value]');
              const itemValue = itemDiv ? itemDiv.getAttribute('item-value') : inp.value;
              return inp.checked && !targetSet.has(String(itemValue));
            });

          console.log('[wyzwalacz] FAZA 2 runda odznaczania', { round, wrongCount: wrong.length });

          if (wrong.length === 0) {
            console.log('[wyzwalacz] selectAndValues END - wszystko poprawne', { listSelector, values });
            resolve(true);
            return;
          }
          if (round > 6) {
            console.log('[wyzwalacz] selectAndValues END - nie udało się odznaczyć wszystkiego', { listSelector, values, wrong: wrong.map((w) => w.id) });
            resolve(false);
            return;
          }

          const queue = wrong.slice();
          const uncheckNext = () => {
            const inp = queue.shift();
            if (!inp) {
              setTimeout(() => uncheckWrongRound(round + 1), 500);
              return;
            }
            clickAndVerify(inp, false, 4, () => uncheckNext());
          };
          uncheckNext();
        };

        checkMissingRound(1);
      };

      waitEnabled();
    });
  }

  function selectNthWeekday(tabContent, nth, weekdayValue) {
    return new Promise((resolve) => {
      const radio = tabContent.querySelector('.c-nth-check input[type="radio"]');
      const everySelect = tabContent.querySelector('select.c-nth-every');
      const weekdaySelect = tabContent.querySelector('select.c-nth-every-weekday');
      if (!radio || !everySelect || !weekdaySelect) { resolve(false); return; }
      if (!radio.checked) radio.click();

      let attempts = 30;
      const tryFill = () => {
        if (everySelect.disabled || weekdaySelect.disabled) {
          attempts -= 1;
          if (attempts <= 0) { resolve(false); return; }
          setTimeout(tryFill, 40);
          return;
        }
        everySelect.value = String(nth);
        everySelect.dispatchEvent(new Event('change', { bubbles: true }));
        weekdaySelect.value = String(weekdayValue);
        weekdaySelect.dispatchEvent(new Event('change', { bubbles: true }));
        resolve(true);
      };
      tryFill();
    });
  }

  async function fillGroupIntoCurrentTrigger(container, group, timeStr) {
    const failures = [];
    const nameInput = container.querySelector('.dx-texteditor-input');
    if (nameInput && !nameInput.value) {
      nameInput.value = 'Co 2 tyg. - ' + group.nth + '. ' + group.weekdayLabel;
      nameInput.dispatchEvent(new Event('input', { bubbles: true }));
      nameInput.dispatchEvent(new Event('change', { bubbles: true }));
    }

    const dayTab = await switchToTab(container, 'DAY');
    if (dayTab) {
      const ok = await selectNthWeekday(dayTab, group.nth, group.weekdayValue);
      if (!ok) failures.push('Dzień');
    } else {
      failures.push('Dzień (nie udało się przełączyć zakładki)');
    }

    const monthTab = await switchToTab(container, 'MONTH');
    if (monthTab) {
      const monthNums = group.months.map(monthLabelToNumber).filter((n) => n !== null);
      const ok = await selectAndValues(monthTab, '.c-and-check input[type="radio"]', '.c-and-list', monthNums);
      if (!ok) failures.push('Miesiąc');
    } else {
      failures.push('Miesiąc (nie udało się przełączyć zakładki)');
    }

    if (timeStr) {
      const parts = timeStr.split(':');
      const hh = parseInt(parts[0], 10);
      const mm = parseInt(parts[1], 10);
      const hourTab = await switchToTab(container, 'HOURS');
      if (hourTab) {
        const ok = await selectAndValues(hourTab, '.c-and-check input[type="radio"]', '.c-and-list', [hh]);
        if (!ok) failures.push('Godzina');
      } else {
        failures.push('Godzina (nie udało się przełączyć zakładki)');
      }
      const minuteTab = await switchToTab(container, 'MINUTES');
      if (minuteTab) {
        const ok = await selectAndValues(minuteTab, '.c-and-check input[type="radio"]', '.c-and-list', [mm]);
        if (!ok) failures.push('Minuta');
      } else {
        failures.push('Minuta (nie udało się przełączyć zakładki)');
      }
    }

    await switchToTab(container, 'DAY');

    if (failures.length) {
      alert('Uwaga: nie udało się automatycznie ustawić: ' + failures.join(', ') +
        '. Sprawdź te zakładki ręcznie przed zatwierdzeniem.');
    }
  }

  // Stan pomocnika trzymany na poziomie skryptu (nie per-modal), żeby przy kolejnym
  // otwarciu "Dodaj wyzwalacz" nie trzeba było przeliczać terminów od nowa.
  let lastBiweekly = null; // { startVal, endVal, timeVal, groups }
  const usedGroupSignatures = new Set();

  function groupSignature(g) {
    return g.nth + '|' + g.weekdayValue + '|' + g.months.join(',');
  }

  function renderBiweeklyResults(resultsEl, groups, container, timeInput) {
    let html = '';
    groups.forEach((g, i) => {
      const used = usedGroupSignatures.has(groupSignature(g));
      html += '<div class="result-group" data-group-index="' + i + '">' +
        '<div><b>Trigger ' + (i + 1) + '</b> — ' + g.nth + '. ' + escapeHtml(g.weekdayLabel) + ' miesiąca' +
        (used ? ' <span style="color:#2e7d32">✓ użyty</span>' : '') + '</div>' +
        '<div>Miesiące: <b>' + escapeHtml(g.months.join(', ')) + '</b></div>' +
        '<button type="button" class="wz-fill-btn" data-group-index="' + i + '">Wypełnij ten formularz</button>' +
        '</div>';
    });
    resultsEl.innerHTML = html;

    resultsEl.querySelectorAll('.wz-fill-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.getAttribute('data-group-index'), 10);
        const group = groups[idx];
        const timeVal = timeInput.value;
        if (!timeVal) {
          alert('Uzupełnij najpierw godzinę uruchomienia.');
          return;
        }
        const allBtns = Array.from(resultsEl.querySelectorAll('.wz-fill-btn'));
        allBtns.forEach((b) => { b.disabled = true; });
        btn.textContent = 'Wypełnianie...';
        fillGroupIntoCurrentTrigger(container, group, timeVal).then(() => {
          usedGroupSignatures.add(groupSignature(group));
          renderBiweeklyResults(resultsEl, groups, container, timeInput);
        });
      });
    });
  }

  function injectBiweeklyHelper(container) {
    if (container.querySelector('.wyzwalacz-biweekly')) return;

    const wrap = document.createElement('div');
    wrap.className = 'wyzwalacz-biweekly';
    wrap.innerHTML =
      '<h4>Harmonogram co 2 tygodnie (pomocnik) ' +
      '<button type="button" class="wz-toggle" style="float:right;padding:1px 8px;font-size:11px;">Zwiń</button></h4>' +
      '<div class="wz-body">' +
      '<div class="row"><label>Data pierwszego wystąpienia:</label>' +
      '<input type="date" class="wz-start-date"></div>' +
      '<div class="row"><label>Koniec planu:</label>' +
      '<input type="date" class="wz-end-date">' +
      '<span class="plan-info wz-plan-info">wykrywanie planu...</span></div>' +
      '<div class="row"><label>Godzina uruchomienia:</label>' +
      '<input type="time" class="wz-time"></div>' +
      '<div class="row"><button type="button" class="wz-calc-btn">Oblicz terminy</button></div>' +
      '<div class="wz-results"></div>' +
      '</div>';

    const toggleBtn = wrap.querySelector('.wz-toggle');
    const bodyEl = wrap.querySelector('.wz-body');
    const setCollapsed = (collapsed) => {
      bodyEl.style.display = collapsed ? 'none' : '';
      toggleBtn.textContent = collapsed ? 'Rozwiń' : 'Zwiń';
    };
    toggleBtn.addEventListener('click', () => {
      setCollapsed(bodyEl.style.display !== 'none');
    });
    setCollapsed(!(lastBiweekly && lastBiweekly.groups));

    const cronRow = container.querySelector('bs5-quartz-cron');
    const targetRow = cronRow ? cronRow.closest('.row') : null;
    if (targetRow && targetRow.parentNode) {
      targetRow.parentNode.insertBefore(wrap, targetRow);
    } else {
      container.insertBefore(wrap, container.firstChild);
    }

    const endInput = wrap.querySelector('.wz-end-date');
    const planInfo = wrap.querySelector('.wz-plan-info');
    const startInput = wrap.querySelector('.wz-start-date');
    const timeInput = wrap.querySelector('.wz-time');
    const resultsEl = wrap.querySelector('.wz-results');
    const calcBtn = wrap.querySelector('.wz-calc-btn');

    startInput.value = new Date().toISOString().slice(0, 10);

    if (lastBiweekly) {
      startInput.value = lastBiweekly.startVal || startInput.value;
      endInput.value = lastBiweekly.endVal || '';
      timeInput.value = lastBiweekly.timeVal || '';
      if (lastBiweekly.groups) {
        const groupsHolder = document.createElement('div');
        groupsHolder.className = 'wz-groups-holder';
        resultsEl.appendChild(groupsHolder);
        renderBiweeklyResults(groupsHolder, lastBiweekly.groups, container, timeInput);
      }
    }

    const domDates = getPlanDatesFromDOM();
    if (domDates.dataDo) {
      const iso = plDateToIso(domDates.dataDo);
      if (iso) endInput.value = iso;
      planInfo.textContent = 'Plan (ze strony): od ' + (domDates.dataOd || '?') + ' do ' + domDates.dataDo +
        ' - sprawdź czy to właściwy plan';
    } else {
      const planId = getPlanIdFromUrl();
      if (planId) {
        fetchPlanDates(planId).then((info) => {
          if (info.validTo) {
            endInput.value = info.validTo.slice(0, 10);
          }
          planInfo.textContent = 'Plan: ' + (info.name || planId) + (info.validTo ? '' : ' (brak daty końca)');
        }).catch(() => {
          planInfo.textContent = 'Nie udało się pobrać planu (id ' + planId + ') - wpisz datę ręcznie';
        });
      } else {
        planInfo.textContent = 'Nie wykryto daty planu na stronie ani ID w adresie - wpisz datę ręcznie';
      }
    }

    calcBtn.addEventListener('click', () => {
      const startVal = startInput.value;
      const endVal = endInput.value;
      if (!startVal || !endVal) {
        resultsEl.innerHTML = '<div style="color:#c62828">Uzupełnij obie daty.</div>';
        return;
      }
      const dates = computeBiweeklyDates(startVal, endVal);
      if (!dates.length) {
        resultsEl.innerHTML = '<div style="color:#c62828">Brak terminów w podanym zakresie.</div>';
        return;
      }
      const groups = groupByNthWeekday(dates);
      const weekday = dates[0].toLocaleDateString('pl-PL', { weekday: 'long' });

      lastBiweekly = { startVal, endVal, timeVal: timeInput.value, groups };

      const info = document.createElement('div');
      info.className = 'plan-info';
      info.innerHTML = 'Dzień tygodnia: <b>' + escapeHtml(weekday) + '</b>, liczba terminów: ' + dates.length;
      resultsEl.innerHTML = '';
      resultsEl.appendChild(info);
      const groupsHolder = document.createElement('div');
      groupsHolder.className = 'wz-groups-holder';
      resultsEl.appendChild(groupsHolder);
      renderBiweeklyResults(groupsHolder, groups, container, timeInput);
    });
  }

  // ---------------------------------------------------------------------
  // Obserwacja DOM
  // ---------------------------------------------------------------------
  function handleAddedNode(node) {
    if (!(node instanceof HTMLElement)) return;

    const containers = node.matches && node.matches('.schedule-trigger-form-container')
      ? [node]
      : Array.from(node.querySelectorAll ? node.querySelectorAll('.schedule-trigger-form-container') : []);

    containers.forEach((container) => {
      if (!container.querySelector('bs5-quartz-cron')) return;
      if (container.dataset.wyzwalaczSummaryInit === '1') return;
      container.dataset.wyzwalaczSummaryInit = '1';
      injectBiweeklyHelper(container);

      const dialog = container.closest('.dx-overlay-content.dx-popup-normal');
      if (dialog && dialog.style.width && parseInt(dialog.style.width, 10) < 760) {
        dialog.style.width = '760px';
      }
    });

    // Domyślne "0" (tylko literalna .c-and-list z Minut/Godzin)
    const lists = node.matches && node.matches('.c-and-list')
      ? [node]
      : Array.from(node.querySelectorAll ? node.querySelectorAll('.c-and-list') : []);
    lists.forEach((list) => {
      const tabContent = list.closest('.c-tab-content');
      const tabName = tabContent ? tabContent.getAttribute('tab-name') : null;
      if (tabName === 'MINUTES' || tabName === 'HOURS') {
        maybeClearDefaultZero(list);
      }
    });
  }

  function handleRemovedNode() {}

  injectStyles();

  const observer = new MutationObserver((mutations) => {
    injectStyles();
    for (const mutation of mutations) {
      mutation.addedNodes.forEach(handleAddedNode);
      mutation.removedNodes.forEach(handleRemovedNode);
    }
  });

  observer.observe(document.documentElement, { childList: true, subtree: true });

  // Wspólny mechanizm dla klikania na zablokowane (disabled) wartości: aktywuje tryb
  // (radio w obrębie .c-segment), czeka aż pole się odblokuje, wykonuje akcję.
  // "Claim" (token per input.id) zapewnia, że jeśli druga operacja (np. dblclick)
  // zacznie się w trakcie oczekiwania pierwszej (np. click), starsza się wycofa,
  // zamiast nadpisać efekt nowszej z opóźnieniem.
  const activeOps = new Map();

  function claimOp(input) {
    const token = {};
    activeOps.set(input.id, token);
    return token;
  }

  function isCurrentOp(input, token) {
    return activeOps.get(input.id) === token;
  }

  function enableAndRun(input, onEnabled) {
    const token = claimOp(input);
    const segment = input.closest('.c-segment');
    const modeRadio = segment ? segment.querySelector('input[type="radio"]') : null;
    if (modeRadio && !modeRadio.checked) {
      modeRadio.click();
    }

    let attempts = 20;
    const step = () => {
      if (!isCurrentOp(input, token)) return; // przejęte przez nowszą operację
      if (!document.body.contains(input)) return;
      attempts -= 1;
      if (!input.disabled) {
        onEnabled(input);
        return;
      }
      if (attempts > 0) setTimeout(step, 30);
    };
    step();
  }

  // Kliknięcie w zablokowaną (disabled) wartość -> aktywuj tryb, zaznacz wartość po odblokowaniu
  document.addEventListener('click', (e) => {
    const label = e.target.closest && e.target.closest('label[class*="-item-label"]');
    if (!label) return;
    const forId = label.getAttribute('for');
    const input = forId ? document.getElementById(forId) : null;
    if (!input || !input.disabled) return;

    e.preventDefault();
    e.stopPropagation();

    enableAndRun(input, (inp) => inp.click());
  }, true);

  // Podwójny klik na wartość -> odznacz resztę w tej samej liście, zostaw tylko klikniętą.
  // Odznaczamy po jednej wartości na raz, za każdym razem na nowo pobierając element z DOM
  // po id (Angular potrafi przebudować listę po każdej zmianie, więc trzymanie starych
  // referencji do wielu elementów naraz gubiło kolejne kliknięcia).
  document.addEventListener('dblclick', (e) => {
    const label = e.target.closest && e.target.closest('label[class*="-item-label"]');
    if (!label) return;
    const list = closestItemList(label);
    if (!list) return;
    const forId = label.getAttribute('for');
    if (!forId) return;

    e.preventDefault();
    console.log('[wyzwalacz] dblclick START', { forId, listClass: list.className });

    const finalize = () => {
      const target = document.getElementById(forId);
      if (!target) { console.log('[wyzwalacz] dblclick target not found', { forId }); return; }

      const doUnchecks = () => {
        if (!document.body.contains(list)) return;
        const others = Array.from(list.querySelectorAll('input[class*="-item-field"]'))
          .filter((inp) => inp.id !== forId && inp.checked && !inp.disabled);
        if (!others.length) { settleAndReconcile(); return; }
        clickAndVerify(others[0], false, 4, () => doUnchecks());
      };

      let reconcileRounds = 0;
      const settleAndReconcile = () => {
        setTimeout(() => {
          if (!document.body.contains(list)) return;
          const targetNow = document.getElementById(forId);
          const wrong = Array.from(list.querySelectorAll('input[class*="-item-field"]'))
            .filter((inp) => (inp.id === forId ? !inp.checked : inp.checked && !inp.disabled));

          reconcileRounds += 1;
          console.log('[wyzwalacz] dblclick reconcile round', { round: reconcileRounds, wrongCount: wrong.length, wrong: wrong.map((w) => w.id) });

          if (wrong.length === 0 || reconcileRounds >= 6) {
            console.log('[wyzwalacz] dblclick DONE', { targetChecked: targetNow && targetNow.checked });
            return;
          }
          const fixNext = () => {
            const inp = wrong.shift();
            if (!inp) { settleAndReconcile(); return; }
            clickAndVerify(inp, inp.id === forId, 4, () => fixNext());
          };
          fixNext();
        }, 700);
      };

      clickAndVerify(target, true, 4, () => doUnchecks());
    };

    const target = document.getElementById(forId);
    if (target && target.disabled) {
      enableAndRun(target, () => finalize());
    } else {
      finalize();
    }
  });

  // Obsłuż okno już otwarte w momencie startu skryptu
  document.querySelectorAll('.schedule-trigger-form-container').forEach((container) => {
    if (!container.querySelector('bs5-quartz-cron')) return;
    if (container.dataset.wyzwalaczSummaryInit === '1') return;
    container.dataset.wyzwalaczSummaryInit = '1';
    injectBiweeklyHelper(container);
  });
  document.querySelectorAll('.c-and-list').forEach((list) => {
    const tabContent = list.closest('.c-tab-content');
    const tabName = tabContent ? tabContent.getAttribute('tab-name') : null;
    if (tabName === 'MINUTES' || tabName === 'HOURS') {
      maybeClearDefaultZero(list);
    }
  });
})();
