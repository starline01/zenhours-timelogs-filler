// ==UserScript==
// @name         Zenhours DTR Filler
// @namespace    starlinesecuritygroup.com
// @version      1.4.0
// @description  Paste a block of timelogs (date + times) and auto-fill the Zenhours timelogs table. Fills only — you click Save.
// @author       Starline Security Group
// @match        *://*.zenoras.com/*
// @match        *://zenoras.com/*
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @require      https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js
// @updateURL    https://raw.githubusercontent.com/starline01/zenhours-timelogs-filler/main/zenhours-dtr-filler.user.js
// @downloadURL  https://raw.githubusercontent.com/starline01/zenhours-timelogs-filler/main/zenhours-dtr-filler.user.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==

// ─────────────────────────────────────────────────────────────────────────
//  UPDATING EVERY PC AT ONCE
// ─────────────────────────────────────────────────────────────────────────
//  Installed from GitHub, Tampermonkey re-checks the @updateURL above and
//  pulls a new copy automatically — but ONLY when @version is higher than the
//  installed one. Editing the code without bumping @version updates nothing.
//
//  So the release routine is: change the code → bump @version → commit → push.
//  Each PC picks it up on its next update check (Tampermonkey's default is
//  daily; the dashboard's "Check for userscript updates" forces it now).
//
//  The raw.githubusercontent.com CDN caches for about five minutes, so a push
//  is not visible instantly.
//
// ─────────────────────────────────────────────────────────────────────────
//  SCOPE — every zenoras.com subdomain
// ─────────────────────────────────────────────────────────────────────────
//  The two @match lines above cover any subdomain (rr.zenoras.com,
//  anything-else.zenoras.com) plus the bare domain. @match lines are OR'd, so
//  add more lines for any site on a different domain entirely — one per line,
//  directly under the others.
//
//  The panel only appears once it detects a timelogs table, so it stays out of
//  the way everywhere else on those sites. To restrict it to the timelogs
//  pages only, swap the two lines for:
//      // @match        *://*.zenoras.com/hr/manage_timelogs/*
//
//  (These notes use // comments on purpose: a /* */ block cannot contain an
//  @match glob, because the */ inside a *://*/* pattern closes it early.)
//
// ─────────────────────────────────────────────────────────────────────────
//  HOW TO USE
// ─────────────────────────────────────────────────────────────────────────
//  1. Open the Zenhours timelogs page and Search your date range.
//  2. A "DTR Filler" panel appears at the top-right. (Ctrl+Shift+D toggles it.)
//  3. Copy your rows out of Excel / Google Sheets and paste into the box:
//
//       8/1/2026<TAB>TUESDAY<TAB>10:20<TAB>12:40<TAB>13:12<TAB>17:00<TAB>17:30<TAB>21:00
//       8/2/2026<TAB>WEDNESDAY<TAB>10:25<TAB>12:45<TAB>13:15<TAB>16:39<TAB>17:09<TAB>21:00
//
//     Column order is the table's own order:
//       Date | Time In | Lunch Out | Lunch In | Break Out | Break In | Time Out
//
//     A day-of-week column (TUESDAY) — or any other label column sitting
//     between the date and the first time — is detected and ignored.
//
//     Times may be 24-hour (13:12) or 12-hour (1:12 PM); both are converted to
//     whatever format Zenhours' own input is already showing.
//
//     Leave a cell empty (two tabs in a row) or put a dash to skip a column —
//     e.g. a guard with no lunch or break. By default the script does not touch
//     those fields, so they keep Zenhours' prefilled 12:00 AM and the log warns
//     you about each one. Tick "Clear the field when my cell is blank/dash" if
//     you would rather have the script empty them instead.
//
//     If your copied block includes a header row ("Date, Time In, ..."), the
//     script reads it and maps the columns by name instead of by position —
//     so a different column order still works.
//
//  3b. OR load a whole workbook: click "Load Excel, CSV or a scan…" and pick a file
//     holding every guard (EmpID · Name · Date · Day · In · LOut · LIn · BOut ·
//     BIn · Out). Editing in Zenoras is per employee, so the script reads the
//     access ID and name shown on the page, finds that guard in the file, and
//     loads only their days. Everyone else is ignored.
//
//     The workbook stays loaded as you move between employees — open the next
//     guard's page and it re-matches automatically. Filled guards are ticked in
//     the list so you can see who is left.
//
//     If the page cannot be matched to exactly one guard, nothing is selected
//     and filling is BLOCKED until you choose from the list. Writing one
//     guard's hours onto another is the one mistake worth stopping for.
//
//  4. Click "Parse" to check what was read, then "Test 1st day" to try a
//     single row, then "Fill all rows".
//  5. The script clicks Edit and types the times. It NEVER clicks Save —
//     review the green-highlighted inputs, then Save each row yourself.
//
//  "Undo fill" restores every input the script touched back to its original
//  value (only works while the rows are still open in edit mode).

(function () {
    'use strict';

    if (window.top !== window.self) return;      // don't run inside iframes
    if (window.__zdfLoaded) return;
    window.__zdfLoaded = true;

    // ── Column vocabulary ────────────────────────────────────────────────
    // Canonical column keys, in the table's left-to-right order.
    const COLUMNS = ['time_in', 'lunch_out', 'lunch_in', 'break_out', 'break_in', 'time_out'];

    const COLUMN_LABELS = {
        time_in: 'Time In',
        lunch_out: 'Lunch Out',
        lunch_in: 'Lunch In',
        break_out: 'Break Out',
        break_in: 'Break In',
        time_out: 'Time Out'
    };

    // Header text (normalized) → canonical key. Covers common wording drift.
    const HEADER_ALIASES = {
        'time in': 'time_in', 'timein': 'time_in', 'in': 'time_in',
        'clock in': 'time_in', 'clockin': 'time_in', 'start': 'time_in',
        'lunch out': 'lunch_out', 'lunchout': 'lunch_out', 'lunch start': 'lunch_out',
        'lunch in': 'lunch_in', 'lunchin': 'lunch_in', 'lunch end': 'lunch_in',
        'break out': 'break_out', 'breakout': 'break_out', 'break start': 'break_out',
        'break in': 'break_in', 'breakin': 'break_in', 'break end': 'break_in',
        'time out': 'time_out', 'timeout': 'time_out', 'out': 'time_out',
        'clock out': 'time_out', 'clockout': 'time_out', 'end': 'time_out',
        // Abbreviated spreadsheet headers: In · LOut · LIn · BOut · BIn · Out
        'lout': 'lunch_out', 'l out': 'lunch_out',
        'lin': 'lunch_in', 'l in': 'lunch_in',
        'bout': 'break_out', 'b out': 'break_out',
        'bin': 'break_in', 'b in': 'break_in'
    };

    // Spreadsheet columns that identify the guard rather than a time.
    const EMPID_ALIASES = new Set([
        'empid', 'emp id', 'employee id', 'employeeid', 'id', 'access id', 'accessid',
        'access no', 'access number', 'access', 'emp no', 'employee no', 'employee number',
        'badge', 'badge id', 'badge no', 'idno', 'id no'
    ]);
    const NAME_ALIASES = new Set([
        'name', 'employee', 'employee name', 'employeename', 'guard', 'guard name',
        'full name', 'fullname', 'personnel', 'staff'
    ]);
    const DATE_ALIASES = new Set(['date', 'log date', 'work date', 'day date']);
    const DAY_ALIASES = new Set(['day', 'weekday', 'day of week', 'dow']);

    const BLANK_MARKERS = ['--:--', '-- : --', '--', '-', '—', '–', '', 'n/a', 'na', 'none'];

    // Label columns that sit between the date and the times and must be ignored.
    const DAY_NAMES = new Set([
        'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday',
        'mon', 'tue', 'tues', 'wed', 'weds', 'thu', 'thur', 'thurs', 'fri', 'sat', 'sun'
    ]);

    const LS_KEY = 'zdf.settings.v1';

    // ── Small utilities ──────────────────────────────────────────────────
    const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
    const normKey = (s) => norm(s).toLowerCase().replace(/[^a-z ]/g, '').replace(/\s+/g, ' ').trim();
    const pad2 = (n) => String(n).padStart(2, '0');
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const isBlankText = (s) => BLANK_MARKERS.includes(norm(s).toLowerCase());

    function loadSettings() {
        try { return JSON.parse(localStorage.getItem(LS_KEY)) || {}; } catch (e) { return {}; }
    }
    function saveSettings(s) {
        try { localStorage.setItem(LS_KEY, JSON.stringify(s)); } catch (e) { /* ignore */ }
    }

    // =====================================================================
    //  PARSING THE PASTED BLOCK
    // =====================================================================

    /**
     * Split a pasted line into fields. Tabs win (spreadsheet paste); otherwise
     * fall back to commas / semicolons / pipes / runs of 2+ spaces.
     */
    function splitFields(line) {
        if (line.includes('\t')) return line.split('\t').map((f) => f.trim());
        if (/[;|]/.test(line)) return line.split(/\s*[;|]\s*/).map((f) => f.trim());
        if (line.includes(',')) return line.split(/\s*,\s*/).map((f) => f.trim());

        const wide = line.trim().split(/\s{2,}/).map((f) => f.trim());
        if (wide.length >= 3) return wide;

        // Single-space separated (hand-typed). Split on every space, then glue
        // an AM/PM back onto the time it belongs to — "7:00 AM" is one field.
        const merged = [];
        for (const token of line.trim().split(/\s+/)) {
            const prev = merged[merged.length - 1];
            if (prev && /^[ap]\.?m\.?$/i.test(token) && /^\d{1,2}([:.]\d{2})?$/.test(prev)) {
                merged[merged.length - 1] = `${prev} ${token}`;
            } else {
                merged.push(token);
            }
        }
        return merged;
    }

    /**
     * Parse a time string into {h, m} (24h). Accepts:
     *   7:00 AM · 7:00am · 07:00 · 19:00 · 7 AM · 7pm · 1900 · 7.00 PM
     * Returns null when the field is blank / a placeholder / unparseable.
     */
    /** A raw spreadsheet number: military integer (1058) or an Excel serial. */
    function numberToTime(v) {
        if (!isFinite(v)) return null;
        if (Number.isInteger(v) && v >= 100 && v <= 2359 && v % 100 < 60 && Math.floor(v / 100) < 24) {
            return { h: Math.floor(v / 100), m: v % 100, plus: 0 };     // 1058 -> 10:58
        }
        const mins = Math.round((v - Math.floor(v)) * 1440);            // fraction of a day
        return { h: Math.floor(mins / 60) % 24, m: mins % 60, plus: 0 };
    }

    /**
     * Parse a time into { h, m, plus }, where `plus` counts days past the row's
     * own date — a night shift's 06:00 Time Out is { h:6, m:0, plus:1 }.
     * Accepts 7:00 AM · 07:00 · 19:00 · 1900 · 1053H · 22;20 (typo) · 06:00+1,
     * plus Date objects and raw Excel numbers straight from a cell.
     */
    function parseTime(raw) {
        if (raw instanceof Date && !isNaN(raw)) return { h: raw.getHours(), m: raw.getMinutes(), plus: 0 };
        if (typeof raw === 'number') return numberToTime(raw);

        let s = norm(raw).toLowerCase();
        if (!s || isBlankText(s)) return null;

        // "+1" suffix marks a punch belonging to the next day (overnight shift).
        let plus = 0;
        const plusMatch = s.match(/\+\s*(\d)\s*$/);
        if (plusMatch) { plus = +plusMatch[1]; s = s.slice(0, plusMatch.index).trim(); }

        if (/^\d*\.\d+$/.test(s)) {
            const t = numberToTime(parseFloat(s));
            return t ? { h: t.h, m: t.m, plus } : null;
        }

        s = s.replace(/^\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}[t ]+/, '');  // drop a leading date
        s = s.replace(/;/g, ':');                                        // "22;20" -> "22:20"
        s = s.replace(/h$/i, '').trim();                                 // military "1053H"

        const ampmMatch = s.match(/([ap])\.?m\.?/);
        const ampm = ampmMatch ? ampmMatch[1] : null;
        s = s.replace(/([ap])\.?m\.?/, '').trim();

        let h = null, m = 0, mt;
        if ((mt = s.match(/^(\d{1,2})\s*[:.h]\s*(\d{2})(?:\s*[:.]\s*\d{2})?$/))) { h = +mt[1]; m = +mt[2]; }
        else if ((mt = s.match(/^(\d{1,2})$/))) { h = +mt[1]; m = 0; }
        else if ((mt = s.match(/^(\d{3,4})$/))) { const n = mt[1]; h = +n.slice(0, n.length - 2); m = +n.slice(-2); }
        else return null;

        if (ampm === 'p' && h < 12) h += 12;
        if (ampm === 'a' && h === 12) h = 0;
        if (!(h >= 0 && h <= 23) || !(m >= 0 && m <= 59)) return null;
        return { h, m, plus };
    }

    /** ISO date shifted by n whole days. */
    function addDays(iso, n) {
        if (!n) return iso;
        const [Y, M, D] = iso.split('-').map(Number);
        const d = new Date(Date.UTC(Y, M - 1, D + n));
        return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
    }

    const MONTHS = {
        jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
        jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12
    };

    /**
     * Parse a date field into an ISO "YYYY-MM-DD" string.
     * `fallbackYear` fills in when the source omits the year ("Aug 1").
     * Ambiguous n/n/nnnn is read as MM/DD/YYYY unless the first number > 12.
     */
    function parseDate(raw, fallbackYear) {
        if (raw instanceof Date && !isNaN(raw)) {
            return `${raw.getFullYear()}-${pad2(raw.getMonth() + 1)}-${pad2(raw.getDate())}`;
        }
        const s = norm(raw);
        if (!s) return null;

        // Excel serial date (days since 1899-12-30), e.g. 46235 → 2026-08-01.
        if (/^\d{5}(\.\d+)?$/.test(s)) {
            const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(+s) * 86400000);
            return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
        }

        let mt = s.match(/(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);        // 2026-08-01
        if (mt) return `${mt[1]}-${pad2(mt[2])}-${pad2(mt[3])}`;

        mt = s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);          // 08/01/2026
        if (mt) {
            let a = +mt[1], b = +mt[2];
            let y = mt[3].length === 2 ? 2000 + +mt[3] : +mt[3];
            let month = a, day = b;
            if (a > 12 && b <= 12) { month = b; day = a; }                  // clearly D/M
            return `${y}-${pad2(month)}-${pad2(day)}`;
        }

        mt = s.match(/([a-z]{3,9})\.?\s+(\d{1,2})(?:\s*,?\s*(\d{4}))?/i);   // Aug 1, 2026
        if (mt && MONTHS[mt[1].slice(0, 3).toLowerCase()]) {
            const month = MONTHS[mt[1].slice(0, 3).toLowerCase()];
            const y = mt[3] ? +mt[3] : fallbackYear;
            if (!y) return null;
            return `${y}-${pad2(month)}-${pad2(+mt[2])}`;
        }

        mt = s.match(/(\d{1,2})[\s\-]+([a-z]{3,9})\.?[\s\-]*(\d{2,4})?/i);  // 1-Aug-26
        if (mt && MONTHS[mt[2].slice(0, 3).toLowerCase()]) {
            const month = MONTHS[mt[2].slice(0, 3).toLowerCase()];
            let y = mt[3] ? (mt[3].length === 2 ? 2000 + +mt[3] : +mt[3]) : fallbackYear;
            if (!y) return null;
            return `${y}-${pad2(month)}-${pad2(+mt[1])}`;
        }

        mt = s.match(/^(\d{1,2})[\/\-.](\d{1,2})$/);                        // 08/01, no year
        if (mt && fallbackYear) {
            let a = +mt[1], b = +mt[2];
            let month = a, day = b;
            if (a > 12 && b <= 12) { month = b; day = a; }
            return `${fallbackYear}-${pad2(month)}-${pad2(day)}`;
        }
        return null;
    }

    /**
     * Drop leading label columns that sit between the date and the first time
     * — day-of-week ("TUESDAY"), employee name, "No Schedule", etc.
     * Stops at the first blank cell, because a blank is a real (skipped) time
     * column and dropping it would shift every column after it.
     */
    function stripLabelColumns(rest, dropped) {
        while (rest.length) {
            const field = rest[0];
            if (norm(field) === '' || isBlankText(field)) break;   // genuine empty time cell
            if (parseTime(field) !== null) break;                  // reached the times
            dropped.push(norm(field));
            rest.shift();
        }
        return rest;
    }

    /**
     * Read the pasted block into entries:
     *   [{ date, times: {col:{h,m}}, blanks: [col], raw }]
     * `blanks` are columns you deliberately left empty or dashed.
     * Detects an optional header line and remaps columns by name when present.
     */
    function parsePaste(text, fallbackYear) {
        const out = { entries: [], warnings: [], mapping: null };
        const lines = String(text || '').split(/\r?\n/).filter((l) => norm(l) !== '');
        if (!lines.length) return out;

        let columnOrder = null;     // null = no header, map positionally after stripping labels
        let startIndex = 0;
        let dateIndex = 0;          // which field holds the date

        // Header line? (first line has no parseable date but does name columns)
        const firstFields = splitFields(lines[0]);
        const headerHits = firstFields.map((f) => HEADER_ALIASES[normKey(f)]).filter(Boolean);
        if (!parseDate(firstFields[0], fallbackYear) && headerHits.length >= 2) {
            // Map by absolute position, so the date needn't be the first column —
            // exports that lead with Guard / Access ID columns still line up.
            columnOrder = firstFields.map((f) => HEADER_ALIASES[normKey(f)] || null);
            const named = firstFields.findIndex((f) => DATE_ALIASES.has(normKey(f)));
            dateIndex = named >= 0 ? named : 0;
            startIndex = 1;
            out.mapping = columnOrder.map((c, i) => (c ? COLUMN_LABELS[c] : (i === dateIndex ? 'Date' : '(ignored)'))).join(' · ');
        }

        for (let i = startIndex; i < lines.length; i++) {
            const line = lines[i];
            const fields = splitFields(line);
            const date = parseDate(fields[dateIndex], fallbackYear);
            if (!date) {
                out.warnings.push(`Line ${i + 1}: no date found — skipped ("${norm(line).slice(0, 48)}")`);
                continue;
            }

            let rest = columnOrder ? fields : fields.slice(1);
            let order = columnOrder;
            const dropped = [];

            if (!order) {
                rest = stripLabelColumns(rest.slice(), dropped);
                order = COLUMNS.slice();
                // Only two values and both are times → Time In / Time Out.
                if (rest.length === 2 && parseTime(rest[0]) && parseTime(rest[1])) {
                    order = ['time_in', 'time_out'];
                }
            }
            if (dropped.length) {
                const labels = dropped.filter((d) => !DAY_NAMES.has(normKey(d)));
                if (labels.length) {
                    out.warnings.push(`Line ${i + 1}: ignored extra column(s) before the times — ${labels.join(', ')}`);
                }
            }

            const times = {};
            const blanks = [];
            rest.forEach((field, idx) => {
                const col = order[idx];
                if (!col) return;                                   // column mapped to nothing
                if (norm(field) === '' || isBlankText(field)) { blanks.push(col); return; }
                const t = parseTime(field);
                if (!t) {
                    out.warnings.push(`Line ${i + 1}: could not read "${norm(field)}" as a time — ${COLUMN_LABELS[col]} skipped`);
                    return;
                }
                // A cell carrying its own full date ("2026/08/02 06:00", as an
                // upstream exporter writes an overnight punch) states which day
                // it belongs to — honour that instead of assuming the row's.
                const stamped = String(field).trim();
                if (/^\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}[T\s]/.test(stamped)) {
                    const own = parseDate(stamped.split(/[T\s]+/)[0], fallbackYear);
                    if (own && own !== date) {
                        const days = Math.round((Date.parse(own) - Date.parse(date)) / 86400000);
                        if (days > 0) t.plus = days;
                    }
                }
                times[col] = t;
            });

            if (rest.length > order.length) {
                out.warnings.push(`Line ${i + 1}: ${rest.length} time columns but the table has ${order.length} — extras ignored`);
            }

            if (!Object.keys(times).length) {
                out.warnings.push(`Line ${i + 1}: date ${date} had no usable times — skipped`);
                continue;
            }
            out.entries.push({ date, times, blanks, raw: line });
        }
        return out;
    }

    // =====================================================================
    //  READING THE PAGE
    // =====================================================================

    const DATE_CELL_RE = /^\s*(\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}|\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{4})\s*$/;

    /** Every table row on the page whose first cells contain a bare date. */
    function indexRows() {
        const map = new Map();
        const rows = document.querySelectorAll('tr');
        for (const tr of rows) {
            const cells = tr.children;
            for (let i = 0; i < Math.min(cells.length, 3); i++) {
                const txt = norm(cells[i].textContent);
                if (!DATE_CELL_RE.test(txt)) continue;
                const iso = parseDate(txt, null);
                if (iso && !map.has(iso)) map.set(iso, tr);
                break;
            }
        }
        return map;
    }

    /** The year most represented in the table — used when the paste omits one. */
    function pageYear() {
        const years = {};
        for (const iso of indexRows().keys()) {
            const y = iso.slice(0, 4);
            years[y] = (years[y] || 0) + 1;
        }
        const best = Object.entries(years).sort((a, b) => b[1] - a[1])[0];
        return best ? +best[0] : new Date().getFullYear();
    }

    /** Header labels of the table this row belongs to, by cell index. */
    function headerFor(tr) {
        const table = tr.closest('table');
        if (!table) return [];
        let headRow = table.querySelector('thead tr');
        if (!headRow) {
            for (const r of table.querySelectorAll('tr')) {
                if (r.querySelector('th')) { headRow = r; break; }
            }
        }
        if (!headRow) headRow = table.querySelector('tr');
        if (!headRow) return [];
        return Array.from(headRow.children).map((c) => normKey(c.textContent));
    }

    /** Find a clickable control in the row whose text is exactly `label`. */
    function findControl(tr, label) {
        const wanted = label.toLowerCase();
        const candidates = tr.querySelectorAll('a, button, input[type="button"], input[type="submit"], span[onclick], [role="button"]');
        for (const el of candidates) {
            const txt = norm(el.value || el.textContent).toLowerCase();
            if (txt === wanted) return el;
        }
        return null;
    }

    /** Text inputs inside the row that are visible and look editable. */
    function rowInputs(tr) {
        return Array.from(tr.querySelectorAll('input, textarea')).filter((el) => {
            if (el.type === 'hidden' || el.type === 'checkbox' || el.type === 'radio') return false;
            if (el.disabled || el.readOnly && el.type !== 'text') return false;
            return el.offsetParent !== null || el.getClientRects().length > 0;
        });
    }

    /** Is this row currently showing display text (not edit inputs)? */
    function isEditing(tr) {
        return rowInputs(tr).length > 0;
    }

    /**
     * Snapshot which time columns already hold real values, read from the
     * DISPLAY cells (before edit mode replaces them with prefilled inputs).
     */
    function displayState(tr) {
        const header = headerFor(tr);
        const offset = headerOffset(tr, header);     // same correction the input mapping uses
        const state = {};
        Array.from(tr.children).forEach((cell, idx) => {
            const col = HEADER_ALIASES[header[idx - offset]];
            if (!col) return;
            state[col] = !isBlankText(cell.textContent);
        });
        return state;
    }

    // Attributes worth interrogating when an input names its own column.
    const ATTR_HINTS = ['name', 'id', 'data-column', 'data-field', 'data-name', 'aria-label', 'title', 'placeholder'];

    /** Does this input name its own column? ("txtTimeIn", "log[lunch_out]", …) */
    function columnFromAttributes(input) {
        for (const attr of ATTR_HINTS) {
            const raw = input.getAttribute(attr);
            if (!raw) continue;
            const spaced = normKey(String(raw)
                .replace(/([a-z0-9])([A-Z])/g, '$1 $2')     // camelCase → camel Case
                .replace(/[_\-.\[\]]+/g, ' '));
            if (HEADER_ALIASES[spaced]) return HEADER_ALIASES[spaced];
            // Substring match, but only on aliases long enough to be unambiguous
            // — never on bare "in"/"out", which appear inside countless names.
            for (const alias of Object.keys(HEADER_ALIASES)) {
                if (alias.length >= 6 && spaced.includes(alias)) return HEADER_ALIASES[alias];
            }
        }
        return null;
    }

    /**
     * How far the row's cells are shifted relative to the header's cells.
     * Some grids give body rows an extra leading cell (a row id, a spacer) that
     * the header does not have, which silently shifts every column lookup.
     * Anchor on the date column, which both rows agree on.
     */
    function headerOffset(tr, header) {
        const hDate = header.indexOf('date');
        if (hDate < 0) return 0;
        let rDate = -1;
        Array.from(tr.children).forEach((cell, i) => {
            if (rDate < 0 && DATE_CELL_RE.test(norm(cell.textContent))) rDate = i;
        });
        return rDate < 0 ? 0 : rDate - hDate;
    }

    /** True when every input got a distinct column, in left-to-right order. */
    function mappingIsSound(pairs, inputCount) {
        if (pairs.length !== inputCount) return false;               // some input unaccounted for
        const cols = pairs.map((p) => p.col);
        if (new Set(cols).size !== cols.length) return false;        // a column claimed twice
        const order = cols.map((c) => COLUMNS.indexOf(c));
        if (order.some((i) => i < 0)) return false;
        return order.every((v, i) => i === 0 || v > order[i - 1]);   // must run left to right
    }

    /**
     * Map the row's edit inputs to canonical column keys.
     * Three strategies, strongest first — each validated before it is accepted,
     * so a partial match can never shift the columns.
     */
    function mapRowInputs(tr) {
        const inputs = rowInputs(tr);
        const build = (pairs) => {
            const map = {};
            pairs.forEach((p) => { map[p.col] = p.input; });
            return map;
        };

        // 1. The inputs name their own columns.
        const byAttr = [];
        for (const input of inputs) {
            const col = columnFromAttributes(input);
            if (col) byAttr.push({ col, input });
        }
        if (mappingIsSound(byAttr, inputs.length)) {
            return { map: build(byAttr), strategy: 'input attributes' };
        }

        // 2. Header labels, corrected for any header/body cell offset.
        const header = headerFor(tr);
        const offset = headerOffset(tr, header);
        const byHeader = [];
        for (const input of inputs) {
            const cell = input.closest('td, th');
            if (!cell || cell.parentElement !== tr) continue;
            const idx = Array.prototype.indexOf.call(tr.children, cell);
            const col = HEADER_ALIASES[header[idx - offset]];
            if (col) byHeader.push({ col, input });
        }
        if (mappingIsSound(byHeader, inputs.length)) {
            return {
                map: build(byHeader),
                strategy: `header labels${offset ? ` (offset ${offset > 0 ? '+' : ''}${offset})` : ''}`
            };
        }

        // 3. Positional: the row's time inputs sit in the table's column order.
        //    Reliable whenever the row exposes exactly one input per time column.
        if (inputs.length && inputs.length <= COLUMNS.length) {
            return {
                map: build(inputs.map((input, i) => ({ col: COLUMNS[i], input }))),
                strategy: `position (${inputs.length} input${inputs.length === 1 ? '' : 's'})`
            };
        }
        if (inputs.length > COLUMNS.length) {
            // More inputs than time columns — keep whatever the header did agree on.
            return { map: build(byHeader), strategy: 'header labels (partial)' };
        }
        return { map: {}, strategy: 'none' };
    }

    // =====================================================================
    //  WRITING VALUES
    // =====================================================================

    /** Rebuild the date half of a value using the template's own convention. */
    function formatDateLike(template, iso) {
        const [Y, M, D] = iso.split('-');
        const nums = template.split(/[^0-9]+/).filter(Boolean);
        const sepMatch = template.match(/[^0-9]/);
        const sep = sepMatch ? sepMatch[0] : '/';
        if (nums.length < 3) return `${M}/${D}/${Y}`;

        let order;
        if (nums[0].length === 4) order = ['Y', 'M', 'D'];
        else if (+nums[0] > 12) order = ['D', 'M', 'Y'];
        else order = ['M', 'D', 'Y'];

        const vals = { Y, M, D };
        return order.map((key, i) => {
            let v = vals[key];
            if (key === 'Y') return nums[i].length === 2 ? Y.slice(2) : Y;
            return nums[i].length === 1 ? String(+v) : v;      // preserve un-padded style
        }).join(sep);
    }

    /** Rebuild the time half of a value using the template's own convention. */
    function formatTimeLike(template, t) {
        const is12h = /[ap]\.?m\.?/i.test(template);
        const nums = template.split(/[^0-9]+/).filter(Boolean);
        const hourPadded = !nums[0] || nums[0].length === 2;
        const wantSeconds = nums.length >= 3;

        let h = t.h;
        let suffix = '';
        if (is12h) {
            const pm = h >= 12;
            h = h % 12 || 12;
            const mt = template.match(/([ap])(\.?)m(\.?)/i);
            const dots = mt ? [mt[2], mt[3]] : ['', ''];
            const upper = mt ? /[AP]/.test(mt[1]) : true;
            suffix = `${pm ? 'P' : 'A'}${dots[0]}M${dots[1]}`;
            if (!upper) suffix = suffix.toLowerCase();
        }

        let s = `${hourPadded ? pad2(h) : String(h)}:${pad2(t.m)}`;
        if (wantSeconds) s += ':00';
        if (suffix) s += (/\s[ap]\.?m/i.test(template) ? ' ' : '') + suffix;
        return s;
    }

    /** Produce the string to type into `input` for this date + time. */
    function buildValue(input, iso, t) {
        const type = (input.getAttribute('type') || 'text').toLowerCase();
        if (type === 'datetime-local') return `${iso}T${pad2(t.h)}:${pad2(t.m)}`;
        if (type === 'time') return `${pad2(t.h)}:${pad2(t.m)}`;
        if (type === 'date') return iso;

        const template = norm(input.value) || norm(input.placeholder) || norm(input.defaultValue);
        const parts = template.match(/^(\S+)([\sT]+)(.+)$/);
        if (parts) {
            return formatDateLike(parts[1], iso) + parts[2] + formatTimeLike(parts[3], t);
        }
        if (template && /^[0-9:]+\s*[ap]?\.?m?\.?$/i.test(template)) {
            return formatTimeLike(template, t);          // time-only field
        }
        // Nothing to imitate — use the format the page showed us in the screenshots.
        const [Y, M, D] = iso.split('-');
        const h12 = t.h % 12 || 12;
        return `${M}/${D}/${Y} ${pad2(h12)}:${pad2(t.m)} ${t.h >= 12 ? 'PM' : 'AM'}`;
    }

    /**
     * Set a value the way a human would, so frameworks notice.
     * Bypasses React/Vue value-setter shadowing, fires the full key/input/change
     * sequence, and nudges jQuery + common datetimepicker plugins.
     */
    function setValue(input, value) {
        input.focus();

        const proto = Object.getPrototypeOf(input);
        const desc = Object.getOwnPropertyDescriptor(proto, 'value');
        const nativeSetter = Object.getOwnPropertyDescriptor(
            input instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype,
            'value'
        ).set;
        if (desc && desc.set && desc.set !== nativeSetter) nativeSetter.call(input, value);
        else input.value = value;

        for (const type of ['keydown', 'keypress', 'keyup']) {
            input.dispatchEvent(new KeyboardEvent(type, { bubbles: true, key: '0' }));
        }
        input.dispatchEvent(new InputEvent('input', { bubbles: true, data: value, inputType: 'insertText' }));
        input.dispatchEvent(new Event('change', { bubbles: true }));

        const $ = window.jQuery || window.$;
        if ($ && $.fn) {
            try {
                const el = $(input);
                el.val(value);
                // Bootstrap / XDSoft / jQuery-UI datetimepickers, when present.
                const bs = el.data && el.data('DateTimePicker');
                if (bs && typeof bs.date === 'function') { try { bs.date(value); } catch (e) { /* ignore */ } }
                if (typeof el.datetimepicker === 'function') { try { el.datetimepicker('setDate', value); } catch (e) { /* ignore */ } }
                el.trigger('input').trigger('change').trigger('blur');
            } catch (e) { /* ignore */ }
        }

        input.dispatchEvent(new Event('blur', { bubbles: true }));
        input.blur();
    }

    /** Click Edit (if needed) and wait for the row's inputs to appear. */
    async function ensureEditMode(iso, timeoutMs = 5000) {
        let tr = indexRows().get(iso);
        if (!tr) return null;
        if (isEditing(tr)) return tr;

        const edit = findControl(tr, 'Edit');
        if (!edit) return isEditing(tr) ? tr : null;
        edit.click();

        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
            await sleep(80);
            tr = indexRows().get(iso);           // the row may have been re-rendered
            if (tr && isEditing(tr)) return tr;
        }
        return null;
    }

    // =====================================================================
    //  THE FILL RUN
    // =====================================================================

    let undoStack = [];      // [{input, previous}]

    async function runFill(entries, opts, log) {
        undoStack = [];
        let filledRows = 0, filledFields = 0, missingRows = 0, skipped = 0, blanksLeft = 0;
        let overnightFields = 0;
        let reportedStrategy = false;

        for (const entry of entries) {
            const existing = indexRows().get(entry.date);
            if (!existing) {
                log(`✗ ${entry.date} — no row on this page (check your date range)`, 'err');
                missingRows++;
                continue;
            }

            // Decide "is this row blank" from the display cells, before edit mode
            // replaces them with inputs prefilled to 12:00 AM.
            let alreadyFilled = {};
            if (!isEditing(existing)) alreadyFilled = displayState(existing);

            const tr = await ensureEditMode(entry.date);
            if (!tr) {
                log(`✗ ${entry.date} — could not open Edit (no inputs appeared)`, 'err');
                missingRows++;
                continue;
            }

            const mapped = mapRowInputs(tr);
            const inputs = mapped.map;
            if (!Object.keys(inputs).length) {
                log(`✗ ${entry.date} — row is in edit mode but no inputs could be mapped`, 'err');
                missingRows++;
                continue;
            }
            if (!reportedStrategy) {
                reportedStrategy = true;
                log(`Columns matched by ${mapped.strategy}: ${Object.keys(inputs).map((c) => COLUMN_LABELS[c]).join(', ')}`, 'info');
            }

            let touched = 0;
            const parts = [];
            const leftAlone = [];

            // Columns you deliberately left blank / dashed in the paste.
            for (const col of (entry.blanks || [])) {
                const input = inputs[col];
                if (!input) continue;
                if (opts.clearBlanks) {
                    undoStack.push({ input, previous: input.value });
                    setValue(input, '');
                    input.classList.add('zdf-touched');
                    touched++;
                    filledFields++;          // a cleared field was still written to
                    parts.push(`${COLUMN_LABELS[col]}=cleared`);
                } else {
                    leftAlone.push(COLUMN_LABELS[col]);
                }
            }

            for (const col of COLUMNS) {
                const t = entry.times[col];
                if (!t) continue;
                const input = inputs[col];
                if (!input) {
                    log(`  · ${entry.date} — no "${COLUMN_LABELS[col]}" field in this row`, 'warn');
                    continue;
                }
                if (opts.onlyBlank && alreadyFilled[col]) {
                    skipped++;
                    parts.push(`${COLUMN_LABELS[col]}=kept`);
                    continue;
                }
                // An overnight punch belongs to a later date than the row's own.
                const effectiveDate = addDays(entry.date, t.plus || 0);
                const value = buildValue(input, effectiveDate, t);
                undoStack.push({ input, previous: input.value });
                setValue(input, value);
                input.classList.add('zdf-touched');
                touched++;
                filledFields++;
                if (t.plus) overnightFields++;
                parts.push(`${COLUMN_LABELS[col]}=${formatTimeLike('12:00 AM', t)}${t.plus ? ` (+${t.plus}d → ${effectiveDate})` : ''}`);
            }

            if (touched) {
                filledRows++;
                tr.classList.add('zdf-row-touched');
                log(`✓ ${entry.date} — ${touched} field${touched === 1 ? '' : 's'}: ${parts.join(', ')}`, 'ok');
                if (filledRows === 1) tr.scrollIntoView({ block: 'center', behavior: 'smooth' });
            } else {
                log(`· ${entry.date} — nothing to write`, 'warn');
            }
            if (leftAlone.length) {
                log(`  ! ${entry.date} — ${leftAlone.join(', ')} left untouched, still showing the prefilled 12:00 AM`, 'warn');
                blanksLeft += leftAlone.length;
            }

            await sleep(opts.delay);
        }

        return { filledRows, filledFields, missingRows, skipped, blanksLeft, overnightFields };
    }

    function undoFill(log) {
        if (!undoStack.length) { log('Nothing to undo.', 'warn'); return; }
        let n = 0;
        for (const { input, previous } of undoStack.slice().reverse()) {
            if (!input.isConnected) continue;
            setValue(input, previous);
            input.classList.remove('zdf-touched');
            n++;
        }
        document.querySelectorAll('.zdf-row-touched').forEach((r) => r.classList.remove('zdf-row-touched'));
        undoStack = [];
        log(`Undid ${n} field${n === 1 ? '' : 's'}.`, 'ok');
    }

    // =====================================================================
    //  DIAGNOSTICS
    // =====================================================================

    function diagnostics() {
        const rows = indexRows();
        const dates = Array.from(rows.keys()).sort();
        const sample = rows.get(dates[0]);
        const lines = [];
        lines.push(`URL: ${location.origin}${location.pathname}`);
        lines.push(`Rows with a date: ${rows.size}`);
        lines.push(`Range: ${dates[0] || '—'} .. ${dates[dates.length - 1] || '—'}`);
        if (sample) {
            const header = headerFor(sample);
            lines.push(`Header cells: ${JSON.stringify(header)}`);
            lines.push(`Header/body cell offset: ${headerOffset(sample, header)}`);
            lines.push(`Row cell count: ${sample.children.length} vs header ${header.length}`);
            lines.push(`Edit control found: ${!!findControl(sample, 'Edit')}`);
            lines.push(`Row in edit mode: ${isEditing(sample)}`);
            const mapped = mapRowInputs(sample);
            lines.push(`Mapping strategy: ${mapped.strategy}`);
            lines.push(`Mapped columns: ${JSON.stringify(Object.keys(mapped.map))}`);
            const inputs = rowInputs(sample);
            lines.push(`Visible inputs: ${inputs.length}`);
            lines.push(`Input name/type/value: ${JSON.stringify(inputs.map((i) => [i.name || i.id || '', i.type, i.value]))}`);
            lines.push('');
            lines.push('--- SAMPLE ROW HTML ---');
            lines.push(sample.outerHTML.slice(0, 4000));
        }
        lines.push('');
        lines.push(`jQuery present: ${!!(window.jQuery || window.$)}`);
        return lines.join('\n');
    }

    // =====================================================================
    //  EMPLOYEE ROSTER — one workbook holding every guard
    // =====================================================================
    //  Editing in Zenoras is per employee, so a workbook covering everyone is
    //  filtered down to whoever this page belongs to. The page is matched by
    //  the access ID and name it displays, never by position in the file.

    const ROSTER_KEY = 'zdf.roster.v1';
    const LINKS_KEY = 'zdf.pagelinks.v1';     // URL path → employee key, once confirmed
    const DONE_KEY = 'zdf.done.v1';

    let ocrActive = false;    // OCR output sits in the box for THIS page's guard
    let roster = null;        // { employees: [{key, id, name, rows:[{date,times,blanks}]}], source }

    function readStore(key, fallback) {
        try { return JSON.parse(localStorage.getItem(key)) || fallback; } catch (e) { return fallback; }
    }
    function writeStore(key, value) {
        try { localStorage.setItem(key, JSON.stringify(value)); return true; }
        catch (e) { return false; }          // quota — keep working from memory
    }

    // ── Reading messy client DTRs ────────────────────────────────────────
    //  Client spreadsheets arrive in whatever shape the site made them. Rather
    //  than demand one layout, sniff which of several known shapes a sheet is
    //  and parse accordingly. Cells are read RAW (numbers, Dates) so military
    //  integers and Excel serials survive.

    const MONTH_NAMES = {
        jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
        jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
    };

    /** Raw value of one cell, or null. */
    function cellAt(ws, r, c) {
        if (c == null || c < 0) return null;
        const x = ws[XLSX.utils.encode_cell({ r, c })];
        return x ? x.v : null;
    }
    const sheetRange = (ws) => XLSX.utils.decode_range(ws['!ref']);
    const cellText = (ws, r, c) => norm(cellAt(ws, r, c));
    const upper = (s) => String(s == null ? '' : s).toUpperCase().replace(/\s+/g, ' ').trim();

    const ROLL_HOURS = 6;    // a backward jump this large means the clock crossed midnight

    /**
     * Turn six raw punches into { times, flags }, rolling each punch onto the
     * next day when it falls before the previous one — this is what makes a
     * night shift's Time Out land on the correct date.
     */
    function sequencePunches(raw) {
        const times = {};
        let prevAbs = null, overnight = false, outOfOrder = false, present = 0;

        for (const col of COLUMNS) {
            const t = raw[col];
            if (!t) continue;
            present++;
            let abs = t.h * 60 + t.m;
            let plus = 0;
            if (prevAbs !== null && abs < prevAbs) {
                if (prevAbs - abs >= ROLL_HOURS * 60) {
                    while (abs < prevAbs) { abs += 1440; plus++; }
                    overnight = true;
                } else {
                    outOfOrder = true;
                }
            }
            times[col] = { h: t.h, m: t.m, plus };
            if (prevAbs === null || abs > prevAbs) prevAbs = abs;
        }

        const flags = [];
        if (overnight) flags.push('overnight');
        if (outOfOrder) flags.push('out-of-order punch');
        if (present > 0 && (!times.time_in || !times.time_out)) flags.push('missing in/out');
        if (present % 2 !== 0) flags.push('unpaired punch');
        return { times, flags, present };
    }

    /** Hours as written in the sheet's own total column, for cross-checking. */
    function statedHours(v) {
        if (v == null) return null;
        if (v instanceof Date) return v.getHours() + v.getMinutes() / 60;   // duration on the 1899 epoch
        if (typeof v === 'number') { if (v > 0 && v <= 1) return v * 24; if (v > 1 && v < 48) return v; return null; }
        const m = String(v).match(/(\d+(?:\.\d+)?)/);                        // "13HRS" -> 13
        return m ? parseFloat(m[1]) : null;
    }

    /** Compare the punches we read against the hours the sheet claims. */
    function reconcile(row) {
        if (row.rest) return row;
        const mins = (t) => t ? (t.h * 60 + t.m + t.plus * 1440) : null;
        const tin = mins(row.times.time_in), tout = mins(row.times.time_out);
        if (tin != null && tout != null) {
            let gross = (tout - tin) / 60;
            let brk = 0;
            const lo = mins(row.times.lunch_out), li = mins(row.times.lunch_in);
            const bo = mins(row.times.break_out), bi = mins(row.times.break_in);
            if (lo != null && li != null) brk += (li - lo) / 60;
            if (bo != null && bi != null) brk += (bi - bo) / 60;
            row.computedHours = Math.round(gross * 10) / 10;
            const sh = row.statedHours;
            if (sh != null && sh > 0 && (sh < gross - brk - 1.5 || sh > gross + 1.5)) {
                row.flags.push(`hours off: sheet says ${sh}, punches give ~${row.computedHours}`);
            }
            if ((gross < 1 || gross > 16) && !row.flags.some((f) => /hours off/.test(f))) {
                row.flags.push(`check span ~${row.computedHours}h`);
            }
        }
        return row;
    }

    const REST_RE = /REST|NO ?DUTY|DAY ?OFF|\bRD\b|OFF\b|LEAVE|ABSENT|VACATION|\bVL\b|\bSL\b/i;

    /** One canonical day row. */
    function makeRow(id, name, date, rawTimes, sheetTotal, extraFlags) {
        const seq = sequencePunches(rawTimes);
        const row = {
            id: String(id || '').replace(/^'/, '').trim(),
            name: norm(name),
            date,
            times: seq.times,
            blanks: COLUMNS.filter((c) => !seq.times[c]),
            flags: (extraFlags || []).concat(seq.flags),
            statedHours: statedHours(sheetTotal),
            rest: false
        };
        return reconcile(row);
    }
    function makeRestRow(id, name, date, reason) {
        return { id: String(id || '').trim(), name: norm(name), date, times: {}, blanks: [], flags: [reason], rest: true };
    }

    // ── Layout A: flat report — "Personnel Name" + "Time In 1..3" ────────
    function parseFlatReport(ws) {
        const rng = sheetRange(ws);
        const nz = (s) => upper(s).replace(/0UT/g, 'OUT');       // OCR'd "0UT" -> "OUT"
        const NAMEH = ['PERSONNEL NAME', 'EMPLOYEE NAME', 'NAME OF EMPLOYEE'];

        let hr = -1;
        for (let r = rng.s.r; r <= rng.e.r && hr < 0; r++) {
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                if (NAMEH.includes(nz(cellAt(ws, r, c)))) { hr = r; break; }
            }
        }
        if (hr < 0) return [];

        const col = {};
        for (let c = rng.s.c; c <= rng.e.c; c++) {
            const k = nz(cellAt(ws, hr, c));
            if (k && col[k] == null) col[k] = c;
        }
        const pick = (names) => { for (const n of names) if (col[n] != null) return col[n]; return null; };
        const cName = pick(NAMEH);
        const cDate = pick(['TRANSACTION DATE', 'DATE', 'LOG DATE']);
        const cId = pick(['ACCESS ID', 'ACCESS ID #', 'EMPLOYEE ID', 'EMP ID']);
        const cHours = pick(['NET HOURS RENDERED', 'GROSS HOURS RENDERED', 'NO. OF HOURS RENDERED',
            'NO. OF HOURS', 'TOTAL HOURS', 'HRS RENDERED']);
        const map = {
            time_in: pick(['TIME IN 1', 'FIRST TIME IN', 'TIME IN']),
            lunch_out: pick(['TIME OUT 1', 'BREAK OUT 1', 'LUNCH OUT']),
            lunch_in: pick(['TIME IN 2', 'BREAK IN 1', 'LUNCH IN']),
            break_out: pick(['TIME OUT 2', 'BREAK OUT 2', 'BREAK OUT']),
            break_in: pick(['TIME IN 3', 'BREAK IN 2', 'BREAK IN']),
            time_out: pick(['TIME OUT 3', 'LAST TIME OUT', 'TIME OUT'])
        };
        if (cDate == null) return [];

        const out = [];
        for (let r = hr + 1; r <= rng.e.r; r++) {
            const name = cellAt(ws, r, cName);
            const date = parseDate(cellAt(ws, r, cDate), null);
            if (!name || !date) continue;
            const raw = {};
            for (const c of COLUMNS) raw[c] = parseTime(cellAt(ws, r, map[c]));
            const id = cellAt(ws, r, cId);
            if (!COLUMNS.some((c) => raw[c])) {
                const marker = upper(cellAt(ws, r, map.time_in));
                out.push(makeRestRow(id, name, date, REST_RE.test(marker) ? 'rest day' : (marker.toLowerCase() || 'no punches')));
                continue;
            }
            out.push(makeRow(id, name, date, raw, cellAt(ws, r, cHours), []));
        }
        return out;
    }

    // ── Layout B: stacked blocks — "SECURITY GUARD: ..." above each table ─
    const STACK_HEADERS = {
        'DATE': 'date', 'SCHEDULE': 'schedule', 'SCHED': 'schedule',
        'TIME IN': 'time_in', 'INITIAL IN': 'time_in',
        'LUNCH OUT': 'lunch_out', 'LUCH OUT': 'lunch_out', 'LB OUT': 'lunch_out', 'L.B OUT': 'lunch_out',
        'LUNCH IN': 'lunch_in', 'LB IN': 'lunch_in', 'L.B IN': 'lunch_in',
        'BREAK OUT': 'break_out', 'CB OUT': 'break_out', 'C.B OUT': 'break_out',
        'BREAK IN': 'break_in', 'CB IN': 'break_in', 'C.B IN': 'break_in',
        'TIME OUT': 'time_out', 'FINAL OUT': 'time_out',
        'HOURS DUTY': 'hours', 'HRS RENDERED': 'hours', 'NO. OF HOURS': 'hours',
        'NO. OF HOURS RENDERED': 'hours', 'TOTAL HOURS': 'hours', 'NET HOURS': 'hours',
        'NET HOURS RENDERED': 'hours'
    };
    const NAME_LABEL = /^\s*(NAME OF (?:GUARD|SECURITY|OFFICER)|SECURITY GUARD|LADY GUARD|LADY GUAR|OFFICER|GUARD NAME|GUARD|NAME)\s*[:.]?\s*(.*)$/i;
    const RANK_PREFIX = /^(SO|SG|LG|SSG|PO|OIC|SIC|SOI)\.?\s+/;

    function parseStackedBlocks(ws, sheetName) {
        const rng = sheetRange(ws);
        const anchors = [];
        for (let r = rng.s.r; r <= rng.e.r; r++) {
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                if (/TIME IN|INN?ITIAL IN/i.test(String(cellAt(ws, r, c) || ''))) { anchors.push(r); break; }
            }
        }
        const clean = (x) => norm(x).replace(/^[.:,\-\s]+/, '').replace(RANK_PREFIX, '');
        const out = [];

        anchors.forEach((hr, bi) => {
            const map = {};
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                const v = cellAt(ws, hr, c);
                if (typeof v === 'string' && STACK_HEADERS[upper(v)]) map[STACK_HEADERS[upper(v)]] = c;
            }
            if (map.date == null) return;

            // Name and Access ID sit in the few rows above the header.
            let id = '', guard = '';
            for (let up = 1; up <= 4 && hr - up >= 0 && !id; up++) {
                for (let c = rng.s.c; c <= rng.e.c; c++) {
                    if (!/ACCESS ID|EMPLOYEE ID/i.test(String(cellAt(ws, hr - up, c) || ''))) continue;
                    for (let cc = c; cc <= rng.e.c; cc++) {
                        const nv = cellAt(ws, hr - up, cc);
                        if (typeof nv === 'number' || (typeof nv === 'string' && /^'?0*\d{3,}$/.test(String(nv).trim()))) {
                            id = String(nv).replace(/^'/, '').trim(); break;
                        }
                    }
                }
            }
            for (let up = 1; up <= 4 && hr - up >= 0 && !guard; up++) {
                for (let c = rng.s.c; c <= rng.e.c; c++) {
                    const v = cellAt(ws, hr - up, c);
                    if (typeof v !== 'string') continue;
                    const m = v.match(NAME_LABEL);
                    if (m && /[A-Za-z]{2}/.test(m[2] || '')) { guard = clean(m[2]); break; }
                    if (m) {                                     // label alone; value is to its right
                        for (let cc = c + 1; cc <= rng.e.c; cc++) {
                            const nv = cellAt(ws, hr - up, cc);
                            if (nv != null && String(nv).trim()) { guard = clean(String(nv)); break; }
                        }
                        if (guard) break;
                    }
                }
            }
            if (!guard && sheetName && !/^sheet\d*$/i.test(String(sheetName).trim())) guard = clean(sheetName);

            const end = bi + 1 < anchors.length ? anchors[bi + 1] : rng.e.r + 1;
            for (let r = hr + 1; r < end; r++) {
                const date = parseDate(cellAt(ws, r, map.date), null);
                if (!date) continue;
                const sched = String(cellAt(ws, r, map.schedule) || '');
                const raw = {};
                for (const c of COLUMNS) raw[c] = map[c] != null ? parseTime(cellAt(ws, r, map[c])) : null;
                if (!COLUMNS.some((c) => raw[c]) || REST_RE.test(sched)) {
                    out.push(makeRestRow(id, guard, date, 'rest day'));
                    continue;
                }
                out.push(makeRow(id, guard, date, raw, map.hours != null ? cellAt(ws, r, map.hours) : null, []));
            }
        });
        return out;
    }

    // ── Layout C: day-number dates, month/year floating above the table ───
    const SKIP_TEXT = /STARLINE|SECURITY|AGENCY|CUT ?OFF|^DATE$|^\d+$|JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER/i;

    function monthYearAbove(ws, r, rng) {
        for (let up = 1; up <= 8 && r - up >= 0; up++) {
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                const v = String(cellAt(ws, r - up, c) || '');
                const mm = v.toUpperCase().match(/(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)/);
                if (mm) {
                    const yy = v.match(/(20\d\d)/);
                    return { mo: MONTH_NAMES[mm[1].toLowerCase()], year: yy ? +yy[1] : null };
                }
            }
        }
        return { mo: null, year: null };
    }
    function nameAbove(ws, r, rng) {
        for (let up = 1; up <= 4 && r - up >= 0; up++) {
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                const v = cellAt(ws, r - up, c);
                if (typeof v === 'string' && v.trim() && !SKIP_TEXT.test(v.trim())) return norm(v.split('/')[0]);
            }
        }
        return '';
    }
    function parseDayNumberBlocks(ws, fallbackYear) {
        const rng = sheetRange(ws);
        const out = [];
        for (let r = rng.s.r; r <= rng.e.r; r++) {
            let isHeader = false;
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                if (/INN?ITIAL IN/i.test(String(cellAt(ws, r, c) || ''))) { isHeader = true; break; }
            }
            if (!isHeader) continue;

            const map = {};
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                const v = cellAt(ws, r, c);
                if (typeof v !== 'string') continue;
                const key = upper(v).replace(/^INNITIAL IN$/, 'INITIAL IN');
                if (STACK_HEADERS[key]) map[STACK_HEADERS[key]] = c;
            }
            if (map.date == null) continue;

            const my = monthYearAbove(ws, r, rng);
            const guard = nameAbove(ws, r, rng);
            const year = my.year || fallbackYear || new Date().getFullYear();
            const yearFlags = my.year ? [] : ['year assumed ' + year];

            for (let d = r + 1; d <= rng.e.r; d++) {
                const dv = cellAt(ws, d, map.date);
                if (typeof dv !== 'number' || dv < 1 || dv > 31) {
                    if (dv == null) continue;
                    break;                                       // next block, or end of this one
                }
                const date = `${year}-${pad2(my.mo || 1)}-${pad2(dv)}`;
                const inCell = cellAt(ws, d, map.time_in);
                if (typeof inCell === 'string' && REST_RE.test(inCell)) {
                    out.push(makeRestRow('', guard, date, 'rest day'));
                    continue;
                }
                const raw = {};
                for (const c of COLUMNS) raw[c] = map[c] != null ? parseTime(cellAt(ws, d, map[c])) : null;
                if (!COLUMNS.some((c) => raw[c])) { out.push(makeRestRow('', guard, date, 'no punches')); continue; }
                out.push(makeRow('', guard, date, raw, map.hours != null ? cellAt(ws, d, map.hours) : null, yearFlags));
            }
        }
        return out;
    }

    // ── Layout D: headerless — name · date · weekday · six punches ───────
    const WEEKDAY_RE = /^(SUN|MON|TUE|WED|THU|FRI|SAT)/i;
    function looksPositional(ws) {
        const rng = sheetRange(ws);
        let hits = 0;
        for (let r = rng.s.r; r <= Math.min(rng.e.r, rng.s.r + 40); r++) {
            const a = cellAt(ws, r, rng.s.c), b = cellAt(ws, r, rng.s.c + 1), c = cellAt(ws, r, rng.s.c + 2);
            if (typeof a === 'string' && a.trim() && parseDate(b, null) && typeof c === 'string' && WEEKDAY_RE.test(c.trim())) hits++;
        }
        return hits >= 3;
    }
    function parsePositional(ws) {
        const rng = sheetRange(ws);
        const c0 = rng.s.c;
        const out = [];
        for (let r = rng.s.r; r <= rng.e.r; r++) {
            const name = cellAt(ws, r, c0);
            const date = parseDate(cellAt(ws, r, c0 + 1), null);
            if (typeof name !== 'string' || !name.trim() || !date) continue;
            const first = cellAt(ws, r, c0 + 3);
            if (typeof first === 'string' && REST_RE.test(first)) {
                out.push(makeRestRow('', name, date, 'rest day'));
                continue;
            }
            const raw = {};
            COLUMNS.forEach((c, i) => { raw[c] = parseTime(cellAt(ws, r, c0 + 3 + i)); });
            if (!COLUMNS.some((c) => raw[c])) { out.push(makeRestRow('', name, date, 'no punches')); continue; }
            out.push(makeRow('', name, date, raw, null, []));
        }
        return out;
    }

    // ── Layout E: our own template / any sheet with a named header row ────
    function detectTableColumns(ws) {
        const rng = sheetRange(ws);
        for (let r = rng.s.r; r <= Math.min(rng.e.r, rng.s.r + 12); r++) {
            const cols = { idCol: -1, nameCol: -1, dateCol: -1, hoursCol: -1, timeCols: {}, headerRow: r };
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                const key = normKey(cellAt(ws, r, c));
                if (!key) continue;
                if (cols.idCol < 0 && EMPID_ALIASES.has(key)) { cols.idCol = c; continue; }
                if (cols.nameCol < 0 && NAME_ALIASES.has(key)) { cols.nameCol = c; continue; }
                if (cols.dateCol < 0 && DATE_ALIASES.has(key)) { cols.dateCol = c; continue; }
                if (DAY_ALIASES.has(key)) continue;
                const col = HEADER_ALIASES[key];
                if (col && cols.timeCols[col] === undefined) cols.timeCols[col] = c;
            }
            // Claim the sheet only when it identifies the guard on the row itself.
            // Without that, identity lives in text around the table — which is
            // the stacked / day-number layouts' job, not this one.
            const identified = cols.idCol >= 0 || cols.nameCol >= 0;
            if (identified && cols.dateCol >= 0 && Object.keys(cols.timeCols).length >= 2) return cols;
        }
        return null;
    }
    function parseTable(ws, cols, sheetName) {
        const rng = sheetRange(ws);
        const out = [];
        let lastId = '', lastName = '';
        for (let r = cols.headerRow + 1; r <= rng.e.r; r++) {
            const date = parseDate(cellAt(ws, r, cols.dateCol), new Date().getFullYear());
            if (!date) continue;
            const id = cols.idCol >= 0 ? (norm(cellAt(ws, r, cols.idCol)) || lastId) : '';
            const nm = cols.nameCol >= 0 ? (norm(cellAt(ws, r, cols.nameCol)) || lastName) : '';
            lastId = id; lastName = nm;
            if (!id && !nm && !sheetName) continue;

            const raw = {};
            for (const c of COLUMNS) raw[c] = cols.timeCols[c] != null ? parseTime(cellAt(ws, r, cols.timeCols[c])) : null;
            if (!COLUMNS.some((c) => raw[c])) {
                const marker = upper(cellAt(ws, r, cols.timeCols.time_in));
                if (marker && REST_RE.test(marker)) { out.push(makeRestRow(id, nm || sheetName, date, 'rest day')); }
                continue;                                        // otherwise just an empty day
            }
            out.push(makeRow(id, nm || sheetName, date, raw,
                cols.hoursCol >= 0 ? cellAt(ws, r, cols.hoursCol) : null, []));
        }
        return out;
    }

    // ── The router ───────────────────────────────────────────────────────
    function classifySheet(ws) {
        if (!ws || !ws['!ref']) return 'empty';
        const rng = sheetRange(ws);
        const lastRow = Math.min(rng.e.r, rng.s.r + 60);
        for (let r = rng.s.r; r <= lastRow; r++) {
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                const v = String(cellAt(ws, r, c) || '');
                if (/SCHEDULE_START_DATE|SCHEDULE_END_DATE|ACTUAL SCHEDULE OF GUARDS/i.test(v)) return 'schedule';
            }
        }
        if (detectTableColumns(ws)) return 'table';
        for (let r = rng.s.r; r <= lastRow; r++) {
            for (let c = rng.s.c; c <= rng.e.c; c++) {
                const v = String(cellAt(ws, r, c) || '');
                if (/Personnel Name|Employee Name|DTR Summary Report/i.test(v)) return 'flat';
                if (/INN?ITIAL IN/i.test(v)) return 'daynumber';
                if (/SECURITY GUARD\s*:|LADY GUARD\s*:|NAME OF GUARD/i.test(v) || /^\s*TIME\s*IN\s*$/i.test(v)) return 'stacked';
            }
        }
        if (looksPositional(ws)) return 'positional';
        return 'unknown';
    }

    function parseSheet(ws, sheetName, fallbackYear) {
        const format = classifySheet(ws);
        let rows = [];
        if (format === 'table') rows = parseTable(ws, detectTableColumns(ws), sheetName);
        else if (format === 'flat') rows = parseFlatReport(ws);
        else if (format === 'stacked') rows = parseStackedBlocks(ws, sheetName);
        else if (format === 'daynumber') rows = parseDayNumberBlocks(ws, fallbackYear);
        else if (format === 'positional') rows = parsePositional(ws);
        return { format, rows };
    }

    const FORMAT_LABELS = {
        table: 'column headers', flat: 'personnel report', stacked: 'per-guard blocks',
        daynumber: 'day-number blocks', positional: 'headerless columns',
        schedule: 'schedule (not punches)', unknown: 'unrecognised', empty: 'empty'
    };

    const nameTokens = (s) => new Set(normKey(s).split(' ').filter(Boolean));

    /**
     * One guard can appear in two sheets — with an access ID in a report and by
     * name only in a raw sheet — which would otherwise leave them split across
     * two roster entries, each holding half their month. Fold the name-only
     * entry into the ID'd one when exactly one name matches; never on a tie.
     */
    function consolidateByName(employees, warnings) {
        const withId = Array.from(employees.values()).filter((e) => e.id);
        const merged = [];
        for (const emp of Array.from(employees.values())) {
            if (emp.id) continue;
            const t = nameTokens(emp.name || emp.key);
            if (!t.size) continue;
            const hits = withId.filter((o) => {
                const ot = nameTokens(o.name);
                return ot.size === t.size && [...t].every((x) => ot.has(x));
            });
            if (hits.length !== 1) continue;              // no match, or ambiguous — leave it alone
            hits[0].rows = hits[0].rows.concat(emp.rows);
            employees.delete(emp.key);
            merged.push(`${emp.name || emp.key} → #${hits[0].id}`);
        }
        if (merged.length) {
            warnings.push(`Same guard found under both a name and an ID — merged: ${merged.join(', ')}`);
        }
    }

    /**
     * Two sheets can both carry the same day for the same guard. Keep whichever
     * row has more punches and flag it, rather than silently importing one.
     */
    function dedupeDays(rows) {
        const byDate = new Map();
        const score = (r) => (r.rest ? -1 : Object.keys(r.times).length);
        for (const row of rows) {
            const prev = byDate.get(row.date);
            if (!prev) { byDate.set(row.date, row); continue; }
            const keep = score(row) > score(prev) ? row : prev;
            keep.flags = (keep.flags || []).concat('same day appears twice in the file');
            byDate.set(row.date, keep);
        }
        return Array.from(byDate.values()).sort((a, b) => a.date.localeCompare(b.date));
    }

    /**
     * Group parsed day rows into employees. Keyed by access ID when the file
     * has one, else by name — matching how the page is matched later.
     */
    function buildRoster(sheets, sourceName) {
        const employees = new Map();
        const warnings = [];
        const formats = [];
        let totalRows = 0, restRows = 0, flaggedRows = 0;

        for (const { name: sheetName, ws } of sheets) {
            if (!ws) continue;
            if (/^\s*(read\s*me|readme|instructions?|notes?|guide|help|legend)\s*$/i.test(sheetName)) continue;

            let parsed;
            try { parsed = parseSheet(ws, sheetName, null); }
            catch (err) { warnings.push(`Sheet "${sheetName}": ${err && err.message ? err.message : 'could not be read'}`); continue; }

            if (parsed.format === 'schedule') {
                warnings.push(`Sheet "${sheetName}" is a schedule of planned shifts, not actual punches — skipped`);
                continue;
            }
            if (!parsed.rows.length) {
                if (parsed.format !== 'empty') {
                    warnings.push(`Sheet "${sheetName}": ${FORMAT_LABELS[parsed.format] || parsed.format} — no day rows found`);
                }
                continue;
            }
            formats.push(`${sheetName}: ${FORMAT_LABELS[parsed.format]}`);

            for (const row of parsed.rows) {
                const key = row.id || row.name || sheetName;
                if (!key) continue;
                if (!employees.has(key)) employees.set(key, { key, id: row.id, name: row.name, rows: [] });
                const emp = employees.get(key);
                if (!emp.name && row.name) emp.name = row.name;
                if (!emp.id && row.id) emp.id = row.id;
                emp.rows.push(row);
            }
        }

        consolidateByName(employees, warnings);

        for (const emp of employees.values()) {
            emp.rows = dedupeDays(emp.rows);
            for (const row of emp.rows) {
                if (row.rest) restRows++;
                else { totalRows++; if (row.flags.length) flaggedRows++; }
            }
        }

        return {
            source: sourceName,
            loadedAt: Date.now(),
            employees: Array.from(employees.values())
                .sort((a, b) => (a.name || a.key).localeCompare(b.name || b.key)),
            totalRows, restRows, flaggedRows, formats, warnings
        };
    }

    /** Read an .xlsx/.xls/.csv File into worksheets the parsers can walk. */
    function readWorkbook(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onerror = () => reject(new Error('could not read the file'));
            const isCsv = /\.csv$/i.test(file.name);

            reader.onload = () => {
                try {
                    if (typeof XLSX === 'undefined') {
                        reject(new Error('the spreadsheet library did not load — check your connection and reload the page'));
                        return;
                    }
                    if (isCsv) {
                        const rows = String(reader.result).split(/\r?\n/)
                            .filter((l) => norm(l) !== '')
                            .map((l) => splitFields(l));
                        resolve([{ name: file.name, ws: XLSX.utils.aoa_to_sheet(rows) }]);
                        return;
                    }
                    // cellDates keeps real times as Date objects; raw values are
                    // read straight from each cell so serials survive intact.
                    const wb = XLSX.read(new Uint8Array(reader.result), { type: 'array', cellDates: true });
                    resolve(wb.SheetNames.map((name) => ({ name, ws: wb.Sheets[name] })));
                } catch (err) { reject(err); }
            };

            if (isCsv) reader.readAsText(file);
            else reader.readAsArrayBuffer(file);
        });
    }

    // ── Working out which employee this page belongs to ──────────────────

    /** Page text with the timelogs grid and our own panel stripped out. */
    function pageIdentityText() {
        const clone = document.body.cloneNode(true);
        clone.querySelectorAll('table, script, style, #zdf-panel, #zdf-fab').forEach((n) => n.remove());
        return normKey(`${document.title} ${clone.textContent}`.replace(/[^\w\s]/g, ' '));
    }

    /** Digit sequences appearing on the page, for access-ID matching. */
    function pageNumbers() {
        const clone = document.body.cloneNode(true);
        clone.querySelectorAll('table, script, style, #zdf-panel, #zdf-fab').forEach((n) => n.remove());
        return new Set((`${document.title} ${clone.textContent}`.match(/\d{3,}/g) || []));
    }

    /**
     * Score each employee against the page. Returns the single confident match,
     * or null — which leaves the user to choose. Guessing wrong here would write
     * one guard's hours onto another, so ties and weak matches never auto-select.
     */
    function matchEmployeeToPage(rosterData) {
        if (!rosterData || !rosterData.employees.length) return null;

        // Strongest signal: this URL was confirmed for someone before.
        const links = readStore(LINKS_KEY, {});
        const remembered = links[location.pathname];
        if (remembered) {
            const hit = rosterData.employees.find((e) => e.key === remembered);
            if (hit) return { employee: hit, reason: 'remembered for this page' };
        }

        const text = pageIdentityText();
        const numbers = pageNumbers();
        const scored = rosterData.employees.map((emp) => {
            let score = 0;
            const why = [];
            if (emp.id && numbers.has(String(emp.id).replace(/\D/g, ''))) {
                score += 100; why.push(`access ID ${emp.id}`);
            }
            const tokens = normKey(emp.name).split(' ').filter((t) => t.length >= 3);
            const matched = tokens.filter((t) => new RegExp(`\\b${t}\\b`).test(text));
            if (tokens.length && matched.length) {
                score += matched.length * 10;
                if (matched.length === tokens.length) score += 10;
                why.push(`name "${matched.join(' ')}"`);
            }
            return { emp, score, why };
        }).sort((a, b) => b.score - a.score);

        const best = scored[0];
        const runnerUp = scored[1];
        if (!best || best.score < 20) return null;                       // too weak
        if (runnerUp && runnerUp.score === best.score) return null;      // ambiguous
        return { employee: best.emp, reason: best.why.join(' + ') };
    }

    /**
     * Render an employee's days into the paste box, reusing the normal flow.
     * A punch that rolled past midnight is written "06:00+1" so the offset
     * survives the round trip through the text box and stays visible to you.
     */
    function employeeToText(emp) {
        return emp.rows.filter((row) => !row.rest).map((row) => {
            const cells = COLUMNS.map((col) => {
                const t = row.times[col];
                if (!t) return '-';
                return `${pad2(t.h)}:${pad2(t.m)}${t.plus ? '+' + t.plus : ''}`;
            });
            return [row.date].concat(cells).join('\t');
        }).join('\n');
    }

    const doneKeyFor = (emp) => `${emp.key}`;
    function markDone(emp) {
        const done = readStore(DONE_KEY, {});
        done[doneKeyFor(emp)] = new Date().toISOString().slice(0, 10);
        writeStore(DONE_KEY, done);
    }
    function isDone(emp) { return !!readStore(DONE_KEY, {})[doneKeyFor(emp)]; }

    // =====================================================================
    //  OCR — reading a scanned or photographed DTR sheet
    // =====================================================================
    //  Tesseract runs entirely in this browser: the image never leaves the PC.
    //
    //  A DTR is nothing but digits, and OCR confuses 0/8, 1/7, 5/6 exactly
    //  there — one misread digit is somebody's pay. So nothing read here is
    //  trusted. A cell the engine is unsure of is written as ??:?? rather than
    //  a plausible guess, which the parser then refuses to fill, and every
    //  uncertain cell is named in the log so you can check it against the page.

    const OCR_IMAGE_RE = /\.(png|jpe?g|webp|bmp|gif|tiff?)$/i;
    // Thresholds are deliberately strict. A clean printed scan reads at 90%+,
    // while a photographed handwritten form lands around 60-70% and produces
    // values that look perfectly plausible and are simply wrong. Refusing a
    // borderline sheet costs a retype; accepting one corrupts someone's pay.
    const OCR_CELL_MIN = 80;      // below this a single value is not trustworthy
    const OCR_PAGE_MIN = 80;      // below this the whole sheet is unreadable
    const TIMEISH_RE = /^\d{1,2}\s*[:.;]\s*\d{2}$|^\d{3,4}\s*h?$/i;
    const UNREADABLE = '??:??';

    function ocrAvailable() { return typeof Tesseract !== 'undefined'; }

    /** Run Tesseract over an image file, reporting progress into the log. */
    async function ocrRecognize(file, onProgress) {
        if (!ocrAvailable()) {
            throw new Error('the OCR library did not load — check the connection and reload the page');
        }
        let last = -1;
        const worker = await Tesseract.createWorker('eng', 1, {
            logger: (m) => {
                if (!m || typeof m.progress !== 'number') return;
                const pct = Math.round(m.progress * 100);
                if (m.status === 'recognizing text' && pct >= last + 25) { last = pct; onProgress(`reading… ${pct}%`); }
            }
        });
        try {
            const { data } = await worker.recognize(file);
            return data;
        } finally {
            try { await worker.terminate(); } catch (e) { /* ignore */ }
        }
    }

    /** Cluster recognised words into visual rows by vertical overlap. */
    function groupWordsIntoRows(words) {
        const rows = [];
        for (const w of words) {
            if (!w || !w.text || !w.text.trim() || !w.bbox) continue;
            const mid = (w.bbox.y0 + w.bbox.y1) / 2;
            const h = Math.max(1, w.bbox.y1 - w.bbox.y0);
            let row = null;
            for (const r of rows) {
                if (Math.abs(r.mid - mid) < Math.max(h, r.h) * 0.6) { row = r; break; }
            }
            if (!row) { row = { mid, h, words: [] }; rows.push(row); }
            row.words.push(w);
            row.h = Math.max(row.h, h);
            row.mid = row.words.reduce((a, x) => a + (x.bbox.y0 + x.bbox.y1) / 2, 0) / row.words.length;
        }
        rows.forEach((r) => r.words.sort((a, b) => a.bbox.x0 - b.bbox.x0));
        return rows.sort((a, b) => a.mid - b.mid);
    }

    /**
     * Find header labels and where they sit horizontally. Used only to NAME
     * columns, never to position them — OCR routinely welds adjacent headers
     * together ("LUNCHIN", "OUBREAK") and mangles the rest.
     */
    function findHeaderColumns(rows) {
        for (const row of rows) {
            const found = [];
            for (let i = 0; i < row.words.length; i++) {
                const w = row.words[i], next = row.words[i + 1];
                const one = normKey(w.text);
                const two = next ? normKey(w.text + ' ' + next.text) : '';
                const pair = two && HEADER_ALIASES[two];
                const col = pair || HEADER_ALIASES[one];
                if (!col || found.some((f) => f.col === col)) continue;
                const span = pair ? [w, next] : [w];
                found.push({ col, x: (span[0].bbox.x0 + span[span.length - 1].bbox.x1) / 2 });
                if (pair) i++;
            }
            if (found.length >= 2) return { headerRow: row, cols: found.sort((a, b) => a.x - b.x) };
        }
        return null;
    }

    const isTimeWord = (w) => TIMEISH_RE.test(String(w.text || '').trim().replace(/[^\dhH:;.]/g, ''));
    const wordX = (w) => (w.bbox.x0 + w.bbox.x1) / 2;

    /**
     * Work out the sheet's time columns from where the VALUES actually sit —
     * the row with the most punches defines the grid — then name those columns
     * from the header wherever a label sits over one, and fill the rest in
     * canonical order. This survives a header OCR mangled beyond recognition.
     */
    function buildColumnGrid(dataRows, header) {
        let template = null;
        for (const row of dataRows) {
            const n = row.words.filter(isTimeWord).length;
            if (!template || n > template.n) template = { row, n };
        }
        if (!template || !template.n) return null;

        const centers = template.row.words.filter(isTimeWord).map(wordX).sort((a, b) => a - b);
        let spacing = Infinity;
        for (let i = 1; i < centers.length; i++) spacing = Math.min(spacing, centers[i] - centers[i - 1]);
        if (!isFinite(spacing)) spacing = 120;
        const tolerance = Math.max(20, spacing * 0.5);

        const assigned = centers.map(() => null);
        if (header) {
            for (const hc of header.cols) {
                let best = -1, bestD = Infinity;
                centers.forEach((x, i) => {
                    const d = Math.abs(x - hc.x);
                    if (d < bestD) { bestD = d; best = i; }
                });
                // A label only names a column it actually sits over.
                if (best >= 0 && bestD <= Math.max(tolerance, spacing * 0.75)
                    && !assigned[best] && assigned.indexOf(hc.col) < 0) {
                    assigned[best] = hc.col;
                }
            }
        }
        const used = new Set(assigned.filter(Boolean));
        let cursor = 0;
        for (let i = 0; i < assigned.length; i++) {
            if (assigned[i]) { cursor = COLUMNS.indexOf(assigned[i]) + 1; continue; }
            while (cursor < COLUMNS.length && used.has(COLUMNS[cursor])) cursor++;
            if (cursor >= COLUMNS.length) break;
            assigned[i] = COLUMNS[cursor];
            used.add(COLUMNS[cursor]);
            cursor++;
        }
        return { centers, assigned, tolerance, named: used.size };
    }

    /** Normalise one OCR'd time token; null when it is not a time at all. */
    function ocrToken(word) {
        const text = String(word.text || '').trim().replace(/[^\dhH:;.]/g, '');
        if (!TIMEISH_RE.test(text)) return null;
        if (word.confidence != null && word.confidence < OCR_CELL_MIN) return { value: UNREADABLE, low: true };
        const t = parseTime(text);
        if (!t) return { value: UNREADABLE, low: true };
        return { value: pad2(t.h) + ':' + pad2(t.m), low: false };
    }

    /**
     * Turn a recognised page into paste-box lines.
     * Dates: a full date is used as-is; a bare day number is matched against
     * the dates already on screen, so "3" becomes the 3rd of the shown month.
     */
    function ocrToLines(data, pageDates) {
        const rows = groupWordsIntoRows((data && data.words) || []);
        const header = findHeaderColumns(rows);
        const dayToDate = new Map();
        for (const iso of pageDates) {
            const day = +iso.slice(8, 10);
            if (dayToDate.has(day)) dayToDate.set(day, null);       // ambiguous across months
            else dayToDate.set(day, iso);
        }

        const onPage = new Set(pageDates);
        const dateOf = (row) => {
            for (const w of row.words) {
                const raw = String(w.text || '').trim();
                const full = parseDate(raw.replace(/[^\d\/\-.]/g, ''), null);
                // Only trust a date the page is actually showing — a misread
                // date is as likely as a misread time, and an off-page row
                // could not be filled regardless.
                if (full && (!onPage.size || onPage.has(full))) return full;
                const n = +raw.replace(/[^\d]/g, '');
                if (n >= 1 && n <= 31 && dayToDate.get(n)) return dayToDate.get(n);
            }
            return null;
        };

        const headerRow = header && header.headerRow;
        const dataRows = rows.filter((r) => r !== headerRow && dateOf(r) && r.words.some(isTimeWord));
        const grid = buildColumnGrid(dataRows, header);
        const base = {
            lines: [], issues: [], timeTokens: 0, lowCells: 0,
            confidence: (data && typeof data.confidence === 'number') ? data.confidence : 0,
            usedHeader: !!header, named: 0, columns: []
        };
        if (!grid) return base;

        const lines = [], issues = [];
        let timeTokens = 0, lowCells = 0;

        for (const row of dataRows) {
            const date = dateOf(row);
            const cells = COLUMNS.map(() => '');
            let placed = 0;

            for (const w of row.words) {
                const tok = ocrToken(w);
                if (!tok) continue;
                timeTokens++;
                if (tok.low) lowCells++;

                const x = wordX(w);
                let best = -1, bestD = Infinity;
                grid.centers.forEach((cx, i) => {
                    const d = Math.abs(cx - x);
                    if (d < bestD) { bestD = d; best = i; }
                });
                if (best < 0 || bestD > grid.tolerance * 2 || !grid.assigned[best]) {
                    issues.push(date + ': a value did not line up under any column — check that row');
                    continue;
                }
                const idx = COLUMNS.indexOf(grid.assigned[best]);
                if (cells[idx]) {
                    issues.push(date + ': two values landed in ' + COLUMN_LABELS[COLUMNS[idx]] + ' — check that row');
                    continue;
                }
                cells[idx] = tok.value;
                if (tok.low) issues.push(date + ': ' + COLUMN_LABELS[COLUMNS[idx]] + ' could not be read clearly');
                placed++;
            }
            if (!placed) continue;
            lines.push([date].concat(cells.map((c) => c || '-')).join('\t'));
        }

        return {
            lines, issues, timeTokens, lowCells,
            confidence: base.confidence, usedHeader: !!header, named: grid.named,
            columns: grid.assigned.map((c) => (c ? COLUMN_LABELS[c] : '?'))
        };
    }

    // =====================================================================
    //  PANEL UI
    // =====================================================================

    const CSS = `
    #zdf-panel { position: fixed; top: 12px; right: 12px; z-index: 2147483600;
        width: 380px; background: #ffffff; color: #1f2933; font: 13px/1.45 "Segoe UI", system-ui, sans-serif;
        border: 1px solid #cbd2d9; border-radius: 10px; box-shadow: 0 10px 34px rgba(0,0,0,.22); overflow: hidden; }
    #zdf-panel * { box-sizing: border-box; font-family: inherit; }
    #zdf-head { display: flex; align-items: center; gap: 8px; padding: 9px 11px;
        background: #6dbe45; color: #fff; cursor: move; user-select: none; }
    #zdf-head b { flex: 1; font-size: 13px; font-weight: 600; letter-spacing: .2px; }
    #zdf-head button { background: rgba(255,255,255,.2); border: 0; color: #fff; width: 22px; height: 22px;
        border-radius: 5px; cursor: pointer; font-size: 14px; line-height: 1; }
    #zdf-head button:hover { background: rgba(255,255,255,.35); }
    #zdf-body { padding: 11px; }
    #zdf-panel.zdf-collapsed #zdf-body { display: none; }
    #zdf-paste { width: 100%; height: 128px; resize: vertical; padding: 7px 8px; border: 1px solid #cbd2d9;
        border-radius: 6px; font-family: Consolas, "Courier New", monospace; font-size: 11.5px; white-space: pre; overflow-x: auto; }
    #zdf-paste:focus { outline: 2px solid #6dbe45; outline-offset: -1px; }
    .zdf-hint { color: #7b8794; font-size: 11px; margin: 5px 0 8px; }
    .zdf-btns { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }
    .zdf-btns button { flex: 1 1 auto; padding: 7px 9px; border-radius: 6px; border: 1px solid #cbd2d9;
        background: #f5f7fa; cursor: pointer; font-size: 12px; font-weight: 500; }
    .zdf-btns button:hover { background: #e4e7eb; }
    .zdf-btns button.zdf-primary { background: #6dbe45; border-color: #5aa838; color: #fff; }
    .zdf-btns button.zdf-primary:hover { background: #5aa838; }
    .zdf-btns button:disabled { opacity: .5; cursor: default; }
    .zdf-opts { display: flex; flex-direction: column; gap: 4px; margin-bottom: 8px; font-size: 11.5px; color: #52606d; }
    .zdf-opts label { display: flex; align-items: center; gap: 6px; cursor: pointer; }
    #zdf-log { background: #12161c; color: #cbd5e0; border-radius: 6px; padding: 8px; height: 150px;
        overflow-y: auto; font-family: Consolas, "Courier New", monospace; font-size: 11px; white-space: pre-wrap; word-break: break-word; }
    #zdf-log div { margin-bottom: 2px; }
    #zdf-log .ok { color: #7ee787; }
    #zdf-log .err { color: #ff7b72; }
    #zdf-log .warn { color: #e3b341; }
    #zdf-log .info { color: #79c0ff; }
    #zdf-status { margin-top: 7px; font-size: 11.5px; color: #52606d; }
    .zdf-file { display: flex; gap: 6px; margin-bottom: 8px; align-items: center; }
    .zdf-file input[type=file] { display: none; }
    .zdf-file label { flex: 1; padding: 7px 9px; border: 1px dashed #a7b0ba; border-radius: 6px;
        background: #f5f7fa; cursor: pointer; font-size: 12px; font-weight: 500; text-align: center; color: #3e4c59; }
    .zdf-file label:hover { border-color: #6dbe45; background: #eefbe7; color: #3d7a24; }
    .zdf-file button { padding: 7px 9px; border-radius: 6px; border: 1px solid #cbd2d9;
        background: #f5f7fa; cursor: pointer; font-size: 12px; }
    #zdf-roster { display: none; margin-bottom: 8px; }
    #zdf-roster select { width: 100%; padding: 6px 7px; border: 1px solid #cbd2d9; border-radius: 6px; font-size: 12px; }
    #zdf-match { margin-top: 5px; font-size: 11px; padding: 5px 7px; border-radius: 5px; }
    #zdf-match.hit { background: #eefbe7; color: #3d7a24; border: 1px solid #b7e39a; }
    #zdf-match.miss { background: #fff6e0; color: #8a6100; border: 1px solid #f0d089; }
    #zdf-fab { position: fixed; top: 12px; right: 12px; z-index: 2147483600; display: none;
        background: #6dbe45; color: #fff; border: 0; border-radius: 20px; padding: 8px 14px;
        font: 600 12px "Segoe UI", system-ui, sans-serif; cursor: pointer; box-shadow: 0 4px 14px rgba(0,0,0,.25); }
    input.zdf-touched { outline: 2px solid #6dbe45 !important; background: #eefbe7 !important; }
    tr.zdf-row-touched > td { background: #f4fcef !important; }
    `;

    function buildPanel() {
        const style = document.createElement('style');
        style.textContent = CSS;
        document.head.appendChild(style);

        const fab = document.createElement('button');
        fab.id = 'zdf-fab';
        fab.textContent = 'DTR Filler';
        document.body.appendChild(fab);

        const panel = document.createElement('div');
        panel.id = 'zdf-panel';
        panel.innerHTML = `
            <div id="zdf-head">
                <b>Zenhours DTR Filler</b>
                <button id="zdf-min" title="Collapse">–</button>
                <button id="zdf-close" title="Hide (Ctrl+Shift+D)">×</button>
            </div>
            <div id="zdf-body">
                <div class="zdf-file">
                    <label for="zdf-xlsx" id="zdf-filelabel">Load Excel, CSV or a scan…</label>
                    <input type="file" id="zdf-xlsx" accept=".xlsx,.xls,.csv,.png,.jpg,.jpeg,.webp,.bmp,.tif,.tiff">
                    <button id="zdf-forget" title="Forget the loaded file">Clear</button>
                </div>
                <div id="zdf-roster">
                    <select id="zdf-emp"></select>
                    <div id="zdf-match"></div>
                </div>
                <textarea id="zdf-paste" spellcheck="false"
                    placeholder="Paste from Excel — one day per line:&#10;8/1/2026&#9;TUESDAY&#9;10:20&#9;12:40&#9;13:12&#9;17:00&#9;17:30&#9;21:00&#10;8/2/2026&#9;WEDNESDAY&#9;10:25&#9;12:45&#9;13:15&#9;16:39&#9;17:09&#9;21:00"></textarea>
                <div class="zdf-hint">Date · <i>(day name — ignored)</i> · Time In · Lunch Out · Lunch In · Break Out · Break In · Time Out<br>
                    24-hour or AM/PM both work. Leave a cell empty or put a dash to skip that column.</div>
                <div class="zdf-btns">
                    <button id="zdf-parse">Parse</button>
                    <button id="zdf-test">Test 1st day</button>
                    <button id="zdf-fill" class="zdf-primary">Fill all rows</button>
                </div>
                <div class="zdf-btns">
                    <button id="zdf-undo">Undo fill</button>
                    <button id="zdf-diag">Copy diagnostics</button>
                    <button id="zdf-clear">Clear log</button>
                </div>
                <div class="zdf-opts">
                    <label><input type="checkbox" id="zdf-onlyblank" checked> Only fill columns that are blank (--:--)</label>
                    <label><input type="checkbox" id="zdf-openedit" checked> Click Edit automatically</label>
                    <label><input type="checkbox" id="zdf-clearblanks"> Clear the field when my cell is blank/dash</label>
                </div>
                <div id="zdf-log"></div>
                <div id="zdf-status">Ready.</div>
            </div>`;
        document.body.appendChild(panel);
        return { panel, fab };
    }

    function init() {
        const { panel, fab } = buildPanel();
        const $id = (id) => panel.querySelector('#' + id);
        const logBox = $id('zdf-log');
        const statusBox = $id('zdf-status');
        const settings = loadSettings();

        if (settings.onlyBlank === false) $id('zdf-onlyblank').checked = false;
        if (settings.openEdit === false) $id('zdf-openedit').checked = false;
        if (settings.clearBlanks === true) $id('zdf-clearblanks').checked = true;

        // Only restore a remembered paste in manual mode. With a workbook loaded
        // the box must always be (re)filled from the employee matched to THIS
        // page — otherwise the previous guard's hours linger in the box and can
        // be written onto whoever's page you land on next.
        const storedRoster = readStore(ROSTER_KEY, null);
        const hasRoster = !!(storedRoster && storedRoster.employees && storedRoster.employees.length);
        if (settings.paste && !hasRoster) $id('zdf-paste').value = settings.paste;

        function log(msg, cls) {
            const line = document.createElement('div');
            line.className = cls || 'info';
            line.textContent = msg;
            logBox.appendChild(line);
            logBox.scrollTop = logBox.scrollHeight;
        }
        function status(msg) { statusBox.textContent = msg; }

        function persist() {
            saveSettings({
                onlyBlank: $id('zdf-onlyblank').checked,
                openEdit: $id('zdf-openedit').checked,
                clearBlanks: $id('zdf-clearblanks').checked,
                // Never carry one employee's times to the next page (see above).
                paste: roster ? '' : $id('zdf-paste').value.slice(0, 20000)
            });
        }

        function currentParse() {
            const result = parsePaste($id('zdf-paste').value, pageYear());
            result.warnings.forEach((w) => log('! ' + w, 'warn'));
            if (result.mapping) log('Header detected → ' + result.mapping, 'info');
            return result;
        }

        // ── Roster: load a workbook, pick the guard this page belongs to ─
        const empSelect = $id('zdf-emp');
        const matchBox = $id('zdf-match');

        function selectedEmployee() {
            if (!roster) return null;
            return roster.employees.find((e) => e.key === empSelect.value) || null;
        }

        function loadSelectedIntoBox(remember) {
            const emp = selectedEmployee();
            if (!emp) return;
            $id('zdf-paste').value = employeeToText(emp);
            if (remember) {
                const links = readStore(LINKS_KEY, {});
                links[location.pathname] = emp.key;
                writeStore(LINKS_KEY, links);
            }
            persist();
        }

        function renderRoster(autoMatch) {
            if (!roster || !roster.employees.length) {
                $id('zdf-roster').style.display = 'none';
                return;
            }
            $id('zdf-roster').style.display = 'block';
            $id('zdf-filelabel').textContent =
                `${roster.source} — ${roster.employees.length} employees, ${roster.totalRows} rows`;

            empSelect.innerHTML = '<option value="">— choose the employee for this page —</option>' +
                roster.employees.map((e) => {
                    const label = [e.name || '(no name)', e.id ? `#${e.id}` : '', `${e.rows.length}d`]
                        .filter(Boolean).join(' · ');
                    return `<option value="${e.key}">${isDone(e) ? '✓ ' : ''}${label}</option>`;
                }).join('');

            const match = autoMatch ? matchEmployeeToPage(roster) : null;
            if (match) {
                empSelect.value = match.employee.key;
                matchBox.className = 'hit';
                const workDays = match.employee.rows.filter((r) => !r.rest).length;
                matchBox.textContent = `Matched this page by ${match.reason} → ${match.employee.name || match.employee.key}`
                    + ` (${workDays} days loaded)`;
                loadSelectedIntoBox(false);
                reportEmployeeRows(match.employee);
            } else if (autoMatch) {
                empSelect.value = '';
                $id('zdf-paste').value = '';        // never leave another guard's times sitting here
                ocrActive = false;
                matchBox.className = 'miss';
                matchBox.textContent = 'Could not tell which employee this page is for — pick them from the list above. '
                    + 'Filling is blocked until you do.';
            }
        }

        /**
         * Read a scanned DTR into the paste box. Deliberately does NOT enter the
         * roster or fill anything: OCR output is a draft for you to check against
         * the sheet in front of you, and unreadable cells are left as ??:?? so
         * the parser refuses them rather than inventing a plausible time.
         */
        async function runOcr(file) {
            const buttons = panel.querySelectorAll('.zdf-btns button');
            buttons.forEach((b) => (b.disabled = true));
            log(`Reading ${file.name} with OCR — this runs on this PC, the image is not uploaded.`, 'info');
            status('Running OCR…');
            try {
                const data = await ocrRecognize(file, (msg) => status(msg));
                const pageDates = Array.from(indexRows().keys());
                const res = ocrToLines(data, pageDates);

                // Quality gate: refuse rather than emit confident-looking nonsense.
                if (res.confidence && res.confidence < OCR_PAGE_MIN) {
                    log(`Overall confidence ${Math.round(res.confidence)}% — too low to trust.`, 'err');
                    log('That usually means a handwritten form, a photo at an angle, or a low-resolution scan.', 'info');
                    log('Nothing was filled. Type these times in by hand, or rescan flat at 300dpi.', 'info');
                    status('OCR refused — sheet not readable.');
                    return;
                }
                if (!res.lines.length) {
                    log('No dated rows could be read from that image.', 'err');
                    log(res.timeTokens
                        ? 'Times were found but no dates — make sure the date column is in the picture, and Search the matching range on this page first.'
                        : 'No times were found at all — if this is a handwritten form, OCR cannot read it.', 'info');
                    status('OCR found nothing to fill.');
                    return;
                }

                ocrActive = true;
                $id('zdf-paste').value = res.lines.join('\n');
                log(`Read ${res.lines.length} row(s) at ${Math.round(res.confidence)}% overall confidence`
                    + (res.usedHeader ? ", using the sheet's own column headers." : ", left-to-right (no header row found)."), 'ok');
                if (res.columns && res.columns.length) {
                    log('  columns read left to right: ' + res.columns.join(' | '), 'info');
                }
                if (!res.usedHeader) {
                    log("! Without a header row, a missing punch can shift a row's columns — check each row.", 'warn');
                }
                res.issues.slice(0, 10).forEach((i) => log('  ! ' + i, 'warn'));
                if (res.issues.length > 10) log(`  ! …and ${res.issues.length - 10} more.`, 'warn');
                if (res.lowCells) {
                    log(`${res.lowCells} cell(s) came out as ${UNREADABLE} — they will NOT be filled until you type them in.`, 'warn');
                }
                log('Check every value against the sheet before filling. OCR misreads digits.', 'warn');
                status(`OCR read ${res.lines.length} rows — review them, then Fill.`);
            } catch (err) {
                log('OCR failed: ' + (err && err.message ? err.message : String(err)), 'err');
                status('OCR failed.');
            } finally {
                buttons.forEach((b) => (b.disabled = false));
            }
        }

        $id('zdf-xlsx').addEventListener('change', async (e) => {
            const file = e.target.files && e.target.files[0];
            if (!file) return;
            logBox.innerHTML = '';
            if (/\.pdf$/i.test(file.name)) {
                log('PDFs are not read directly yet — export the page as a PNG or JPG and load that.', 'err');
                e.target.value = '';
                return;
            }
            if (OCR_IMAGE_RE.test(file.name) || /^image\//.test(file.type || '')) {
                await runOcr(file);
                e.target.value = '';
                return;
            }
            log(`Reading ${file.name}…`, 'info');
            try {
                ocrActive = false;
                const sheets = await readWorkbook(file);
                roster = buildRoster(sheets, file.name);
                if (!roster.employees.length) {
                    log('No employee rows found. Check that the sheet has a Date column and time columns.', 'err');
                    roster.warnings.slice(0, 6).forEach((w) => log('! ' + w, 'warn'));
                    return;
                }
                if (!writeStore(ROSTER_KEY, roster)) {
                    log('File is too large to remember between pages — it stays loaded for this page only.', 'warn');
                }
                log(`Loaded ${roster.employees.length} employees, ${roster.totalRows} day rows`
                    + (roster.restRows ? `, ${roster.restRows} rest/leave days skipped` : '') + '.', 'ok');
                (roster.formats || []).slice(0, 8).forEach((f) => log('  layout — ' + f, 'info'));
                if (roster.flaggedRows) {
                    log(`${roster.flaggedRows} day row(s) need a look — select an employee to see theirs.`, 'warn');
                }
                roster.warnings.slice(0, 6).forEach((w) => log('! ' + w, 'warn'));
                if (roster.warnings.length > 6) log(`! …and ${roster.warnings.length - 6} more warnings.`, 'warn');
                renderRoster(true);
                status(`${roster.employees.length} employees loaded.`);
            } catch (err) {
                log('Could not read that file: ' + (err && err.message ? err.message : String(err)), 'err');
            } finally {
                e.target.value = '';           // allow re-picking the same file
            }
        });

        empSelect.addEventListener('change', () => {
            const emp = selectedEmployee();
            if (!emp) return;
            ocrActive = false;          // roster data supersedes an earlier scan
            matchBox.className = 'hit';
            matchBox.textContent = `${emp.name || emp.key} — ${emp.rows.length} days loaded into the box below.`;
            loadSelectedIntoBox(true);        // remember: you told us who this page is
            reportEmployeeRows(emp);
        });

        /** Say what was read for this guard, and which days the importer distrusts. */
        function reportEmployeeRows(emp) {
            const work = emp.rows.filter((r) => !r.rest);
            const rest = emp.rows.length - work.length;
            log(`Selected ${emp.name || emp.key}${emp.id ? ` (#${emp.id})` : ''} — ${work.length} day(s)`
                + (rest ? `, ${rest} rest/leave day(s) not loaded` : '') + '.', 'info');
            const flagged = work.filter((r) => r.flags && r.flags.length);
            flagged.slice(0, 8).forEach((r) => log(`  ! ${r.date} — ${r.flags.join('; ')}`, 'warn'));
            if (flagged.length > 8) log(`  ! …and ${flagged.length - 8} more day(s) worth checking.`, 'warn');
        }

        $id('zdf-forget').addEventListener('click', () => {
            roster = null;
            try { localStorage.removeItem(ROSTER_KEY); } catch (err) { /* ignore */ }
            $id('zdf-roster').style.display = 'none';
            $id('zdf-filelabel').textContent = 'Load Excel, CSV or a scan…';
            log('Forgot the loaded file. The paste box still works on its own.', 'info');
        });

        // ── Buttons ──────────────────────────────────────────────────────
        $id('zdf-parse').addEventListener('click', () => {
            logBox.innerHTML = '';
            const { entries } = currentParse();
            if (!entries.length) { log('No rows parsed.', 'err'); status('Nothing to fill.'); return; }
            const rows = indexRows();
            let found = 0;
            for (const e of entries) {
                const on = rows.has(e.date);
                if (on) found++;
                const preview = COLUMNS
                    .filter((c) => e.times[c] || (e.blanks || []).includes(c))
                    .map((c) => e.times[c]
                        ? `${COLUMN_LABELS[c]} ${formatTimeLike('12:00 AM', e.times[c])}`
                        : `${COLUMN_LABELS[c]} —`)
                    .join(' · ');
                log(`${on ? '✓' : '✗'} ${e.date} → ${preview}`, on ? 'ok' : 'err');
            }
            log(`Table has ${rows.size} dated row(s).`, 'info');
            status(`${entries.length} line(s) parsed · ${found} matched on page.`);
            persist();
        });

        async function doFill(limitToFirst) {
            logBox.innerHTML = '';

            // With a workbook loaded, filling requires knowing whose page this is.
            // Writing one guard's hours onto another is the expensive mistake here,
            // so this refuses rather than guesses.
            if (roster && roster.employees.length && !selectedEmployee() && !ocrActive) {
                log('No employee selected for this page.', 'err');
                log('Pick the guard from the list above — or click Clear to drop the file and paste manually.', 'info');
                status('Pick the employee first.');
                return;
            }

            const { entries } = currentParse();
            if (!entries.length) { log('No rows parsed.', 'err'); status('Nothing to fill.'); return; }

            const list = limitToFirst ? entries.slice(0, 1) : entries;
            const opts = {
                onlyBlank: $id('zdf-onlyblank').checked,
                openEdit: $id('zdf-openedit').checked,
                clearBlanks: $id('zdf-clearblanks').checked,
                delay: 120
            };
            if (!opts.openEdit) log('Auto-Edit is off — only rows already in edit mode will fill.', 'warn');

            const buttons = panel.querySelectorAll('.zdf-btns button');
            buttons.forEach((b) => (b.disabled = true));
            status(limitToFirst ? 'Filling first row…' : `Filling ${list.length} row(s)…`);

            try {
                if (!opts.openEdit) {
                    // Respect the toggle: skip rows that aren't already editable.
                    for (const e of list) {
                        const tr = indexRows().get(e.date);
                        if (tr && !isEditing(tr)) e.__skip = true;
                    }
                }
                const runnable = list.filter((e) => !e.__skip);
                const r = await runFill(runnable, opts, log);
                log('—', 'info');
                log(`Done: ${r.filledFields} field(s) across ${r.filledRows} row(s)` +
                    (r.skipped ? `, ${r.skipped} kept (already had a value)` : '') +
                    (r.blanksLeft ? `, ${r.blanksLeft} left at 12:00 AM (blank in your paste)` : '') +
                    (r.missingRows ? `, ${r.missingRows} row(s) not filled` : '') + '.', 'ok');
                if (r.overnightFields) {
                    log(`${r.overnightFields} overnight punch(es) were dated to the following day — check those rows before saving.`, 'warn');
                }
                log('Nothing was saved — review the highlighted inputs, then click Save on each row.', 'info');
                status(`Filled ${r.filledRows} row(s). Click Save yourself.`);

                // Track progress through the roster so you can pick up where you left off.
                const emp = selectedEmployee();
                if (emp && r.filledRows && !limitToFirst) {
                    markDone(emp);
                    const done = roster.employees.filter(isDone).length;
                    const next = roster.employees.find((e) => !isDone(e));
                    log(`Roster progress: ${done} of ${roster.employees.length} employees filled.`
                        + (next ? ` Next up: ${next.name || next.key}${next.id ? ` (#${next.id})` : ''}.` : ' All done.'), 'info');
                }
            } catch (err) {
                log('Error: ' + (err && err.message ? err.message : String(err)), 'err');
                status('Stopped on an error.');
            } finally {
                buttons.forEach((b) => (b.disabled = false));
                persist();
            }
        }

        $id('zdf-test').addEventListener('click', () => doFill(true));
        $id('zdf-fill').addEventListener('click', () => doFill(false));
        $id('zdf-undo').addEventListener('click', () => undoFill(log));
        $id('zdf-clear').addEventListener('click', () => { logBox.innerHTML = ''; status('Ready.'); });

        $id('zdf-diag').addEventListener('click', async () => {
            const text = diagnostics();
            try {
                await navigator.clipboard.writeText(text);
                log('Diagnostics copied to clipboard — paste them to Claude if a row will not fill.', 'ok');
            } catch (e) {
                log('Clipboard blocked. Diagnostics printed below:', 'warn');
                log(text, 'info');
            }
        });

        $id('zdf-onlyblank').addEventListener('change', persist);
        $id('zdf-openedit').addEventListener('change', persist);
        $id('zdf-paste').addEventListener('change', persist);

        // ── Collapse / hide / drag ───────────────────────────────────────
        $id('zdf-min').addEventListener('click', () => panel.classList.toggle('zdf-collapsed'));
        $id('zdf-close').addEventListener('click', () => {
            panel.style.display = 'none';
            fab.style.display = 'block';
        });
        fab.addEventListener('click', () => {
            panel.style.display = '';
            fab.style.display = 'none';
        });
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.shiftKey && (e.key === 'D' || e.key === 'd')) {
                e.preventDefault();
                const hidden = panel.style.display === 'none';
                panel.style.display = hidden ? '' : 'none';
                fab.style.display = hidden ? 'none' : 'block';
            }
        });

        (function makeDraggable() {
            const head = $id('zdf-head');
            let dragging = false, ox = 0, oy = 0;
            head.addEventListener('mousedown', (e) => {
                if (e.target.tagName === 'BUTTON') return;
                dragging = true;
                const r = panel.getBoundingClientRect();
                ox = e.clientX - r.left; oy = e.clientY - r.top;
                e.preventDefault();
            });
            document.addEventListener('mousemove', (e) => {
                if (!dragging) return;
                panel.style.left = Math.max(0, e.clientX - ox) + 'px';
                panel.style.top = Math.max(0, e.clientY - oy) + 'px';
                panel.style.right = 'auto';
            });
            document.addEventListener('mouseup', () => { dragging = false; });
        })();

        const rows = indexRows();
        log(`Timelogs table detected — ${rows.size} dated row(s).`, 'ok');
        status(`${rows.size} row(s) on page. Paste your logs and click Parse.`);

        // A workbook loaded on an earlier employee's page is still available here.
        const stored = readStore(ROSTER_KEY, null);
        if (stored && stored.employees && stored.employees.length) {
            roster = stored;
            const doneCount = roster.employees.filter(isDone).length;
            log(`Workbook "${roster.source}" still loaded — ${roster.employees.length} employees`
                + `${doneCount ? `, ${doneCount} already filled` : ''}.`, 'ok');
            renderRoster(true);
        }
        if (typeof XLSX === 'undefined') {
            log('Excel support unavailable (the library did not load) — CSV files and pasting still work.', 'warn');
        }
    }

    // ── Only show up on a page that actually has the timelogs table ──────
    function looksLikeTimelogsPage() {
        const rows = indexRows();
        if (rows.size < 2) return false;
        for (const tr of rows.values()) {
            const header = headerFor(tr);
            const hits = header.filter((h) => HEADER_ALIASES[h]).length;
            if (hits >= 2) return true;
        }
        return false;
    }

    let attempts = 0;
    (function waitForTable() {
        if (looksLikeTimelogsPage()) { init(); return; }
        if (attempts++ > 60) return;                 // give up after ~30s
        setTimeout(waitForTable, 500);
    })();

    // The table reloads after "Search" — re-check so the panel appears then too.
    const mo = new MutationObserver(() => {
        if (!window.__zdfInited && looksLikeTimelogsPage() && !document.getElementById('zdf-panel')) {
            window.__zdfInited = true;
            init();
        }
    });
    if (document.body) mo.observe(document.body, { childList: true, subtree: true });
})();
