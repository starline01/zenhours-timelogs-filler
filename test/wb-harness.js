// Runs the userscript's workbook readers over a real .xlsx under Node.
//
//   node test/read-workbooks.test.js
//
// SheetJS is fetched once into test/vendor/ (gitignored) — the same build the
// userscript @requires in the browser, so the reader under test is the shipped
// one, not a stand-in.
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const HERE = __dirname;
const SRC = process.env.ZDF_SRC || path.join(HERE, '..', 'zenhours-dtr-filler.user.js');
const VENDOR = path.join(HERE, 'vendor');
const SHEETJS = path.join(VENDOR, 'xlsx.full.min.js');
const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';

if (!fs.existsSync(SHEETJS)) {
    fs.mkdirSync(VENDOR, { recursive: true });
    process.stderr.write('fetching SheetJS (once) … ');
    execFileSync('curl', ['-sL', SHEETJS_URL, '-o', SHEETJS]);
    process.stderr.write('done\n');
}

// Load SheetJS in THIS realm. Inside a vm sandbox it gets its own Date
// constructor, and `cell.v instanceof Date` then fails across the realm
// boundary — which looks exactly like a parser bug and is not one.
const XLSX = require(SHEETJS);
if (!XLSX || !XLSX.read) throw new Error('SheetJS did not load');

const src = fs.readFileSync(SRC, 'utf8').split('\r\n').join('\n');
const slice = (a, b) => {
    const i = src.indexOf(a), j = src.indexOf(b, i);
    if (i < 0 || j < 0) throw new Error('marker not found: ' + (i < 0 ? a : b));
    return src.slice(i, j);
};

const ROSTER = '    // =====================================================================\n    //  EMPLOYEE ROSTER';
const PANEL = '    // =====================================================================\n    //  PANEL UI';

// The roster section persists to localStorage and reads the live table for a
// fallback year; both are stubbed so the readers themselves run unchanged.
const localStorage = {
    _d: {},
    getItem(k) { return Object.prototype.hasOwnProperty.call(this._d, k) ? this._d[k] : null; },
    setItem(k, v) { this._d[k] = String(v); },
    removeItem(k) { delete this._d[k]; }
};
const indexRows = () => new Map();

const code =
    slice('const COLUMNS = ', 'function loadSettings') +
    slice('    function splitFields', '    //  READING THE PAGE') +
    slice(ROSTER, PANEL) +
    '\nreturn { buildRoster, classifySheet, employeeToText, parseTime, parseDate, applyOvernightRoll };';

const api = new Function('XLSX', 'localStorage', 'indexRows', code)(XLSX, localStorage, indexRows);

function sheetsOf(file) {
    const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true });
    return wb.SheetNames.map((name) => ({ name, ws: wb.Sheets[name] }));
}

module.exports = { XLSX, api, sheetsOf, fixture: (n) => path.join(HERE, 'fixtures', n) };
