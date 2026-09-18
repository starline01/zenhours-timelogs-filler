// Workbook layouts, checked against fixtures that mirror real client DTRs.
//
//   node test/read-workbooks.test.js
//
// Fixtures carry invented names and times (see make-fixtures.py); no client
// data lives in this repo.
const { api, sheetsOf, fixture } = require('./wb-harness');

let pass = 0, fail = 0;
function check(label, actual, expected) {
    const a = JSON.stringify(actual), e = JSON.stringify(expected);
    if (a === e) { pass++; console.log('  ok   ' + label); }
    else { fail++; console.log('  FAIL ' + label + '\n         got      ' + a + '\n         expected ' + e); }
}
const p2 = (n) => String(n).padStart(2, '0');
const COLS = ['time_in', 'lunch_out', 'lunch_in', 'break_out', 'break_in', 'time_out'];
const row = (r) => COLS.map((c) => (r.times[c]
    ? `${p2(r.times[c].h)}:${p2(r.times[c].m)}` + (r.times[c].plus ? '+' + r.times[c].plus : '')
    : '--')).join(' ');

// ── Per-guard blocks with the access ID printed bare ─────────────────────
// Shape of "ACTUAL SCHED AND TIME LOGS": one block per guard, each headed by
// NAME and a header row, the access ID sitting past the last header column
// with nothing labelling it, and the SCHEDULE column labelled in the first
// block but left blank in the others.
console.log('\n== Per-guard blocks, access ID printed bare ==');
const blocks = api.buildRoster(sheetsOf(fixture('bare-id-blocks.xlsx')), 'bare-id-blocks.xlsx');
const by = (n) => blocks.employees.find((e) => (e.name || '').indexOf(n) === 0);
const day = (emp, d) => emp.rows.find((r) => r.date === d);

check('one entry per guard', blocks.employees.length, 3);
check('5 worked days, 1 rest', [blocks.totalRows, blocks.restRows], [5, 1]);
check('read without complaint', blocks.warnings, []);

check('ID taken from the name row', by('GUARD ALPHA').id, '166166');
check('ID taken from the header row', by('GUARD BRAVO').id, '153254');
check('ID found for every block', by('GUARD CHARLIE').id, '162269');

check('punches land in their own columns',
    row(day(by('GUARD ALPHA'), '2026-09-02')), '10:05 13:25 13:55 18:15 18:45 22:30');
check('the schedule column is not read as a punch',
    row(day(by('GUARD BRAVO'), '2026-09-01')), '09:20 12:15 12:45 17:20 17:50 21:30');
check('nor is "no. of hours rendered"', day(by('GUARD ALPHA'), '2026-09-02').times.time_out, { h: 22, m: 30, plus: 0 });

check('a finish after midnight rolls to the next day',
    row(day(by('GUARD ALPHA'), '2026-09-01')), '10:10 13:15 13:45 18:45 19:15 01:00+1');
check('and does so from a late break too',
    row(day(by('GUARD CHARLIE'), '2026-09-01')), '10:24 14:10 14:40 18:45 19:15 02:00+1');

check('a day with nothing recorded is rest', day(by('GUARD ALPHA'), '2026-09-03').rest, true);
check('and carries no times', row(day(by('GUARD ALPHA'), '2026-09-03')), '-- -- -- -- -- --');

// ── A quote typed where a colon was meant ────────────────────────────────
// ; and : share a key, as do ' and ". A client sheet had 18"45 for 18:45,
// which otherwise dropped out of the row entirely.
console.log('\n== A quote typed where a colon was meant ==');
check('the fixture cell reads 18"45', day(by('GUARD CHARLIE'), '2026-09-01').times.break_out, { h: 18, m: 45, plus: 0 });
check('straight quote', api.parseTime('18"45'), { h: 18, m: 45, plus: 0 });
check('apostrophe', api.parseTime("18'45"), { h: 18, m: 45, plus: 0 });
check('curly quote', api.parseTime('18”45'), { h: 18, m: 45, plus: 0 });
check('semicolon still works', api.parseTime('22;20'), { h: 22, m: 20, plus: 0 });
check('text is still refused', api.parseTime('DAYOFF'), null);
check('so is a nonsense time', api.parseTime('99"99'), null);

// ── The other layouts still read ─────────────────────────────────────────
console.log('\n== Other layouts are unaffected ==');
const stacked = api.buildRoster(sheetsOf(fixture('stacked-with-schedule.xlsx')), 'stacked-with-schedule.xlsx');
check('punches win over a schedule banner', stacked.totalRows > 0, true);
const matrix = api.buildRoster(sheetsOf(fixture('roster-matrix.xlsx')), 'roster-matrix.xlsx');
check('day-across roster grid still reads', matrix.employees.length > 0, true);

console.log(`\n${pass} passed, ${fail} failed\n`);
process.exit(fail ? 1 : 0);
