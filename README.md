# Zenhours DTR Filler

A Tampermonkey userscript that fills the Zenhours timelogs table from a block of
timelogs pasted out of Excel. It finds each date's row, clicks **Edit**, and types
the times into the six time fields.

**It saves each row as it fills it.** Use **Test 1st day** to check a new client
file before committing anything — that one fills a single row and saves nothing.

---

## Install (each PC, once)

1. Install [Tampermonkey](https://www.tampermonkey.net/) in your browser.
2. Open this link — Tampermonkey shows its install page:

   <https://raw.githubusercontent.com/starline01/zenhours-timelogs-filler/main/zenhours-dtr-filler.user.js>

3. Click **Install**. Done — that PC now auto-updates from this repo.

> Installing from the raw link is what wires up auto-update. Pasting the code
> into a new script by hand does **not** — that copy is local and will never
> update, even though it otherwise behaves identically.

### Pushing a change to every PC

Tampermonkey only updates when the version number goes **up**. Editing the code
alone changes nothing on the other machines — which is why there's a script for it:

```powershell
.\release.ps1 -Message "fix the lunch column"
```

That bumps `@version`, syntax-checks the file, commits and pushes. If the syntax
check fails it reverts the bump and pushes nothing. Use `-Version 1.3.0` to set an
exact number instead of bumping the patch digit.

Doing it by hand is the same three steps: make the change, bump `@version` on
line 4, commit and push to `main`.

Each PC picks it up on its next check — Tampermonkey's default is roughly daily.
To pull it immediately on a given PC: Tampermonkey dashboard → **Installed
userscripts** → *Check for userscript updates*. Note that `raw.githubusercontent.com`
caches for about five minutes, so a push isn't visible the same second.

Confirm what a PC is running by checking the version in its Tampermonkey dashboard.

### Manual install (fallback)

If a PC can't reach GitHub: dashboard → **+** → delete the template → paste the
whole of `zenhours-dtr-filler.user.js` → Ctrl+S. That copy won't auto-update.

The script is scoped to **every `zenoras.com` subdomain**:

```
// @match        *://*.zenoras.com/*
// @match        *://zenoras.com/*
```

That covers `rr.zenoras.com` and any other subdomain without further edits.
`@match` lines are OR'd — add another line for any site on a different domain.

It only shows the panel once it detects a timelogs table, so it stays out of the
way everywhere else. To restrict it to timelogs pages only, swap both lines for:

```
// @match        *://*.zenoras.com/hr/manage_timelogs/*
```

---

## Use

1. Open the timelogs page — e.g. `rr.zenoras.com/hr/manage_timelogs/employee/121914`
   — and **Search** your date range.
2. The **DTR Filler** panel appears top-right. `Ctrl+Shift+D` hides/shows it.
3. Copy your rows from Excel and paste into the box:

   ```
   8/1/2026	TUESDAY	10:20	12:40	13:12	17:00	17:30	21:00
   8/2/2026	WEDNESDAY	10:25	12:45	13:15	16:39	17:09	21:00
   ```

4. **Parse** — shows what was read and which dates exist on the page.
   A ✗ means that date is not in the current search range.
5. **Test 1st day** — fills one row and saves nothing, so you can check a new
   client file before anything is committed.
6. **Fill & Save all** — fills every matched row and saves it.

**Undo fill** puts every field the script touched back to its original value. It
only works while a row is still open, so it applies to a test run rather than to
rows already saved.

### Saving

Every row is committed to Zenoras as it is filled. What that means in practice:

- **Blanks are cleared, not left at `12:00 AM`.** Untouched fields keep Zenoras'
  prefill, and saving that records a real midnight punch on a guard who simply
  had no lunch break — wrong hours, with no review step left to catch it.
- **A row is only saved if every value meant for it went in.** If a column had
  nowhere to go — an unreadable OCR cell, an unparseable time — the row is left
  open and listed under `not saved` for you to finish by hand.
- **Each save is confirmed** by waiting for the row to leave edit mode. A Save
  that never completes is reported, not counted.
- **Undo stops working** for rows already committed. The log warns before the
  run starts.
- **Test 1st day never saves** — it exists to let you check one row, and a whole
  client file, before committing anything.

### Copying straight off the Zenoras page

Select a stretch of the timelogs grid, copy, and paste it in. It copies as one
day per block — the date alone on its line, then the schedule, then one punch
per line — and that layout is read directly:

```
2026-09-06

9:45 AM to 5:45 PM      09/06/2026 09:37 AM
09/06/2026 11:47 AM
09/06/2026 12:16 PM
09/06/2026 03:52 PM
--:--
09/06/2026 04:21 PM
```

Two things this format gives you that a spreadsheet does not:

- **`--:--` is a gap, not a blank.** It means that punch was never recorded, so
  the row goes to the gap grid below instead of being filled with a hole in it.
- **Each punch carries its own date**, so a night shift needs no guessing: a
  Time Out stamped `09/11/2026` on a `2026-09-10` row is written to the 11th.
  The overnight roll only runs on blocks the page gave bare times for.

The schedule line (`9:45 AM to 5:45 PM`) is recognised and ignored — it is a
shift pattern, not a punch. So are the row's own controls: selecting rows on the
page drags along their **Edit** link (and **Save** / **Cancel** / **No Schedule**
if they are showing). Leave them in — they are dropped, and they never take a
column from a real punch.

The rule behind that: a value with no digit in it was never a time. Known page
words go quietly; anything else unexpected is dropped too but named in the log,
so stray text is never swallowed silently. The one exception is `--:--`, which
has no digits either and must keep its column — it is the gap itself.

The page also copies a whole day onto **one tab-separated line**, and that is
read the same way:

```
2026-09-01   09:00-17:00   --:--  --:--  --:--  --:--  --:--  --:--   Edit
```

A day with **no punches at all** is still a row — it goes to the grid with all
six cells open, rather than being discarded as an empty result. Both schedule
notations are recognised (`09:00-17:00` and `9:00 AM to 5:00 PM`).

This shape is told apart from a spreadsheet row by the page's own marks: an
**Edit** link, a schedule range, or the literal `--:--`. A spreadsheet row
carries none of those, so it keeps its own meaning — there, a bare `-` still
says *clear this column*, not *this is missing*.

### Filling the gaps

Any row with a `--:--` in it appears in an editable grid, with the missing cells
outlined. Type the real times, then use the grid's own button:

```
        In      L.Out   L.In    B.Out   B.In    Out
09-06   09:37   11:47   12:16   15:52   [    ]  16:21
09-04   09:33   11:46   12:15   15:32   [    ]  16:01
```

- **Every cell is editable**, not just the gaps — correct a wrong punch the same
  way you fill a missing one. `16:05`, `3:45 PM` and `1615H` are all accepted.
- **Filling from the grid overwrites** what those rows currently show on the
  page. That is the point of it: these are rows you have just corrected. Leave a
  cell empty to clear that column instead.
- **A value it cannot read stops the whole run.** The cell is outlined in red
  and nothing is written — a typo in a payroll field is not worth a partial fill.
- **Complete rows are not put in the grid** and are not filled by its button.
  They go through **Fill & Save all** as usual, so the two never collide.
- **Rows with gaps are held back from Fill & Save all**, which names them and
  fills the rest. Filling one there would clear the gap column and commit the
  row, spending the one chance to enter what is actually missing.

### Edit + Save (no edits)

Clicks **Edit** then **Save** on every row that already has all six times,
changing nothing — the same thing you would do by hand to commit a row.

**It only touches complete rows, and that is the whole point.** Zenoras prefills
an empty field with the row's date at 12:00 AM, so clicking Edit then Save on a
row with any blank column records a midnight punch that reads like a real one.
A row missing even one column is skipped, and the log names the columns that
stopped it:

```
✓ 2026-08-01 — re-saved, values unchanged
· 2026-08-04 — skipped, Lunch Out, Lunch In, Break Out, Break In still blank
3 row(s) re-saved, 11 skipped as incomplete.
```

It reads the page and nothing else — no paste box, no loaded workbook, no
employee selection. A row already open for editing is left alone, because once
inputs replace the display text there is no way to tell a real value from the
prefill. Each save is confirmed the same way as a fill run: by waiting for the
row to leave edit mode.

There is nothing to undo, since the values written back are the ones already
there.

---

## Loading a whole workbook (all employees)

Editing in Zenoras is per employee, so a file covering everyone gets filtered
down to whoever the current page belongs to.

**`DTR Template.xlsx` in this folder is a ready-made blank** — fill in the
*Timelogs* tab and upload it. Its *READ ME* tab carries the column reference and
a worked example; sheets named READ ME / Instructions / Notes / Guide are skipped
by the importer, so documentation can live in the same workbook.

1. Click **Load Excel, CSV or a scan…** and pick the file. **You do not have to reformat
   a client's DTR first** — each sheet is sniffed and read according to its own
   shape (see *Messy client DTRs* below).

2. The script reads the **access ID and name displayed on the page**, finds that
   guard in the file, and loads only their days. You'll see:

   ```
   Matched this page by access ID 121914 + name "dela cruz juan" → Dela Cruz, Juan (14 days loaded)
   ```

3. Fill and save as normal, then open the next employee's page. **The workbook
   stays loaded** — no re-upload. Guards you've filled are ticked (`✓`) in the
   list, and the log tells you who's next.

### The safety rule

If the page can't be matched to exactly one guard — not in the file, ambiguous
name, or a tie — then **nothing is selected, the box is cleared, and filling is
blocked** until you pick from the list. The script will not guess which employee
a page belongs to, because writing one guard's hours onto another is expensive to
find and undo.

For the same reason, the paste box is never carried from one employee's page to
the next while a workbook is loaded.

### Messy client DTRs

Every sheet is classified before it's read, and the panel logs what it decided
(`layout — Report: personnel report`). Eight shapes are recognised:

| Layout | Recognised by | Identity comes from |
|---|---|---|
| **column headers** | a header row naming Date + times **and** the guard | that row |
| **personnel report** | `Personnel Name` / `DTR Summary Report`, `Time In 1..3` | the row |
| **per-guard blocks** | `ACTUAL TIME LOGS`, `NAME:`, `SECURITY GUARD:` above each small table | the text above, or a bare number past the table |
| **day-number blocks** | `INNITIAL IN`, `L.B OUT`, `C.B OUT`, dates as 1–31 | name + month/year above |
| **headerless columns** | name · date · weekday · six punches, no header at all | column A |
| **biometric export** | repeated `Time In` / `Time Out` pairs, `Enroll No` | the row |
| **device scan log** | one row per punch with a full timestamp, several per day | the row |
| **day-across roster grid** | day numbers along the top, `IN`/`OUT` pairs beneath, a `SHIFT` column | the row |

A sheet of `SCHEDULE_START_DATE` / `ACTUAL SCHEDULE OF GUARDS` is recognised as
**planned shifts, not punches**, and refused rather than imported.

#### Access IDs with nothing labelling them

Per-guard blocks normally name the ID (`ACCESS ID: 166166`). Some sheets just
park the number in a column past the last header — on the guard's name row, or
on the header row itself — with no label at all. That is now picked up, which
matters because the access ID is the reliable way to tell which guard's page you
are on; without it, matching falls back to the name alone.

To avoid inventing an ID, only a plain whole number of four or more digits is
taken, only from a column the table does not use, and never from a row that is
itself a day of data. A bare year (`2026`) is ignored.

**A sheet holding both is read correctly.** Many client DTRs put the planned
roster at the top and the real punches (`ACTUAL TIME LOGS`) underneath, in the
same sheet. The punches win; the schedule half is skipped with a note. Where the
sheet names its store and cut-off, both appear in the log so an old period is
obvious before you fill anything:

```
layout — Sheet1: per-guard blocks — RSC MAGNOLIA — AUGUST 16-31, 2026
```

Rest markers spelled one letter per cell — `D | A | Y | O | F | F`, as these
sheets often do — are rebuilt into the real reason (`dayoff`, `leaved`,
`absent`) rather than guessed at.

**A day-across roster grid** turns the table sideways: one row per guard, one
column PAIR per day (`IN`/`OUT`), day numbers along the top, and a `SHIFT`
column reading `DS` or `NS`. Only two punches a day, so lunch and break stay
empty. `X` becomes *no duty*; any other marker (a posting such as `BRICKSTONE`)
is kept as written. A row marked `NS` whose Time Out is not after its Time In is
dated to the next morning even when the gap is too small to roll on its own.

Such sheets often name only the day number, so the month comes from the sheet
title, then the **file name**, then the cut-off already open on the page — the
night sheet in one real workbook misspells August as "Auust", and the file name
carried it.

**A raw device scan log has no time columns at all.** A fingerprint terminal
exports every punch as its own row with a full timestamp — six rows make one
working day, and the device often logs the same scan twice. Those are grouped by
guard and date, de-duplicated, sorted, and read in order: first scan is Time In,
last is Time Out. A day with an odd number of scans cannot be paired into breaks,
so only Time In and Time Out are set and the row says how many scans it saw.

**An unlabelled date column is found by its contents.** Real DTRs routinely
leave that header blank — naming the weekday (`Transaction Day`) while the date
sits in a blank-headed column beside it. Read by header alone, every row in such
a sheet is discarded. Columns of Excel *time* values are excluded from that
search, since those are Dates too (Excel puts them on 1899-12-31).

**Biometric exports** label every punch column just `Time In` / `Time Out` and
repeat the pair once per break, so six columns carry only two distinct names.
Those are read positionally — in, out, in, out, in, out — so the day's Time Out
is the *last* one, not the first. Read by name instead, a 10:05 PM Time Out came
back as the 12:03 PM lunch break and the breaks vanished entirely.

**Real Excel time cells are read in UTC**, the way SheetJS builds them. Read with
local getters instead, every such time shifts by the machine's timezone offset —
a `20:00` Time In came back as `04:00` at UTC+8. Files whose times are text
(`0737H`) were never affected, which is why this hid for so long.

Messy cells are handled too: military integers (`1058`), military strings
(`1053H`), Excel serials, real Date cells, `22;20` typed with a semicolon,
`0UT` mis-typed for `OUT` in a header, and rest markers (`REST`, `NO DUTY`,
`DAY OFF`, `RD`, `LEAVE`, `ABSENT`) which are counted and skipped rather than
filled.

**One guard split across two sheets gets merged.** If a report lists them by
Access ID and a raw sheet lists them by name only, both sets of days end up on
one person — but only when exactly one name matches, never on a tie. If the same
day appears twice, the row with more punches wins and is flagged.

### Overnight shifts

A punch is only rolled when the result stays within a plausible shift length,
and the cap depends on **where** the punch sits:

- a punch in the **middle** of a row may stretch the shift to 18h
- the **final** punch may stretch it to 24h

That distinction matters. One mistyped punch — `15:07` keyed as `05:07` — rolls
itself *and* drags every punch after it onto the next day, turning a 12-hour
shift into a 36-hour one; the mid-row cap stops that, leaving the punch on its
own date flagged `out-of-order punch`. But a guard on a double really can clock
out 22 hours after clocking in, and that is always the *last* punch, so the
looser cap lets a genuine overnight through.

A negative span is reported as `Time Out is before Time In` rather than as
"-1.9 hours".

A punch that falls *before* the one preceding it by more than six hours is taken
to have crossed midnight, and is written to the **next day's date**. So a
19:00 → 06:00 night shift fills as:

```
Time In    08/01/2026 07:00 PM
Lunch Out  08/01/2026 11:00 PM
Break Out  08/02/2026 02:00 AM   ← rolled
Time Out   08/02/2026 06:00 AM   ← rolled
```

In the paste box these read `02:00+1`, so the offset stays visible and editable,
and the run summary tells you how many punches were re-dated. Rows the importer
distrusts — overnight, missing in/out, an odd number of punches, a span under 1h
or over 16h, or hours that disagree with the sheet's own total by more than 1.5h
— are listed per guard when you select them.

### Scanned sheets (OCR)

Load a **.png / .jpg / .webp / .tif** of a printed DTR and it is read on this PC —
Tesseract runs in the browser, the image is never uploaded anywhere.

The result goes into the paste box for you to check, and **nothing is filled
automatically**. OCR confuses `0/8`, `1/7` and `5/6`, and a DTR is nothing but
digits, so every value is a draft until you have compared it to the sheet.

What protects you:

- **Columns come from the values, not the headings.** The row with the most
  punches defines the grid; headings only name the columns. OCR often welds
  headings together (`LUNCHIN`, `OUBREAK`) — that no longer matters, and a
  missing punch leaves a gap in the right column instead of shifting the row.
  The log prints the order it read: `Time In | Lunch Out | Lunch In | …`
- **Unreadable cells become `??:??`**, which the filler refuses. They are never
  guessed, and the log names each one.
- **A sheet below 80% confidence is refused outright.** A clean printed scan
  reads at 90%+; a photographed handwritten form lands near 60–70% and produces
  values that look perfectly valid and are wrong. Refusing costs you a retype;
  accepting corrupts a guard's pay.
- **Dates must already be on the page.** A row whose date isn't in the range you
  searched is dropped, since a misread date is as likely as a misread time.

**Handwritten DTRs will be refused**, and no scan quality changes that — this
engine reads print. A handwritten punch card photographed straight, cropped and
squared still comes back around 40% confidence, well under the gate, with output
like `[29s` for `1233`. The refusal message says so explicitly rather than
suggesting a rescan that cannot help.

Scanning tips: flat on the glass, 300dpi, straight (not a phone photo at an
angle). PDFs aren't read directly yet — export the page as PNG first.

### Excel support

`.xlsx`, `.xls`, and `.csv` all work. Excel parsing uses SheetJS, pulled in by the
`@require` line at the top of the script and cached by Tampermonkey after first
load. If your network blocks it, the panel says so — CSV files and pasting keep
working without it, so saving the sheet as CSV is the fallback.

---

## Paste format

| Position | Column | Notes |
|---|---|---|
| 1 | Date | `8/1/2026`, `2026-08-01`, `Aug 1`, `1-Aug-26` all work |
| 2 | *(day name)* | `TUESDAY` etc. — detected and ignored. Optional. |
| 3–8 | Time In, Lunch Out, Lunch In, Break Out, Break In, Time Out | |

- **Times** may be 24-hour (`13:12`), 12-hour (`1:12 PM`), or military (`1312`).
  They are written in whatever format the Zenhours field is already showing.
- **Guards with no lunch or break:** leave the cell empty or put a dash (`-`, `--:--`).
  The column stays aligned and the script leaves that field alone — the log warns
  you that it still holds Zenhours' prefilled `12:00 AM`. Tick
  **"Clear the field when my cell is blank/dash"** if you want those emptied instead.
- **Only two times per day?** `date, day, time in, time out` works too.
- **A header row** in your paste (`Date, Time In, ...`) is detected and used to map
  columns by name, so a different column order still works.

### Options

| Option | Default | Effect |
|---|---|---|
| Only fill columns that are blank (`--:--`) | on | Never overwrites a time already saved in Zenhours |
| Click Edit automatically | on | Off = only fills rows you already opened yourself |


Columns with no punch are always cleared rather than left at `12:00 AM`, and every
row is saved as it is filled — neither is optional, because saving a prefilled
`12:00 AM` would record a midnight punch that never happened.

---

## Checking the column mapping

The first line of every run's log tells you how the six fields were matched:

```
Columns matched by header labels (offset +1): Time In, Lunch Out, Lunch In, Break Out, Break In, Time Out
```

All six listed, in that order, means the mapping is right. The script tries three
strategies and validates each before accepting it, so a partial match can never
shift columns:

1. **input attributes** — the field names its own column (`name="time_in"`).
2. **header labels** — matched against the header row, corrected for any
   header/body cell offset. Zenoras' body rows carry an extra leading cell the
   header does not have, so you will normally see `(offset +1)` here.
3. **position** — the row's time inputs in left-to-right column order.

If a column is missing from that line, send me the **Copy diagnostics** output.

---

## If a row will not fill

Click **Copy diagnostics** and send me the result. It captures the table's header
names, whether the Edit control was found, the input types and values, and one
sample row's HTML — which is what I need to adjust the column mapping.

---

## Testing

`test/make-fixtures.py` regenerates the sample workbooks under `test/fixtures/`.
They mirror the *shape* of real client DTRs — schedule above time logs, letter-
per-cell day-offs, night shifts, a mistyped punch, unlabelled access IDs — with
invented names, so the layouts stay covered without client data in the repo:

```bash
python test/make-fixtures.py
```

Real client files dropped into `test/` are git-ignored.

The workbook readers run against those fixtures under Node:

```bash
node test/read-workbooks.test.js
```

It loads the shipped userscript and the same SheetJS build the browser uses
(fetched once into the git-ignored `test/vendor/`), so what is tested is the
reader that ships, not a stand-in.

`test/` holds a mock of the timelogs page that mirrors the real one (blank
`--:--` grid, Edit → six inputs prefilled to `MM/DD/YYYY 12:00 AM`, Save/Cancel).
Use it to try changes without touching live client data:

```bash
node "test/serve.js"
```

Then open <http://localhost:8731/test/zenhours-mock.html>. The mock loads the
userscript directly, so the panel behaves exactly as it does on Zenhours.
