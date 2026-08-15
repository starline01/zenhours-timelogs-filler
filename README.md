# Zenhours DTR Filler

A Tampermonkey userscript that fills the Zenhours timelogs table from a block of
timelogs pasted out of Excel. It finds each date's row, clicks **Edit**, and types
the times into the six time fields.

**It never clicks Save.** You review the highlighted values and save each row yourself.

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
5. **Test 1st day** — fills only the first row, so you can sanity-check one row.
6. **Fill all rows** — fills every matched row and highlights the fields in green.
7. Review, then click **Save** on each row.

**Undo fill** puts every field the script touched back to its original value
(works while the rows are still open in edit mode).

---

## Loading a whole workbook (all employees)

Editing in Zenoras is per employee, so a file covering everyone gets filtered
down to whoever the current page belongs to.

**`DTR Template.xlsx` in this folder is a ready-made blank** — fill in the
*Timelogs* tab and upload it. Its *READ ME* tab carries the column reference and
a worked example; sheets named READ ME / Instructions / Notes / Guide are skipped
by the importer, so documentation can live in the same workbook.

1. Click **Load Excel / CSV…** and pick the file. Expected layout — one sheet,
   one row per employee per day:

   | EmpID | Name | Date | Day | In | LOut | LIn | BOut | BIn | Out |
   |---|---|---|---|---|---|---|---|---|---|
   | 121914 | Dela Cruz, Juan | 8/1/2026 | TUESDAY | 10:20 | 12:40 | 13:12 | 17:00 | 17:30 | 21:00 |

   Column names are matched by header text, so order doesn't matter and extra
   columns are ignored. Merged/blank ID cells inherit from the row above.

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
| Clear the field when my cell is blank/dash | off | On = empties the field instead of leaving `12:00 AM` |

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

`test/` holds a mock of the timelogs page that mirrors the real one (blank
`--:--` grid, Edit → six inputs prefilled to `MM/DD/YYYY 12:00 AM`, Save/Cancel).
Use it to try changes without touching live client data:

```bash
node "test/serve.js"
```

Then open <http://localhost:8731/test/zenhours-mock.html>. The mock loads the
userscript directly, so the panel behaves exactly as it does on Zenhours.
