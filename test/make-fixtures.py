"""Regenerate the test workbooks.

    python test/make-fixtures.py

Fixtures mirror the SHAPE of real client DTRs with invented names and times, so
the layouts stay covered without any client data in the repo.
"""
from openpyxl import Workbook
import os

HERE = os.path.dirname(os.path.abspath(__file__))


def stacked_with_schedule(path):
    """A planned roster ABOVE the real punches, in one sheet.

    This is the shape that was silently rejected: the 'ACTUAL SCHEDULE OF
    GUARDS' banner made the whole sheet look like a schedule, hiding the
    'ACTUAL TIME LOGS' blocks below it.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    ws['C3'] = 'STARLINE SECURITY AGENCY '
    ws['B4'] = 'ACTUAL SCHEDULE OF GUARDS'
    ws['B5'] = 'STORE: TEST BRANCH'
    ws['B6'] = 'CUT OFF: '
    ws['C6'] = 'AUGUST 16-31, 2026'

    ws['B8'], ws['C8'], ws['D8'], ws['E8'] = 'NO.', 'NAME ', 'ACCESS ID ', 'POSITION '
    for i, day in enumerate(range(16, 32)):
        ws.cell(row=8, column=6 + i, value=day)

    guards = [('DAY GUARD ONE', '0800H-2000H'), ('NIGHT GUARD TWO', '2230H-1030H')]
    for gi, (name, shift) in enumerate(guards):
        ws.cell(row=10 + gi, column=2, value=gi + 1)
        ws.cell(row=10 + gi, column=3, value=name)
        ws.cell(row=10 + gi, column=5, value='GUARD')
        for i in range(16):
            ws.cell(row=10 + gi, column=6 + i, value='DAYOFF' if i == 1 else shift)

    # ── ACTUAL TIME LOGS: one stacked block per guard ────────────────────
    ws['B31'] = 'STARLINE SECURITY AGENCY  '
    ws['B33'] = 'ACTUAL TIME LOGS '

    HEAD = ['DATE ', 'TIME IN ', ' LUNCH OUT ', 'LUNCH IN ', 'BREAK OUT ',
            'BREAK IN ', 'TIME OUT ', ' NO. OF HOURS RENDERED', 'REMARKS ']

    def block(top, name, rows):
        ws.cell(row=top, column=2, value='NAME:')
        ws.cell(row=top, column=3, value=name)
        for i, h in enumerate(HEAD):
            ws.cell(row=top + 1, column=2 + i, value=h)
        for j, cells in enumerate(rows):
            r = top + 2 + j
            for i, v in enumerate(cells):
                if v is not None:
                    ws.cell(row=r, column=2 + i, value=v)

    # Day guard: normal shifts, a letter-per-cell DAYOFF, and one mistyped
    # punch (05:07 for 15:07) that must NOT drag the rest of the row.
    day_rows = [
        [' 8/16/2026', '0737H', '1112H', '1141H', '1502H', '1529H', '2013H', 12],
        [' 8/17/2026', 'D', 'A', 'Y', 'O', 'F', 'F'],
        [' 8/18/2026', '0958H', '1207H', '1237H', '0507H', '1537H', '2205H', 12],
    ]
    block(35, 'DAY GUARD ONE', day_rows)

    # Night guard: 22xx in, 11xx out next day, with the banner text that real
    # sheets stretch across the lunch/break columns.
    BANNER = ' ' * 60 + 'N    I    G    H    T       S    H    I    F   T'
    night_rows = [
        [' 8/16/2026', '2226H', BANNER, None, None, None, '1132H', 13, '1 HR OT'],
        [' 8/17/2026', '2229H', BANNER, None, None, None, '1135H', 13, '1 HR OT'],
        [' 8/18/2026', 'D', 'A', 'Y', 'O', 'F', 'F'],
    ]
    block(60, 'NIGHT GUARD TWO', night_rows)
    ws.cell(row=66, column=6, value='TOTAL NO. OF HOURS RENDERED')
    ws.cell(row=66, column=9, value=26)

    wb.save(path)
    return path


def _main():
    print('wrote', stacked_with_schedule(os.path.join(HERE, 'fixtures', 'stacked-with-schedule.xlsx')))
    print('wrote', roster_matrix(os.path.join(HERE, 'fixtures', 'roster-matrix.xlsx')))


def roster_matrix(path):
    """Wide roster grid: guards down, days across, IN/OUT per day.

    Mirrors a detachment roster — a DS/NS shift column, 'X' and posting
    markers, names in a merged column beside the row numbers, and a second
    sheet whose title misspells the month so it has to come from elsewhere.
    """
    from openpyxl import Workbook
    import datetime
    wb = Workbook()

    def sheet(ws, title, staff):
        ws['A1'] = title
        ws['A3'], ws['C3'] = 'NAME OF PERSONNEL', 'SHIFT'
        for i, day in enumerate(range(16, 21)):
            ws.cell(row=3, column=4 + i * 2, value=day)
            ws.cell(row=4, column=4 + i * 2, value='IN')
            ws.cell(row=4, column=5 + i * 2, value='OUT')
        for gi, (name, shift, days) in enumerate(staff):
            r = 5 + gi * 2
            ws.cell(row=r, column=1, value=gi + 1)      # row number, NOT the name
            ws.cell(row=r, column=2, value=name)
            ws.cell(row=r, column=3, value=shift)
            for i, pair in enumerate(days):
                a, b = pair
                ws.cell(row=r, column=4 + i * 2, value=a)
                ws.cell(row=r, column=5 + i * 2, value=b)

    t = datetime.time
    ws1 = wb.active
    ws1.title = '0930H-2130H'
    sheet(ws1, 'TEST MALL: August 16-31, 2026    DAY SHIFT', [
        ('DAY GUARD ONE', 'DS', [(t(8, 32), t(21, 40)), (t(8, 38), t(21, 44)),
                                 ('X', 'X'), ('BRICKSTONE ', 'BRICKSTONE '), (t(8, 25), t(21, 34))]),
        ('NIGHT ON DAYSHEET', 'NS', [(t(20, 10), t(9, 29)), (t(20, 19), t(9, 34)),
                                     (t(20, 36), t(9, 34)), ('X', 'X'), ('X', 'X')]),
    ])
    ws2 = wb.create_sheet('2130H-0930H')
    # Month deliberately misspelled, as the real file has it.
    sheet(ws2, 'TEST MALL:  Auust 16-31 , 2026     2130H-0930H  NIGHT SHIFT', [
        ('NIGHT GUARD TWO', 'N/S', [(t(20, 51), t(9, 31)), (t(20, 54), t(9, 34)),
                                    (t(20, 53), t(9, 37)), (t(20, 55), t(9, 40)), (t(20, 47), t(9, 36))]),
    ])
    wb.save(path)
    return path


if __name__ == '__main__':
    _main()
