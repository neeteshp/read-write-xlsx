# Generates test fixtures with openpyxl (and LibreOffice when available).
# Run: python3 scripts/make-fixtures.py
import datetime, os, subprocess, openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

out = os.path.join(os.path.dirname(__file__), '..', 'test', 'fixtures')

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'People'
ws.append(['Name', 'Age', 'Born', 'Active', 'Score', 'Time'])
ws.append(['Alice', 34, datetime.datetime(1992, 5, 17), True, 0.875, datetime.time(13, 30)])
ws.append(['Bob', '41', datetime.datetime(1985, 1, 2, 8, 15), False, None, None])
ws.append([CellRichText(['Ca', TextBlock(InlineFont(b=True), 'rol')]), 29, None, None, '=1/0', None])
ws['E2'].number_format = '0.0%'
ws['C3'].number_format = 'dd.mm.yyyy hh:mm'
ws['H10'] = 'far away'
other = wb.create_sheet('Hidden stuff')
other['A1'] = 'second sheet'
other.sheet_state = 'hidden'
third = wb.create_sheet('Ünïcødé & <xml>')
third['B2'] = 'ok'
wb.save(os.path.join(out, 'openpyxl.xlsx'))

wb = openpyxl.Workbook()
wb.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904
wb.active.append(['Date'])
wb.active.append([datetime.datetime(2020, 2, 29)])
wb.save(os.path.join(out, 'date1904.xlsx'))

# LibreOffice-produced file (different XML layout/prefixes than openpyxl).
try:
    subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML',
                    '--outdir', '/tmp/lo-fixture', os.path.join(out, 'openpyxl.xlsx')],
                   check=True, timeout=120, capture_output=True)
    os.replace('/tmp/lo-fixture/openpyxl.xlsx', os.path.join(out, 'libreoffice.xlsx'))
except Exception as e:
    print('LibreOffice fixture skipped:', e)
print('fixtures written')
