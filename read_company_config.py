import openpyxl
from pathlib import Path

path = Path(r"C:\Users\razor\Desktop\22 10 2026\SMART MOVE CO v8.4.14 ADVANCED SALES & COMMUNICATION SYSTEM.xlsx")
wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
for name in ['Company Configuration', 'AI Service Card Description']:
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    print('SHEET', name)
    for r in rows[:20]:
        print(r)
    print('---END---')
