"""Cherche, dans toutes les formules d un .xlsx, les références à des lignes FUTURES (ligne > ligne de la formule). Usage : python3 check_future_refs.py fichier.xlsx"""
import sys
import openpyxl, re, collections
from openpyxl.utils import get_column_letter as L
f=sys.argv[1]
wf=openpyxl.load_workbook(f,data_only=False)
ref=re.compile(r"(?:(\w+)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)")
tot=0
for ws in wf:
    fut=collections.Counter(); ex={}; n=0
    for row in ws.iter_rows(min_row=2):
        for c in row:
            v=c.value
            if isinstance(v,str) and v.startswith('='):
                n+=1
                s=re.sub(r'"[^"]*"','',v)
                for m in ref.finditer(s):
                    if m.group(4)=='$': continue
                    d=int(m.group(5))-c.row
                    if d>0: fut[L(c.column)]+=1; ex.setdefault(L(c.column),(c.coordinate,m.group(0)))
    print(ws.title,'formules',n,'refs futures:',dict(fut),ex)
