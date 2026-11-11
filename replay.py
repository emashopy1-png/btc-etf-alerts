"""Rejoue une journée (Excel réenregistré) : moteur Python -> machine d'alertes -> tableau des événements.

Usage : python3 replay.py fichier.xlsx [flux=IEX]
"""
import sys

import openpyxl
from openpyxl.utils import column_index_from_string as CI

from alerts import AlertMachine
from engine import ETFS, Engine
from validate import ETF_COLS, FIRST, LAST, clean, minutes, f_excluded_for


def replay(path, feed="IEX"):
    wb = openpyxl.load_workbook(path, data_only=True)
    conf = wb[f"CONFLUENCE_{feed}"]
    etf = {s: wb[f"{s}_{feed}"] for s in ETFS}
    rows = {s: {c: [clean(etf[s].cell(r, CI(c)).value) for r in range(FIRST, LAST + 1)] for c in ETF_COLS} for s in ETFS}
    eng, am = Engine(f_excluded=f_excluded_for(etf)), AlertMachine()
    events, series = [], []
    for k in range(LAST - FIRST + 1):
        r = FIRST + k
        bars = {s: {c: rows[s][c][k] for c in ETF_COLS} for s in ETFS}
        refs = {"EQ": clean(conf.cell(r, CI("EQ")).value), "ER": clean(conf.cell(r, CI("ER")).value)}
        mn = minutes(conf.cell(r, CI("C")).value)
        out = eng.push(bars, refs, mn)
        bar_no = conf.cell(r, 1).value
        series.append((bar_no, mn, out))
        events += am.push(bar_no, mn, out)
    return events, series


if __name__ == "__main__":
    ev, _ = replay(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "IEX")
    print(f"{'N°':>4} {'NY':<6}{'mode':<5}{'événement':<10}{'sens':<6}détail")
    for e in ev:
        print(f"{e.bar:>4} {e.ny:<6}{e.mode:<5}{e.kind:<10}{e.direction:<6}{e.info}")
    from collections import Counter
    print("\nRésumé :", dict(Counter((e.mode, e.kind) for e in ev)))
