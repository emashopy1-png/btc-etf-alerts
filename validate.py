"""Compare le moteur Python (engine.py) aux valeurs de CONFLUENCE_<flux> d'un Excel réenregistré (valeurs en cache), cellule par cellule.

Usage : python3 validate.py fichier.xlsx [flux=IEX] [tolérance relative=1e-9]
Entrées du moteur = valeurs des feuilles ETF ; sorties comparées = colonnes D..P de CONFLUENCE + colonnes d'aide (JF, KD, KP, WC, GJ, GK).
"""
import datetime as dt
import sys
from collections import defaultdict

import openpyxl
from openpyxl.utils import column_index_from_string as CI, get_column_letter as L

from engine import ERR, ETFS, Engine, is_num, f_excluded_from_refs

ETF_COLS = "F G H I J K L M Y Z AT BA BB BC BD BE BJ BN CR CV".split()
CM_COL = "CM"   # réf. tape de chaque ETF (sert à déduire la liste des ETF de F)


def f_excluded_for(etf_sheets, first=2, last=289):
    """ETF exclus de F pour la journée : ceux sans réf. tape (CM) numérique > 0."""
    refs = {s: [clean(ws.cell(r, CI(CM_COL)).value) for r in range(first, last + 1)] for s, ws in etf_sheets.items()}
    return f_excluded_from_refs(refs)
TARGETS = {"D": "D", "E": "E", "F": "F", "G": "G", "H": "H", "I": "I", "J": "J", "K": "K", "L": "L", "M": "M", "P": "P"}
FIRST, LAST = 2, 289          # lignes de données (288 bougies de 5 min)
HELP = {  # colonnes d'aide de CONFLUENCE : (première colonne du bloc de 12 ETF, attribut du moteur)
    "JF": ("JF", "JF"), "KD": ("KD", "KD"), "KP": ("KP", "KP"), "WC": ("HD", "WC"),
}


def clean(v):
    if v is None or v == "":
        return None
    if isinstance(v, str) and v.startswith("#"):
        return ERR
    return v


def same(a, b, tol):
    if a is None or b is None:
        return a is None and b is None
    if a == ERR or b == ERR:
        return a == b
    if is_num(a) and is_num(b):
        return abs(a - b) <= tol * max(1.0, abs(a), abs(b))
    return a == b


def minutes(t):
    if isinstance(t, dt.datetime):
        t = t.time()
    if isinstance(t, dt.time):
        return t.hour * 60 + t.minute
    raise ValueError(f"heure illisible : {t!r}")


def main(path, feed="IEX", tol=1e-9):
    wb = openpyxl.load_workbook(path, data_only=True)
    conf = wb[f"CONFLUENCE_{feed}"]
    etf = {s: wb[f"{s}_{feed}"] for s in ETFS}
    cidx = {c: CI(c) for c in list(TARGETS) + ["C", "EQ", "ER", "GJ", "GK"]}
    fx = f_excluded_for(etf)
    print('ETF exclus de F (déduits de CM) :', sorted(fx))
    eng = Engine(f_excluded=fx)
    mism = defaultdict(list)
    stats = defaultdict(lambda: [0, 0])      # colonne -> [comparées, identiques]
    help_cols = {name: [L(CI(first) + k) for k in range(12)] for name, (first, _) in HELP.items()}

    # lecture en mémoire (plus rapide que cellule par cellule)
    etf_rows = {s: {c: [clean(etf[s].cell(r, CI(c)).value) for r in range(FIRST, LAST + 1)] for c in ETF_COLS} for s in ETFS}
    nrows = LAST - FIRST + 1
    for k in range(nrows):
        r = FIRST + k
        bars = {s: {c: etf_rows[s][c][k] for c in ETF_COLS} for s in ETFS}
        refs = {"EQ": clean(conf.cell(r, cidx["EQ"]).value), "ER": clean(conf.cell(r, cidx["ER"]).value)}
        mn = minutes(conf.cell(r, cidx["C"]).value)
        out = eng.push(bars, refs, mn)
        for col in TARGETS:
            exp = clean(conf.cell(r, cidx[col]).value)
            got = out[col]
            ok = same(got, exp, tol)
            stats[col][0] += 1
            stats[col][1] += ok
            if not ok:
                mism[col].append((r, k + 1, got, exp))
        # colonnes d'aide
        for name, cols in help_cols.items():
            for j, c in enumerate(cols):
                exp = clean(conf.cell(r, CI(c)).value)
                got = getattr(eng, HELP[name][1])[ETFS[j]][k]
                if name == "JF" and not is_num(got):
                    got = None
                ok = same(got, exp, tol) if not (name == "JF" and exp is None) else got is None
                if name in ("KD", "KP", "WC") and exp is None:
                    ok = False
                key = f"{name}[{c}]"
                stats[name][0] += 1
                stats[name][1] += ok
                if not ok:
                    mism[name].append((r, ETFS[j], got, exp))
        for col, series in (("GJ", eng.GJ), ("GK", eng.GK)):
            exp = clean(conf.cell(r, cidx[col]).value)
            ok = same(series[k], exp if exp is not None else 0.0, tol)
            stats[col][0] += 1
            stats[col][1] += ok
            if not ok:
                mism[col].append((r, k + 1, series[k], exp))

    print(f"{path}\nflux {feed}, {nrows} bougies, tolérance relative {tol}\n")
    print(f"{'colonne':<8}{'comparées':>11}{'identiques':>12}{'écarts':>8}")
    total = bad = 0
    for col in list(TARGETS) + list(HELP) + ["GJ", "GK"]:
        c, o = stats[col]
        total += c
        bad += c - o
        print(f"{col:<8}{c:>11}{o:>12}{c - o:>8}")
    print(f"\nTOTAL {total} cellules comparées, {bad} écart(s)")
    for col, lst in mism.items():
        print(f"\n--- écarts colonne {col} (jusqu'à 12) : (ligne, bougie/ETF, python, excel)")
        for e in lst[:12]:
            print("   ", e)
    return bad


if __name__ == "__main__":
    p = sys.argv[1]
    f = sys.argv[2] if len(sys.argv) > 2 else "IEX"
    t = float(sys.argv[3]) if len(sys.argv) > 3 else 1e-9
    sys.exit(1 if main(p, f, t) else 0)
