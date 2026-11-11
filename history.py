"""Agrège l'historique des JSON journaliers (dossier history/) en références par ETF et par bougie de 5 min.

Format d'un JSON (ex. 20261110_IEX.json) :
    {"date": "2026-11-10", "feed": "IEX", "etfs": {"IBIT": {"09:30": [volume, tape, carnet, x4], ...}, ...}}
Hypothèse sur les 4 valeurs de chaque bougie : [0] volume, [1] speed of tape, [2] carnet (déséquilibre %), [3] x4 (4e colonne, non nommée).

Usage : python3 history.py [dossier=history] [flux=IEX] [sortie=history_stats]
Sorties (dans le dossier courant) :
    <sortie>_daily.csv  : un total par jour et par ETF
    <sortie>_slots.csv  : moyenne / médiane par ETF et par bougie (HH:MM), sur tous les jours
    <sortie>.json       : mêmes références par bougie, lisible par un autre script
La ligne "ALL" = somme des 12 ETF bougie par bougie, puis moyenne sur les jours.
"""
import csv
import glob
import json
import os
import statistics
import sys
from collections import defaultdict

FIELDS = ["volume", "tape", "carnet", "x4"]
SUMMED = {"volume", "tape", "x4"}          # champs additifs entre ETF ; le carnet se moyenne


def num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool) and x == x


def load_days(folder, feed):
    days = []
    for path in sorted(glob.glob(os.path.join(folder, "*.json"))):
        try:
            with open(path, encoding="utf-8") as f:
                d = json.load(f)
        except (OSError, json.JSONDecodeError) as e:
            print(f"ignoré {path} : {e}", file=sys.stderr)
            continue
        if feed and d.get("feed") != feed:
            continue
        if "etfs" not in d or "date" not in d:
            print(f"ignoré {path} : format inattendu", file=sys.stderr)
            continue
        days.append(d)
    days.sort(key=lambda d: d["date"])
    return days


def mean(v):
    v = [x for x in v if num(x)]
    return sum(v) / len(v) if v else None


def median(v):
    v = [x for x in v if num(x)]
    return statistics.median(v) if v else None


def build(days):
    # series[etf][slot][champ] = [valeur jour 1, valeur jour 2, ...]
    series = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    daily = []
    for d in days:
        all_by_slot = defaultdict(lambda: {k: [] for k in FIELDS})
        for etf, bars in d["etfs"].items():
            tot = {k: 0.0 for k in FIELDS}
            carnets = []
            for slot, vals in bars.items():
                vals = (list(vals) + [None] * 4)[:4]
                for k, v in zip(FIELDS, vals):
                    series[etf][slot][k].append(v)
                    all_by_slot[slot][k].append(v)
                    if k in SUMMED and num(v):
                        tot[k] += v
                if num(vals[2]) and (num(vals[0]) and vals[0] > 0 or num(vals[1]) and vals[1] > 0):
                    carnets.append(vals[2])
            daily.append({"date": d["date"], "feed": d.get("feed", ""), "etf": etf,
                          "volume": tot["volume"], "tape": tot["tape"],
                          "carnet_moy": mean(carnets), "x4": tot["x4"]})
        # ligne ALL du jour : somme des ETF (carnet = moyenne des ETF)
        for slot, ch in all_by_slot.items():
            for k in FIELDS:
                v = [x for x in ch[k] if num(x)]
                series["ALL"][slot][k].append((sum(v) if k in SUMMED else mean(v)) if v else None)
    return series, daily


def main(folder="history", feed="IEX", out="history_stats"):
    days = load_days(folder, feed)
    if not days:
        sys.exit(f"Aucun JSON exploitable dans '{folder}' (flux {feed}).")
    series, daily = build(days)

    with open(f"{out}_daily.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(daily[0].keys()))
        w.writeheader()
        w.writerows(daily)

    header = ["etf", "slot", "n_jours"]
    for k in FIELDS:
        header += [f"{k}_moy", f"{k}_med"]
    ref = {}
    with open(f"{out}_slots.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        for etf in sorted(series, key=lambda e: (e == "ALL", e)):
            for slot in sorted(series[etf]):
                ch = series[etf][slot]
                row = [etf, slot, max(len([x for x in ch[k] if num(x)]) for k in FIELDS)]
                ref.setdefault(etf, {})[slot] = {}
                for k in FIELDS:
                    m, md = mean(ch[k]), median(ch[k])
                    row += ["" if m is None else round(m, 6), "" if md is None else round(md, 6)]
                    ref[etf][slot][k] = {"moy": m, "med": md}
                w.writerow(row)

    with open(f"{out}.json", "w", encoding="utf-8") as f:
        json.dump({"jours": [d["date"] for d in days], "feed": feed, "champs": FIELDS, "refs": ref}, f)

    print(f"{len(days)} jour(s) : {days[0]['date']} -> {days[-1]['date']} (flux {feed})")
    print(f"écrit : {out}_daily.csv, {out}_slots.csv, {out}.json")
    last = [r for r in daily if r["date"] == days[-1]["date"]]
    print(f"\n{'ETF':<6}{'volume':>12}{'tape':>12}{'carnet moy':>12}")
    for r in sorted(last, key=lambda r: -r["volume"]):
        cm = r["carnet_moy"]
        print(f"{r['etf']:<6}{r['volume']:>12.0f}{r['tape']:>12.0f}{'' if cm is None else format(cm, '.4f'):>12}")


if __name__ == "__main__":
    a = sys.argv[1:]
    main(a[0] if len(a) > 0 else "history", a[1] if len(a) > 1 else "IEX", a[2] if len(a) > 2 else "history_stats")
