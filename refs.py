"""Références « normales » par tranche horaire pour le live (mêmes règles que le script Excel v7/v9).

Sources :
  - history/*.json  : jours passés extraits des Excel ({date, etfs: {SYM: {"HH:MM": [vol, churn, carnet, trades]}}})
  - live_refs.json  : jours accumulés par le bot lui-même (même contenu, tableau de 288 slots par ETF)

Pour un jour J, seules les séances STRICTEMENT AVANT J servent (comme _hist_baseline du script Excel).
Références par ETF et par slot (288 slots de 5 min) :
  vol   : médiane du volume par slot sur >= 3 jours, lissée (moyenne slot ±1), plancher 10 % de la moyenne  -> BA / colonne D
  churn : médiane du churn (quotes/min) par slot, lissée (médiane ±1), plancher 10 % de la moyenne         -> EQ / colonne E
  tape  : médiane des trades/min par slot, lissée (médiane ±1), plancher 10 % de la moyenne                -> ER / colonne F
EQ(slot) = Σ churn des ETF dont la réf. churn est valide à CE slot ; ER(slot) = idem avec le tape.
"""
from __future__ import annotations

import json
import math
import os
from pathlib import Path
from statistics import mean, median

from engine import ETFS

BAR_MINUTES = 5
SLOTS = 24 * 60 // BAR_MINUTES          # 288
HIST_MIN_DAYS = 3                       # comme le script Excel
REF_SMOOTH = 1
REF_FLOOR_FRAC = 0.10                   # planchers identiques au script Excel
CHURN_REF_FLOOR_FRAC = 0.10
TAPE_REF_FLOOR_FRAC = 0.10

_FIN = float("inf")


def _fin(x) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool) and math.isfinite(x)


def _smooth(a, fn):
    out = [None] * len(a)
    for j in range(len(a)):
        w = [x for x in a[max(0, j - REF_SMOOTH): j + REF_SMOOTH + 1] if _fin(x)]
        if w:
            out[j] = float(fn(w))
    return out


def _slot_medians(days):
    """days : liste de tableaux [288][3] (vol, churn, trades ; None autorisé). -> 3 listes de médianes par slot."""
    need = min(HIST_MIN_DAYS, len(days))
    out = []
    for k in range(3):
        col = []
        for s in range(SLOTS):
            vals = [d[s][k] for d in days if s < len(d) and _fin(d[s][k])]
            col.append(float(median(vals)) if len(vals) >= need else None)
        out.append(col)
    return out


def _baseline(days):
    """-> (vol_slots, churn_slots, tape_slots) ou None si < HIST_MIN_DAYS séances. Reproduit _hist_baseline."""
    if len(days) < HIST_MIN_DAYS:
        return None
    v, c, tp = _slot_medians(days)
    finite_v = [x for x in v if _fin(x)]
    if not finite_v:
        return None
    floor = REF_FLOOR_FRAC * mean(finite_v)
    vol = [max(x, floor) if _fin(x) else floor for x in _smooth(v, mean)]
    churn = None
    if any(_fin(x) for x in c):
        cf = CHURN_REF_FLOOR_FRAC * mean(x for x in c if _fin(x))
        churn = [max(x, cf) if _fin(x) else None for x in _smooth(c, median)]
    tape = None
    tp5 = [x / BAR_MINUTES if _fin(x) else None for x in tp]
    if any(_fin(x) for x in tp5):
        tf = TAPE_REF_FLOOR_FRAC * mean(x for x in tp5 if _fin(x))
        tape = [max(x, tf) if _fin(x) else None for x in _smooth(tp5, median)]
    return {"vol": vol, "churn": churn, "tape": tape}


class RefBook:
    """Charge l'historique, calcule les références d'un jour, accumule les données live des jours suivants."""

    def __init__(self, history_dir="history", live_file=None):
        self.history_dir = Path(history_dir)
        self.live_file = Path(live_file or os.environ.get("REFS_FILE", "live_refs.json"))
        self.days: dict[str, dict[str, list]] = {}     # date -> {etf: [[vol, churn, trades] x SLOTS]}
        self._load_history()
        self._load_live()
        self.today: str | None = None

    # ---------------------------------------------------------- chargement
    @staticmethod
    def _blank_day():
        return {e: [[None, None, None] for _ in range(SLOTS)] for e in ETFS}

    def _load_history(self):
        for p in sorted(self.history_dir.glob("*.json")):
            try:
                d = json.loads(p.read_text())
                day = d.get("date", "")
                if len(day) != 10 or day[4] != "-":          # date parfois corrompue -> nom de fichier AAAA MM JJ
                    day = f"{(s := p.stem.split('_')[0])[:4]}-{s[4:6]}-{s[6:8]}"
                arr = self._blank_day()
                for e, rows in d.get("etfs", {}).items():
                    if e not in arr:
                        continue
                    for hhmm, v in rows.items():
                        s = (int(hhmm[:2]) * 60 + int(hhmm[3:])) // BAR_MINUTES
                        if 0 <= s < SLOTS and isinstance(v, (list, tuple)) and len(v) >= 4:
                            arr[e][s] = [v[0], v[1], v[3]]          # vol, churn, trades (carnet inutile ici)
                self.days[day] = arr
            except Exception as ex:      # noqa: BLE001
                print(f"[!] référence ignorée {p} : {ex}")

    def _load_live(self):
        try:
            d = json.loads(self.live_file.read_text())
            for day, etfs in d.get("days", {}).items():
                self.days[day] = {e: (etfs.get(e) or [[None, None, None]] * SLOTS) for e in ETFS}
        except FileNotFoundError:
            pass
        except Exception as ex:          # noqa: BLE001
            print(f"[!] {self.live_file} illisible : {ex}")

    def save(self):
        tmp = self.live_file.with_suffix(".tmp")
        tmp.write_text(json.dumps({"days": self.days}))
        os.replace(tmp, self.live_file)

    # ---------------------------------------------------------- références d'un jour
    def baselines(self, day: str) -> dict:
        """{etf: {"vol": [288], "churn": [288]|None, "tape": [288]|None}} calculé sur les séances < day."""
        past = sorted(d for d in self.days if d < day)
        out = {}
        for e in ETFS:
            mats = [self.days[d][e] for d in past if e in self.days[d]]
            b = _baseline(mats)
            if b is not None:
                out[e] = b
        return out

    # ---------------------------------------------------------- accumulation live
    def start_day(self, day: str):
        if day not in self.days:
            self.days[day] = self._blank_day()
        self.today = day

    def record_bar(self, slot: int, per_etf):
        """per_etf : {etf: (volume, nb_quotes, nb_trades)} de la bougie close."""
        if self.today is None:
            return
        arr = self.days[self.today]
        for e, (vol, nq, nt) in per_etf.items():
            if e in arr and 0 <= slot < SLOTS:
                cur = arr[e][slot]
                vol = vol if _fin(vol) else 0.0
                ch = nq / BAR_MINUTES if _fin(nq) else None        # churn = quotes / minute (colonne BC)
                arr[e][slot] = [(cur[0] or 0.0) + vol,
                                None if ch is None and cur[1] is None else (cur[1] or 0.0) + (ch or 0.0),
                                (cur[2] or 0.0) + (nt if _fin(nt) else 0.0)]
