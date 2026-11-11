"""Machine à états des alertes (Mode 1 et Mode 2 « Vito »), alimentée bougie après bougie par les sorties de engine.py.

Chaque bougie i n'utilise que les valeurs des bougies <= i : aucune donnée future.

HYPOTHÈSES À CONFIRMER (toutes modifiables dans CONFIG) :
  - « delta » = colonnes G, K, L, M, P. Un delta est « dans le sens » s'il est strictement du signe de la direction ;
    une case vide (None) compte comme « pas dans le sens ».
  - Ordre sur une bougie : 1) la mise en attente existante est évaluée (annulation d'abord, puis confirmation), 2) ensuite
    seulement, la bougie peut déclencher une nouvelle mise en attente.
  - Mode 1 : une nouvelle mise en attente REMPLACE l'ancienne (fenêtre repartant de zéro). Mode 2 : ignorée pendant l'attente.
  - La bougie déclencheuse n'est jamais la bougie de confirmation : la confirmation se cherche dans les N bougies suivantes.
  - Session : déclenchement et confirmation uniquement de 9h30 à 15h55 NY (heure de début de bougie).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional

from engine import ERR, is_num

CONFIG = {
    "session_start": 9 * 60 + 30,      # minutes NY depuis minuit
    "session_end": 15 * 60 + 55,       # dernière bougie (début 15h55)
    "delta_cols": ("G", "K", "L", "M", "P"),
    "min_deltas": 2,
    "m1": {"trigger_cols": ("D", "E", "F"), "trigger_min": 1.8, "window": 3},
    "m2": {"trigger_m": 0.20, "window": 4, "j_min": 0.05, "delta_cols": ("G", "K", "L", "M", "P")},
}


def sgn(x) -> int:
    return (x > 0) - (x < 0) if is_num(x) else 0


def n_deltas_in_dir(out: dict, cols, d: int) -> int:
    return sum(1 for c in cols if is_num(out.get(c)) and sgn(out[c]) == d)


@dataclass
class Event:
    bar: int            # N° de bougie (colonne A)
    ny: str             # heure NY de début de bougie
    mode: str           # "M1" / "M2"
    kind: str           # TRIGGER / ALERT / CANCEL / EXPIRE / REPLACED / IGNORED
    direction: str      # LONG / SHORT
    info: str = ""


@dataclass
class _Pending:
    d: int
    trig_bar: int
    left: int


@dataclass
class AlertMachine:
    cfg: dict = field(default_factory=lambda: CONFIG)
    p1: Optional[_Pending] = None
    p2: Optional[_Pending] = None

    @staticmethod
    def _dir(d):
        return "LONG" if d > 0 else "SHORT"

    def push(self, bar_no: int, minutes_ny: int, out: dict) -> List[Event]:
        c = self.cfg
        ev: List[Event] = []
        hh = f"{minutes_ny // 60:02d}:{minutes_ny % 60:02d}"
        in_session = c["session_start"] <= minutes_ny <= c["session_end"]
        if not in_session:
            return ev

        def E(mode, kind, d, info=""):
            ev.append(Event(bar_no, hh, mode, kind, self._dir(d), info))

        # ------------------------------------------------ Mode 1 : attente en cours
        if self.p1:
            p = self.p1
            if sgn(out.get("I")) != p.d:                       # I change de signe (ou devient nul / vide) -> annulation
                E("M1", "CANCEL", p.d, f"I={out.get('I')} a changé de signe (déclencheur N°{p.trig_bar})")
                self.p1 = None
            else:
                nd = n_deltas_in_dir(out, c["delta_cols"], p.d)
                j_ok = is_num(out.get("J")) and sgn(out["J"]) == p.d
                p.left -= 1
                if nd >= c["min_deltas"] and j_ok:
                    E("M1", "ALERT", p.d, f"{nd} delta dans le sens, J={out['J']:+.3f} (déclencheur N°{p.trig_bar})")
                    self.p1 = None
                elif p.left == 0:
                    E("M1", "EXPIRE", p.d, f"non confirmé ({nd} delta, J={out.get('J')}) (déclencheur N°{p.trig_bar})")
                    self.p1 = None
        # Mode 1 : déclenchement
        m1 = c["m1"]
        trig = [k for k in m1["trigger_cols"] if is_num(out.get(k)) and out[k] >= m1["trigger_min"]]
        if trig and sgn(out.get("I")) != 0:
            d = sgn(out["I"])
            if self.p1:
                E("M1", "REPLACED", self.p1.d, f"remplacée par le déclencheur N°{bar_no}")
            self.p1 = _Pending(d, bar_no, m1["window"])
            E("M1", "TRIGGER", d, f"{','.join(f'{k}={out[k]:.2f}' for k in trig)} ; I={out['I']:+.3f}")

        # ------------------------------------------------ Mode 2 (Vito) : attente en cours (I ignorée)
        m2 = c["m2"]
        if self.p2:
            p = self.p2
            nd = n_deltas_in_dir(out, m2["delta_cols"], p.d)
            j_ok = is_num(out.get("J")) and p.d * out["J"] >= m2["j_min"]
            p.left -= 1
            if nd >= c["min_deltas"] and j_ok:
                E("M2", "ALERT", p.d, f"{nd} delta dans le sens, J={out['J']:+.3f} (déclencheur N°{p.trig_bar})")
                self.p2 = None
            elif p.left == 0:
                E("M2", "EXPIRE", p.d, f"non confirmé ({nd} delta, J={out.get('J')}) (déclencheur N°{p.trig_bar})")
                self.p2 = None
        # Mode 2 : déclenchement (ignoré pendant l'attente)
        m = out.get("M")
        if is_num(m) and abs(m) >= m2["trigger_m"]:
            d = sgn(m)
            if self.p2:
                E("M2", "IGNORED", d, f"nouveau déclencheur ignoré (M={m:+.3f}), attente N°{self.p2.trig_bar} en cours")
            else:
                self.p2 = _Pending(d, bar_no, m2["window"])
                E("M2", "TRIGGER", d, f"M={m:+.3f}")
        return ev
