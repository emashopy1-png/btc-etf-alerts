"""Machine à états des alertes (Mode 1 et Mode 2 « Vito ») + envoi Telegram.

Alimentée bougie après bougie par les sorties de engine.py (replay.py pour le backtest,
live.py en temps réel via le hook `on_bar`). Chaque bougie i n'utilise que les valeurs
des bougies <= i : aucune donnée future.

RÈGLES VALIDÉES (mises à jour) :
  - « delta » = colonnes G, K, L, O, P. Un delta est « dans le sens » s'il est strictement
    du signe de la direction ; une case vide (None) compte comme « pas dans le sens ».
  - Mode 1 : déclencheur max(D,E,F) >= 1,8 avec I signé (0/vide = pas de déclencheur) ;
    direction = signe de I ; fenêtre 3 bougies ; annulation UNIQUEMENT si I change de signe
    (I = 0 ou vide N'annule PAS) ; un nouveau déclencheur remplace l'attente ;
    confirmation = >= 2 deltas dans le sens + J dans le sens sur une même bougie suivante.
  - Mode 2 « Vito » : M >= +0,15 long / M <= -0,15 short ; fenêtre 4 bougies ; I ignorée ;
    nouveaux déclencheurs ignorés pendant l'attente ;
    confirmation = >= 2 deltas dans le sens + J >= 0,05 dans le sens.
  - Ordre sur une bougie : 1) l'attente existante est évaluée (annulation, puis confirmation),
    2) ensuite seulement la bougie peut déclencher une nouvelle attente.
  - La bougie déclencheuse n'est jamais sa propre confirmation.
  - Session : déclenchement et confirmation uniquement de 9h30 à 14h00 NY (début de bougie).
  - Max 2 alertes confirmées par mode et par séance ; pas de cooldown entre elles.
    Une fois le max atteint, le mode ne déclenche plus rien jusqu'au lendemain.
  - État réinitialisé à chaque nouveau jour NY.

Envoi Telegram (hook live) : variables d'environnement
  TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID  -> si absentes, les événements sont seulement logués.
  ALERT_KINDS = "ALERT" par défaut ; ex. "TRIGGER,ALERT,CANCEL,EXPIRE" pour tout recevoir.
"""
from __future__ import annotations

import json
import logging
import os
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

try:
    from engine import ERR, is_num
except Exception:      # alerts.py doit aussi fonctionner sans engine.py (montage fusionné avec le script Excel)
    ERR = None

    def is_num(x):
        return isinstance(x, (int, float)) and not isinstance(x, bool) and x == x and x not in (float("inf"), float("-inf"))

log = logging.getLogger("alerts")

CONFIG = {
    "session_start": 9 * 60 + 30,      # minutes NY depuis minuit
    "session_end": 14 * 60,            # 14h00 NY max (début de bougie) ; 13*60+55 pour exclure 14h00
    "max_alerts": 2,                   # alertes confirmées max par mode et par séance
    "delta_cols": ("G", "K", "L", "O", "P"),
    "min_deltas": 2,
    "m1": {"trigger_cols": ("D", "E", "F"), "trigger_min": 1.8, "window": 3},
    "m2": {"trigger_m": 0.15, "window": 4, "j_min": 0.05},
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
    out: Optional[dict] = None   # valeurs D..P de la bougie (détail du message Telegram)


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
    n1: int = 0      # alertes M1 confirmées aujourd'hui
    n2: int = 0      # alertes M2 confirmées aujourd'hui

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
            ev.append(Event(bar_no, hh, mode, kind, self._dir(d), info, out))

        # ------------------------------------------------ Mode 1 : attente en cours
        if self.p1:
            p = self.p1
            if sgn(out.get("I")) == -p.d:                      # annulation SEULEMENT si I change de signe
                E("M1", "CANCEL", p.d, f"I={out.get('I')} a changé de signe (déclencheur N°{p.trig_bar})")
                self.p1 = None
            else:
                nd = n_deltas_in_dir(out, c["delta_cols"], p.d)
                j_ok = is_num(out.get("J")) and sgn(out["J"]) == p.d
                p.left -= 1
                if nd >= c["min_deltas"] and j_ok:
                    E("M1", "ALERT", p.d, f"déclencheur N°{p.trig_bar}")
                    self.n1 += 1
                    self.p1 = None
                elif p.left == 0:
                    E("M1", "EXPIRE", p.d, f"non confirmé ({nd} delta, J={out.get('J')}) (déclencheur N°{p.trig_bar})")
                    self.p1 = None
        # Mode 1 : déclenchement (bloqué une fois le max d'alertes atteint)
        m1 = c["m1"]
        trig = [k for k in m1["trigger_cols"] if is_num(out.get(k)) and out[k] >= m1["trigger_min"]]
        if trig and sgn(out.get("I")) != 0 and self.n1 < c["max_alerts"]:
            d = sgn(out["I"])
            if self.p1:
                E("M1", "REPLACED", self.p1.d, f"remplacée par le déclencheur N°{bar_no}")
            self.p1 = _Pending(d, bar_no, m1["window"])
            E("M1", "TRIGGER", d, f"{','.join(f'{k}={out[k]:.2f}' for k in trig)} ; I={out['I']:+.3f}")

        # ------------------------------------------------ Mode 2 (Vito) : attente en cours (I ignorée)
        m2 = c["m2"]
        if self.p2:
            p = self.p2
            nd = n_deltas_in_dir(out, c["delta_cols"], p.d)
            j_ok = is_num(out.get("J")) and p.d * out["J"] >= m2["j_min"]
            p.left -= 1
            if nd >= c["min_deltas"] and j_ok:
                E("M2", "ALERT", p.d, f"déclencheur N°{p.trig_bar}")
                self.n2 += 1
                self.p2 = None
            elif p.left == 0:
                E("M2", "EXPIRE", p.d, f"non confirmé ({nd} delta, J={out.get('J')}) (déclencheur N°{p.trig_bar})")
                self.p2 = None
        # Mode 2 : déclenchement (ignoré pendant l'attente ; bloqué une fois le max atteint)
        m = out.get("M")
        if is_num(m) and abs(m) >= m2["trigger_m"] and self.n2 < c["max_alerts"]:
            d = sgn(m)
            if self.p2:
                E("M2", "IGNORED", d, f"nouveau déclencheur ignoré (M={m:+.3f}), attente N°{self.p2.trig_bar} en cours")
            else:
                self.p2 = _Pending(d, bar_no, m2["window"])
                E("M2", "TRIGGER", d, f"M={m:+.3f}")
        return ev


# ------------------------------------------------------------------ hook live + Telegram
_KIND_FR = {"TRIGGER": "Déclencheur", "ALERT": "ALERTE", "CANCEL": "Annulation",
            "EXPIRE": "Expiré", "REPLACED": "Remplacé", "IGNORED": "Ignoré"}
_KIND_EMOJI = {"TRIGGER": "⏳", "ALERT": "🚨", "CANCEL": "✖️", "EXPIRE": "⌛",
               "REPLACED": "🔁", "IGNORED": "⏭️"}


def _fmt(x, nd=2, sign=False) -> str:
    if not is_num(x):
        return "—"
    return f"{x:+.{nd}f}" if sign else f"{x:.{nd}f}"


def format_event(e: Event) -> str:
    side = "🟢 LONG" if e.direction == "LONG" else "🔴 SHORT"
    head = f"{_KIND_EMOJI.get(e.kind, '')} {_KIND_FR.get(e.kind, e.kind)} {side} — Mode {e.mode[-1]} — {e.ny} NY (bougie {e.bar})"
    if e.kind != "ALERT" or not e.out:
        return f"{head}\n{e.info}" if e.info else head
    o, d = e.out, 1 if e.direction == "LONG" else -1
    parts, nd = [], 0
    for c in CONFIG["delta_cols"]:
        v = o.get(c)
        ok = is_num(v) and sgn(v) == d
        nd += ok
        parts.append(f"{c} {_fmt(v, 2, True)} {'✅' if ok else '❌'}")
    lines = [head,
             f"📊 Ratio volume (D) : {_fmt(o.get('D'))} | Ratio carnet (E) : {_fmt(o.get('E'))} | "
             f"Speed of tape (F) : {_fmt(o.get('F'))}",
             f"📈 CVD carnet (I) : {_fmt(o.get('I'), 3, True)} | CVD carnet prix (J) : {_fmt(o.get('J'), 3, True)}",
             f"Deltas ({nd}/5 dans le sens) : " + "  ".join(parts)]
    if e.info:
        lines.append(f"({e.info})")
    return "\n".join(lines)


def send_telegram(text: str) -> None:
    token, chat = os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat:
        return
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = json.dumps({"chat_id": chat, "text": text}).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            r.read()
    except Exception as ex:      # noqa: BLE001  (ne jamais tuer la boucle live pour Telegram)
        log.warning("telegram : %s", ex)


class _LiveAlerts:
    """Singleton alimenté par live.py : on_bar(ny_iso, out, bars). Nouvelle machine à chaque jour NY."""

    def __init__(self):
        self.day: Optional[str] = None
        self.machine = AlertMachine()
        kinds = os.environ.get("ALERT_KINDS", "ALERT")
        self.send_kinds = {k.strip().upper() for k in kinds.split(",") if k.strip()}

    def on_bar(self, ny_iso: str, out: dict, bars=None) -> List[Event]:
        dt = datetime.fromisoformat(ny_iso)
        day = dt.date().isoformat()
        if day != self.day:                       # minuit NY : état (dont compteurs n1/n2) remis à zéro
            self.day, self.machine = day, AlertMachine()
        minutes = dt.hour * 60 + dt.minute
        events = self.machine.push(minutes // 5 + 1, minutes, out)
        for e in events:
            log.info("%s %s %s %s %s", e.ny, e.mode, e.kind, e.direction, e.info)
            if e.kind in self.send_kinds:
                send_telegram(format_event(e))
        return events


_live = _LiveAlerts()


def on_bar(ny_iso: str, out: dict, bars=None) -> List[Event]:
    return _live.on_bar(ny_iso, out, bars)
