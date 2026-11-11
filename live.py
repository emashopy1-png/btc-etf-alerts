#!/usr/bin/env python3
"""live.py : pilote temps réel du moteur CONFLUENCE (engine.py) sur le flux Alpaca (feed IEX par défaut).

    pip install websockets
    export APCA_API_KEY_ID=...  APCA_API_SECRET_KEY=...
    python live.py                                   # live, feed iex
    python live.py --replay trades.jsonl             # replay hors-ligne (messages Alpaca : T=t / T=q)
    python live.py --selftest 20261110_IEX.json      # test de bout en bout depuis le JSON de référence
    python live.py --seed 20261110_IEX.json          # démarrage à froid : réfs tape J-1 prises dans ce JSON

Principes
  * 1 instance d'Engine par jour NY (créée à minuit NY, bougies de 00:00 à 23:55, comme le classeur).
  * Une bougie de 5 min n'est émise qu'une fois close (+ GRACE s) : aucune donnée future dans le calcul.
  * Les réfs (EQ/ER) viennent de la veille (tape_refs.json relu à CHAQUE changement de jour) : correctif multi-jours inclus.
  * Colonnes non calculables depuis le tape (AS, AT, BB, BC, BD, BE, BJ, BN, CE, CF) : absentes -> défauts du moteur.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import logging
import os
import re
import statistics
import time
from collections import deque
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import engine as eng
from engine import ETFS, F_EXCLUDED, Engine, is_num

NY = ZoneInfo("America/New_York")
BAR = 300                      # secondes
GRACE = 2.0                    # attente (s) après la fin d'une bougie pour les trades en retard
SESSION = (9 * 60 + 30, 16 * 60)   # bougies servant de réf. tape (minutes NY)
REF_KEEP = 390                 # bougies de réf. conservées par ETF (≈ 5 séances)
BA_WIN = 20                    # BA = médiane du volume des 20 bougies précédentes
HIST_DIR = Path("history")
REFS_FILE = Path("tape_refs.json")
log = logging.getLogger("live")

try:                           # hook optionnel
    import alerts  # type: ignore
except Exception:              # noqa: BLE001
    alerts = None


# ------------------------------------------------------------------------------------------------ utilitaires
_FRAC = re.compile(r"(\.\d{6})\d+")


def parse_ts(s: str) -> float:
    """RFC3339 Alpaca (nanosecondes, 'Z') -> epoch secondes."""
    s = _FRAC.sub(r"\1", s.replace("Z", "+00:00"))
    return datetime.fromisoformat(s).timestamp()


def bucket_of(ts: float) -> int:
    return int(ts // BAR) * BAR


def ny_midnight(ts: float) -> int:
    d = datetime.fromtimestamp(ts, NY)
    return int(datetime(d.year, d.month, d.day, tzinfo=NY).timestamp())


def med(seq):
    seq = [x for x in seq if is_num(x)]
    return statistics.median(seq) if seq else None


# ------------------------------------------------------------------------------------------------ agrégation 5 min
def _blank():
    return dict(hi=None, lo=None, last=None, vol=0.0, n=0, pv=0.0, buy=0.0, sell=0.0,
                bs=0.0, bn=0, as_=0.0, an=0, cq=0.0, cr=0.0, cu=0.0, cv=0.0)


class Agg:
    """Agrège trades/quotes d'une bougie ; L/M par tick rule (Alpaca ne donne pas le sens de l'agresseur)."""

    def __init__(self, etfs):
        self.etfs = list(etfs)
        self.px = {}           # dernier prix (persiste d'une bougie / d'un jour à l'autre)
        self.sg = {}           # dernier signe non nul
        self.cur = {e: _blank() for e in self.etfs}

    def trade(self, sym, p, sz):
        c = self.cur.get(sym)
        if c is None or not sz or sz <= 0:
            return
        prev = self.px.get(sym)
        s = 0 if prev is None else (1 if p > prev else -1 if p < prev else self.sg.get(sym, 0))
        if s:
            self.sg[sym] = s
        self.px[sym] = p
        c["n"] += 1
        c["vol"] += sz
        c["pv"] += p * sz
        c["hi"] = p if c["hi"] is None else max(c["hi"], p)
        c["lo"] = p if c["lo"] is None else min(c["lo"], p)
        c["last"] = p
        if s > 0:
            c["buy"] += sz
        elif s < 0:
            c["sell"] += sz
        if 500 <= sz < 2000:                      # buckets de taille -> colonnes O (CQ/CU) et P (CR/CV)
            c["cu"] += sz
            c["cq"] += s * sz
        elif sz >= 2000:
            c["cv"] += sz
            c["cr"] += s * sz

    def quote(self, sym, bs, as_):
        c = self.cur.get(sym)
        if c is None:
            return
        c["bs"] += bs
        c["bn"] += 1
        c["as_"] += as_
        c["an"] += 1

    def finalize(self, ba):
        """-> dict des clés de colonnes attendues par Engine.push ; remet les compteurs à zéro."""
        bars = {}
        for e in self.etfs:
            c = self.cur[e]
            has = c["n"] > 0
            bars[e] = {
                "F": c["hi"], "G": c["lo"],
                "H": c["last"] if has else self.px.get(e),     # clôture reportée sur bougie vide
                "I": c["vol"], "J": float(c["n"]),
                "K": c["pv"] / c["vol"] if has and c["vol"] > 0 else None,
                "L": c["buy"], "M": c["sell"],
                "Y": c["bs"] / c["bn"] if c["bn"] else None,
                "Z": c["as_"] / c["an"] if c["an"] else None,
                "BA": ba.get(e),
                "CQ": c["cq"], "CR": c["cr"], "CU": c["cu"], "CV": c["cv"],
            }
        self.cur = {e: _blank() for e in self.etfs}
        return bars


# ------------------------------------------------------------------------------------------------ réfs persistées
class RefStore:
    """tape_refs.json = {"hist": {etf: [trades/bougie, ...≤390]}, "today": {"day": "...", "bars": {etf: [...]}}}.
    hist = séances terminées (réfs J-1 et antérieures) ; today = séance en cours (jamais utilisée comme réf. du jour)."""

    def __init__(self, path=REFS_FILE):
        self.path = Path(path)
        self.d = {"hist": {}, "today": {"day": None, "bars": {}}}
        try:
            self.d = json.loads(self.path.read_text())
        except Exception:      # noqa: BLE001
            pass

    def save(self):
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.d))
        os.replace(tmp, self.path)

    def rollover(self, day):
        """Nouveau jour NY : verse la séance précédente dans hist (bornée) et repart à vide."""
        t = self.d["today"]
        if t.get("day") and t["day"] != day:
            for e, v in t["bars"].items():
                h = self.d["hist"].setdefault(e, [])
                h.extend(v)
                self.d["hist"][e] = h[-REF_KEEP:]
            self.d["today"] = {"day": None, "bars": {}}
        if self.d["today"].get("day") is None:
            self.d["today"]["day"] = day
        self.save()

    def add_bar(self, minutes, trades):
        if SESSION[0] <= minutes < SESSION[1]:
            b = self.d["today"]["bars"]
            for e, n in trades.items():
                b.setdefault(e, []).append(n)

    def seed_json(self, path):
        """Démarrage à froid depuis un JSON de référence (même format que 20261110_IEX.json)."""
        ref = json.loads(Path(path).read_text())
        hist = {}
        for e, rows in ref["etfs"].items():
            hist[e] = [rows[k][3] for k in sorted(rows) if SESSION[0] <= int(k[:2]) * 60 + int(k[3:]) < SESSION[1]]
        self.d["hist"] = hist
        self.save()

    def er(self, excl, etfs):
        """ER = médiane (séances J-1 et avant) de Σ_{ETF non exclus} trades/5 min, par bougie de séance."""
        h = self.d["hist"]
        ln = min((len(h[e]) for e in etfs if e in h), default=0)
        if ln == 0:
            return None
        tot = [sum(h[e][-ln + k] for e in etfs if e in h and e not in excl) / eng.BAR_MINUTES for k in range(ln)]
        m = med(tot)
        return m if m and m > 0 else None


# ------------------------------------------------------------------------------------------------ pilote
class Runner:
    def __init__(self, f_excluded=None, store=None, write_history=True):
        self.agg = Agg(ETFS)
        self.store = store or RefStore()
        self.f_excl_cli = set(f_excluded) if f_excluded else None
        self.write_history = write_history
        self.vol = {e: deque(maxlen=BA_WIN) for e in ETFS}
        self.bucket = None        # début de la bougie en cours
        self.first = None         # première bougie « réelle » (avant : rattrapage depuis minuit, sans alertes)
        self.day = None
        self.eng = None
        self.er = None
        self.late = 0
        self.hooks = []           # fn(bucket, bars, out) (tests)

    # ---- entrées
    def on_trade(self, sym, ts, p, sz):
        b = bucket_of(ts)
        self._advance(b)
        if b < self.bucket:
            self.late += 1
            return
        self.agg.trade(sym, p, sz)

    def on_quote(self, sym, ts, bs, as_, bp=1, ap=1):
        if not bp or not ap:
            return
        b = bucket_of(ts)
        self._advance(b)
        if b < self.bucket:
            return
        self.agg.quote(sym, bs, as_)

    def tick(self, now):                 # appelé chaque seconde en live
        if self.bucket is not None:
            self._advance(bucket_of(now - GRACE))

    def flush(self, ts):                 # ferme la bougie contenant ts (fin de replay / arrêt)
        self._advance(bucket_of(ts) + BAR)
        self.store.save()

    # ---- interne
    def _advance(self, target):
        if self.bucket is None:
            self.first = target
            self.bucket = ny_midnight(target)
        while self.bucket < target:
            self._close(self.bucket)
            self.bucket += BAR

    def _new_day(self, day):
        self.store.rollover(day)                         # relit/verse les réfs de la veille (correctif multi-jours)
        excl = self.f_excl_cli if self.f_excl_cli is not None else set(F_EXCLUDED)
        self.er = self.store.er(excl, ETFS)
        self.eng = Engine(ETFS, f_excluded=excl)
        self.day = day
        log.info("jour %s : f_excluded=%s ER=%s", day, sorted(excl), None if self.er is None else round(self.er, 2))

    def _close(self, b):
        dt = datetime.fromtimestamp(b, NY)
        day = dt.date().isoformat()
        if day != self.day:
            self._new_day(day)
        minutes = dt.hour * 60 + dt.minute
        ba = {e: (med(self.vol[e]) if len(self.vol[e]) >= 3 else None) for e in ETFS}   # bougies PRÉCÉDENTES seulement
        bars = self.agg.finalize(ba)
        refs = {"EQ": None, "ER": self.er}
        try:
            out = self.eng.push(bars, refs, minutes)
        except Exception as ex:      # noqa: BLE001  (ne jamais tuer la boucle live)
            log.exception("push a échoué sur %s : %s", dt.isoformat(), ex)
            out = {k: eng.ERR for k in "DEFGHIJKLMNOP"}
        for e in ETFS:
            self.vol[e].append(bars[e]["I"])
        self.store.add_bar(minutes, {e: int(bars[e]["J"]) for e in ETFS})
        if minutes % 60 == 0:
            self.store.save()
        catchup = b < (self.first or 0)
        rec = {"t": dt.isoformat(), "out": out,
               "bars": {e: {k: bars[e][k] for k in ("H", "I", "J", "L", "M")} for e in ETFS if bars[e]["J"]}}
        if self.write_history and not catchup:
            HIST_DIR.mkdir(exist_ok=True)
            with open(HIST_DIR / f"{day}.jsonl", "a") as f:
                f.write(json.dumps(rec) + "\n")
        if not catchup:
            log.info("%s %s", dt.strftime("%H:%M"), {k: (round(v, 4) if is_num(v) else v) for k, v in out.items()})
            if alerts is not None:
                try:
                    alerts.on_bar(dt.isoformat(), out, bars)      # <- adapter si la signature diffère
                except Exception as ex:   # noqa: BLE001
                    log.warning("alerts.on_bar : %s", ex)
        for h in self.hooks:
            h(b, bars, out)


# ------------------------------------------------------------------------------------------------ live
async def run_live(runner, feed):
    import websockets                                   # pip install websockets
    key, sec = os.environ["APCA_API_KEY_ID"], os.environ["APCA_API_SECRET_KEY"]
    url = f"wss://stream.data.alpaca.markets/v2/{feed}"

    async def clock():
        while True:
            runner.tick(time.time())
            await asyncio.sleep(1)

    task = asyncio.create_task(clock())
    backoff = 1
    try:
        while True:
            try:
                async with websockets.connect(url, ping_interval=20, max_size=None) as ws:
                    await ws.recv()
                    await ws.send(json.dumps({"action": "auth", "key": key, "secret": sec}))
                    resp = json.loads(await ws.recv())
                    if not any(m.get("msg") == "authenticated" for m in resp):
                        raise RuntimeError(f"auth refusée : {resp}")
                    await ws.send(json.dumps({"action": "subscribe", "trades": ETFS, "quotes": ETFS}))
                    log.info("connecté (%s)", feed)
                    backoff = 1
                    async for raw in ws:
                        for m in json.loads(raw):
                            _dispatch(runner, m)
            except asyncio.CancelledError:
                raise
            except Exception as ex:    # noqa: BLE001
                log.warning("flux interrompu (%s) ; reconnexion dans %ss", ex, backoff)
                await asyncio.sleep(backoff)
                backoff = min(60, backoff * 2)
    finally:
        task.cancel()
        runner.store.save()


def _dispatch(runner, m):
    t = m.get("T")
    if t == "t":
        runner.on_trade(m["S"], parse_ts(m["t"]), float(m["p"]), float(m["s"]))
    elif t == "q":
        runner.on_quote(m["S"], parse_ts(m["t"]), float(m.get("bs", 0)), float(m.get("as", 0)),
                        m.get("bp", 1), m.get("ap", 1))
    elif t == "error":
        log.error("flux : %s", m)


# ------------------------------------------------------------------------------------------------ replay / selftest
def run_replay(runner, path):
    last = None
    with open(path) as f:
        for line in f:
            line = line.strip()
            if line:
                m = json.loads(line)
                _dispatch(runner, m)
                last = parse_ts(m["t"])
    if last:
        runner.flush(last)


def run_selftest(runner, ref_path):
    """Fabrique des trades à partir du JSON de référence (volume I et nb de trades J par bougie), les rejoue
    et compare I et J bougie par bougie. Ne valide NI le sens (tick rule) NI Y/Z : le JSON ne les contient pas."""
    ref = json.loads(Path(ref_path).read_text())
    d0 = datetime.fromisoformat(ref["date"]).replace(tzinfo=NY)
    base = d0.timestamp()
    ev = []
    for e, rows in ref["etfs"].items():
        for k, (vol, _q, _x, n) in rows.items():
            n, t0 = int(n), base + (int(k[:2]) * 60 + int(k[3:])) * 60
            if n <= 0:
                continue
            sz, rem = divmod(int(vol), n)
            for j in range(n):
                ev.append((t0 + (j + 0.5) * BAR / n, e, 100.0 + 0.01 * (j % 2), sz + (rem if j == n - 1 else 0)))
    ev.sort()
    got = {}
    runner.hooks.append(lambda b, bars, out: got.__setitem__(b, {e: (bars[e]["I"], bars[e]["J"]) for e in ETFS}))
    for ts, e, p, sz in ev:
        runner.on_trade(e, ts, p, sz)
    runner.flush(base + 86399)
    bad = tot = 0
    for e, rows in ref["etfs"].items():
        for k, (vol, _q, _x, n) in rows.items():
            b = int(base + (int(k[:2]) * 60 + int(k[3:])) * 60)
            tot += 1
            if got.get(b, {}).get(e) != (float(vol), float(n)):
                bad += 1
    print(f"selftest : {len(ev)} trades rejoués, {tot} (ETF x bougie) comparés, {bad} écarts sur (I, J) ; "
          f"{len(got)} bougies émises ; en retard : {runner.late}")
    return bad == 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--feed", default=os.environ.get("APCA_DATA_FEED", "iex"))
    ap.add_argument("--replay")
    ap.add_argument("--selftest")
    ap.add_argument("--seed", help="JSON de référence pour amorcer les réfs tape (démarrage à froid)")
    ap.add_argument("--f-excluded", default=None, help="ex. BTCO,BTCW (défaut : engine.F_EXCLUDED)")
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    excl = a.f_excluded.split(",") if a.f_excluded else None
    if a.selftest:
        global HIST_DIR
        import tempfile
        tmp = Path(tempfile.mkdtemp())
        r = Runner(excl, RefStore(tmp / "refs.json"), write_history=False)
        if a.seed:
            r.store.seed_json(a.seed)
        raise SystemExit(0 if run_selftest(r, a.selftest) else 1)
    store = RefStore()
    if a.seed and not store.d["hist"]:
        store.seed_json(a.seed)
    runner = Runner(excl, store)
    if a.replay:
        run_replay(runner, a.replay)
    else:
        try:
            asyncio.run(run_live(runner, a.feed))
        except KeyboardInterrupt:
            runner.store.save()


if __name__ == "__main__":
    main()
