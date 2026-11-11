#!/usr/bin/env python3
"""live.py : pilote temps réel du moteur CONFLUENCE (engine.py) sur le flux Alpaca (feed IEX par défaut).

    pip install websockets
    export APCA_API_KEY_ID=...  APCA_API_SECRET_KEY=...
    export TELEGRAM_BOT_TOKEN=... TELEGRAM_CHAT_ID=...     # optionnel (sinon : logs seulement)
    python live.py                                   # live, feed iex
    python live.py --replay trades.jsonl             # replay hors-ligne (messages Alpaca : T=t / T=q)
    python live.py --selftest history/20261110_IEX.json   # test de bout en bout (volumes et nb de trades)

Principes
  * 1 instance d'Engine par jour NY (créée à minuit NY, bougies de 00:00 à 23:55, comme le classeur).
  * Une bougie de 5 min n'est émise qu'une fois close (+ GRACE s) : aucune donnée future dans le calcul.
  * Références D / E / F par tranche horaire (médiane sur l'historique, comme l'Excel v7/v9) via refs.py :
    history/*.json (jours Excel) + live_refs.json (jours accumulés par le bot). Sur Railway, monter un
    volume et mettre REFS_FILE=/data/live_refs.json pour conserver les jours live entre redéploiements.
  * Sens des trades : règle du milieu (bid/ask) comme le script Excel, tick rule en secours.
  * ISO (colonne M, déclencheur du Mode 2) : trades avec condition « F » du stream, direction bid/ask
    (sans la détection de sweep du script Excel : légère différence possible sur M).
  * Colonnes non calculables depuis le tape (AS, CE, CF, BB, BD, BE, BJ, BN) : absentes -> défauts du moteur.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import logging
import os
import re
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import engine as eng
from engine import ETFS, Engine, is_num
from refs import BAR_MINUTES, SLOTS, RefBook

NY = ZoneInfo("America/New_York")
BAR = BAR_MINUTES * 60         # secondes
GRACE = 2.0                    # attente (s) après la fin d'une bougie pour les trades en retard
HIST_DIR = Path("history")
log = logging.getLogger("live")

# conditions de trade « non régulières » (comme le script Excel) : jamais classées « certaines »
NON_REGULAR_CONDS = {"B", "C", "N", "R", "V", "W", "Z", "U", "4", "P", "M", "Q"}

try:                           # hook optionnel (machine à états + Telegram)
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


def _conds(c) -> set:
    if not c:
        return set()
    if isinstance(c, str):
        return set(c.replace("[", " ").replace("]", " ").replace("'", " ").replace(",", " ").split())
    try:
        return {str(x) for x in c}
    except TypeError:
        return set()


# ------------------------------------------------------------------------------------------------ agrégation 5 min
def _blank():
    return dict(hi=None, lo=None, last=None, vol=0.0, n=0, pv=0.0, buy=0.0, sell=0.0,
                bs=0.0, bn=0, as_=0.0, an=0, qn=0, iso_v=0.0, iso_d=0.0,
                cq=0.0, cr=0.0, cu=0.0, cv=0.0)


class Agg:
    """Agrège trades/quotes d'une bougie. Sens des trades : règle du milieu (comme le script Excel),
    tick rule en secours. ISO : condition « F », direction certaine bid/ask (prix >= ask / <= bid)."""

    def __init__(self, etfs):
        self.etfs = list(etfs)
        self.px = {}           # dernier prix (persiste d'une bougie / d'un jour à l'autre)
        self.sg = {}           # dernier signe non nul (tick rule)
        self.q = {}            # dernière quote (bid, ask) par ETF
        self.cur = {e: _blank() for e in self.etfs}

    def trade(self, sym, p, sz, cond=()):
        c = self.cur.get(sym)
        if c is None or not sz or sz <= 0:
            return
        q = self.q.get(sym)
        prev = self.px.get(sym)
        tick = 0 if prev is None else (1 if p > prev else -1 if p < prev else self.sg.get(sym, 0))
        if q and q[0] > 0 and q[1] > q[0]:
            mid = (q[0] + q[1]) / 2
            s = 1 if p > mid else (-1 if p < mid else tick)     # règle du milieu, tick rule en secours
        else:
            s = tick
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
        conds = _conds(cond)
        if "F" in conds and not (conds & NON_REGULAR_CONDS):    # trade ISO « régulier » -> colonne M (delta ISO)
            d = 0
            if q and q[0] > 0 and q[1] > q[0]:
                d = 1 if p >= q[1] else (-1 if p <= q[0] else 0)
            c["iso_v"] += sz
            c["iso_d"] += d * sz

    def quote(self, sym, bs, as_, bp=1, ap=1):
        c = self.cur.get(sym)
        if c is None:
            return
        if bp and ap and bp > 0 and ap > bp:
            self.q[sym] = (bp, ap)
        c["bs"] += bs
        c["bn"] += 1
        c["as_"] += as_
        c["an"] += 1
        c["qn"] += 1                               # nb de mises à jour de quote -> churn (colonne BC)

    def finalize(self):
        """-> dict des clés de colonnes attendues par Engine.push ; remet les compteurs à zéro.
        BA est posée ensuite par le Runner (référence par tranche, refs.py)."""
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
                "BC": c["qn"] / BAR_MINUTES,                   # churn live (quotes / minute)
                "AT": c["iso_d"] / c["iso_v"] if c["iso_v"] > 0 else None,
                "QN": c["qn"],                                 # interne (accumulation des références)
                "CQ": c["cq"], "CR": c["cr"], "CU": c["cu"], "CV": c["cv"],
            }
        self.cur = {e: _blank() for e in self.etfs}
        return bars


# ------------------------------------------------------------------------------------------------ pilote
class Runner:
    def __init__(self, f_excluded=None, refbook=None, write_history=True):
        self.agg = Agg(ETFS)
        self.refbook = refbook or RefBook()
        self.f_excl_cli = set(f_excluded) if f_excluded else None
        self.write_history = write_history
        self.bucket = None        # début de la bougie en cours
        self.first = None         # première bougie « réelle » (avant : rattrapage depuis minuit, sans alertes)
        self.day = None
        self.eng = None
        self.ba = {}              # etf -> [288] réf. volume (colonne BA)
        self.eq = [None] * SLOTS  # réf. churn agrégée par slot (colonne EQ)
        self.eq_syms = [set() for _ in range(SLOTS)]
        self.er = [None] * SLOTS  # réf. tape agrégée par slot (colonne ER)
        self.late = 0
        self.hooks = []           # fn(bucket, bars, out) (tests)

    # ---- entrées
    def on_trade(self, sym, ts, p, sz, cond=()):
        b = bucket_of(ts)
        self._advance(b)
        if b < self.bucket:
            self.late += 1
            return
        self.agg.trade(sym, p, sz, cond)

    def on_quote(self, sym, ts, bs, as_, bp=1, ap=1):
        b = bucket_of(ts)
        self._advance(b)
        if b < self.bucket:
            return
        self.agg.quote(sym, bs, as_, bp, ap)

    def tick(self, now):                 # appelé chaque seconde en live
        if self.bucket is not None:
            self._advance(bucket_of(now - GRACE))

    def flush(self, ts):                 # ferme la bougie contenant ts (fin de replay / arrêt)
        self._advance(bucket_of(ts) + BAR)
        self.refbook.save()

    # ---- interne
    def _advance(self, target):
        if self.bucket is None:
            self.first = target
            self.bucket = ny_midnight(target)
        while self.bucket < target:
            self._close(self.bucket)
            self.bucket += BAR

    def _new_day(self, day):
        self.refbook.start_day(day)
        base = self.refbook.baselines(day)
        self.ba = {e: b["vol"] for e, b in base.items()}
        self.eq, self.er = [None] * SLOTS, [None] * SLOTS
        self.eq_syms = [set() for _ in range(SLOTS)]
        for s in range(SLOTS):
            ch = {e: b["churn"][s] for e, b in base.items()
                  if b["churn"] and is_num(b["churn"][s]) and b["churn"][s] > 0}
            if ch:
                self.eq[s] = sum(ch.values())
                self.eq_syms[s] = set(ch)
            tp = [b["tape"][s] for b in base.values()
                  if b["tape"] and is_num(b["tape"][s]) and b["tape"][s] > 0]
            if tp:
                self.er[s] = sum(tp)
        excl = self.f_excl_cli if self.f_excl_cli is not None else {
            e for e in ETFS if e not in base or not base[e]["tape"]
            or not any(is_num(x) and x > 0 for x in base[e]["tape"])}
        self.eng = Engine(ETFS, f_excluded=excl)
        self.day = day
        log.info("jour %s : réfs %d ETF, f_excluded=%s", day, len(base), sorted(excl))

    def _close(self, b):
        dt = datetime.fromtimestamp(b, NY)
        day = dt.date().isoformat()
        if day != self.day:
            self._new_day(day)
        minutes = dt.hour * 60 + dt.minute
        slot = minutes // BAR_MINUTES
        bars = self.agg.finalize()
        for e in ETFS:
            v = self.ba.get(e)
            bars[e]["BA"] = v[slot] if v and slot < len(v) else None
        refs = {"EQ": self.eq[slot] if slot < SLOTS else None,
                "ER": self.er[slot] if slot < SLOTS else None,
                "EQ_SYMS": self.eq_syms[slot] if slot < SLOTS else set()}
        try:
            out = self.eng.push(bars, refs, minutes)
        except Exception as ex:      # noqa: BLE001  (ne jamais tuer la boucle live)
            log.exception("push a échoué sur %s : %s", dt.isoformat(), ex)
            out = {k: eng.ERR for k in "DEFGHIJKLMNOP"}
        self.refbook.record_bar(slot, {e: (bars[e]["I"], bars[e]["QN"], bars[e]["J"]) for e in ETFS})
        if minutes % 60 == 0:
            self.refbook.save()
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
                    alerts.on_bar(dt.isoformat(), out, bars)
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
        runner.refbook.save()


def _dispatch(runner, m):
    t = m.get("T")
    if t == "t":
        runner.on_trade(m["S"], parse_ts(m["t"]), float(m["p"]), float(m["s"]), m.get("c"))
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
    """Fabrique des trades à partir d'un JSON d'historique (volume et nb de trades par bougie), les rejoue
    et compare I et J bougie par bougie. Ne valide NI le sens (bid/ask) NI Y/Z : le JSON ne les contient pas."""
    ref = json.loads(Path(ref_path).read_text())
    d0 = datetime.strptime(Path(ref_path).stem.split("_")[0], "%Y%m%d").replace(tzinfo=NY)
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
    ap.add_argument("--f-excluded", default=None, help="ex. BTCO,BTCW (défaut : déduit des réfs du jour)")
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    excl = a.f_excluded.split(",") if a.f_excluded else None
    if a.selftest:
        import tempfile
        tmp = Path(tempfile.mkdtemp())
        r = Runner(excl, RefBook(HIST_DIR, tmp / "live_refs.json"), write_history=False)
        raise SystemExit(0 if run_selftest(r, a.selftest) else 1)
    runner = Runner(excl)
    if a.replay:
        run_replay(runner, a.replay)
    else:
        try:
            asyncio.run(run_live(runner, a.feed))
        except KeyboardInterrupt:
            runner.refbook.save()


if __name__ == "__main__":
    main()
