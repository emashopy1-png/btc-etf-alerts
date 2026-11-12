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
  * Horodatages en NANOSECONDES entières (int), comme le script Excel (int64 ns) : les durées d'affichage
    des états du carnet (BE / BJ / BN) et la règle du milieu sont alors identiques au centième de ns.
  * Références D / E / F par tranche horaire (médiane sur l'historique, comme l'Excel v7/v9) via refs.py :
    history/*.json (jours Excel) + live_refs.json (jours accumulés par le bot). Sur Railway, monter un
    volume et mettre REFS_FILE=/data/live_refs.json pour conserver les jours live entre redéploiements.
  * FIDÉLITÉ TOTALE au script Excel (alpaca_excel_BTC_ETF_confluence_24h.py) :
      - sens des trades : règle du milieu (bid/ask BRUT, comme classify), tick rule en secours ;
      - ISO : condition « F » ; AT = (ISO acheteur − ISO vendeur) ÷ volume TOTAL de la bougie (sens = milieu) ;
        CE / CF = volumes ISO à direction CERTAINE (règle bid/ask sur la dernière cotation STRICTEMENT
        antérieure + détection de sweep <= 100 ms sur >= 3 niveaux ; contradiction ou non régulier -> 0) ;
      - carnet dynamique : BB (qualité du spread), BD (rotation des tailles), BE (déséquilibre pondéré
        temps), BJ (persistance), BN (déséquilibre x persistance) reconstruits ÉTAT PAR ÉTAT, durées
        plafonnées à 300 s, médianes causales sur les 12 bougies précédentes (_book_metrics / _book_metrics_v3) ;
      - état remis à zéro à chaque jour NY (le script traite une journée à la fois).
    Seule approximation : une rafale ISO qui franchit la fin d'une bougie est figée avec sa direction
    connue à la clôture (le script, en batch, connaît la rafale complète). Sans effet sur D..M, O, P.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import logging
import os
import queue
import re
import statistics
import threading
import time
from collections import deque
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import engine as eng
from engine import ETFS, Engine, is_num
from refs import BAR_MINUTES, SLOTS, RefBook

NY = ZoneInfo("America/New_York")
BAR = BAR_MINUTES * 60         # secondes
NS = 1_000_000_000
BAR_NS = BAR * NS              # durée d'une bougie en nanosecondes
GRACE_NS = 2 * NS              # attente après la fin d'une bougie pour les trades en retard
HIST_DIR = Path("history")
log = logging.getLogger("live")

# conditions de trade « non régulières » (comme le script Excel) : jamais classées « certaines »
NON_REGULAR_CONDS = {"B", "C", "N", "R", "V", "W", "Z", "U", "4", "P", "M", "Q"}

# constantes des indicateurs carnet (identiques au script Excel)
SWEEP_GAP_NS = 100_000_000     # SWEEP_GAP_MS = 100 : trades ISO séparés de moins de 100 ms = même rafale
SWEEP_MIN_LEVELS = 3           # un sweep traverse au moins 3 niveaux de prix distincts
BOOK_QUALITY_WINDOW = 12       # BB : médiane du spread moyen sur 12 bougies (bougie courante exclue)
BOOK_MIN_HISTORY = 3           # BB : nb minimum de bougies d'historique avec spread valide
TW_CAP_NS = 300 * NS           # durée retenue max pour un même état du carnet (5 min)
REF_WINDOW_NS = 12 * BAR_NS    # BJ/BN : médiane des durées d'état sur 12 bougies (courante exclue)
REF_MIN_STATES = 20            # BJ/BN : nb minimum d'états dans la fenêtre pour une médiane

try:                           # hook optionnel (machine à états + Telegram)
    import alerts  # type: ignore
except Exception:              # noqa: BLE001
    alerts = None

try:                           # exécution Bybit démo des alertes confirmées (EXEC_ENABLED=1)
    import bybit_exec as executor  # type: ignore
except Exception:              # noqa: BLE001
    executor = None


# ------------------------------------------------------------------------------------------------ data Excel (montage hybride)
# Le service Excel (alpaca_excel_BTC_ETF_confluence_24h.py en mode live) publie sur GitHub, à chaque clôture
# de bougie, les valeurs recalculées du classeur. Ici, à chaque bougie close : on attend le JSON de CETTE
# bougie, on remplace G, I, J, K, L, M, O, P (les colonnes qui dérivaient) par les valeurs de l'Excel, et on
# garde D, E, F calculés par le bot comme avant. L'attente tourne dans UN thread dédié avec une file FIFO :
# la boucle websocket n'est jamais bloquée et les bougies arrivent à la machine à alertes dans l'ordre.
EXCEL_COLS = ("G", "I", "J", "K", "L", "M", "O", "P")
_EXCEL_ON = False
_EXCEL_URL = None
_EXCEL_WAIT_S = int(os.environ.get("EXCEL_WAIT_S", "180"))   # attente max du JSON de la bougie (le service Excel met ~1-2 min)
_EXCEL_POLL_S = float(os.environ.get("EXCEL_POLL_S", "5"))   # intervalle entre deux lectures du JSON
_excel_q = None                                              # queue.Queue créée au démarrage live


def _excel_feed_setup():
    """Appelé au démarrage live (jamais en replay/selftest). Actif seulement si EXCEL_FEED=1 + URL disponible."""
    global _EXCEL_ON, _EXCEL_URL, _excel_q
    if os.environ.get("EXCEL_FEED", "0") != "1":
        log.info("data Excel : désactivée (EXCEL_FEED != 1) — toutes les colonnes sont calculées par le bot.")
        return
    url = os.environ.get("EXCEL_JSON_URL", "").strip()
    if not url:
        repo = os.environ.get("GITHUB_REPO", "").strip()
        branch = os.environ.get("GITHUB_BRANCH", "live-data").strip() or "live-data"
        if repo:
            url = f"https://raw.githubusercontent.com/{repo}/{branch}/live/confluence_derniere_bougie.json"
    if not url:
        log.warning("EXCEL_FEED=1 mais ni EXCEL_JSON_URL ni GITHUB_REPO : data Excel désactivée.")
        return
    _EXCEL_URL = url
    _excel_q = queue.Queue()
    threading.Thread(target=_excel_worker, daemon=True, name="excel-feed").start()
    _EXCEL_ON = True
    log.info("data Excel : ACTIVÉE (%s) — %s viennent de l'Excel ; attente max %ss.",
             url, ", ".join(EXCEL_COLS), _EXCEL_WAIT_S)


def _excel_fetch():
    """Télécharge le JSON de la dernière bougie (paramètre anti-cache) -> dict ou None."""
    import urllib.request
    try:
        with urllib.request.urlopen(f"{_EXCEL_URL}?t={time.time_ns()}", timeout=10) as r:
            return json.loads(r.read().decode())
    except Exception:              # noqa: BLE001  (réseau, 404 avant le 1er tick du service Excel...)
        return None


def _excel_wait_for(ny_key: str):
    """Attend le JSON de la bougie ny_key ('AAAA-MM-JJTHH:MM') -> payload, ou None (timeout / bougie dépassée)."""
    t0 = time.time()
    while time.time() - t0 < _EXCEL_WAIT_S:
        p = _excel_fetch()
        if p and p.get("day") == ny_key[:10]:
            got = str(p.get("ny_iso", ""))[:16]
            if got == ny_key:
                return p
            if got > ny_key:       # le service Excel a déjà publié une bougie plus récente : celle-ci n'existera pas
                log.warning("data Excel : bougie %s sautée par le service (il est à %s) — valeurs du bot.", ny_key, got)
                return None
        time.sleep(_EXCEL_POLL_S)
    log.warning("data Excel : pas de JSON pour %s après %ss — valeurs du bot utilisées pour cette bougie.",
                ny_key, _EXCEL_WAIT_S)
    return None


def _excel_worker():
    """Un seul thread : pour chaque bougie close, merge Excel puis alertes + exécution (ordre des bougies garanti)."""
    while True:
        dt, out, bars = _excel_q.get()
        try:
            p = _excel_wait_for(dt.isoformat()[:16])
            if p and isinstance(p.get("values"), dict):
                vals = p["values"]
                for k in EXCEL_COLS:
                    out[k] = vals.get(k)           # None = cellule vide dans l'Excel (c'est la valeur de référence)
                log.info("%s : %s <- Excel (recalculées)", dt.strftime("%H:%M"), "+".join(EXCEL_COLS))
            events = []
            if alerts is not None:
                try:
                    events = alerts.on_bar(dt.isoformat(), out, bars) or []
                except Exception as ex:   # noqa: BLE001
                    log.warning("alerts.on_bar : %s", ex)
            if events and executor is not None:
                try:
                    executor.on_events(events)
                except Exception as ex:   # noqa: BLE001
                    log.warning("bybit_exec.on_events : %s", ex)
        except Exception as ex:           # noqa: BLE001  (ne jamais tuer le thread)
            log.warning("excel worker : %s", ex)
        finally:
            _excel_q.task_done()


# ------------------------------------------------------------------------------------------------ utilitaires
_TS_RE = re.compile(r"^([^.]+)(?:\.(\d+))?(Z|[+-]\d{2}:?\d{2})$")


def parse_ts(s: str) -> int:
    """RFC3339 Alpaca (nanosecondes, 'Z') -> epoch NANOSECONDES (int exact, comme le script Excel)."""
    m = _TS_RE.match(s)
    frac = (m.group(2) or "").ljust(9, "0")[:9]
    dt = datetime.fromisoformat(m.group(1) + ("+00:00" if m.group(3) == "Z" else m.group(3)))
    return int(dt.timestamp()) * NS + int(frac or 0)


def _dt_ns(ts: int) -> datetime:
    """epoch ns -> datetime NY exacte (pas de float)."""
    return datetime.fromtimestamp(ts // NS, NY).replace(microsecond=(ts % NS) // 1000)


def _iso_ms(ts: int) -> str:
    """epoch ns -> ISO 8601 à la milliseconde (troncature), pour l'API historique Alpaca."""
    return _dt_ns(ts).isoformat(timespec="milliseconds")


def bucket_of(ts: int) -> int:
    return (ts // BAR_NS) * BAR_NS


def ny_midnight(ts: int) -> int:
    d = _dt_ns(ts)
    return int(datetime(d.year, d.month, d.day, tzinfo=NY).timestamp()) * NS


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
                bs=0.0, bn=0, as_=0.0, an=0, qn=0, spr=0.0,
                iso_v=0.0, iso_b=0.0, iso_s=0.0, ce=0.0, cf=0.0,
                rot_min=None, rot_max=None, rot_n=0, seg=[],
                cq=0.0, cr=0.0, cu=0.0, cv=0.0)


class Agg:
    """Agrège trades/quotes d'une bougie de 5 min — réplique exacte des feuilles ETF du script Excel.

    États du carnet : chaque quote ouvre un état (affiché jusqu'à la quote suivante). Segments de la
    bougie = état affiché au début + tous les états ouverts dans la bougie, durée = vie dans la bougie,
    âge = vie depuis l'apparition (plafonnés à 300 s), en ns entières — exactement comme
    _book_metrics / _book_metrics_v3 (qui travaillent en int64 ns).
    """

    def __init__(self, etfs):
        self.etfs = list(etfs)
        self.bucket = None                       # début (epoch ns) de la bougie en cours, posé par le Runner
        self.cur = {e: _blank() for e in self.etfs}
        self.reset_day()

    def reset_day(self):
        """État par ETF valable UNE journée (le script Excel traite un jour à la fois : tick rule,
        fenêtres de médianes et états du carnet ne traversent pas minuit). La bougie en cours (cur)
        n'est PAS touchée : le reset est piloté par le Runner dès que la bougie courante change de jour."""
        self.px = {}           # dernier prix du jour (tick rule)
        self.sg = {}           # dernier signe non nul (ffill de la tick rule)
        self.qcur = {}         # etf -> (ts, bp, ap, bs, as) dernière quote du jour
        self.qpre = {}         # etf -> dernière quote d'instant STRICTEMENT antérieur à qcur
        self.st = {}           # etf -> (ts début, bid_size, ask_size) état carnet affiché
        self.hist = {e: deque() for e in self.etfs}     # (début, fin) des états terminés (médianes BJ/BN)
        self.spr_hist = {e: [] for e in self.etfs}      # spread_avg des bougies closes (médiane BB)
        self.grp = {}          # etf -> rafale ISO en cours (sweep)

    # ---- trades
    def trade(self, sym, ts, p, sz, cond=()):
        c = self.cur.get(sym)
        if c is None or not sz or sz <= 0:
            return
        # classify() : dernière quote à ts <= trade (égalité incluse), SANS filtre de validité
        qi = self.qcur.get(sym)
        if qi is not None and qi[0] > ts:
            qi = self.qpre.get(sym)
        if qi is not None and qi[0] > ts:
            qi = None
        prev = self.px.get(sym)
        tick = 0 if prev is None else (1 if p > prev else -1 if p < prev else self.sg.get(sym, 0))
        if qi is not None:
            mid = (qi[1] + qi[2]) / 2
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
        if "F" in conds:                          # trade ISO (script : AUCUN filtre « régulier » pour AS/AT)
            c["iso_v"] += sz
            if s > 0:                             # sens = règle du milieu (colonne AT : ÷ volume TOTAL)
                c["iso_b"] += sz
            elif s < 0:
                c["iso_s"] += sz
            # direction « certaine » (_sure_direction) : bid/ask de la dernière quote STRICTEMENT avant
            qs = self.qcur.get(sym)
            if qs is not None and qs[0] >= ts:
                qs = self.qpre.get(sym)
            if qs is not None and qs[0] >= ts:
                qs = None
            ba = 0
            if qs is not None:
                bp, ap = qs[1], qs[2]
                if bp > 0 and ap > bp:            # marché locked / croisé : non classé
                    ba = 1 if p >= ap else (-1 if p <= bp else 0)
            # rafale (sweep) : trades ISO à moins de 100 ms ; la rafale inclut aussi les non réguliers
            g = self.grp.get(sym)
            if g is None or ts - g["ts"] > SWEEP_GAP_NS:
                if g is not None:
                    self._fin_grp(sym)            # rafale terminée : direction définitive
                g = {"first": p, "last": p, "dmin": None, "dmax": None, "px": {p}, "ts": ts, "pend": []}
                self.grp[sym] = g
            else:
                d = p - g["last"]
                g["dmin"] = d if g["dmin"] is None else min(g["dmin"], d)
                g["dmax"] = d if g["dmax"] is None else max(g["dmax"], d)
                g["px"].add(p)
                g["last"] = p
                g["ts"] = ts
            if not (conds & NON_REGULAR_CONDS):   # non régulier : dans la rafale, mais jamais « certain »
                g["pend"].append((ba, sz))

    # ---- quotes
    def quote(self, sym, ts, bs, as_, bp=1, ap=1):
        c = self.cur.get(sym)
        if c is None:
            return
        qc = self.qcur.get(sym)
        if qc is None or ts > qc[0]:
            self.qpre[sym] = qc                   # dernière quote d'instant strictement antérieur
            self.qcur[sym] = (ts, bp, ap, bs, as_)
        else:                                     # même instant (rafale) : la dernière compte, qpre inchangé
            self.qcur[sym] = (ts, bp, ap, bs, as_)
        prev = self.st.get(sym)                   # l'état précédent cesse d'être affiché à ts
        if prev is not None:
            self.hist[sym].append((prev[0], ts))
            dur = ts - max(prev[0], self.bucket)  # vie de l'état DANS la bougie (ns)
            if dur > 0:
                tot = prev[1] + prev[2]
                imb = (prev[1] - prev[2]) / tot if tot > 0 else None
                c["seg"].append((min(dur, TW_CAP_NS), min(ts - prev[0], TW_CAP_NS), imb))
        self.st[sym] = (ts, bs, as_)
        c["bs"] += bs                             # moyennes Y / Z (toutes les quotes, comme le script)
        c["bn"] += 1
        c["as_"] += as_
        c["an"] += 1
        c["qn"] += 1                              # nb de mises à jour -> churn (colonne BC)
        c["spr"] += ap - bp                       # spread moyen (colonne AA des feuilles -> BB)
        tot = bs + as_
        if tot > 0:                               # BD : rotation des tailles affichées (quotes de la bougie)
            c["rot_min"] = tot if c["rot_min"] is None else min(c["rot_min"], tot)
            c["rot_max"] = tot if c["rot_max"] is None else max(c["rot_max"], tot)
            c["rot_n"] += 1

    # ---- rafales ISO (sweeps)
    @staticmethod
    def _grp_dir(g):
        """Direction de la rafale (script _sure_direction) : >= 3 prix distincts, variation monotone."""
        if len(g["px"]) < SWEEP_MIN_LEVELS or g["dmin"] is None:
            return 0
        if g["dmin"] >= 0 and g["last"] > g["first"]:
            return 1
        if g["dmax"] <= 0 and g["last"] < g["first"]:
            return -1
        return 0

    def _fin_grp(self, sym):
        """Cumule les volumes « certains » (CE / CF) des trades de la rafale en attente, avec la
        direction courante de la rafale. Appelé quand la rafale se termine (direction définitive)
        ou à la clôture de la bougie (direction connue à cet instant — seule approximation du live)."""
        g = self.grp.get(sym)
        if not g or not g["pend"]:
            return
        sw = self._grp_dir(g)
        c = self.cur[sym]
        for ba, sz in g["pend"]:
            sure = 0 if (ba and sw and ba != sw) else (ba if ba else sw)   # contradiction -> 0
            if sure > 0:
                c["ce"] += sz
            elif sure < 0:
                c["cf"] += sz
        g["pend"] = []

    # ---- clôture de bougie
    def finalize(self, b):
        """Clôt la bougie [b, b+BAR_NS) -> dict des clés attendues par Engine.push ; remet les compteurs
        à zéro. BA est posée ensuite par le Runner (référence par tranche, refs.py)."""
        s, e = b, b + BAR_NS
        bars = {}
        for etf in self.etfs:
            c = self.cur[etf]
            self._fin_grp(etf)                    # rafale ISO encore ouverte : figée à la clôture
            stt = self.st.get(etf)                # l'état affiché à la clôture participe à la bougie
            if stt is not None:
                dur = e - max(stt[0], s)
                if dur > 0:
                    tot = stt[1] + stt[2]
                    imb = (stt[1] - stt[2]) / tot if tot > 0 else None
                    c["seg"].append((min(dur, TW_CAP_NS), min(e - stt[0], TW_CAP_NS), imb))
            # BE : déséquilibre du carnet pondéré par la durée d'affichage
            num = den = 0
            ages = []
            segs = [sg for sg in c["seg"] if sg[2] is not None]
            for dur, age, imb in segs:
                num += imb * dur                  # dur : int ns -> produit float identique au script
                den += dur
                if age > 0:
                    ages.append(age)
            be = num / den if den > 0 else None
            # BJ / BN : persistance = min(1, âge ÷ médiane des durées d'état des 12 bougies précédentes)
            ws = s - REF_WINDOW_NS
            ds = []
            for t0, t1 in self.hist[etf]:
                if t0 >= s:
                    break                         # deque chronologique
                if t0 >= ws:
                    ds.append(min(min(t1, s) - t0, TW_CAP_NS))
            if stt is not None and ws <= stt[0] < s:
                ds.append(min(s - stt[0], TW_CAP_NS))
            ref = None
            if len(ds) >= REF_MIN_STATES:
                dpos = [d for d in ds if d > 0]
                if len(dpos) >= REF_MIN_STATES:
                    ref = statistics.median(dpos)
            if not (ref is not None and ref > 0) and len(ages) >= 3:   # repli : médiane des âges de la bougie
                ref = statistics.median(ages)
            bj = bn = None
            if ref is not None and ref > 0 and segs:
                pers = [min(1.0, age / ref) for _dur, age, _imb in segs]
                bj = sum(pers) / len(pers)        # par comptage d'états (script : np.mean)
                wnum = wden = 0.0
                for (dur, _age, imb), ps in zip(segs, pers):
                    if ps > 0:
                        wnum += imb * dur * ps
                        wden += dur * ps
                bn = wnum / wden if wden > 0 else None
            # BB : médiane mobile du spread moyen (12 bougies précédentes) ÷ spread de la bougie
            spr = c["spr"] / c["qn"] if c["qn"] else None
            hist = [x for x in self.spr_hist[etf][-BOOK_QUALITY_WINDOW:] if x is not None and x > 0]
            bb = min(1.0, statistics.median(hist) / spr) if spr is not None and spr > 0 \
                and len(hist) >= BOOK_MIN_HISTORY else None
            self.spr_hist[etf].append(spr)
            # BD : rotation des tailles affichées
            bd = c["rot_max"] / c["rot_min"] if c["rot_n"] >= 2 else None
            # nettoyage de la fenêtre des états (12 bougies de la PROCHAINE bougie au plus loin)
            wn = s + BAR_NS - REF_WINDOW_NS
            h = self.hist[etf]
            while h and h[0][0] < wn:
                h.popleft()
            has = c["n"] > 0
            bars[etf] = {
                "F": c["hi"], "G": c["lo"],
                "H": c["last"] if has else None,   # clôture VIDE sur bougie vide (comme l'Excel)
                "I": c["vol"], "J": float(c["n"]),
                "K": c["pv"] / c["vol"] if has and c["vol"] > 0 else None,
                "L": c["buy"], "M": c["sell"],
                "Y": c["bs"] / c["bn"] if c["bn"] else None,
                "Z": c["as_"] / c["an"] if c["an"] else None,
                "BB": bb, "BD": bd, "BE": be, "BJ": bj, "BN": bn,
                "BC": c["qn"] / BAR_MINUTES,                   # churn (quotes / minute)
                "AS": c["iso_v"] / c["vol"] if c["vol"] > 0 else None,
                "AT": (c["iso_b"] - c["iso_s"]) / c["vol"] if c["vol"] > 0 else None,
                "CE": c["ce"], "CF": c["cf"],
                "QN": c["qn"],                                 # interne (accumulation des références)
                "CQ": c["cq"], "CR": c["cr"], "CU": c["cu"], "CV": c["cv"],
            }
        self.cur = {e2: _blank() for e2 in self.etfs}
        return bars


# ------------------------------------------------------------------------------------------------ pilote
class Runner:
    def __init__(self, f_excluded=None, refbook=None, write_history=True):
        self.agg = Agg(ETFS)
        self.refbook = refbook or RefBook()
        self.f_excl_cli = set(f_excluded) if f_excluded else None
        self.write_history = write_history
        self.bucket = None        # début de la bougie en cours (epoch ns)
        self.first = None         # première bougie « réelle » (avant : rattrapage depuis minuit, sans alertes)
        self.catchup_until = None # posé par backfill_today : tout ce qui est rejoué reste « catchup »
        self.day = None
        self.eng = None
        self.ba = {}              # etf -> [288] réf. volume (colonne BA)
        self.eq = [None] * SLOTS  # réf. churn agrégée par slot (colonne EQ)
        self.eq_syms = [set() for _ in range(SLOTS)]
        self.er = [None] * SLOTS  # réf. tape agrégée par slot (colonne ER)
        self.late = 0
        self._agg_day = None      # jour NY de l'état de l'agrégateur (reset à minuit)
        self.hooks = []           # fn(bucket, bars, out) (tests)

    # ---- entrées
    def on_trade(self, sym, ts, p, sz, cond=()):
        b = bucket_of(ts)
        self._advance(b)
        if b < self.bucket:
            self.late += 1
            return
        self.agg.trade(sym, ts, p, sz, cond)

    def on_quote(self, sym, ts, bs, as_, bp=1, ap=1):
        b = bucket_of(ts)
        self._advance(b)
        if b < self.bucket:
            return
        self.agg.quote(sym, ts, bs, as_, bp, ap)

    def tick(self, now):                 # appelé chaque seconde en live (now en secondes)
        if self.bucket is not None:
            self._advance(bucket_of(int(now * NS) - GRACE_NS))

    def flush(self, ts):                 # ferme la bougie contenant ts (fin de replay / arrêt)
        self._advance(bucket_of(ts) + BAR_NS)
        self.refbook.save()

    # ---- interne
    def _advance(self, target):
        if self.bucket is None:
            self.first = target
            self.bucket = ny_midnight(target)
        while self.bucket < target:
            self._close(self.bucket)
            self.bucket += BAR_NS
            day = _dt_ns(self.bucket).date().isoformat()
            if day != self._agg_day:     # minuit NY : l'état carnet/prix ne traverse pas les jours
                self.agg.reset_day()     # (le script Excel traite une journée à la fois) — AVANT tout
                self._agg_day = day      # événement ou clôture de bougie du nouveau jour
        self.agg.bucket = self.bucket

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
        dt = _dt_ns(b)
        day = dt.date().isoformat()
        if day != self.day:
            self._new_day(day)
        minutes = dt.hour * 60 + dt.minute
        slot = minutes // BAR_MINUTES
        bars = self.agg.finalize(b)
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
        catchup = b < max(self.first or 0, self.catchup_until or 0)
        if not catchup:                   # ne jamais enregistrer les bougies rejouées (double comptage)
            self.refbook.record_bar(slot, {e: (bars[e]["I"], bars[e]["QN"], bars[e]["J"]) for e in ETFS})
            if minutes % 60 == 0:
                self.refbook.save()
        rec = {"t": dt.isoformat(), "out": out,
               "bars": {e: {k: bars[e][k] for k in ("H", "I", "J", "L", "M")} for e in ETFS if bars[e]["J"]}}
        if self.write_history and not catchup:
            HIST_DIR.mkdir(exist_ok=True)
            with open(HIST_DIR / f"{day}.jsonl", "a") as f:
                f.write(json.dumps(rec) + "\n")
        if not catchup:
            log.info("%s %s", dt.strftime("%H:%M"), {k: (round(v, 4) if is_num(v) else v) for k, v in out.items()})
            if _EXCEL_ON:
                _excel_q.put((dt, out, bars))   # merge Excel + alertes + exécution dans le thread dédié
            else:
                events = []
                if alerts is not None:
                    try:
                        events = alerts.on_bar(dt.isoformat(), out, bars) or []
                    except Exception as ex:   # noqa: BLE001
                        log.warning("alerts.on_bar : %s", ex)
                if events and executor is not None:
                    try:
                        executor.on_events(events)
                    except Exception as ex:   # noqa: BLE001
                        log.warning("bybit_exec.on_events : %s", ex)
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


# ------------------------------------------------------------------------------------------------ rattrapage du jour (API historique)
# Au (re)démarrage en cours de journée : rejoue trades + quotes depuis minuit NY pour reconstruire
# H, I, J et les fenêtres internes du moteur comme si le bot tournait depuis minuit. Sans alertes
# ni enregistrement dans les références (catchup). BACKFILL_TODAY=0 pour désactiver.
BACKFILL_TODAY = os.environ.get("BACKFILL_TODAY", "1") == "1"
REST_URL = "https://data.alpaca.markets"
_PAGE = 10000


def _rest_json(path, params, key, sec):
    """GET data.alpaca.markets avec retries (réseau, 429)."""
    import urllib.parse
    import urllib.request
    url = f"{REST_URL}{path}?{urllib.parse.urlencode(params)}"
    wait = 1.0
    for attempt in range(5):
        try:
            req = urllib.request.Request(url, headers={"APCA-API-KEY-ID": key,
                                                       "APCA-API-SECRET-KEY": sec})
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read())
        except Exception as ex:      # noqa: BLE001
            if attempt == 4:
                raise
            log.warning("rattrapage %s : %s (retry dans %.0fs)", path, ex, wait)
            time.sleep(wait)
            wait = min(30.0, wait * 2)


def _paged(path, d0, t1, feed, key, sec, field):
    """Itérateur paresseux sur toutes les pages d'un endpoint historique (trades / quotes)."""
    token = None
    while True:
        params = {"start": d0, "end": t1, "limit": _PAGE, "feed": feed, "sort": "asc"}
        if token:
            params["page_token"] = token
        r = _rest_json(path, params, key, sec)
        yield from (r.get(field) or [])
        token = r.get("next_page_token")
        if not token:
            return
        time.sleep(0.15)             # marge sous la limite Alpaca (200 req/min)


def backfill_today(runner, feed):
    """Rejoue le jour en cours depuis minuit NY, puis des passes de rattrapage jusqu'au présent
    (le téléchargement de la 1re passe prend du temps : sans passes suivantes, les bougies couvrant
    ce délai resteraient vides). Les flux paginés des 12 ETF sont fusionnés par heapq.merge :
    mémoire bornée (~1 page par flux) et ordre global chronologique respecté. À timestamp égal,
    la quote passe AVANT le trade (comme le merge_asof « égalité incluse » du script Excel)."""
    import heapq
    key, sec = os.environ["APCA_API_KEY_ID"], os.environ["APCA_API_SECRET_KEY"]
    t0 = ny_midnight(int(time.time() * NS))
    if int(time.time() * NS) - t0 < BAR_NS:
        return                                    # journée qui vient de commencer : rien à rejouer

    start = t0
    for passe in range(1, 5):
        end = int(time.time() * NS)
        if end - start < 45 * NS:                 # moins de 45 s de trou : inutile de refaire une passe
            break
        d0 = _iso_ms(start)
        t1 = _iso_ms(end)

        def quotes(sym):
            for m in _paged(f"/v2/stocks/{sym}/quotes", d0, t1, feed, key, sec, "quotes"):
                yield (parse_ts(m["t"]), 0, sym, m)

        def trades(sym):
            for m in _paged(f"/v2/stocks/{sym}/trades", d0, t1, feed, key, sec, "trades"):
                yield (parse_ts(m["t"]), 1, sym, m)

        runner.catchup_until = end                # rejoué = pas d'alerte, pas de record (double comptage)
        streams = [g for sym in ETFS for g in (quotes(sym), trades(sym))]
        log.info("rattrapage (passe %d) : rejeu %s -> %s ...", passe, d0[:19], t1[:19])
        n = 0
        for ts, kind, sym, m in heapq.merge(*streams, key=lambda e: (e[0], e[1])):
            if kind == 1:
                runner.on_trade(sym, ts, float(m["p"]), float(m["s"]), m.get("c"))
            else:
                runner.on_quote(sym, ts, float(m.get("bs", 0)), float(m.get("as", 0)),
                                float(m.get("bp", 1)), float(m.get("ap", 1)))
            n += 1
            if n % 500000 == 0:
                log.info("rattrapage : %d événements rejoués...", n)
        log.info("rattrapage passe %d terminée : %d événements", passe, n)
        start = end                               # passe suivante : seulement le trou créé par celle-ci


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
    base = int(d0.timestamp()) * NS
    ev = []
    for e, rows in ref["etfs"].items():
        for k, (vol, _q, _x, n) in rows.items():
            n, t0 = int(n), base + (int(k[:2]) * 60 + int(k[3:])) * 60 * NS
            if n <= 0:
                continue
            sz, rem = divmod(int(vol), n)
            for j in range(n):
                ev.append((t0 + int((j + 0.5) * BAR_NS / n), e, 100.0 + 0.01 * (j % 2),
                           sz + (rem if j == n - 1 else 0)))
    ev.sort()
    got = {}
    runner.hooks.append(lambda b, bars, out: got.__setitem__(b, {e: (bars[e]["I"], bars[e]["J"]) for e in ETFS}))
    for ts, e, p, sz in ev:
        runner.on_trade(e, ts, p, sz)
    runner.flush(base + 86399 * NS)
    bad = tot = 0
    for e, rows in ref["etfs"].items():
        for k, (vol, _q, _x, n) in rows.items():
            b = base + (int(k[:2]) * 60 + int(k[3:])) * 60 * NS
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
        if BACKFILL_TODAY:
            try:
                backfill_today(runner, a.feed)
            except Exception as ex:   # noqa: BLE001  (mieux vaut démarrer sans rattrapage que pas du tout)
                log.warning("rattrapage impossible (%s) : démarrage live sans historique du jour", ex)
                runner.catchup_until = None
        try:
            _excel_feed_setup()      # montage hybride : actif seulement si EXCEL_FEED=1 (voir variables Railway)
            asyncio.run(run_live(runner, a.feed))
        except KeyboardInterrupt:
            runner.refbook.save()


if __name__ == "__main__":
    main()
