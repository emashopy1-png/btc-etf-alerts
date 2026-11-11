"""Exécution des alertes confirmées sur compte DEMO Bybit (site principal, api-demo.bybit.com).

Paramètres validés (12/11) :
  - 3 perpétuels USDT : BTCUSDT, ETHUSDT, XRPUSDT — le signal BTC-ETF pilote les 3.
  - Levier visé : 150x BTC / 150x ETH / 100x XRP, PLAFONNÉ automatiquement au max de Bybit
    (lu via instruments-info ; ex. 100x / 100x / 75x).
  - Notional : 100 USDT par trade et par paire (marge = notional / levier).
  - TP = +100 % de ROI sur marge, SL = -40 % de ROI sur marge
    -> mouvement de prix = ROI / levier (ex. 100x : TP +1,00 %, SL -0,40 %), hors frais/funding.
  - On n'ouvre que si la paire est À PLAT. Même sens (M1 + M2) : position nette unique.
    Signal OPPOSÉ en position : ignoré (la position court jusqu'à son TP/SL).
  - Seuls les événements ALERT (confirmés) exécutent ; TRIGGER/CANCEL/EXPIRE = Telegram seul.

FILTRE D'ENTRÉE (validé 12/11) — « jamais en sens inverse » :
  - LONG uniquement si la dernière bougie 1 min CLOSE de la paire est VERTE (close > open) ;
    SHORT uniquement si elle est ROUGE (close < open). Doji (close = open) = pas d'entrée.
  - Si la bougie n'est pas dans le sens : ATTENTE (par paire) jusqu'à la 1re bougie 1 min qui
    ferme dans le sens, au maximum ENTRY_WAIT_MAX_S (défaut 600 s = 10 min), puis abandon.
  - Une nouvelle alerte REMPLACE l'attente en cours sur la même paire (sens et délai repartent).
  - Au moment de l'entrée, la position est re-vérifiée : si une position est apparue entre-temps
    (autre alerte), l'attente est abandonnée.
  - L'attente tourne dans un thread de fond : la boucle live n'est jamais bloquée.
  - En mode SIMULATION (EXEC_ENABLED absent), pas de connexion API : le filtre bougie n'est pas
    appliqué, le message 🧪 indique l'intention.

Variables d'environnement (Railway) :
  BYBIT_API_KEY, BYBIT_API_SECRET   (clés du mode Demo Trading du site principal)
  EXEC_ENABLED=1                     -> exécution réelle ; absent/autre = SIMULATION (log + Telegram 🧪)
  BYBIT_ENV=demo                     -> "demo" (défaut) ou "testnet"
  BYBIT_NOTIONAL_USDT=100            -> notional par trade et par paire
  ENTRY_CANDLE_CHECK=1               -> 0 pour désactiver le filtre bougie
  ENTRY_WAIT_MAX_S=600               -> attente max d'une bougie dans le sens (secondes)
  ENTRY_POLL_S=4                     -> période de surveillance des bougies (secondes)
Exposition max théorique : 3 paires x 100 USDT = 300 USDT de notional.
Les erreurs Bybit sont loguées et signalées sur Telegram ; elles ne tuent jamais la boucle live.
"""
from __future__ import annotations

import logging
import os
import threading
import time
from decimal import Decimal, ROUND_DOWN

log = logging.getLogger("bybit_exec")

# levier VISÉ par symbole (plafonné au maximum réel de la plateforme au démarrage)
SYMBOLS = {"BTCUSDT": 150, "ETHUSDT": 150, "XRPUSDT": 100}
NOTIONAL_USDT = Decimal(os.environ.get("BYBIT_NOTIONAL_USDT", "100"))
TP_ROI = Decimal("1.00")          # +100 % de ROI sur marge
SL_ROI = Decimal("0.40")          # -40 % de ROI sur marge
CANDLE_CHECK = os.environ.get("ENTRY_CANDLE_CHECK", "1") == "1"
WAIT_MAX_S = int(os.environ.get("ENTRY_WAIT_MAX_S", "600"))
POLL_S = int(os.environ.get("ENTRY_POLL_S", "4"))


# ------------------------------------------------------------------ calculs purs (testables hors-ligne)
def tp_sl_prices(side: str, entry: Decimal, lev: int):
    """Prix de TP/SL pour un ROI sur marge donné : prix = entrée x (1 ± ROI/levier)."""
    tp_move = TP_ROI / Decimal(lev)
    sl_move = SL_ROI / Decimal(lev)
    if side == "Buy":
        return entry * (1 + tp_move), entry * (1 - sl_move)
    return entry * (1 - tp_move), entry * (1 + sl_move)


def round_step(x: Decimal, step: Decimal) -> Decimal:
    """Arrondi vers le bas au pas de cotation/quantité (filtres Bybit)."""
    if step <= 0:
        return x
    return (x / step).to_integral_value(rounding=ROUND_DOWN) * step


def fmt(x: Decimal) -> str:
    return format(x.normalize(), "f")


def candle_color(o: Decimal, c: Decimal) -> str:
    return "green" if c > o else ("red" if c < o else "doji")


def candle_ok(color: str, side: str) -> bool:
    """Jamais en sens inverse : LONG = bougie verte, SHORT = bougie rouge."""
    return (color == "green" and side == "Buy") or (color == "red" and side == "Sell")


# ------------------------------------------------------------------ exécuteur
class Executor:
    def __init__(self, api=None):
        self.enabled = os.environ.get("EXEC_ENABLED", "0") == "1"
        self.api = api                       # injectable pour les tests ; None -> pybit réel
        self.info = {}                       # sym -> {"step","min","tick","lev"}
        self._ready = False
        self._lock = threading.Lock()
        self.pending = {}                    # sym -> {"side": Buy/Sell, "deadline": ts, "seen": début de bougie}
        self._watcher: threading.Thread | None = None
        self._wake = threading.Event()

    # ---- connexion et préparation (une fois)
    def _connect(self):
        if self.api is None:
            from pybit.unified_trading import HTTP
            env = os.environ.get("BYBIT_ENV", "demo").strip().lower()
            kw = {"testnet": True} if env == "testnet" else {"demo": True}
            self.api = HTTP(api_key=os.environ.get("BYBIT_API_KEY", ""),
                            api_secret=os.environ.get("BYBIT_API_SECRET", ""), **kw)
        for sym, target in SYMBOLS.items():
            r = self.api.get_instruments_info(category="linear", symbol=sym)
            it = self._ok(r, f"instruments {sym}")["result"]["list"][0]
            max_lev = int(Decimal(it["leverageFilter"]["maxLeverage"]))
            lev = min(target, max_lev)
            self.info[sym] = {"step": Decimal(it["lotSizeFilter"]["qtyStep"]),
                              "min": Decimal(it["lotSizeFilter"]["minOrderQty"]),
                              "tick": Decimal(it["priceFilter"]["tickSize"]),
                              "lev": lev}
            log.info("%s : levier visé %sx -> appliqué %sx (max Bybit %sx)", sym, target, lev, max_lev)
            try:
                self.api.switch_position_mode(category="linear", symbol=sym, mode=0)     # one-way
            except Exception as ex:      # noqa: BLE001  (déjà en one-way, ou position ouverte)
                log.info("%s : mode one-way non modifié (%s)", sym, ex)
            try:
                self.api.switch_isolated(category="linear", symbol=sym, tradeMode=1,
                                         buyLeverage=str(lev), sellLeverage=str(lev))     # marge isolée
            except Exception as ex:      # noqa: BLE001
                log.info("%s : marge isolée non appliquée (%s)", sym, ex)
            try:
                self._ok(self.api.set_leverage(category="linear", symbol=sym,
                                               buyLeverage=str(lev), sellLeverage=str(lev)),
                         f"levier {sym}")
            except Exception as ex:      # noqa: BLE001  (110043 = inchangé, acceptable)
                log.info("%s : set_leverage (%s)", sym, ex)
        self._ready = True

    @staticmethod
    def _ok(resp, what):
        if not isinstance(resp, dict) or resp.get("retCode") != 0:
            raise RuntimeError(f"{what} : {resp.get('retMsg') if isinstance(resp, dict) else resp}")
        return resp

    # ---- état du compte / marché
    def _position(self, sym):
        """-> (taille>0, 'Buy'/'Sell'/None, prix d'entrée). À plat = (0, None, None)."""
        r = self._ok(self.api.get_positions(category="linear", symbol=sym), f"position {sym}")
        for p in r["result"]["list"]:
            size = Decimal(p["size"] or "0")
            if size > 0:
                return size, p["side"], Decimal(p["avgPrice"] or "0")
        return Decimal(0), None, None

    def _last_price(self, sym):
        r = self._ok(self.api.get_tickers(category="linear", symbol=sym), f"prix {sym}")
        return Decimal(r["result"]["list"][0]["lastPrice"])

    def _last_closed_candle(self, sym):
        """-> (couleur, début de la bougie) de la dernière bougie 1 min CLOSE (list[1] : [0] = en cours)."""
        r = self._ok(self.api.get_kline(category="linear", symbol=sym, interval="1", limit=2),
                     f"kline {sym}")
        k = r["result"]["list"][1]
        return candle_color(Decimal(k[1]), Decimal(k[4])), k[0]

    # ---- exécution
    def _open(self, sym, side):
        """side = 'Buy' (LONG) / 'Sell' (SHORT). Ouvre au marché puis pose TP/SL en ROI marge."""
        inf = self.info[sym]
        price = self._last_price(sym)
        qty = round_step(NOTIONAL_USDT / price, inf["step"])
        if qty < inf["min"]:
            raise RuntimeError(f"quantité {fmt(qty)} < minimum {fmt(inf['min'])} pour {sym}")
        self._ok(self.api.place_order(category="linear", symbol=sym, side=side,
                                      orderType="Market", qty=fmt(qty), positionIdx=0),
                 f"ordre {sym}")
        _, _, entry = self._position(sym)
        if not entry:                                 # repli : dernier prix si la position tarde
            entry = price
        tp, sl = tp_sl_prices(side, entry, inf["lev"])
        tp, sl = round_step(tp, inf["tick"]), round_step(sl, inf["tick"])
        self._ok(self.api.set_trading_stop(category="linear", symbol=sym, tpslMode="Full",
                                           takeProfit=fmt(tp), stopLoss=fmt(sl),
                                           tpOrderType="Market", slOrderType="Market",
                                           positionIdx=0),
                 f"TP/SL {sym}")
        return qty, entry, tp, sl

    def _execute(self, sym, side):
        """Exécute (position vérifiée à plat à cet instant) + message Telegram."""
        qty, entry, tp, sl = self._open(sym, side)
        d = "LONG" if side == "Buy" else "SHORT"
        msg = (f"⚡ EXÉCUTION {d} {sym}\n"
               f"Quantité {fmt(qty)} (≈{fmt(NOTIONAL_USDT)} USDT) @ {fmt(entry)}\n"
               f"Levier {self.info[sym]['lev']}x — TP {fmt(tp)} (+100 % ROI) / SL {fmt(sl)} (-40 % ROI)")
        log.info(msg.replace("\n", " | "))
        _telegram(msg)

    # ---- filtre d'entrée : bougie 1 min dans le sens, attente max WAIT_MAX_S
    def _request_entry(self, sym, side, now=None):
        with self._lock:
            self.pending[sym] = {"side": side, "deadline": (now or time.time()) + WAIT_MAX_S,
                                 "seen": None}
            self._ensure_watcher()
            self._wake.set()

    def _drop(self, sym):
        with self._lock:
            self.pending.pop(sym, None)

    def _ensure_watcher(self):
        """Thread unique, jamais arrêté (daemon) : aucune course entre sortie et nouvelle attente.
        Appelé avec self._lock déjà pris."""
        if self._watcher is None or not self._watcher.is_alive():
            self._watcher = threading.Thread(target=self._watch_loop, name="bybit-entry-watch",
                                             daemon=True)
            self._watcher.start()

    def _watch_loop(self):
        while True:
            self._wake.wait(POLL_S)     # réveil immédiat à chaque nouvelle attente, sinon périodique
            self._wake.clear()
            with self._lock:
                alive = bool(self.pending)
            if not alive:
                continue
            try:
                self._watch_once(time.time())
            except Exception as ex:      # noqa: BLE001  (erreur transitoire : on réessaie au tour suivant)
                log.warning("surveillance entrées : %s", ex)

    def _watch_once(self, now):
        with self._lock:
            items = list(self.pending.items())
        for sym, p in items:
            if now > p["deadline"]:
                self._drop(sym)
                d = "LONG" if p["side"] == "Buy" else "SHORT"
                log.info("%s : attente expirée (%ds), entrée %s abandonnée", sym, WAIT_MAX_S, d)
                _telegram(f"⏱️ {sym} : aucune bougie 1 min dans le sens en {WAIT_MAX_S // 60} min, "
                          f"entrée {d} abandonnée")
                continue
            color, cstart = self._last_closed_candle(sym)
            if cstart == p["seen"]:
                continue
            p["seen"] = cstart
            if not candle_ok(color, p["side"]):
                log.info("%s : bougie 1 min %s, attente d'une bougie %s", sym, color,
                         "verte" if p["side"] == "Buy" else "rouge")
                continue
            size, pside, _ = self._position(sym)
            if size > 0:
                log.info("%s : position %s apparue pendant l'attente -> attente abandonnée", sym, pside)
                self._drop(sym)
                continue
            self._drop(sym)
            self._execute(sym, p["side"])

    # ---- point d'entrée appelé par live.py
    def on_alert(self, direction: str):
        """direction = 'LONG' / 'SHORT' (alerte confirmée, Mode 1 ou 2)."""
        side = "Buy" if direction == "LONG" else "Sell"
        for sym in SYMBOLS:
            try:
                size, pside, _ = self._position(sym)
                if size > 0 and pside == side:
                    log.info("%s : déjà en position %s -> ignoré (position nette unique)", sym, pside)
                    continue
                if size > 0 and pside != side:
                    log.info("%s : signal %s ignoré, position %s en cours (court jusqu'au TP/SL)",
                             sym, side, pside)
                    continue
                if CANDLE_CHECK:
                    self._request_entry(sym, side)
                    log.info("%s : alerte %s enregistrée, attente d'une bougie 1 min %s (max %ds)",
                             sym, direction, "verte" if side == "Buy" else "rouge", WAIT_MAX_S)
                else:
                    self._execute(sym, side)
            except Exception as ex:      # noqa: BLE001
                log.exception("exécution %s %s : %s", direction, sym, ex)
                _telegram(f"⚠️ ÉCHEC exécution {direction} {sym} : {ex}")

    def on_events(self, events):
        for e in events:
            if getattr(e, "kind", None) != "ALERT":
                continue
            if not self.enabled:
                _telegram(f"🧪 [SIMU] ouvrirait {e.direction} sur BTCUSDT/ETHUSDT/XRPUSDT "
                          f"({e.mode} {e.ny} NY, bougie {e.bar}) — EXEC_ENABLED=1 pour exécuter")
                continue
            try:
                if not self._ready:
                    self._connect()
                self.on_alert(e.direction)
            except Exception as ex:      # noqa: BLE001
                log.exception("executor : %s", ex)
                _telegram(f"⚠️ Erreur exécuteur Bybit : {ex}")


def _telegram(text):
    try:
        import alerts
        alerts.send_telegram(text)
    except Exception as ex:      # noqa: BLE001
        log.warning("telegram : %s", ex)


_EXEC: Executor | None = None


def on_events(events):
    """Hook appelé par live.py avec les événements de la bougie close."""
    global _EXEC
    if _EXEC is None:
        _EXEC = Executor()
    _EXEC.on_events(events)
