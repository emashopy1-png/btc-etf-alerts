"""Exécution des alertes confirmées sur compte DEMO Bybit (site principal, api-demo.bybit.com).

Paramètres validés (12/11) :
  - 3 perpétuels USDT : BTCUSDT, ETHUSDT, XRPUSDT — le signal BTC-ETF pilote les 3.
  - Levier visé : 150x BTC / 150x ETH / 100x XRP, PLAFONNÉ automatiquement au max de Bybit
    (lu via instruments-info ; ex. 100x / 100x / 75x).
  - Notional : 100 USDT par trade et par paire (marge = notional / levier).
  - TP = +100 % de ROI sur marge, SL = -40 % de ROI sur marge
    -> mouvement de prix = ROI / levier (ex. 100x : TP +1,00 %, SL -0,40 %), hors frais/funding.
  - On n'ouvre que si la paire est À PLAT. Même sens (M1 + M2) : position nette unique, pas de
    renfort. Signal OPPOSÉ en position : ignoré (la position court jusqu'à son TP/SL).
  - Seuls les événements ALERT (confirmés) exécutent ; TRIGGER/CANCEL/EXPIRE = Telegram seul.

Variables d'environnement (Railway) :
  BYBIT_API_KEY, BYBIT_API_SECRET   (clés du mode Demo Trading du site principal)
  EXEC_ENABLED=1                     -> exécution réelle ; absent/autre = SIMULATION (log + Telegram 🧪)
  BYBIT_ENV=demo                     -> "demo" (défaut) ou "testnet"
  BYBIT_NOTIONAL_USDT=100            -> notional par trade et par paire
Exposition max théorique : 3 paires x 100 USDT = 300 USDT de notional.
Les erreurs Bybit sont loguées et signalées sur Telegram ; elles ne tuent jamais la boucle live.
"""
from __future__ import annotations

import logging
import os
from decimal import Decimal, ROUND_DOWN

log = logging.getLogger("bybit_exec")

# levier VISÉ par symbole (plafonné au maximum réel de la plateforme au démarrage)
SYMBOLS = {"BTCUSDT": 150, "ETHUSDT": 150, "XRPUSDT": 100}
NOTIONAL_USDT = Decimal(os.environ.get("BYBIT_NOTIONAL_USDT", "100"))
TP_ROI = Decimal("1.00")          # +100 % de ROI sur marge
SL_ROI = Decimal("0.40")          # -40 % de ROI sur marge


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
    s = format(x.normalize(), "f")
    return s


# ------------------------------------------------------------------ exécuteur
class Executor:
    def __init__(self, api=None):
        self.enabled = os.environ.get("EXEC_ENABLED", "0") == "1"
        self.api = api                       # injectable pour les tests ; None -> pybit réel
        self.info = {}                       # sym -> {"step": Decimal, "min": Decimal, "tick": Decimal, "lev": int}
        self._ready = False

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

    # ---- état du compte
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
                qty, entry, tp, sl = self._open(sym, side)
                msg = (f"⚡ EXÉCUTION {direction} {sym}\n"
                       f"Quantité {fmt(qty)} (≈{fmt(NOTIONAL_USDT)} USDT) @ {fmt(entry)}\n"
                       f"Levier {self.info[sym]['lev']}x — TP {fmt(tp)} (+100 % ROI) / SL {fmt(sl)} (-40 % ROI)")
                log.info(msg.replace("\n", " | "))
                _telegram(msg)
            except Exception as ex:      # noqa: BLE001
                log.exception("exécution %s %s : %s", direction, sym, ex)
                _telegram(f"⚠️ ÉCHEC exécution {direction} {sym} : {ex}")

    # ---- point d'entrée appelé par live.py
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
