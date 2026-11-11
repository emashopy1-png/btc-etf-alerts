"""Port Python des colonnes CONFLUENCE (D, E, F, G, H, I, J, K, L, M, N, O, P) du script alpaca_excel_BTC_ETF_confluence_24h.py.

Chaque fonction reproduit la formule Excel de la feuille CONFLUENCE_<flux> (formules lues dans le .xlsx, ligne par ligne).
Règle d'or : la valeur de la bougie i n'utilise que les bougies <= i (aucune donnée future). Voir `check_no_lookahead.py`.

Conventions Excel reproduites :
  - cellule vide ou texte ("") -> None ; ISNUMBER(x) -> is_num(x) ; N(x) -> n(x) (texte = 0)
  - SUM/AVERAGE/MEDIAN/COUNT ignorent les textes et les vides ; un vide dans une opération arithmétique vaut 0
  - une cellule « "" » de la feuille est rendue par None
"""
from __future__ import annotations

import math
import statistics
from typing import Dict, List, Optional, Sequence

ETFS = ["IBIT", "FBTC", "BITB", "ARKB", "BTCO", "EZBC", "BRRR", "HODL", "BTCW", "MSBT", "GBTC", "BTC"]
F_EXCLUDED = {"BTCO", "BTCW"}      # la formule F (Speed of tape) n'additionne que 10 ETF : ni BTCO ni BTCW
BAR_MINUTES = 5                    # CL = trades / 5 dans les feuilles ETF

# --- constantes lues dans les formules ---
CVD_DECAY_I = 0.925875             # colonne I
CVD_DECAY_J = 0.890899             # colonne J
WIN = 60                           # fenêtre glissante de 60 bougies (J, AW, JR)
WC_CAP = 0.4                       # plafond de poids : 40 % (3 itérations, HP/HQ/HR)
MIN_TRADES_WIN = 20                # KP : au moins 20 trades sur 60 bougies
TVP_PRIOR_N = 10                   # KD : prior (0,5 pondéré 10)
TVP_MIN_OBS = 12                   # KD : au moins 12 observations validées
RX_MIN_ABS_IMB = 0.05              # JR : |carnet| >= 5 %
RX_MOVE_FRAC = 0.15                # JR : mouvement >= 15 % de l'amplitude moyenne
RX_LAG = 3                         # JR : on regarde 3 bougies en arrière
BD_MED_WIN = 12                    # KP : médiane de rotation des 12 bougies précédentes
CLIP_DELTA = 0.5                   # G : delta % / 0,5 borné à ±1
TIME_GATE_MIN = 9 * 60 + 30        # I, J : vides avant 9h30 NY

ERR = "#ERR"                       # erreur Excel (#VALUE!, #DIV/0!) : on la propage telle quelle


def is_num(x) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool) and not (isinstance(x, float) and math.isnan(x))


def n(x) -> float:
    """Fonction N() d'Excel : nombre -> lui-même, sinon 0."""
    return float(x) if is_num(x) else 0.0


def v0(x) -> float:
    """Valeur d'une cellule dans une opération arithmétique : vide -> 0 (un texte propagerait #VALUE!)."""
    if x is None:
        return 0.0
    if is_num(x):
        return float(x)
    raise ValueError("texte dans une opération arithmétique")


def sign(x: float) -> int:
    return (x > 0) - (x < 0)


def clip(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


def _nums(seq) -> List[float]:
    return [float(x) for x in seq if is_num(x)]


def avg(seq) -> Optional[float]:
    a = _nums(seq)
    return sum(a) / len(a) if a else None       # AVERAGE sans nombre -> #DIV/0 (géré par IFERROR côté appelant)


# ----------------------------------------------------------------------------------------------------------
# Feuilles ETF (formules de calcul des colonnes dérivées)
# ----------------------------------------------------------------------------------------------------------
def etf_O(I, L, M):
    """O : Delta % = IF(I=0,0,N/I) avec N = L-M."""
    if n(I) == 0 and (I is None or is_num(I)):
        return 0.0
    return (v0(L) - v0(M)) / v0(I)


def etf_AB(Y, Z):
    """AB : déséquilibre carnet moyen."""
    if Y is None or Z is None:
        return None
    if Y + Z == 0:
        return 0.0
    return (Y - Z) / (Y + Z)


def etf_AF(I, BA):
    """AF : ratio volume vs référence."""
    if BA is None or n(BA) <= 0:
        return None
    return v0(I) / BA


def etf_AW(F, G, K):
    """AW : amplitude / VWAP."""
    if F is None or G is None or K is None or K == 0:
        return None
    return (F - G) / K


def etf_CL(J):
    """CL : vitesse du tape (trades / min)."""
    return v0(J) / BAR_MINUTES


# ----------------------------------------------------------------------------------------------------------
# Moteur
# ----------------------------------------------------------------------------------------------------------
class Engine:
    """Calcule, bougie après bougie, les colonnes CONFLUENCE à partir des lignes des 12 feuilles ETF.

    Entrée de `push` : bars[sym] = dict avec les clés de colonnes de la feuille ETF
        F, G, H, I, J, K, L, M, Y, Z, AS, AT, BA, BB, BC, BD, BE, BJ, BN, CE, CF, CQ, CR, CU, CV
    (valeurs numériques ou None) et refs = {"EQ": réf. churn J-1 de la bougie, "ER": réf. tape de la bougie}.
    Les colonnes dérivées (O, AB, AF, AW, CL) sont recalculées ici, comme dans le classeur.
    `minutes_ny` = heure NY de début de la bougie, en minutes depuis minuit (sert à la porte de 9h30).
    """

    def __init__(self, etfs: Sequence[str] = ETFS):
        self.etfs = list(etfs)
        self.k = len(self.etfs)
        self.i = -1
        # séries par ETF (indexées par la bougie)
        z = lambda: {s: [] for s in self.etfs}
        self.H = z(); self.J = z(); self.AW = z(); self.BD = z()
        self.JF = z(); self.JR = z(); self.KD = z(); self.KP = z(); self.WC = z()
        # séries agrégées nécessaires aux fenêtres de 2 bougies (J) et au CVD (I, J)
        self.GJ: List[float] = []
        self.GK: List[float] = []
        self.Hconf: List[Optional[float]] = []
        self.I: List[Optional[float]] = []
        self.Jcvd: List[Optional[float]] = []

    # -- helpers fenêtres ------------------------------------------------------------------------------
    @staticmethod
    def _win(seq, i, size=WIN):
        return seq[max(0, i - size + 1): i + 1]

    def push(self, bars: Dict[str, dict], refs: dict, minutes_ny: int) -> dict:
        self.i += 1
        i = self.i
        E = self.etfs
        w: Dict[str, Optional[float]] = {}      # W.. : delta % de chaque ETF
        x: Dict[str, Optional[float]] = {}      # X.. : carnet % de chaque ETF
        af: Dict[str, Optional[float]] = {}
        cl: Dict[str, float] = {}
        for s in E:
            b = bars.get(s) or {}
            w[s] = etf_O(b.get("I"), b.get("L"), b.get("M"))
            x[s] = etf_AB(b.get("Y"), b.get("Z"))
            af[s] = etf_AF(b.get("I"), b.get("BA"))
            cl[s] = etf_CL(b.get("J"))
            self.H[s].append(b.get("H"))
            self.J[s].append(b.get("J"))
            self.AW[s].append(etf_AW(b.get("F"), b.get("G"), b.get("K")))
            self.BD[s].append(b.get("BD"))

        # --- JF (signal carnet de chaque ETF) : BN, sinon BE, sinon carnet moyen (X) ---
        for s in E:
            b = bars.get(s) or {}
            fl = b.get("BE") if is_num(b.get("BE")) else (x[s] if is_num(x[s]) else None)
            self.JF[s].append(b.get("BN") if is_num(b.get("BN")) else fl)

        # --- JR (le carnet d'il y a 3 bougies a-t-il vu juste ?) puis KD (taux de validation prix) ---
        for s in E:
            jf, h, aw = self.JF[s], self.H[s], self.AW[s]
            r = None
            if i >= RX_LAG:
                j0 = jf[i - RX_LAG]
                if is_num(j0) and is_num(h[i]) and is_num(h[i - RX_LAG]):
                    if abs(j0) >= RX_MIN_ABS_IMB and h[i - RX_LAG] > 0:
                        a = avg(self._win(aw, i))
                        if a is not None and abs(h[i] / h[i - RX_LAG] - 1) >= RX_MOVE_FRAC * a:
                            r = 1 if sign(j0) == sign(h[i] - h[i - RX_LAG]) else 0
            self.JR[s].append(r)
            win = _nums(self._win(self.JR[s], i))
            cnt = len(win)
            kd = 0.5 if cnt < TVP_MIN_OBS else (sum(1 for q in win if q == 1) + 0.5 * TVP_PRIOR_N) / (cnt + TVP_PRIOR_N)
            self.KD[s].append(kd)

        # --- KP (poids dynamique de chaque ETF) ---
        for s in E:
            b = bars.get(s) or {}
            jf = self.JF[s][i]
            kp = 0.0
            if is_num(jf) and sum(_nums(self._win(self.J[s], i))) >= MIN_TRADES_WIN:
                q = b.get("BB") if is_num(b.get("BB")) else 0.5
                stab = 1.0
                if i >= 1:
                    prev = _nums(self.BD[s][max(0, i - BD_MED_WIN): i])
                    bd = b.get("BD")
                    if is_num(bd) and len(prev) >= 3:
                        stab = min(1.0, statistics.median(prev) / bd)
                pers = b.get("BJ") if is_num(b.get("BJ")) else 0.5
                rel = (0.1 + 0.9 * max(0.0, min(1.0, (self.KD[s][i] - 0.5) / 0.15))) ** 2
                kp = q * stab * pers * rel
            self.KP[s].append(kp)

        # --- plafonnement des poids : 3 itérations à 40 % (HP, HQ, HR) -> WC (HD..HO) ---
        kps = [self.KP[s][i] for s in E]
        cap = sum(kps)
        for _ in range(3):
            cap = sum(min(p, WC_CAP * cap) for p in kps)
        # NB : HP = Σ min(KP, 0,4 ΣKP) ; HQ = Σ min(KP, 0,4 HP) ; HR = Σ min(KP, 0,4 HQ). `cap` = HR après 3 tours.
        wc = {s: min(self.KP[s][i], WC_CAP * cap) for s in E}
        for s in E:
            self.WC[s].append(wc[s])

        # --- GJ / GK : rendement et amplitude pondérés (servent au CVD prix, colonne J) ---
        gj = gk = 0.0
        for s in E:
            h = self.H[s]
            if i >= 1 and is_num(h[i]) and is_num(h[i - 1]) and wc[s] > 0:
                try:
                    gj += wc[s] * (h[i] / h[i - 1] - 1)
                except ZeroDivisionError:
                    pass
                a = avg(self._win(self.AW[s], i))
                if a is not None:
                    gk += wc[s] * a
        self.GJ.append(gj)
        self.GK.append(gk)

        out = {}
        # ---------------- D : ratio volume ----------------
        den = sum(v0((bars.get(s) or {}).get("BA")) for s in E if is_num(af[s]))
        num = sum(v0((bars.get(s) or {}).get("I")) for s in E if is_num(af[s]))
        out["D"] = None if den == 0 else num / den

        # ---------------- E : ratio churn carnet ----------------
        eq = refs.get("EQ")
        if n(eq) == 0:
            out["E"] = None
        else:
            try:
                out["E"] = sum(v0((bars.get(s) or {}).get("BC")) for s in E) / eq
            except ValueError:
                out["E"] = ERR

        # ---------------- F : speed of tape ----------------
        er = refs.get("ER")
        if n(er) == 0:
            out["F"] = None
        else:
            out["F"] = sum(cl[s] for s in E if s not in F_EXCLUDED) / er

        # ---------------- G : delta aligné pondéré volume (clip à ±1 de delta/0,5) ----------------
        vol_w = sum(v0((bars.get(s) or {}).get("I")) for s in E if is_num(w[s]))
        if vol_w == 0:
            out["G"] = None
        else:
            out["G"] = sum(clip(w[s] / CLIP_DELTA, -1, 1) * v0((bars.get(s) or {}).get("I")) for s in E if is_num(w[s])) / vol_w

        # ---------------- H : carnet % dynamique ----------------
        if sum(kps) == 0:
            hconf = None
        else:
            d = sum(wc[s] for s in E)
            nu = sum(self.JF[s][i] * wc[s] for s in E if is_num(self.JF[s][i]))
            hconf = nu / d
        self.Hconf.append(hconf)
        out["H"] = hconf

        # ---------------- I, J : CVD carnet (décroissance 0,925875) et CVD carnet prix (0,890899) ----------------
        if minutes_ny < TIME_GATE_MIN:
            self.I.append(None)
            self.Jcvd.append(None)
            out["I"] = out["J"] = None
        else:
            hv = hconf if is_num(hconf) else 0.0
            prev_i = self.I[i - 1] if i >= 1 and is_num(self.I[i - 1]) else 0.0
            ival = CVD_DECAY_I * prev_i + hv
            self.I.append(ival)
            out["I"] = ival
            prev_j = self.Jcvd[i - 1] if i >= 1 and is_num(self.Jcvd[i - 1]) else 0.0
            if is_num(hconf):
                s2 = 0.0
                for q in ((i - 1, i) if i >= 1 else (i,)):
                    s2 += (self.GJ[q] / self.GK[q]) if n(self.GK[q]) > 0 else 0.0
                add = hconf * clip(sign(hconf) * s2 / 1, -1, 1)
            else:
                add = 0.0
            jval = CVD_DECAY_J * prev_j + add
            self.Jcvd.append(jval)
            out["J"] = jval

        # ---------------- K : delta % de l'ETF au poids dynamique maximal ----------------
        if sum(kps) == 0:
            out["K"] = None
        else:
            idx = max(range(len(E)), key=lambda q: (kps[q], -q))   # MATCH(MAX(...),...,0) : premier maximum
            out["K"] = w[E[idx]] if is_num(w[E[idx]]) else None

        # ---------------- L : delta % dynamique ----------------
        den_l = sum(wc[s] for s in E if is_num(w[s]))
        out["L"] = None if den_l == 0 else sum(w[s] * wc[s] for s in E if is_num(w[s])) / den_l

        # ---------------- M : delta ISO ----------------
        iso = [(s, (bars.get(s) or {}).get("AT"), v0((bars.get(s) or {}).get("I"))) for s in E]
        den_m = sum(vol for _, at, vol in iso if is_num(at))
        out["M"] = None if den_m == 0 else sum(at * vol for _, at, vol in iso if is_num(at)) / den_m

        # ---------------- P : delta % exécuté >= 2 000 actions ----------------
        cv = sum(n((bars.get(s) or {}).get("CV")) for s in E)
        cr = sum(n((bars.get(s) or {}).get("CR")) for s in E)
        out["P"] = None if cv == 0 else cr / cv

        # ---------------- N : % volume ISO CERTAIN (Σ(CE+CF) / Σ(AS x volume)) ----------------
        den_n = sum(n((bars.get(s) or {}).get("AS")) * n((bars.get(s) or {}).get("I")) for s in E)
        out["N"] = None if den_n == 0 else sum(n((bars.get(s) or {}).get("CE")) + n((bars.get(s) or {}).get("CF")) for s in E) / den_n

        # ---------------- O : delta % exécuté trades 500-2 000 actions (ΣCQ / ΣCU) ----------------
        cu = sum(n((bars.get(s) or {}).get("CU")) for s in E)
        cq = sum(n((bars.get(s) or {}).get("CQ")) for s in E)
        out["O"] = None if cu == 0 else cq / cu
        return out
