"""
VERSION 24h/24 – ETF BITCOIN SPOT UNIQUEMENT (12 ETF) – v2 « qualité du carnet ».
La fenêtre couvre toute la journée New York (00:00 -> 24:00).
v7 : les références « normales » des colonnes D (ratio volume), E (ratio churn) et F (écart déséquilibre) ne sont plus J-1 :
c'est la MÉDIANE, tranche par tranche (même heure New York), de toutes les séances (ou des HIST_DAYS plus proches si limité) lues dans le dossier HIST_DIR
(Excel quotidiens déjà produits par ce script, feuilles du flux traité, ex. IEX). Aucun téléchargement d'historique ; l'Excel du jour
est copié dans HIST_DIR en fin d'exécution et alimente les calculs du lendemain.
Mettre FULL_DAY = False pour revenir à la séance 9h30-16h.

Un fichier Excel, pour IEX ET SIP (mêmes formules partagées) :
  - feuille 1 : CONFLUENCE_IEX      feuille 2 : CONFLUENCE_SIP
  - puis une feuille par ETF et par flux : IBIT_IEX, FBTC_IEX, ... (12) puis IBIT_SIP, FBTC_SIP, ... (12)
  - une feuille GUIDE à la fin

ETF suivis : IBIT, FBTC, BITB, ARKB, BTCO, EZBC, BRRR, HODL, BTCW, MSBT, GBTC, BTC (voir ETFS).

Feuilles ETF, colonnes A -> BI : données brutes + delta + carnet + pression (A -> AH), indicateurs avancés
(AI -> AZ : taille moyenne, ratios 12 bougies, score whale, score flux...), référence J-1 (BA), puis les 8 indicateurs
v2 « qualité du carnet » (BB -> BI : coef qualité, churn, rotation des tailles, déséquilibre pondéré temps, impact
post-trade, score absorption, confirmation carnet -> exécution, verdict carnet). AS / AT = % volume ISO et Delta ISO
(signé), ajoutés sans toucher au delta existant. Les colonnes BB -> BH sont calculées en seconde passe sur les quotes
et trades bruts de CHAQUE flux (IEX et SIP séparément) ; BI est une formule.

Feuilles CONFLUENCE_<flux> : les ETF côte à côte, bougie par bougie (delta % et carnet % de chacun), puis
delta % / carnet % moyens, variation de l'Open Interest BTC agrégé (Coinalyze, perp + futures, 5 min),
confluence DELTA + CARNET, confluence FLUX GLOBAL, et un bloc « moyenne des ETF » qui reprend les indicateurs
AG -> AZ des feuilles ETF (le ratio volume J-1 moyen est placé juste après les 12 blocs ETF ; heures UTC et New York en B et C).
Bloc final v2 : ETF divergent + écart max + VERDICT DIVERGENCE (12 colonnes helper masquées _gap_*), carnet qualité moyen,
carnet % qualité-pondéré, absorption moyenne, confirmation carnet moyenne, enchères 9h30 / 16h00 (si disponibles).
Récapitulatifs en bas.
La colonne OI nécessite COINALYZE_API_KEY ; sans clé elle reste vide.
v12 : nouvelle colonne « CVD carnet prix » à droite de « CVD carnet » (qui reste inchangée) : même carnet dynamique, mais chaque bougie
n'est créditée que dans la mesure où le PRIX a réagi dans son sens ; carnet affiché sans réaction du prix = 0, prix à contre-sens = contribution inversée
(absorption). Paramètres PX_* en haut du script.
v3 « taille affichée crédible » (feuilles ETF BJ -> CD, à droite de BI ; rien de l'existant n'est modifié) : persistance des
états du carnet, murs anormaux, déséquilibre temps x persistance, déséquilibre par classe de taille, consommation / annulation
des retraits de taille, alerte « piège retiré -> sweep », écart du déséquilibre vs médiane J-1 (même tranche). Feuilles
CONFLUENCE : agrégats pondérés volume juste à droite de « Déséquilibre pondéré temps » (colonne M), dont une version pondérée
volume x fiabilité et l'écart IEX vs SIP.
v4/v5 « carnet dynamique » (feuilles CONFLUENCE) : chaque ETF est pondéré par son TAUX DE VALIDATION PRIX (le carnet de la bougie j
a-t-il annoncé le sens du prix à j+3 ?) x qualité d'affichage (spread, stabilité des tailles, persistance), SANS le volume exécuté ;
plafond par ETF, CVD carnet à demi-vie, porte de validation et verdict carnet final. Un ETF dont le carnet ment (ex. IBIT) perd son poids.

IMPORTANT : une « whale » n'est pas identifiable avec des trades agrégés. Le score détecte
une activité anormale compatible avec de gros ordres ; il ne prouve ni l'identité ni l'intention
d'un acteur. Le flux IEX est un échantillon ; le SIP offre une couverture plus complète.

Installation :  pip install alpaca-py pandas openpyxl requests
Clés Windows :  set "ALPACA_API_KEY=..." et set "ALPACA_SECRET_KEY=..."
                Les noms APCA_API_KEY_ID / APCA_API_SECRET_KEY sont aussi acceptés.
Usage        :  python alpaca_excel_BTC_ETF_confluence_24h.py                    -> date par défaut, IEX + SIP
                python alpaca_excel_BTC_ETF_confluence_24h.py 2026-10-26         -> autre date, IEX + SIP
                python alpaca_excel_BTC_ETF_confluence_24h.py 2026-10-26 iex     -> IEX seulement
                python alpaca_excel_BTC_ETF_confluence_24h.py 2026-10-26 sip     -> SIP seulement
Si le SIP est refusé par ton abonnement (403), le script l'ignore et garde l'IEX.
"""
import json
import os
import random
import re
import shutil
import sys
import warnings
import time
from collections import deque
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

# ----------------------------------------------------------------- config
# ETF Bitcoin spot suivis (ordre du tableau de flux). IBIT est obligatoire : s'il échoue, le script s'arrête ;
# un autre ETF sans données (pas coté, trop récent, trop peu liquide) est simplement ignoré.
ETFS = ["IBIT", "FBTC", "BITB", "ARKB", "BTCO", "EZBC", "BRRR", "HODL", "BTCW", "MSBT", "GBTC", "BTC"]
MANDATORY = "IBIT"
if len(ETFS) != len(set(ETFS)):
    raise SystemExit("Config ETF invalide : doublon dans ETFS.")
FEEDS = ["IEX"]                   # flux traités (IEX = gratuit/échantillon ; SIP = marché complet, selon plan)
BAR_MINUTES = 5                          # 1 bougie = 5 min (288 bougies sur 24h ; 78 si séance seule) ; mets 1 pour des bougies 1 min
FULL_DAY = True                          # True = 24h/24 (00:00-24:00 New York) ; False = séance US 9h30-16h
NY = ZoneInfo("America/New_York")

# ---- Chemins : configurables par variables d'environnement (Railway : monter un Volume sur /data) ----
def _default_dir(env_name, volume_subdir, local_dir, windows_fallback=None):
    v = os.environ.get(env_name)
    if v and v.strip():
        return v.strip()
    if os.path.isdir("/data"):                       # Volume Railway monté sur /data
        return os.path.join("/data", volume_subdir)
    if windows_fallback and os.name == "nt" and os.path.isdir(os.path.dirname(windows_fallback)):
        return windows_fallback                      # PC Windows : comportement d'origine
    return local_dir

OUT_DIR = _default_dir("DATA_DIR", "data_alpaca", "data_alpaca")
DEFAULT_DAY = "2026-10-23"               # journée traitée par défaut (AAAA-MM-JJ) en mode batch ; ignorée en mode live

# ---- Mode LIVE (python ...py live) : bougie par bougie, alertes + push GitHub ----
LIVE_MODE = False                        # activé par l'argument « live » (ne pas modifier ici)
LIVE_SETTLE_S = int(os.environ.get("LIVE_SETTLE_S", "45"))   # délai après la clôture de bougie avant calcul (laisse Alpaca stabiliser la donnée)
LIVE_SESSION_ONLY = os.environ.get("LIVE_SESSION_ONLY", "1") == "1"   # 1 = ne calcule que la séance d'alerte 9h30-14h00 NY en live
                                                                      # (grosse économie de RAM/coût : ~4,5 h de données au lieu de 24 h ; 0 = journée complète)
LIVE_ALERTS = os.environ.get("LIVE_ALERTS", "0") == "1"      # montage hybride : les alertes restent sur le bot (il lit le JSON) ; 1 = le service Excel alerte lui-même
GITHUB_REPO = os.environ.get("GITHUB_REPO", "").strip()      # ex. "emashopy1-png/btc-etf-alerts" ; vide = pas de push
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "").strip()    # token GitHub (classic, droit « repo »)
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "live-data").strip() or "live-data"   # branche dédiée aux données (main reste propre)
GIT_DIR = _default_dir("GIT_DIR", "repo", "repo_clone")      # clone local du dépôt (sur le Volume = survit aux redéploiements)

# Seuils de la colonne "Pression" :
SEUIL_DELTA = 0.10      # 10 %  (Delta %, colonne O)
SEUIL_CARNET = 0.10     # 10 %  (Déséquilibre carnet, colonne AB)

# Alertes de la colonne "Ratio volume vs moyenne J-1"
RATIO_ALERTE = 1.8       # recalibré : le ratio est maintenant mesuré vs la MÊME tranche horaire de J-1 (médiane ~1)
RATIO_FORT = 2.5
SIGNAL_RATIO_J1 = RATIO_ALERTE   # seuil du ratio J-1 dans SIGNAL FLUX GLOBAL (colonnes AZ / CZ)
# ---- v7 : historique local (dossier d'Excel quotidiens) ----
HIST_DIR = _default_dir("HIST_DIR", "historique", "historique",
                        windows_fallback=r"C:\Users\HP\Downloads\Historique Alpaca")   # dossier des Excel quotidiens (lu au démarrage, alimenté en fin d'exécution)
HIST_DAYS = None          # None = TOUTES les séances valides du dossier HIST_DIR (médiane par tranche horaire) ; un entier (ex. 8) limite aux N plus proches
HIST_CONF_FULL = min(8, HIST_DAYS) if HIST_DAYS else 8   # nb de séances à partir duquel la confiance du z-score carnet (CONFLUENCE, colonne « Z-score déséquilibre vs réf. normale ») atteint 100 % (inchangé : 8)
HIST_LABEL = "tout l'historique" if HIST_DAYS is None else f"{HIST_DAYS} j"   # libellé des en-têtes (colonnes D / E)
HIST_MIN_DAYS = 3         # minimum de séances valides ; en dessous -> repli sur J-1 téléchargé chez Alpaca (comportement d'avant)
HIST_ALLOW_AFTER = True   # analyse d'une date ANCIENNE : si l'historique AVANT cette date est insuffisant (< HIST_MIN_DAYS), on complète avec les séances les plus proches APRÈS (le jour analysé est toujours exclu)
# ---- v8 : z-score robuste de l'écart du carnet vs historique (feuille CONFLUENCE : colonne « Z-score déséquilibre vs réf. normale », à droite des blocs ETF) ----
Z_MAD_FLOOR = 0.02        # plancher de la dispersion (2 points de %) : évite les z explosifs sur les carnets ultra-stables
Z_EMA_ALPHA = 0.5         # lissage EMA du z agrégé (1 = aucune mémoire ; 0,3 = mémoire d'environ 2 bougies)
Z_ANOMALIE = 1.2          # |z lissé| >= 1,2 sur 2 bougies d'affilée : « anormalement acheteur / vendeur » (environ P90 des |z| sur 8 jours d'IEX)
Z_FORT = 1.75             # |z lissé| >= 1,75 sur 2 bougies d'affilée : « très » (environ P97)
HIST_SAVE_TODAY = True    # copie l'Excel du jour dans HIST_DIR en fin d'exécution (il servira de référence demain)
REF_SMOOTH = 1          # lissage de la référence J-1 : moyenne de la tranche et de ses ±1 voisines
REF_FLOOR_FRAC = 0.10   # plancher de la référence = 10 % du volume moyen par bougie de J-1 (évite les ratios explosifs la nuit)
CHURN_RATIO_ALERTE = 1.8   # ratio churn carnet vs J-1 (même tranche) : alerte (orange clair) - à recalibrer après quelques jours
CHURN_RATIO_FORT = 2.5     # fort (orange foncé)
# ---- v9 : SPEED OF TAPE (vitesse du ruban = trades / minute) vs la médiane historique de la MÊME heure ----
TAPE_RATIO_ALERTE = 1.8    # ratio vitesse tape vs réf. normale : alerte (orange clair)
TAPE_RATIO_FORT = 2.5      # fort (orange foncé)
TAPE_REF_FLOOR_FRAC = 0.10 # plancher de la réf. = 10 % de la vitesse moyenne de l'historique (évite les ratios explosifs la nuit)
CHURN_REF_FLOOR_FRAC = 0.10   # plancher de la réf. churn J-1 = 10 % du churn moyen par bougie de J-1 (évite les ratios explosifs la nuit)

# Paramètres des indicateurs avancés (fenêtre glissante, hors bougie courante) :
WHALE_WINDOW = 12
WHALE_MIN_HISTORY = 3
WHALE_MIN_TRADES = 3
WHALE_ALERTE = 50
WHALE_FORT = 75
WHALE_CONCENTRATION_MAX = 0.25
WHALE_AVG_SIZE_MAX = 3.0
WHALE_MAX_TRADE_MAX = 4.0
WHALE_VOLUME_MAX = 3.0
FLUX_VWAP_SCALE = 0.005

# ---- v6 : direction CERTAINE des trades (règle bid/ask au lieu du milieu + détection de sweep ISO) ----
SWEEP_GAP_MS = 100          # trades ISO séparés de moins de 100 ms = même rafale (sweep)
SWEEP_MIN_LEVELS = 3        # un sweep doit traverser au moins 3 niveaux de prix distincts
# conditions de trade « non régulières » (prix moyen, hors séquence, dérivé, référence...) : jamais classées « certaines » (liste modifiable)
NON_REGULAR_CONDS = {"B", "C", "N", "R", "V", "W", "Z", "U", "4", "P", "M", "Q"}

# Confluence des ETF : un ETF est "acheteur" si son score flux >= seuil, "vendeur" si <= -seuil
CONF_SEUIL = 20
# Feuilles CONFLUENCE_<flux> (N ETF) : parts d'ETF alignés (net acheteurs − vendeurs, divisé par le nb d'ETF avec donnée)
COIN_CONF_FORT = 0.75     # >= 75 % des ETF alignés : confluence forte
COIN_CONF_SIMPLE = 0.50   # >= 50 % : confluence simple
DC_FORT = 0.80            # delta + carnet : >= 80 % du maximum possible = fort
DC_SIMPLE = 0.50          # >= 50 % = simple
ALIGN_SCALE = 0.50        # Delta / Carnet aligné pondéré : valeur de l'ETF = % ÷ 50 %, limitée à ±1
DCW_FORT = 0.25           # verdict pondéré : moyenne (delta aligné pondéré + carnet aligné pondéré)/2 >= 25 % = fort
DCW_SIMPLE = 0.15         # >= 15 % = simple (et seuil de « divergence »)

# ---- nouveautés v2 : qualité du carnet, churn, absorption, confirmation, divergence ----
BOOK_QUALITY_WINDOW = 12    # BB : médiane mobile du spread de l'ETF sur 12 bougies (bougie courante exclue)
BOOK_MIN_HISTORY = 3        # BB : nb minimum de bougies d'historique avec un spread valide
TW_CAP_SECONDS = 300        # BE : plafond de durée retenue pour un même état du carnet (5 min)
IMPACT_HORIZON_S = 2        # BF : horizon du déplacement du mid après un trade (secondes)
TRACK_FLOOR = 0.10         # composite : poids plancher du « track record » (BH <= VERDICT_PEU_FIABLE)
TRACK_NEUTRAL = 0.50       # composite : track record neutre quand BH est vide (historique insuffisant)
# ---- v4/v5 : carnet DYNAMIQUE pondéré par le TAUX DE VALIDATION PRIX (TVP) de chaque ETF (feuilles CONFLUENCE) ----
# Le volume exécuté n'entre PAS dans le poids du carnet (il décrit l'exécution, pas la qualité de l'affichage).
REACT_H = 3                # horizon de validation : le carnet de la bougie j est jugé sur le prix à j+H (3 x 5 min = 15 min)
REACT_TRIGGER = 0.05       # le carnet ne « parle » que si |déséquilibre x persistance (BN)| >= 5 % (à défaut BE, puis Carnet %)
REACT_THR_AMP = 0.15       # mouvement de prix significatif = 0,15 x amplitude moyenne des bougies (colonne AW, sur la fenêtre)
REACT_WINDOW = 60          # nb de bougies de la fenêtre glissante de mesure (les signaux jugés y sont comptés)
REACT_MIN_CASES = 12       # nb minimum de signaux jugés avant de sortir de la fiabilité neutre (50 %)
REACT_PRIOR = 10           # pseudo-observations à 50 % (lissage : peu de cas => TVP proche de 50 %)
REACT_FULL = 0.65          # TVP lissé à partir duquel la fiabilité vaut 100 % (0 à 50 %)
REACT_FLOOR = 0.10         # fiabilité plancher
REACT_GAMMA = 2            # exposant sur la fiabilité-prix
REACT_MIN_TRADES = 20      # éligibilité (pas un poids) : nb de trades de l'ETF sur la fenêtre, sinon exclu du vote carnet
REACT_CAP = 0.40           # plafond de contribution d'un ETF au carnet agrégé (une passe, renormalisée)
REACT_HALF_LIFE_MIN = 45   # demi-vie (minutes) du CVD carnet dynamique : un carnet ne pèse pas plus d'1-2 h
REACT_GATE = 0.55          # le carnet agrégé n'est « validé » que si le TVP moyen pondéré des contributeurs est >= 55 %
REACT_VERDICT = 0.15       # carnet dynamique >= +15 % : ACHAT ; <= -15 % : VENTE (si validé)
# ---- v12 : « CVD carnet prix » (nouvelle colonne à côté de « CVD carnet », qui n'est PAS modifiée) ----
# Chaque bougie : contribution = carnet dynamique x k, k = réaction du prix dans le sens du carnet, bornée entre -1 et +1 :
#   k = +1 : le prix a bougé dans le sens du carnet d'au moins PX_TAU x amplitude moyenne (carnet confirmé) ;
#   k =  0 : le prix n'a pas bougé (carnet affiché mais sans effet : ignoré) ;
#   k = -1 : le prix a bougé à CONTRE-SENS (absorption / carnet qui ment) : la contribution change de signe.
# Le prix est celui des ETF, pondéré par les mêmes poids dynamiques que le carnet, mesuré sur PX_BARS bougies (courante incluse, aucune donnée future).
PX_BARS = 2                # nb de bougies sur lesquelles on mesure la réaction du prix (la bougie courante + les précédentes)
PX_TAU = 1.0               # mouvement cumulé (en x amplitude moyenne d'une bougie) qui confirme à 100 % : 1,0 sur 2 bougies = 0,5 par bougie
PX_HALF_LIFE_MIN = 30      # demi-vie (min) du CVD carnet prix : plus courte que celle de CVD carnet (45 min) pour sortir plus vite d'un retournement
# Source du DELTA dynamique (colonne « Delta % dynamique », celle qui entre dans la confluence DELTA + CARNET) :
#   "carnet" = mêmes ETF et mêmes poids (_wc_, plafonnés) que le carnet dynamique -> « Delta le plus fiable » = « Carnet le plus fiable »
#   "top"    = delta du seul ETF dont le carnet est le plus fiable (et le TVP delta de cet ETF)
#   "own"    = ancien calcul : poids propres au delta (racine du volume x fiabilité du delta)
DELTA_SOURCE = "carnet"
# même exercice pour le DELTA : TVP delta = le Delta % de la bougie j (>= SEUIL_DELTA) a-t-il annoncé le sens du close à j+H ?
# poids delta = racine(volume de la fenêtre) x fiabilité-prix^GAMMA (le volume décrit l'exécution : il garde un rôle modéré), plafond REACT_CAP.
# ---- organisation des feuilles CONFLUENCE_<flux> (les feuilles ETF ne changent pas) ----
# Colonnes d'analyse placées juste après « Heure New York », dans cet ordre ; toutes les autres colonnes (blocs ETF,
# moyennes des ETF, divergence, enchères...) sont décalées à droite, dans leur ordre d'origine. [] = ancienne organisation.
# Clés : ratio, ndw, cvd_d, ncw, cvd_c, cvd_sip, iso_p, iso_d, oi, nd, nc, vdc, vdcw, avg, net, wh, ver, dm, cm,
#        div_etf, div_gap, div_ver, qm, qc, abs, conf, auc_o, auc_c, auc_r, chn, rot, imb, imp, vdq, qcc,
#        a33 ... a52 (colonnes « moyenne des ETF » = colonne AG ... AZ des feuilles ETF : a33 = AG, a34 = AH, ...).
# ---- v11 : ORGANISATION DE LA FEUILLE CONFLUENCE (modifiable) ----
# Colonnes placées juste après « Heure New York », dans CET ordre, avec CET en-tête et CETTE largeur (en caractères Excel).
# Toutes les autres colonnes (z-score, blocs ETF, moyennes, divergence, enchères, colonnes d'aide masquées...) sont décalées à droite,
# dans leur ordre d'origine : la feuille couvre toujours l'ensemble des éléments. Jetons : {N} = nb d'ETF ; {T3} / {T4} = libellés des classes de taille de trade.
# Une clé absente de la feuille (ex. « oi » sans Coinalyze) est ignorée.
CONFLUENCE_LAYOUT = [
    ("ratio",     "Ratio volume", 9.5),
    ("chn_ratio", "Ratio churn carnet", 13),
    ("tape_ratio", "Speed of tape", 7.5),
    ("ndw",       "Delta aligné pondéré", 8),
    ("dyn_c",     "Carnet % dynamique", 10),
    ("dyn_cvd",   "CVD carnet", 9),
    ("dyn_cvd_px", "CVD carnet prix", 10),
    ("dtop_val",  "Delta % le plus fiable", 8.5),
    ("dd_c",      "Delta % dynamique", 11.5),
    ("iso_d",     "Delta ISO", 6),
    ("iso_pc",    "% volume ISO CERTAIN ({N} ETF BTC)", 8),
    ("tx_cls3",   "Delta % EXÉCUTÉ – trades {T3} actions (Σ {N} ETF)", 10),
    ("tx_cls4",   "Delta % EXÉCUTÉ {T4}", 10),
    ("ncw",       "Carnet aligné", 6.5),
    ("oi",        "OI BTC agrégé (var. / bougie)", 7.5),
    ("nd",        "Delta aligné", 7.5),
    ("imb_ema",   "Z-score EMA (mémoire)", 7),
    ("imb_ver",   "Carnet vs historique (persistance 2 bougies)", 28),
]
CONFLUENCE_FRONT = [k for k, _, _ in CONFLUENCE_LAYOUT]
CONF_ABC_WIDTHS = (9, 8, 9.5)    # largeurs de « N° bougie », « Heure UTC », « Heure New York »
CONF_HEADER_HEIGHT = 118         # hauteur de la ligne d'en-tête (points)
CONFIRM_WINDOW = 36         # BH : nb de bougies (résultats) examinées pour la confirmation carnet -> exécution
CONFIRM_MIN_SAMPLES = 5     # BH : nb minimum de cas (carnet déséquilibré >= SEUIL_CARNET) pour afficher un taux
VERDICT_FIABLE = 0.60       # BI : « Fiable » si qualité >= 60 % ET confirmation >= 60 %
VERDICT_PEU_FIABLE = 0.40   # BI : « Peu fiable » si qualité OU confirmation < 40 % (sinon « Moyen »)
DIV_GAP = 0.15              # CONFLUENCE : écart minimum (15 pts) entre le delta % d'un ETF et le consensus
DIV_CONSENSUS_MIN = 0.10    # CONFLUENCE : |consensus| minimum (10 %) pour parler de divergence
REF_COL = 53                # colonne BA des feuilles ETF (référence J-1) : ne change pas quand on ajoute des colonnes après


# ---- v3 : taille affichée « crédible », murs, classes de taille, destin des murs, baseline J-1 ----
REF_WINDOW_BARS = 12        # médianes mobiles (durée d'un état, tailles bid / ask) : 12 bougies précédentes, bougie courante exclue
REF_MIN_STATES = 20         # nb minimum de quotes / d'états dans cette fenêtre pour calculer une médiane
WALL_MULT = 3.0             # mur anormal : taille d'un côté > 3 x la médiane de ce côté
WALL_GONE_MULT = 1.0        # un mur est « retiré » quand sa taille retombe sous 1 x la médiane (ou que le niveau de prix disparaît)
SIZE_CLASSES = (1_000, 10_000, 50_000)   # bornes (actions) des 4 classes : < 1k / 1k-10k / 10k-50k / >= 50k
PULL_WINDOW_S = 30          # alerte piège : retrait dans les 30 dernières secondes de la bougie
PULL_PRESENCE = 0.60        # alerte piège : mur présent >= 60 % de la bougie
PULL_CANCEL_FRAC = 0.50     # alerte piège : le retrait doit être au moins à 50 % une annulation (non consommée par des trades)
BOOK_J1_BASELINE = True     # télécharge les quotes de J-1 (12 ETF x flux) pour la réf. J-1 du déséquilibre ; False = colonnes vides


# ---- v10 : delta EXÉCUTÉ par taille de TRADE (volume réellement échangé, pas les ordres affichés qui apparaissent / disparaissent) ----
TRADE_SIZE_CLASSES = (100, 500, 2_000)   # bornes (actions par trade) des 4 classes : < 100 / 100-500 / 500-2 000 / >= 2 000


def _trade_class_labels():
    b = TRADE_SIZE_CLASSES
    f = lambda x: f"{x:,}".replace(",", " ")
    return [f"< {f(b[0])}", f"{f(b[0])}–{f(b[1])}", f"{f(b[1])}–{f(b[2])}", f"≥ {f(b[2])}"]


def _class_labels():
    b = SIZE_CLASSES
    f = lambda x: f"{x:,}".replace(",", " ")
    return [f"< {f(b[0])}", f"{f(b[0])}–{f(b[1])}", f"{f(b[1])}–{f(b[2])}", f"≥ {f(b[2])}"]


class FeedUnavailable(Exception):
    """Flux indisponible (ex. SIP non inclus dans l'abonnement)."""


# ----------------------------------------------------------------- fenêtre
def day_bounds(d):
    """Début et fin (UTC) de la fenêtre analysée pour la date d (heure de New York)."""
    if FULL_DAY:
        start = datetime(d.year, d.month, d.day, 0, 0, tzinfo=NY).astimezone(timezone.utc)
        nxt = d + timedelta(days=1)
        close = datetime(nxt.year, nxt.month, nxt.day, 0, 0, tzinfo=NY).astimezone(timezone.utc)
    else:
        start = datetime(d.year, d.month, d.day, 9, 30, tzinfo=NY).astimezone(timezone.utc)
        close = datetime(d.year, d.month, d.day, 16, 0, tzinfo=NY).astimezone(timezone.utc)
    return start, close


def session_window(day, feed):
    d = datetime.strptime(day, "%Y-%m-%d").date() if day else datetime.now(NY).date()
    if LIVE_MODE and LIVE_SESSION_ONLY:
        # live allégé (plan gratuit) : uniquement la séance d'alerte 9h30 -> 14h05 NY (la bougie de 14h00 est incluse)
        start = datetime(d.year, d.month, d.day, 9, 30, tzinfo=NY).astimezone(timezone.utc)
        close = datetime(d.year, d.month, d.day, 14, 5, tzinfo=NY).astimezone(timezone.utc)
    else:
        start, close = day_bounds(d)
    now = datetime.now(timezone.utc)
    end = min(close, now)
    if feed == "SIP":
        end = min(end, now - timedelta(minutes=16))   # plans gratuits : SIP retardé de 15 min
    if LIVE_MODE:
        # live : on arrête la fenêtre à la dernière bougie CLÔTURÉE (la dernière ligne de l'Excel est toujours définitive ;
        # les données entre cette borne et « maintenant » seront rattrapées par le téléchargement incrémental du tick suivant).
        step = timedelta(minutes=BAR_MINUTES)
        aligned = datetime.fromtimestamp(int(now.timestamp()) // (BAR_MINUTES * 60) * (BAR_MINUTES * 60), tz=timezone.utc)
        end = min(end, aligned)
    if end <= start:
        raise ValueError("La fenêtre n'a pas encore commencé (ou date future / SIP encore retardé).")
    return start, end


# ----------------------------------------------------------------- téléchargement
def _env_first(names):
    for name in names:
        value = os.environ.get(name)
        if value and value.strip():
            return value.strip(), name
    return None, None


def _client():
    from alpaca.data.historical import StockHistoricalDataClient

    key, key_name = _env_first(("ALPACA_API_KEY", "APCA_API_KEY_ID"))
    secret, secret_name = _env_first(
        ("ALPACA_SECRET_KEY", "ALPACA_API_SECRET_KEY", "APCA_API_SECRET_KEY")
    )
    if not key or not secret:
        raise SystemExit(
            "Clés Alpaca absentes. Dans le même terminal Windows, utilise :\n"
            '  set "ALPACA_API_KEY=ta_cle"\n'
            '  set "ALPACA_SECRET_KEY=ton_secret"\n'
            "Puis relance le script dans ce terminal."
        )
    print(f"Identifiants chargés via {key_name} / {secret_name} (valeurs masquées).")
    return StockHistoricalDataClient(api_key=key, secret_key=secret)


def _feed(feed):
    from alpaca.data.enums import DataFeed
    return DataFeed.SIP if feed == "SIP" else DataFeed.IEX


def fetch_trades(client, sym, start, end, feed):
    from alpaca.data.requests import StockTradesRequest
    res = client.get_stock_trades(
        StockTradesRequest(symbol_or_symbols=sym, start=start, end=end, feed=_feed(feed)))
    if sym not in res.data:
        return pd.DataFrame()
    return res.df.reset_index().sort_values("timestamp").reset_index(drop=True)


def fetch_quotes(client, sym, start, end, feed):
    from alpaca.data.requests import StockQuotesRequest
    res = client.get_stock_quotes(
        StockQuotesRequest(symbol_or_symbols=sym, start=start, end=end, feed=_feed(feed)))
    if sym not in res.data:
        return pd.DataFrame()
    return res.df.reset_index().sort_values("timestamp").reset_index(drop=True)


def prev_session_avg_volume_api(client, sym, day, feed, max_back=7):
    """Dernière journée (24h ou séance) précédant `day` : (moyenne du volume par bougie, date, référence par tranche).

    La référence par tranche = volume de J-1 à la MÊME position dans la fenêtre (donc même heure New York),
    lissé sur ±REF_SMOOTH bougies et plafonné par le bas (REF_FLOOR_FRAC × moyenne)."""
    for k in range(1, max_back + 1):
        d = day - timedelta(days=k)
        if d.weekday() >= 5:
            continue
        s, e = day_bounds(d)
        tr = fetch_trades(client, sym, s, e, feed)
        if tr.empty:
            continue
        tr = _utc_ns(tr).set_index("timestamp")
        freq = f"{BAR_MINUTES}min"
        idx = pd.date_range(s, e, freq=freq, inclusive="left")
        vol = tr["size"].resample(freq, origin=pd.Timestamp(s), label="left", closed="left").sum()
        v = vol.reindex(idx).fillna(0).to_numpy(dtype=float)
        mean = float(v.mean())
        floor = REF_FLOOR_FRAC * mean
        slots = [float(max(v[max(0, j - REF_SMOOTH): j + REF_SMOOTH + 1].mean(), floor)) for j in range(len(v))]
        return mean, d, slots
    return None, None, None



# ----------------------------------------------------------------- v7 : historique local (dossier d'Excel quotidiens)
_HIST_COLS = {"bar": "N° bougie", "time": "Heure New York", "vol": "Volume",
              "chn": "Churn carnet (MàJ NBBO / min)", "imb": "Déséquilibre pondéré temps"}
_HIST_TRADES_COL = "Trades"   # v9 : optionnelle (un vieux fichier sans cette colonne reste utilisable pour le reste)
_HIST_CACHE = {}      # (date, flux) -> (dates utilisées, {sym: [ {heure NY: (volume, churn, déséquilibre, trades)} par séance ]})
_HIST_BASE = {}       # (sym, date, flux) -> références calculées
HIST_INFO = {}        # flux -> dates des séances utilisées (pour le libellé de la feuille)
HIST_RETRO = {}       # flux -> nombre de séances POSTÉRIEURES ajoutées (analyse d'une date ancienne)


def _hist_cap_reached(n):
    """True si on a déjà HIST_DAYS séances (jamais vrai quand HIST_DAYS is None = tout l'historique)."""
    return HIST_DAYS is not None and n >= HIST_DAYS


def _hist_candidates(day_date):
    """Fichiers flux_BTC_ETF_AAAAMMJJ_*.xlsx du dossier HIST_DIR antérieurs à day_date -> [(date, chemin)] récents d'abord."""
    best = {}
    if not HIST_DIR or not os.path.isdir(HIST_DIR):
        return []
    for name in os.listdir(HIST_DIR):
        m = re.match(r"flux_BTC_ETF_(\d{8})_.*\.xlsx$", name, re.I)
        if not m or name.startswith("~$"):
            continue
        try:
            d = datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            continue
        if d >= day_date:
            continue                                   # jamais le jour analysé ni le futur
        p = os.path.join(HIST_DIR, name)
        if d not in best or os.path.getmtime(p) > os.path.getmtime(best[d]):
            best[d] = p                                # plusieurs versions du même jour : la plus récente
    return sorted(best.items(), reverse=True)


def _hnum(x):
    return float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else np.nan


def _read_hist_file(path, feed):
    """-> {sym: {heure NY 'HH:MM': (volume, churn, déséquilibre)}} pour les feuilles <SYM>_<feed> complètes (bougie 15:55 présente)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    out = {}
    try:
        for sym in ETFS:
            sh = f"{sym}_{feed}"
            if sh not in wb.sheetnames:
                continue
            it = wb[sh].iter_rows(values_only=True)
            hdr = next(it, None)
            if not hdr:
                continue
            pos = {h: i for i, h in enumerate(hdr) if h}
            if any(c not in pos for c in _HIST_COLS.values()):
                continue                               # ancien format (sans churn / déséquilibre pondéré temps)
            ib, it_, iv, ic, ii = (pos[_HIST_COLS[k]] for k in ("bar", "time", "vol", "chn", "imb"))
            itd = pos.get(_HIST_TRADES_COL)
            rows = {}
            for r in it:
                if len(r) <= max(ib, it_, iv, ic, ii) or not isinstance(r[ib], (int, float)):
                    continue                           # lignes de récapitulatif sous le tableau
                t = r[it_]
                lab = t.strftime("%H:%M") if hasattr(t, "strftime") else (str(t)[:5] if t else None)
                if not lab:
                    continue
                v = _hnum(r[iv])
                ntd = _hnum(r[itd]) if (itd is not None and len(r) > itd) else np.nan
                rows[lab] = (0.0 if np.isnan(v) else v, _hnum(r[ic]), _hnum(r[ii]), ntd)
            if "15:55" in rows:                        # séance complète
                out[sym] = rows
    finally:
        wb.close()
    return out


def _hist_candidates_after(day_date):
    """v8 : fichiers du dossier HIST_DIR POSTÉRIEURS à day_date, les plus proches d'abord (analyse d'une date ancienne)."""
    best = {}
    if not HIST_DIR or not os.path.isdir(HIST_DIR):
        return []
    for name in os.listdir(HIST_DIR):
        m = re.match(r"flux_BTC_ETF_(\d{8})_.*\.xlsx$", name, re.I)
        if not m or name.startswith("~$"):
            continue
        try:
            d = datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            continue
        if d <= day_date:
            continue
        p = os.path.join(HIST_DIR, name)
        if d not in best or os.path.getmtime(p) > os.path.getmtime(best[d]):
            best[d] = p
    return sorted(best.items())


def _load_hist(day_date, feed):
    key = (day_date, feed)
    if key in _HIST_CACHE:
        return _HIST_CACHE[key]
    days, data = [], {s: [] for s in ETFS}
    for d, path in _hist_candidates(day_date):
        if _hist_cap_reached(len(days)):
            break
        try:
            res = _read_hist_file(path, feed)
        except Exception as exc:
            print(f"[!] historique : {os.path.basename(path)} illisible ({exc}) -> ignoré")
            continue
        if MANDATORY not in res:
            print(f"[!] historique : {os.path.basename(path)} ignoré (feuille {MANDATORY}_{feed} absente, incomplète ou ancien format)")
            continue
        days.append(d)
        for s, m in res.items():
            data[s].append(m)
    n_before = len(days)
    if HIST_ALLOW_AFTER and n_before < HIST_MIN_DAYS:      # date ancienne : on complète avec les séances les plus proches APRÈS
        for d, path in _hist_candidates_after(day_date):
            if _hist_cap_reached(len(days)):
                break
            try:
                res = _read_hist_file(path, feed)
            except Exception as exc:
                print(f"[!] historique : {os.path.basename(path)} illisible ({exc}) -> ignoré")
                continue
            if MANDATORY not in res:
                continue
            days.append(d)
            for s_, m_ in res.items():
                data[s_].append(m_)
        if len(days) > n_before:
            order = sorted(range(len(days)), key=lambda q: days[q], reverse=True)
            days = [days[q] for q in order]
            for s_ in data:
                if len(data[s_]) == len(order):
                    data[s_] = [data[s_][q] for q in order]
            HIST_RETRO[feed] = len(days) - n_before
            print(f"\n[i] Date ancienne : {n_before} séance(s) avant, complétées par {len(days) - n_before} séance(s) APRÈS (référence rétrospective).")
    _HIST_CACHE[key] = (days, data)
    if days:
        HIST_INFO[feed] = days
        print(f"\nHistorique local ({feed}) : {len(days)} séance(s) -> " + ", ".join(x.strftime("%d/%m") for x in days)
              + (f"  [moins que HIST_DAYS={HIST_DAYS}]" if HIST_DAYS is not None and len(days) < HIST_DAYS else ""))
    else:
        print(f"\n[!] Historique local ({feed}) : aucun fichier exploitable dans {HIST_DIR!r}. Le script continue : colonnes D et E mesurées vs J-1 (téléchargé), "
              "colonne « Z-score déséquilibre vs réf. normale » laissée vide faute de dispersion historique.")
    return _HIST_CACHE[key]


def _window_labels(day_date):
    s, e = day_bounds(day_date)
    idx = pd.date_range(s, e, freq=f"{BAR_MINUTES}min", inclusive="left")
    return [t.tz_convert(NY).strftime("%H:%M") for t in idx]


def _smooth(a, fn):
    out = np.full(len(a), np.nan)
    for j in range(len(a)):
        w = a[max(0, j - REF_SMOOTH): j + REF_SMOOTH + 1]
        w = w[np.isfinite(w)]
        if len(w):
            out[j] = float(fn(w))
    return out


def _hist_baseline(sym, day_date, feed):
    """Références « normales » de sym : médiane par heure NY des séances de l'historique local, lissée ±REF_SMOOTH.
    -> dict(mean, slots (volume), churn, imb, n, last) ou None si moins de HIST_MIN_DAYS séances."""
    key = (sym, day_date, feed)
    if key in _HIST_BASE:
        return _HIST_BASE[key]
    days, data = _load_hist(day_date, feed)
    mats = data.get(sym) or []
    res = None
    if len(mats) >= HIST_MIN_DAYS:
        labels = _window_labels(day_date)
        need = min(HIST_MIN_DAYS, len(mats))
        nan3 = (np.nan, np.nan, np.nan, np.nan)

        def med(k):
            M = np.array([[m.get(l, nan3)[k] for l in labels] for m in mats], dtype=float)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                md = np.nanmedian(M, axis=0)
            return np.where(np.isfinite(M).sum(axis=0) >= need, md, np.nan)

        v = med(0)
        mean = float(np.nanmean(v)) if np.isfinite(v).any() else 0.0
        floor = REF_FLOOR_FRAC * mean
        slots = _smooth(v, np.mean)
        slots = np.where(np.isfinite(slots), np.maximum(slots, floor), floor)
        c = med(1)
        cs = _smooth(c, np.median)
        if np.isfinite(c).any():
            cs = np.where(np.isfinite(cs), np.maximum(cs, CHURN_REF_FLOOR_FRAC * float(np.nanmean(c))), cs)
        def mad_(k):      # v8 : dispersion robuste = médiane des écarts absolus à la médiane, par tranche
            M = np.array([[m.get(l, nan3)[k] for l in labels] for m in mats], dtype=float)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                md = np.nanmedian(M, axis=0)
                dev = np.nanmedian(np.abs(M - md), axis=0)
            return np.where(np.isfinite(M).sum(axis=0) >= need, dev, np.nan)

        tp = med(3) / float(BAR_MINUTES)          # v9 : trades / minute, médiane par tranche horaire
        tps = _smooth(tp, np.median)
        if np.isfinite(tp).any():
            tps = np.where(np.isfinite(tps), np.maximum(tps, TAPE_REF_FLOOR_FRAC * float(np.nanmean(tp))), tps)
        res = dict(mean=mean, slots=[float(x) for x in slots], churn=cs, imb=_smooth(med(2), np.median),
                   imb_mad=_smooth(mad_(2), np.median), tape=tps, n=len(mats), last=days[0])
    _HIST_BASE[key] = res
    return res


def prev_session_avg_volume(client, sym, day, feed, max_back=7):
    """Référence « normale » du volume : (moyenne par bougie, date de la séance la plus récente, référence par tranche).
    v7 : médiane de toutes les séances (ou des HIST_DAYS plus proches) lues dans HIST_DIR ; repli sur J-1 (API Alpaca) si l'historique est insuffisant."""
    try:
        b = _hist_baseline(sym, day, feed)
    except Exception as exc:
        print(f"[!] historique local indisponible pour {sym} ({feed}) : {exc}")
        b = None
    if b is not None:
        return b["mean"], b["last"], b["slots"]
    print(f"[!] {sym} ({feed}) : historique local insuffisant (< {HIST_MIN_DAYS} séances) -> repli sur J-1 téléchargé.")
    return prev_session_avg_volume_api(client, sym, day, feed, max_back)


# ----------------------------------------------------------------- enchères (optionnel)
ALPACA_DATA_URL = "https://data.alpaca.markets"


def _biggest_print(prints):
    """Parmi les prints d'une enchère, celui de plus grande taille (= la croix) -> (prix, taille) ou None."""
    best = None
    for p in prints or []:
        try:
            size = float(p.get("s", p.get("S", 0)) or 0)
            price = float(p.get("p", 0) or 0)
        except (AttributeError, TypeError, ValueError):
            continue
        if size > 0 and (best is None or size > best[1]):
            best = (price, size)
    return best


def fetch_auctions(symbols, day_date):
    """Enchères d'ouverture (9h30) et de clôture (16h00) via l'API REST Alpaca, si ton plan les fournit.

    -> {sym: {"open_price", "open_size", "close_price", "close_size"}}. Toute erreur (plan, endpoint, format) donne {}
    et le script continue : les colonnes d'enchères sont alors simplement absentes."""
    key, _ = _env_first(("ALPACA_API_KEY", "APCA_API_KEY_ID"))
    secret, _ = _env_first(("ALPACA_SECRET_KEY", "ALPACA_API_SECRET_KEY", "APCA_API_SECRET_KEY"))
    if not key or not secret or not symbols:
        return {}
    out = {}
    try:
        import requests
        s, e = day_bounds(day_date)
        fmt = "%Y-%m-%dT%H:%M:%SZ"
        headers = {"APCA-API-KEY-ID": key, "APCA-API-SECRET-KEY": secret}
        token = None
        for _ in range(10):
            params = {"symbols": ",".join(symbols), "start": s.strftime(fmt), "end": e.strftime(fmt), "limit": 1000}
            if token:
                params["page_token"] = token
            r = requests.get(f"{ALPACA_DATA_URL}/v2/stocks/auctions", params=params, headers=headers, timeout=30)
            r.raise_for_status()
            js = r.json()
            for sym, items in (js.get("auctions") or {}).items():
                for it in items or []:
                    rec = out.setdefault(sym, {})
                    for key_name, label in (("o", "open"), ("c", "close")):
                        best = _biggest_print(it.get(key_name))
                        if best:
                            rec[f"{label}_price"], rec[f"{label}_size"] = best
            token = js.get("next_page_token")
            if not token:
                break
    except Exception as exc:
        print(f"\n[!] Enchères indisponibles ({exc}) : colonnes d'enchères omises.")
        return {}
    if not out:
        print("\n[!] Aucune donnée d'enchères renvoyée par Alpaca : colonnes d'enchères omises.")
    else:
        print(f"Enchères : {len(out)} ETF renseigné(s).")
    return out


# ----------------------------------------------------------------- classification du côté
def _utc_ns(df):
    df = df.copy()
    df["timestamp"] = pd.to_datetime(df["timestamp"], utc=True).astype("datetime64[ns, UTC]")
    return df


def classify(trades, quotes):
    trades = _utc_ns(trades)
    q = _utc_ns(quotes)[["timestamp", "bid_price", "bid_size", "ask_price", "ask_size"]]
    df = pd.merge_asof(trades.sort_values("timestamp"), q.sort_values("timestamp"),
                       on="timestamp", direction="backward")
    mid = (df["bid_price"] + df["ask_price"]) / 2
    side = pd.Series(0, index=df.index, dtype="int8")
    side[df["price"] > mid] = 1
    side[df["price"] < mid] = -1
    tick = df["price"].diff().apply(lambda x: 1 if x > 0 else (-1 if x < 0 else 0))
    side = side.where(side != 0, tick)
    side = side.mask(side == 0).ffill().fillna(0).astype("int8")
    df["side"] = side
    try:
        df["sure_dir"] = _sure_direction(df, quotes)
    except Exception as exc:                      # la direction « certaine » est un plus : en cas d'erreur, le reste du fichier est produit
        print(f"[!] direction certaine indisponible ({exc}) : colonnes « certain » vides.")
        df["sure_dir"] = np.int8(0)
    return df


def _has_non_regular(cond):
    """True si la liste de conditions du trade contient une condition non régulière (voir NON_REGULAR_CONDS)."""
    if cond is None:
        return False
    if isinstance(cond, str):
        toks = cond.replace("[", " ").replace("]", " ").replace("'", " ").replace(",", " ").split()
    else:
        try:
            toks = [str(c) for c in list(cond)]
        except TypeError:
            return False
    return any(t_ in NON_REGULAR_CONDS for t_ in toks)


def _sure_direction(df, quotes):
    """Direction CERTAINE (+1 achat / -1 vente / 0 non classé) de chaque trade de `df` (déjà trié par timestamp, issu de classify).

    1) règle BID/ASK : prix >= ask -> acheteur agressif ; prix <= bid -> vendeur agressif, avec la dernière cotation STRICTEMENT AVANT le trade
       (une cotation datée du même instant peut être la conséquence du trade). Marché locked / croisé (ask <= bid) : non classé.
    2) sweep ISO : rafale de trades ISO (écart < SWEEP_GAP_MS) sur >= SWEEP_MIN_LEVELS prix distincts, prix toujours croissants -> achat,
       toujours décroissants -> vente.
    Si les deux règles se contredisent, ou si le trade a une condition non régulière : 0. Le milieu / la tick rule n'entrent JAMAIS ici."""
    q = _utc_ns(quotes)[["timestamp", "bid_price", "ask_price"]].sort_values("timestamp")
    t = pd.DataFrame({"timestamp": df["timestamp"].to_numpy(), "price": df["price"].to_numpy(dtype=float)})
    m = pd.merge_asof(t, q, on="timestamp", direction="backward", allow_exact_matches=False)
    px = t["price"].to_numpy(dtype=float)
    bid = m["bid_price"].to_numpy(dtype=float)
    ask = m["ask_price"].to_numpy(dtype=float)
    valid = np.isfinite(bid) & np.isfinite(ask) & (bid > 0) & (ask > bid)
    ba = np.zeros(len(df), dtype=np.int8)
    ba[valid & (px >= ask)] = 1
    ba[valid & (px <= bid)] = -1
    # sweep ISO
    sw = np.zeros(len(df), dtype=np.int8)
    if "conditions" in df.columns:
        iso = df["conditions"].apply(_has_iso).to_numpy(dtype=bool)
        reg = ~df["conditions"].apply(_has_non_regular).to_numpy(dtype=bool)
    else:
        iso = np.zeros(len(df), dtype=bool)
        reg = np.ones(len(df), dtype=bool)
    if iso.any():
        ts_ns = pd.to_datetime(df["timestamp"]).astype("int64").to_numpy()
        pos_i = np.flatnonzero(iso)
        tsi, pxi = ts_ns[pos_i], px[pos_i]
        newg = np.ones(len(pos_i), dtype=bool)
        newg[1:] = (tsi[1:] - tsi[:-1]) > SWEEP_GAP_MS * 1_000_000
        g = pd.DataFrame({"gid": np.cumsum(newg), "p": pxi})
        g["d"] = g["p"].diff()
        g.loc[newg, "d"] = np.nan
        a = g.groupby("gid").agg(lv=("p", "nunique"), first=("p", "first"), last=("p", "last"),
                                 dmin=("d", "min"), dmax=("d", "max"))
        enough = a["lv"] >= SWEEP_MIN_LEVELS
        dirg = pd.Series(0, index=a.index, dtype="int8")
        dirg[enough & (a["dmin"] >= 0) & (a["last"] > a["first"])] = 1
        dirg[enough & (a["dmax"] <= 0) & (a["last"] < a["first"])] = -1
        sw[pos_i] = dirg.reindex(g["gid"].to_numpy()).to_numpy(dtype=np.int8)
    out = np.where((ba != 0) & (sw != 0) & (ba != sw), 0,
                   np.where(ba != 0, ba, sw)).astype(np.int8)
    out[~reg] = 0
    return out


# ----------------------------------------------------------------- bougies
def _has_iso(cond):
    """True si la liste de conditions du trade contient « F » (Intermarket Sweep Order)."""
    if cond is None:
        return False
    if isinstance(cond, str):
        return "F" in cond.replace("[", " ").replace("]", " ").replace("'", " ").replace(",", " ").split()
    try:
        return "F" in list(cond)
    except TypeError:
        return False


BOOK_KEYS = ("q_coef", "churn", "size_rot", "imb_tw", "impact", "absorb", "confirm")


def _book_metrics(bars, trades_cl, quotes, start, end):
    """Indicateurs v2 « qualité du carnet » (colonnes BB -> BH), calculés en seconde passe sur les quotes et les
    trades bruts du flux traité (donc séparément pour IEX et pour SIP). Retourne {nom: tableau numpy aligné sur bars}."""
    idx = bars.index
    nb = len(idx)
    out = {k: np.full(nb, np.nan) for k in BOOK_KEYS}
    step = BAR_MINUTES * 60 * 10 ** 9
    s_ns = pd.Timestamp(start).value
    e_ns = pd.Timestamp(end).value
    lo = s_ns + step * np.arange(nb, dtype="int64")          # début de chaque bougie (ns UTC)
    hi = np.minimum(lo + step, e_ns)                          # fin (la dernière bougie peut être partielle)

    q = _utc_ns(quotes).sort_values("timestamp")
    qt = q["timestamp"].values.astype("int64")
    bp = q["bid_price"].to_numpy(dtype=float)
    ap = q["ask_price"].to_numpy(dtype=float)
    bsz = q["bid_size"].to_numpy(dtype=float)
    asz = q["ask_size"].to_numpy(dtype=float)

    with np.errstate(invalid="ignore", divide="ignore"):
        # BB : coef qualité = médiane mobile du spread moyen / spread moyen de la bougie, plafonné à 100 %
        spr = bars["spread_avg"].to_numpy(dtype=float)
        for b in range(nb):
            if not (np.isfinite(spr[b]) and spr[b] > 0):
                continue
            hist = spr[max(0, b - BOOK_QUALITY_WINDOW): b]
            hist = hist[np.isfinite(hist) & (hist > 0)]
            if len(hist) >= BOOK_MIN_HISTORY:
                out["q_coef"][b] = min(1.0, float(np.median(hist)) / float(spr[b]))

        if len(qt):
            # BC : churn = mises à jour du NBBO par minute (durée réelle de la bougie)
            minutes = np.maximum(hi - lo, 0) / 6e10
            cnt = np.searchsorted(qt, hi, "left") - np.searchsorted(qt, lo, "left")
            out["churn"] = np.where(minutes > 0, cnt / np.where(minutes > 0, minutes, 1), np.nan)

            # BD (rotation des tailles) et BE (déséquilibre pondéré par la durée d'affichage)
            tot = bsz + asz
            imb = np.where(tot > 0, (bsz - asz) / np.where(tot > 0, tot, 1), np.nan)
            qnext = np.append(qt[1:], max(e_ns, int(qt[-1])))
            cap = TW_CAP_SECONDS * 10 ** 9
            for b in range(nb):
                s, e = lo[b], hi[b]
                if e <= s:
                    continue
                i0 = np.searchsorted(qt, s, "left")
                i1 = np.searchsorted(qt, e, "left")
                seg = tot[i0:i1]
                seg = seg[np.isfinite(seg) & (seg > 0)]
                if len(seg) >= 2:
                    out["size_rot"][b] = float(seg.max() / seg.min())
                k0 = max(int(np.searchsorted(qt, s, "right")) - 1, 0)    # état déjà affiché au début de la bougie
                ks = np.arange(k0, max(int(i1), k0 + 1))
                dur = np.clip(np.minimum(qnext[ks], e) - np.maximum(qt[ks], s), 0, cap).astype(float)
                ok = np.isfinite(imb[ks]) & (dur > 0)
                if ok.any():
                    out["imb_tw"][b] = float(np.average(imb[ks][ok], weights=dur[ok]))

            # BF : déplacement du mid dans les IMPACT_HORIZON_S secondes après chaque trade, en × spread,
            # pondéré par la taille des trades de la bougie (signe = sens du mid : + monte, − baisse)
            good_q = np.isfinite(bp) & np.isfinite(ap) & (ap > bp)
            if good_q.any() and len(trades_cl):
                t = _utc_ns(trades_cl)
                tt = t["timestamp"].values.astype("int64")
                sz = t["size"].to_numpy(dtype=float)
                qv, mv, sv = qt[good_q], ((bp + ap) / 2)[good_q], (ap - bp)[good_q]
                ib = np.searchsorted(qv, tt, "right") - 1
                ia = np.searchsorted(qv, tt + IMPACT_HORIZON_S * 10 ** 9, "right") - 1
                ok = (ib >= 0) & (ia >= 0)
                val = np.full(len(tt), np.nan)
                val[ok] = (mv[ia[ok]] - mv[ib[ok]]) / sv[ib[ok]]
                bidx = np.clip((tt - s_ns) // step, 0, nb - 1).astype(int)
                m = np.isfinite(val)
                num = np.bincount(bidx[m], weights=(sz * val)[m], minlength=nb)
                den = np.bincount(bidx[m], weights=sz[m], minlength=nb)
                out["impact"] = np.where(den > 0, num / np.where(den > 0, den, 1), np.nan)

        # carnet affiché de la bougie : déséquilibre pondéré temps (BE), à défaut déséquilibre moyen (colonne AB)
        vol = bars["volume"].to_numpy(dtype=float)
        delta = (bars["buy_vol"] - bars["sell_vol"]).to_numpy(dtype=float)
        dpct = np.where(vol > 0, delta / np.where(vol > 0, vol, 1), np.nan)
        s_avg = (bars["bid_size_avg"] + bars["ask_size_avg"]).to_numpy(dtype=float)
        ab = np.where(s_avg > 0, (bars["bid_size_avg"] - bars["ask_size_avg"]).to_numpy(dtype=float)
                      / np.where(s_avg > 0, s_avg, 1), np.nan)
        book = np.where(np.isfinite(out["imb_tw"]), out["imb_tw"], ab)
        ntr = bars["trades"].to_numpy(dtype=float)

        # BG : score d'absorption (0-100) = exécution À CONTRE du carnet affiché × intensité du déséquilibre × qualité
        for b in range(nb):
            bk, d, qc = book[b], dpct[b], out["q_coef"][b]
            if not (np.isfinite(bk) and np.isfinite(d) and np.isfinite(qc)) or vol[b] <= 0 or ntr[b] < WHALE_MIN_TRADES:
                continue
            if abs(bk) < SEUIL_CARNET:
                out["absorb"][b] = 0.0
                continue
            against = max(0.0, -float(np.sign(bk)) * d)          # delta signé contre le côté affiché
            out["absorb"][b] = float(round(100 * min(1.0, against / ALIGN_SCALE)
                                           * min(1.0, abs(bk) / ALIGN_SCALE) * qc))

        # BH : confirmation carnet -> exécution. Pour chaque bougie j des CONFIRM_WINDOW dernières (résultat connu),
        # carnet(j) déséquilibré (>= SEUIL_CARNET) : le delta de la bougie j+1 a-t-il le même signe ? -> taux de réussite
        if nb >= 3:
            valid = (np.isfinite(book[:-1]) & (np.abs(book[:-1]) >= SEUIL_CARNET)
                     & np.isfinite(dpct[1:]) & (dpct[1:] != 0))
            hit = valid & (np.sign(book[:-1]) == np.sign(dpct[1:]))
            cv = np.concatenate([[0], np.cumsum(valid)])
            ch = np.concatenate([[0], np.cumsum(hit)])
            for i in range(2, nb):
                a = max(0, i - CONFIRM_WINDOW - 1)
                n_ok = cv[i - 1] - cv[a]
                if n_ok >= CONFIRM_MIN_SAMPLES:
                    out["confirm"][i] = float(ch[i - 1] - ch[a]) / float(n_ok)
    return out


BOOK3_KEYS = ("pers", "wall_b", "wall_a", "wall_x", "imb_p", "cls1", "cls2", "cls3", "cls4",
              "rm_b", "con_b", "rm_a", "con_a", "pull")


def _book_metrics_v3(bars, trades_cl, quotes, start, end):
    """Indicateurs v3 (colonnes BJ -> CA), seconde passe indépendante de _book_metrics (une erreur ici ne touche pas BB -> BH).

    - pers / imb_p : persistance d'un état = min(1, âge de l'état ÷ médiane des durées d'état des REF_WINDOW_BARS bougies
      précédentes). L'âge court depuis l'apparition de l'état (avant-bougie inclus) jusqu'à la fin de sa vie dans la bougie :
      aucune information future. imb_p = déséquilibre où chaque état pèse durée x persistance (durée plafonnée à TW_CAP_SECONDS).
    - wall_b / wall_a / wall_x : mur anormal = taille d'un côté > WALL_MULT x médiane de ce côté (bougies précédentes).
    - cls1..cls4 : déséquilibre pondéré temps restreint aux états dont la plus grande taille (bid ou ask) est dans la classe.
    - rm_* / con_* : retraits de taille (même prix : baisse de taille ; niveau qui disparaît : taille entière) et part
      consommée par des trades au bid (ventes) / à l'ask (achats) pendant que l'état était affiché ; le reste = annulé.
    - pull : retrait de mur pré-impact (1 bid, 2 ask, 3 les deux)."""
    nb = len(bars)
    out = {k: np.full(nb, np.nan) for k in BOOK3_KEYS}
    if nb == 0 or len(quotes) == 0:
        return out
    step = BAR_MINUTES * 60 * 10 ** 9
    s_ns = pd.Timestamp(start).value
    e_ns = pd.Timestamp(end).value
    lo = s_ns + step * np.arange(nb, dtype="int64")
    hi = np.minimum(lo + step, e_ns)
    cap = TW_CAP_SECONDS * 10 ** 9
    eps = 1e-9

    q = _utc_ns(quotes).sort_values("timestamp")
    qt = q["timestamp"].values.astype("int64")
    nq = len(qt)
    if nq < 2:
        return out
    bp = q["bid_price"].to_numpy(dtype=float)
    ap = q["ask_price"].to_numpy(dtype=float)
    bsz = q["bid_size"].to_numpy(dtype=float)
    asz = q["ask_size"].to_numpy(dtype=float)
    qnext = np.append(qt[1:], max(e_ns, int(qt[-1])))
    tot = bsz + asz
    with np.errstate(invalid="ignore", divide="ignore"):
        imb = np.where(tot > 0, (bsz - asz) / np.where(tot > 0, tot, 1), np.nan)
    wall_sz = np.fmax(bsz, asz)
    cls = np.digitize(np.nan_to_num(wall_sz, nan=0.0), SIZE_CLASSES)      # 0 .. 3

    # médianes mobiles causales (REF_WINDOW_BARS bougies précédentes, bougie courante exclue)
    i_lo = np.searchsorted(qt, lo, "left")
    med_age = np.full(nb, np.nan)
    med_b = np.full(nb, np.nan)
    med_a = np.full(nb, np.nan)
    for b in range(1, nb):
        a0, a1 = int(i_lo[max(0, b - REF_WINDOW_BARS)]), int(i_lo[b])
        if a1 - a0 < REF_MIN_STATES:
            continue
        d = np.clip(np.minimum(qnext[a0:a1], lo[b]) - qt[a0:a1], 0, cap)
        d = d[d > 0]
        if len(d) >= REF_MIN_STATES:
            med_age[b] = float(np.median(d))
        sb = bsz[a0:a1]
        sb = sb[np.isfinite(sb) & (sb > 0)]
        if len(sb) >= REF_MIN_STATES:
            med_b[b] = float(np.median(sb))
        sa = asz[a0:a1]
        sa = sa[np.isfinite(sa) & (sa > 0)]
        if len(sa) >= REF_MIN_STATES:
            med_a[b] = float(np.median(sa))

    # ---- par bougie : persistance, déséquilibre x persistance, murs, classes de taille ----
    for b in range(nb):
        s, e = int(lo[b]), int(hi[b])
        if e <= s:
            continue
        i1 = int(np.searchsorted(qt, e, "left"))
        k0 = max(int(np.searchsorted(qt, s, "right")) - 1, 0)        # état déjà affiché au début de la bougie
        ks = np.arange(k0, max(i1, k0 + 1))
        t_to = np.minimum(qnext[ks], e)
        dur = np.clip(t_to - np.maximum(qt[ks], s), 0, cap).astype(float)
        age = np.clip(t_to - qt[ks], 0, cap).astype(float)
        pos = dur > 0
        ik = imb[ks]
        ok = np.isfinite(ik) & pos
        if ok.any():
            ref = med_age[b]
            if not (np.isfinite(ref) and ref > 0):                   # repli (début de fenêtre) : médiane des âges de la bougie
                pa = age[ok & (age > 0)]
                ref = float(np.median(pa)) if len(pa) >= 3 else np.nan
            if np.isfinite(ref) and ref > 0:
                pers = np.minimum(1.0, age / ref)
                out["pers"][b] = float(np.mean(pers[ok]))   # par comptage d'états : les flashs pèsent autant que les murs installés
                m = ok & (pers > 0)
                if m.any():
                    out["imb_p"][b] = float(np.average(ik[m], weights=(dur * pers)[m]))
            wk, ck = wall_sz[ks], cls[ks]
            for c in range(4):
                m = ok & np.isfinite(wk) & (ck == c)
                if m.any():
                    out[f"cls{c + 1}"][b] = float(np.average(ik[m], weights=dur[m]))
        mults = []
        for sz_arr, med, key in ((bsz, med_b[b], "wall_b"), (asz, med_a[b], "wall_a")):
            if not (np.isfinite(med) and med > 0):
                continue
            v = sz_arr[ks]
            mv = np.isfinite(v) & pos
            if mv.any():
                out[key][b] = float(dur[mv & (v > WALL_MULT * med)].sum() / dur[mv].sum())
                mults.append(float(v[mv].max() / med))
        if mults:
            out["wall_x"][b] = max(mults)

    # ---- transitions quote k -> k+1 : retraits de taille, consommation par les trades, retrait de mur ----
    t1 = qt[1:]
    bi = (t1 - s_ns) // step                                          # bougie de l'instant du changement
    bic = np.clip(bi, 0, nb - 1).astype(int)
    in_win = (bi >= 0) & (bi < nb) & (t1 < hi[bic])
    b0, b1, a0, a1 = bsz[:-1], bsz[1:], asz[:-1], asz[1:]
    pb0, pb1, pa0, pa1 = bp[:-1], bp[1:], ap[:-1], ap[1:]
    with np.errstate(invalid="ignore"):
        same_b = np.abs(pb1 - pb0) < eps
        same_a = np.abs(pa1 - pa0) < eps
        rem_b = np.where(same_b & (b1 < b0), b0 - b1, np.where(pb1 < pb0 - eps, b0, 0.0))
        rem_a = np.where(same_a & (a1 < a0), a0 - a1, np.where(pa1 > pa0 + eps, a0, 0.0))
    rem_b = np.nan_to_num(rem_b, nan=0.0)
    rem_a = np.nan_to_num(rem_a, nan=0.0)

    hit_b = np.zeros(nq)
    hit_a = np.zeros(nq)
    if trades_cl is not None and len(trades_cl):
        t = _utc_ns(trades_cl)
        tt = t["timestamp"].values.astype("int64")
        px = t["price"].to_numpy(dtype=float)
        sz = t["size"].to_numpy(dtype=float)
        sd = t["side"].to_numpy(dtype=float)
        st = np.searchsorted(qt, tt, "right") - 1                     # état affiché au moment du trade (comme classify)
        okt = st >= 0
        stc = np.maximum(st, 0)
        with np.errstate(invalid="ignore"):
            mb = okt & (sd < 0) & (px <= bp[stc] + eps)               # vente au bid (ou en dessous)
            ma = okt & (sd > 0) & (px >= ap[stc] - eps)               # achat à l'ask (ou au-dessus)
        hit_b = np.bincount(st[mb], weights=sz[mb], minlength=nq)
        hit_a = np.bincount(st[ma], weights=sz[ma], minlength=nq)
    con_b = np.minimum(rem_b, hit_b[:-1])
    con_a = np.minimum(rem_a, hit_a[:-1])

    valid_bar = np.searchsorted(qt, hi, "left") > 0                   # au moins un état affiché avant la fin de la bougie
    for key, arr in (("rm_b", rem_b), ("con_b", con_b), ("rm_a", rem_a), ("con_a", con_a)):
        agg = np.bincount(bic[in_win], weights=arr[in_win], minlength=nb)
        out[key] = np.where(valid_bar, agg, np.nan)

    thr_b, thr_a = WALL_MULT * med_b[bic], WALL_MULT * med_a[bic]
    gone_b_lvl, gone_a_lvl = WALL_GONE_MULT * med_b[bic], WALL_GONE_MULT * med_a[bic]
    late = t1 >= hi[bic] - PULL_WINDOW_S * 10 ** 9
    with np.errstate(invalid="ignore"):
        gone_b = np.where(same_b, b1 < gone_b_lvl, pb1 < pb0 - eps)
        gone_a = np.where(same_a, a1 < gone_a_lvl, pa1 > pa0 + eps)
        pull_b = (in_win & late & (b0 > thr_b) & gone_b & (rem_b > 0)
                  & ((rem_b - con_b) >= PULL_CANCEL_FRAC * rem_b))
        pull_a = (in_win & late & (a0 > thr_a) & gone_a & (rem_a > 0)
                  & ((rem_a - con_a) >= PULL_CANCEL_FRAC * rem_a))
    pb_bar = np.bincount(bic[pull_b], minlength=nb) > 0
    pa_bar = np.bincount(bic[pull_a], minlength=nb) > 0
    with np.errstate(invalid="ignore"):
        pres_b = out["wall_b"] >= PULL_PRESENCE
        pres_a = out["wall_a"] >= PULL_PRESENCE
    code = (pb_bar & pres_b).astype(float) + 2.0 * (pa_bar & pres_a)
    has_ref = np.isfinite(med_b) | np.isfinite(med_a)
    out["pull"] = np.where(valid_bar & has_ref, code, np.nan)
    return out


def _imb_tw_series(quotes, start, end):
    """Déséquilibre bid/ask pondéré par la durée d'affichage, par bougie : même calcul que la colonne BE (utilisé pour J-1)."""
    nb = len(pd.date_range(start, end, freq=f"{BAR_MINUTES}min", inclusive="left"))
    out = np.full(nb, np.nan)
    if nb == 0 or len(quotes) == 0:
        return out
    step = BAR_MINUTES * 60 * 10 ** 9
    s_ns = pd.Timestamp(start).value
    e_ns = pd.Timestamp(end).value
    lo = s_ns + step * np.arange(nb, dtype="int64")
    hi = np.minimum(lo + step, e_ns)
    q = _utc_ns(quotes).sort_values("timestamp")
    qt = q["timestamp"].values.astype("int64")
    bsz = q["bid_size"].to_numpy(dtype=float)
    asz = q["ask_size"].to_numpy(dtype=float)
    tot = bsz + asz
    with np.errstate(invalid="ignore", divide="ignore"):
        imb = np.where(tot > 0, (bsz - asz) / np.where(tot > 0, tot, 1), np.nan)
    qnext = np.append(qt[1:], max(e_ns, int(qt[-1])))
    cap = TW_CAP_SECONDS * 10 ** 9
    for b in range(nb):
        s, e = lo[b], hi[b]
        if e <= s:
            continue
        i1 = np.searchsorted(qt, e, "left")
        k0 = max(int(np.searchsorted(qt, s, "right")) - 1, 0)
        ks = np.arange(k0, max(int(i1), k0 + 1))
        dur = np.clip(np.minimum(qnext[ks], e) - np.maximum(qt[ks], s), 0, cap).astype(float)
        ok = np.isfinite(imb[ks]) & (dur > 0)
        if ok.any():
            out[b] = float(np.average(imb[ks][ok], weights=dur[ok]))
    return out


def _churn_ref_series(quotes, start, end):
    """Réf. J-1 du churn : MàJ NBBO / min de chaque tranche de J-1, médiane lissée sur +/- REF_SMOOTH bougies,
    avec un plancher (CHURN_REF_FLOOR_FRAC x churn moyen de J-1). Même formule que la colonne BC (churn du jour)."""
    nb = len(pd.date_range(start, end, freq=f"{BAR_MINUTES}min", inclusive="left"))
    out = np.full(nb, np.nan)
    if nb == 0 or len(quotes) == 0:
        return out
    step = BAR_MINUTES * 60 * 10 ** 9
    s_ns = pd.Timestamp(start).value
    e_ns = pd.Timestamp(end).value
    lo = s_ns + step * np.arange(nb, dtype="int64")
    hi = np.minimum(lo + step, e_ns)
    q = _utc_ns(quotes).sort_values("timestamp")
    qt = q["timestamp"].values.astype("int64")
    minutes = np.maximum(hi - lo, 0) / 6e10
    cnt = np.searchsorted(qt, hi, "left") - np.searchsorted(qt, lo, "left")
    with np.errstate(invalid="ignore", divide="ignore"):
        raw = np.where(minutes > 0, cnt / np.where(minutes > 0, minutes, 1), np.nan)
    for j in range(nb):
        w = raw[max(0, j - REF_SMOOTH): j + REF_SMOOTH + 1]
        w = w[np.isfinite(w)]
        if len(w):
            out[j] = float(np.median(w))
    if np.isfinite(raw).any():
        floor = CHURN_REF_FLOOR_FRAC * float(np.nanmean(raw))
        out = np.where(np.isfinite(out), np.maximum(out, floor), out)
    return out


def _attach_book_ref(client, sym, feed, ref, bars):
    """Colonne « réf. J-1 » du déséquilibre : médiane de J-1 sur la même tranche (+/- REF_SMOOTH bougies). Jamais bloquant."""
    try:
        if BOOK_J1_BASELINE and len(bars):
            hb = _hist_baseline(sym, bars.index[0].tz_convert(NY).date(), feed)
            if hb is not None:                              # v7 : réf. déséquilibre + churn = médiane de l'historique local (aucun téléchargement)
                n_ = len(bars)
                a1 = np.full(n_, np.nan); a2 = np.full(n_, np.nan)
                m1 = min(n_, len(hb["imb"])); a1[:m1] = hb["imb"][:m1]
                m2 = min(n_, len(hb["churn"])); a2[:m2] = hb["churn"][:m2]
                bars["imb_ref"], bars["churn_ref"] = a1, a2
                a3 = np.full(n_, np.nan); m3 = min(n_, len(hb["imb_mad"])); a3[:m3] = hb["imb_mad"][:m3]
                bars["imb_mad"] = a3
                a4 = np.full(n_, np.nan); m4 = min(n_, len(hb["tape"])); a4[:m4] = hb["tape"][:m4]
                bars["tape_ref"] = a4                       # v9 : réf. vitesse du tape (trades / min)
                print(f"réf. déséquilibre / churn ({sym}) : médiane de {hb['n']} séance(s) de l'historique local.")
                return
        if not BOOK_J1_BASELINE or not ref or ref[1] is None:
            return
        s, e = day_bounds(ref[1])
        q = fetch_quotes(client, sym, s, e, feed)
        if q.empty:
            print(f"[!] {sym} ({feed}) : pas de quotes J-1 ({ref[1]}), réf. J-1 du déséquilibre vide.")
            return
        v = _imb_tw_series(q, s, e)
        slots = np.full(len(v), np.nan)
        for j in range(len(v)):
            w = v[max(0, j - REF_SMOOTH): j + REF_SMOOTH + 1]
            w = w[np.isfinite(w)]
            if len(w):
                slots[j] = float(np.median(w))
        arr = np.full(len(bars), np.nan)
        n = min(len(bars), len(slots))
        arr[:n] = slots[:n]
        bars["imb_ref"] = arr
        # réf. J-1 du churn (sert au ratio churn de la feuille CONFLUENCE uniquement) : jamais bloquant
        try:
            cr = _churn_ref_series(q, s, e)
            carr = np.full(len(bars), np.nan)
            m_ = min(len(bars), len(cr))
            carr[:m_] = cr[:m_]
            bars["churn_ref"] = carr
        except Exception as exc2:
            print(f"[!] réf. J-1 du churn indisponible pour {sym} ({feed}) : {exc2}")
        print(f"réf. J-1 du déséquilibre ({ref[1]}) : {int(np.isfinite(arr).sum())} tranche(s) renseignée(s).")
    except Exception as exc:
        print(f"[!] réf. J-1 du déséquilibre indisponible pour {sym} ({feed}) : {exc}")


def build_bars(trades_cl, quotes, start, end):
    freq = f"{BAR_MINUTES}min"
    idx = pd.date_range(start, end, freq=freq, inclusive="left")
    if len(idx) == 0:
        idx = pd.DatetimeIndex([start])

    t = _utc_ns(trades_cl).set_index("timestamp")
    quotes = _utc_ns(quotes)
    rs = dict(origin=start, label="left", closed="left")

    bars = pd.DataFrame(index=idx)
    ohlc = t["price"].resample(freq, **rs).ohlc()
    bars["open"], bars["high"], bars["low"], bars["close"] = ohlc["open"], ohlc["high"], ohlc["low"], ohlc["close"]
    bars["volume"] = t["size"].resample(freq, **rs).sum()
    bars["trades"] = t["size"].resample(freq, **rs).count()
    pv = (t["price"] * t["size"]).resample(freq, **rs).sum()
    bars["vwap"] = pv / bars["volume"].where(bars["volume"] > 0)
    bars["buy_vol"] = t["size"].where(t["side"] > 0, 0).resample(freq, **rs).sum()
    bars["sell_vol"] = t["size"].where(t["side"] < 0, 0).resample(freq, **rs).sum()
    bars["max_trade"] = t["size"].resample(freq, **rs).max()
    # v10 : volume exécuté acheteur / vendeur par classe de taille de trade (sens = `side` déjà utilisé pour le delta)
    _tc = np.digitize(t["size"].to_numpy(dtype=float), TRADE_SIZE_CLASSES)
    for _k in range(4):
        _m = pd.Series(_tc == _k, index=t.index)
        bars[f"tcb{_k + 1}"] = t["size"].where(_m & (t["side"] > 0), 0).resample(freq, **rs).sum()
        bars[f"tcs{_k + 1}"] = t["size"].where(_m & (t["side"] < 0), 0).resample(freq, **rs).sum()

    # ISO = Intermarket Sweep Order (condition de trade « F »). Absent des données -> colonnes à 0.
    if "conditions" in t.columns:
        is_iso = t["conditions"].apply(_has_iso)
    else:
        is_iso = pd.Series(False, index=t.index)
    bars["iso_vol"] = t["size"].where(is_iso, 0).resample(freq, **rs).sum()
    bars["iso_buy"] = t["size"].where(is_iso & (t["side"] > 0), 0).resample(freq, **rs).sum()
    bars["iso_sell"] = t["size"].where(is_iso & (t["side"] < 0), 0).resample(freq, **rs).sum()
    # v6 : ISO à direction CERTAINE (bid/ask ou sweep) et ISO « estimé » (milieu / tick rule : sens historique `side`)
    sd = t["sure_dir"] if "sure_dir" in t.columns else pd.Series(0, index=t.index)
    bars["iso_buy_sure"] = t["size"].where(is_iso & (sd > 0), 0).resample(freq, **rs).sum()
    bars["iso_sell_sure"] = t["size"].where(is_iso & (sd < 0), 0).resample(freq, **rs).sum()
    bars["iso_buy_est"] = t["size"].where(is_iso & (sd == 0) & (t["side"] > 0), 0).resample(freq, **rs).sum()
    bars["iso_sell_est"] = t["size"].where(is_iso & (sd == 0) & (t["side"] < 0), 0).resample(freq, **rs).sum()

    tmp = t[["size", "side"]].copy()
    step = pd.Timedelta(freq)
    tmp["bin"] = start + ((tmp.index - start) // step) * step
    big = tmp.sort_values("size", kind="stable").groupby("bin")["side"].last()
    bars["max_trade_side"] = big.reindex(idx)

    q = quotes.set_index("timestamp")
    last = q[["bid_price", "bid_size", "ask_price", "ask_size"]].resample(freq, **rs).last()
    avg = q[["bid_price", "bid_size", "ask_price", "ask_size"]].resample(freq, **rs).mean()
    bars["bid_last"], bars["bid_size_last"] = last["bid_price"], last["bid_size"]
    bars["ask_last"], bars["ask_size_last"] = last["ask_price"], last["ask_size"]
    bars["bid_avg"], bars["ask_avg"] = avg["bid_price"], avg["ask_price"]
    bars["bid_size_avg"], bars["ask_size_avg"] = avg["bid_size"], avg["ask_size"]
    bars["spread_avg"] = (q["ask_price"] - q["bid_price"]).resample(freq, **rs).mean()
    bars["n_quotes"] = q["bid_price"].resample(freq, **rs).count()

    bars = bars.reindex(idx)
    for c in ["bid_last", "bid_size_last", "ask_last", "ask_size_last"]:
        bars[c] = bars[c].ffill()
    for c in ["volume", "trades", "buy_vol", "sell_vol", "n_quotes", "iso_vol", "iso_buy", "iso_sell",
              "iso_buy_sure", "iso_sell_sure", "iso_buy_est", "iso_sell_est"] + \
             [f"tc{x}{k}" for x in ("b", "s") for k in range(1, 5)]:
        bars[c] = bars[c].fillna(0)

    # indicateurs v2 (BB -> BH) : si leur calcul échoue, les colonnes restent vides et le reste du fichier est produit
    try:
        extra = _book_metrics(bars, trades_cl, quotes, start, end)
    except Exception as exc:
        print(f"[!] indicateurs carnet v2 indisponibles ({exc}) : colonnes BB -> BH vides.")
        extra = {k: np.full(len(bars), np.nan) for k in BOOK_KEYS}
    for k, v in extra.items():
        bars[k] = v
    # indicateurs v3 (BJ -> CA) : calcul indépendant ; en cas d'erreur seules ces colonnes restent vides
    try:
        extra3 = _book_metrics_v3(bars, trades_cl, quotes, start, end)
    except Exception as exc:
        print(f"[!] indicateurs carnet v3 indisponibles ({exc}) : colonnes BJ -> CA vides.")
        extra3 = {k: np.full(len(bars), np.nan) for k in BOOK3_KEYS}
    for k, v in extra3.items():
        bars[k] = v
    bars["imb_ref"] = np.nan          # rempli ensuite par _attach_book_ref (réf. J-1)
    bars["imb_mad"] = np.nan          # v8 : dispersion historique du déséquilibre (même tranche), pour le z-score de F
    bars["churn_ref"] = np.nan        # idem : réf. J-1 du churn (ratio churn de CONFLUENCE)
    bars["tape_ref"] = np.nan         # v9 : réf. historique de la vitesse du tape (trades / min)
    return bars


# ----------------------------------------------------------------- Open Interest BTC agrégé (Coinalyze)
# 100 % BTC (aucun suivi ETH). Même logique que le module coinalyze_validator :
#  - perp : un seul contrat par exchange (celui qui a le plus d'OI), puis les TOP_PERP exchanges ;
#  - futures à échéance : les TOP_FUTURES contrats avec le plus d'OI ;
#  - OI en USD, historique 5 min, somme de tous les contrats retenus.
# Les options ne sont pas dans l'API « future-markets » de Coinalyze : elles ne sont pas incluses.
CZ_BASE = "https://api.coinalyze.net/v1"
CZ_EXCHANGES = [x.strip() for x in os.environ.get(
    "COINALYZE_EXCHANGES", "Binance,Bybit,OKX,Bitget,Gate,BingX,Deribit,Hyperliquid").split(",") if x.strip()]
CZ_TOP_PERP = int(os.environ.get("COINALYZE_TOP_PERP", "8"))
CZ_TOP_FUTURES = int(os.environ.get("COINALYZE_TOP_FUTURES", "4"))
CZ_RATE_PER_MIN = int(os.environ.get("COINALYZE_RATE_LIMIT_PER_MIN", "20"))   # plafond Coinalyze : 40 / min
CZ_CACHE_HOURS = 6                         # la liste des contrats retenus est réutilisée pendant 6 h
CZ_BASES = ("BTC", "1000BTC")
_cz_calls = deque()                        # (instant, coût) : fenêtre glissante d'une minute


def _cz_wait(cost):
    while True:
        now = time.monotonic()
        while _cz_calls and now - _cz_calls[0][0] > 60:
            _cz_calls.popleft()
        if not _cz_calls or sum(c for _, c in _cz_calls) + cost <= CZ_RATE_PER_MIN:
            _cz_calls.append((now, cost))
            return
        time.sleep(max(0.2, 60 - (now - _cz_calls[0][0]) + 0.05))


def _cz_get(path, params, cost=1, retries=4):
    """GET Coinalyze ; chaque symbole demandé consomme 1 appel du quota."""
    import requests
    q = dict(params)
    q["api_key"] = os.environ.get("COINALYZE_API_KEY", "").strip()
    for _ in range(retries + 1):
        _cz_wait(cost)
        r = requests.get(CZ_BASE + path, params=q, timeout=30)
        if r.status_code == 429:
            try:
                wait = float(r.headers.get("Retry-After", "5"))
            except ValueError:
                wait = 5.0
            time.sleep(min(wait + random.uniform(0, 1), 65))
            continue
        r.raise_for_status()
        return r.json()
    raise RuntimeError("limite de requêtes Coinalyze atteinte")


def _cz_btc_symbols():
    """Contrats BTC retenus (perp + futures), avec cache local pour ne pas refaire le classement à chaque run."""
    os.makedirs(OUT_DIR, exist_ok=True)
    cache = os.path.join(OUT_DIR, "coinalyze_btc_markets.json")
    if os.path.exists(cache) and time.time() - os.path.getmtime(cache) < CZ_CACHE_HOURS * 3600:
        try:
            with open(cache, encoding="utf-8") as fh:
                saved = json.load(fh)
            if saved.get("symbols"):
                return saved["symbols"]
        except (OSError, ValueError):
            pass

    names = {ex["code"]: ex["name"] for ex in _cz_get("/exchanges", {})}
    target = {code for code, name in names.items() if any(t.lower() in name.lower() for t in CZ_EXCHANGES)}
    if not target:
        raise RuntimeError(f"aucun exchange cible trouvé parmi {CZ_EXCHANGES}")
    infos = [{"symbol": m["symbol"], "exchange": names.get(m.get("exchange"), str(m.get("exchange"))),
              "perp": bool(m.get("is_perpetual"))}
             for m in _cz_get("/future-markets", {})
             if m.get("base_asset") in CZ_BASES and m.get("exchange") in target]
    if not infos:
        raise RuntimeError("aucun contrat BTC trouvé sur les exchanges cibles")

    oi = {}
    syms = [i["symbol"] for i in infos]
    for k in range(0, len(syms), 20):
        chunk = syms[k:k + 20]
        for row in _cz_get("/open-interest", {"symbols": ",".join(chunk), "convert_to_usd": "true"},
                           cost=len(chunk)) or []:
            if row.get("symbol") is not None and row.get("value") is not None:
                oi[row["symbol"]] = float(row["value"])

    def best(group, top_n):
        group = sorted(group, key=lambda i: oi.get(i["symbol"], -1.0), reverse=True)
        return group[:top_n] if top_n > 0 else group

    by_exchange = {}
    for i in infos:
        if i["perp"]:
            by_exchange.setdefault(i["exchange"], []).append(i)
    perps = best([max(g, key=lambda i: oi.get(i["symbol"], -1.0)) for g in by_exchange.values()], CZ_TOP_PERP)
    futures = best([i for i in infos if not i["perp"]], CZ_TOP_FUTURES)
    chosen = [i["symbol"] for i in perps + futures]
    print(f"Coinalyze BTC : {len(perps)} perp + {len(futures)} future(s) retenus : {', '.join(chosen)}")
    with open(cache, "w", encoding="utf-8") as fh:
        json.dump({"symbols": chosen, "perps": [i["symbol"] for i in perps]}, fh)
    return chosen


def _cz_history(path, symbols, params):
    """Historique Coinalyze pour une liste de symboles (20 par appel) -> {symbole: [points triés par t]}."""
    out = {}
    for k in range(0, len(symbols), 20):
        chunk = symbols[k:k + 20]
        for row in _cz_get(path, dict(params, symbols=",".join(chunk)), cost=len(chunk)) or []:
            if row.get("history"):
                out[row["symbol"]] = sorted(row["history"], key=lambda p: int(p["t"]))
    return out


def fetch_btc_oi(start, end):
    """-> {epoch_début_de_bougie : variation de l'OI BTC agrégé sur la bougie (OI fin / OI début − 1)}.

    Un contrat manquant sur une bougie ne crée pas de faux saut (ouverture et clôture portent sur les mêmes contrats).
    """
    if not os.environ.get("COINALYZE_API_KEY", "").strip():
        print("\n[!] COINALYZE_API_KEY absente : colonne OI laissée vide.")
        return {}
    interval = "5min" if BAR_MINUTES % 5 == 0 else "1min"
    frm, to = int(start.timestamp()), int(end.timestamp())
    try:
        symbols = _cz_btc_symbols()
        oi_hist = _cz_history("/open-interest-history", symbols,
                              {"interval": interval, "from": frm, "to": to, "convert_to_usd": "true"})
    except Exception as exc:
        msg = str(exc)
        if "401" in msg or "403" in msg:
            msg += " (clé Coinalyze refusée ?)"
        print(f"\n[!] OI Coinalyze indisponible : {msg}. Colonne OI laissée vide.")
        return {}

    step = BAR_MINUTES * 60
    agg = {}
    for sym, pts in oi_hist.items():
        per_bar = {}
        for p in pts:
            t = int(p["t"])
            if t < frm:
                continue
            o, c = float(p.get("o") or 0.0), float(p.get("c") or 0.0)
            b = (t - frm) // step
            if b not in per_bar:
                per_bar[b] = [o, c]            # ouverture = 1er bucket de la bougie
            else:
                per_bar[b][1] = c              # clôture = dernier bucket
        for b, (o, c) in per_bar.items():
            if o > 0 and c > 0:
                a = agg.setdefault(b, [0.0, 0.0])
                a[0] += o
                a[1] += c
    print(f"OI BTC agrégé : {len(agg)} bougie(s) renseignée(s) sur {len(oi_hist)} contrat(s).")
    return {frm + b * step: a[1] / a[0] - 1 for b, a in agg.items() if a[0] > 0}


# ----------------------------------------------------------------- Excel : styles & colonnes
NAVY = PatternFill("solid", fgColor="1F3864")
RED = PatternFill("solid", fgColor="C00000")
TEAL = PatternFill("solid", fgColor="1F6F78")
PURPLE = PatternFill("solid", fgColor="5B2C83")
F_HEAD = Font(name="Arial", bold=True, color="FFFFFF", size=10)
F = Font(name="Arial", size=10)
FB = Font(name="Arial", size=10, bold=True)
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
GREEN_FONT = Font(name="Arial", size=10, bold=True, color="007A33")
RED_FONT = Font(name="Arial", size=10, bold=True, color="C00000")
FILL_G = PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")
FILL_R = PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")
FILL_Y = PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")
FONT_Y = Font(name="Arial", size=10, bold=True, color="9C5700")
FILL_O = PatternFill("solid", bgColor="F4B183", fgColor="F4B183")
FONT_O = Font(name="Arial", size=10, bold=True, color="C55A11")

# (titre, largeur, couleur d'entête, format)
COLS = [
    ("Date", 12, NAVY, "DD/MM/YYYY"),                       # A
    ("N° bougie", 9, NAVY, "0"),                            # B
    ("Heure UTC", 10, NAVY, "HH:MM"),                       # C
    ("Heure New York", 12, NAVY, "HH:MM"),                  # D
    ("Open", 10, NAVY, "0.0000"),                           # E
    ("High", 10, NAVY, "0.0000"),                           # F
    ("Low", 10, NAVY, "0.0000"),                            # G
    ("Close", 10, NAVY, "0.0000"),                          # H
    ("Volume", 12, NAVY, "#,##0"),                          # I
    ("Trades", 9, NAVY, "#,##0"),                           # J
    ("VWAP bougie", 11, NAVY, "0.0000"),                    # K
    ("Volume achat estimé", 14, NAVY, "#,##0"),             # L
    ("Volume vente estimé", 14, NAVY, "#,##0"),             # M
    ("DELTA (achat − vente)", 15, RED, "#,##0;-#,##0;0"),   # N
    ("Delta % (delta/volume)", 13, RED, "0.0%"),            # O
    ("DELTA CUMULÉ du jour", 15, RED, "#,##0;-#,##0;0"),    # P
    ("Best bid (fin)", 11, TEAL, "0.0000"),                 # Q
    ("Taille bid (fin)", 11, TEAL, "#,##0"),                # R
    ("Best ask (fin)", 11, TEAL, "0.0000"),                 # S
    ("Taille ask (fin)", 11, TEAL, "#,##0"),                # T
    ("Spread (fin)", 10, TEAL, "0.0000"),                   # U
    ("Mid (fin)", 10, TEAL, "0.0000"),                      # V
    ("Bid moyen", 10, TEAL, "0.0000"),                      # W
    ("Ask moyen", 10, TEAL, "0.0000"),                      # X
    ("Taille bid moy.", 11, TEAL, "#,##0"),                 # Y
    ("Taille ask moy.", 11, TEAL, "#,##0"),                 # Z
    ("Spread moyen", 10, TEAL, "0.0000"),                   # AA
    ("Déséquilibre carnet moy.", 13, TEAL, "0.0%"),         # AB
    ("Nb quotes", 9, TEAL, "#,##0"),                        # AC
    ("Plus gros trade", 11, NAVY, "#,##0"),                 # AD
    ("% gros trade / volume", 12, NAVY, "0.0%"),            # AE
    (f"Ratio volume vs réf. normale (médiane {HIST_LABEL}, même tranche)", 13, NAVY, '0.00"x"'),      # AF
    ("Sens plus gros trade (estimé)", 14, NAVY, "General"), # AG
    ("PRESSION (exécution + carnet)", 16, RED, "General"),  # AH
    # ---- indicateurs avancés : à partir de la colonne AI ----
    ("Taille moyenne / trade", 13, NAVY, "#,##0.0"),         # AI
    ("Trades / 1 000 titres", 13, NAVY, "0.00"),             # AJ
    ("Taille moy. vs 12 bougies", 13, NAVY, '0.00"x"'),     # AK
    ("Nb trades vs 12 bougies", 13, NAVY, '0.00"x"'),       # AL
    ("Volume vs 12 bougies", 13, NAVY, '0.00"x"'),          # AM
    ("Plus gros trade vs 12 bougies", 14, NAVY, '0.00"x"'), # AN
    ("SCORE WHALE (proxy 0–100)", 14, RED, "0"),             # AO
    ("Présence whale estimée", 16, RED, "General"),          # AP
    ("Biais whale estimé", 18, RED, "General"),              # AQ
    ("Plus gros trade signé", 13, RED, "#,##0;-#,##0;0"),   # AR
    ("% volume ISO", 12, NAVY, "0.0%"),                      # AS  (nouveau)
    ("Delta ISO (signé)", 13, RED, "0.0%;-0.0%;0.0%"),       # AT  (nouveau, séparé du delta existant)
    ("Close vs VWAP", 12, NAVY, "0.00%"),                    # AU
    ("Impact |C−O| / 10k titres ($)", 15, NAVY, "0.0000"), # AV
    ("Amplitude / VWAP", 12, NAVY, "0.00%"),                 # AW
    ("Part volume achat", 12, RED, "0.0%"),                  # AX
    ("SCORE FLUX GLOBAL (-100/+100)", 15, RED, "0;-0;0"),  # AY
    ("SIGNAL FLUX GLOBAL", 18, RED, "General"),              # AZ
    ("Volume réf. normale (médiane, même tranche, lissé)", 15, NAVY, "#,##0"),   # BA (sert au ratio AF)
    # ---- v2 « qualité du carnet » : colonnes BB -> BI ----
    ("Coef qualité carnet (0–100 %)", 14, PURPLE, "0%"),                       # BB
    ("Churn carnet (MàJ NBBO / min)", 13, PURPLE, "0.0"),                      # BC
    ("Rotation tailles (max/min bid+ask)", 14, PURPLE, '0.00"x"'),            # BD
    ("Déséquilibre pondéré temps", 14, PURPLE, "0.0%;-0.0%;0.0%"),            # BE
    ("Impact post-trade (× spread, 2 s)", 14, PURPLE, "0.000;-0.000;0.000"),  # BF
    ("Score absorption (0–100)", 13, PURPLE, "0"),                             # BG
    ("Confirmation carnet→exécution", 15, PURPLE, "0%"),                      # BH
    ("Verdict carnet", 19, PURPLE, "General"),                                 # BI
    # ---- v3 « taille affichée crédible » : colonnes BJ -> CD ----
    ("Persistance affichage (0–100 %)", 13, PURPLE, "0%"),                      # BJ
    (f"Mur bid (> {WALL_MULT:g}x médiane) – % bougie", 13, PURPLE, "0%"),       # BK
    (f"Mur ask (> {WALL_MULT:g}x médiane) – % bougie", 13, PURPLE, "0%"),       # BL
    ("Mur max (x médiane de taille)", 13, PURPLE, '0.0"x"'),                    # BM
    ("Déséquilibre temps × persistance", 15, PURPLE, "0.0%;-0.0%;0.0%"),       # BN
    ("Écart BN − BE (flashs)", 13, PURPLE, "+0.0%;-0.0%;0.0%"),                 # BO
    (f"Déséquilibre {_class_labels()[0]}", 13, PURPLE, "0.0%;-0.0%;0.0%"),      # BP
    (f"Déséquilibre {_class_labels()[1]}", 13, PURPLE, "0.0%;-0.0%;0.0%"),      # BQ
    (f"Déséquilibre {_class_labels()[2]}", 13, PURPLE, "0.0%;-0.0%;0.0%"),      # BR
    (f"Déséquilibre {_class_labels()[3]}", 13, PURPLE, "0.0%;-0.0%;0.0%"),      # BS
    ("Bid retiré (actions)", 13, PURPLE, "#,##0"),                              # BT
    ("dont consommé par trades (bid)", 14, PURPLE, "#,##0"),                    # BU
    ("Ask retiré (actions)", 13, PURPLE, "#,##0"),                              # BV
    ("dont consommé par trades (ask)", 14, PURPLE, "#,##0"),                    # BW
    ("% consommé bid", 12, PURPLE, "0%"),                                       # BX
    ("% consommé ask", 12, PURPLE, "0%"),                                       # BY
    ("Consommation bid − ask", 13, PURPLE, "+0%;-0%;0%"),                       # BZ
    ("Retrait de mur (30 dernières s)", 16, PURPLE, "General"),                 # CA
    ("Alerte piège", 24, PURPLE, "General"),                                    # CB
    ("Réf. déséquilibre (médiane historique, même tranche)", 15, PURPLE, "0.0%;-0.0%;0.0%"), # CC
    ("Écart déséquilibre vs réf. normale", 14, PURPLE, "+0.0%;-0.0%;0.0%"),         # CD
    # ---- v6 : direction CERTAINE des trades ISO (règle bid/ask + sweep) : colonnes CE -> CI ----
    ("Volume ISO acheteur CERTAIN", 14, TEAL, "#,##0"),                         # CE
    ("Volume ISO vendeur CERTAIN", 14, TEAL, "#,##0"),                          # CF
    ("Delta ISO CERTAIN", 13, TEAL, "+0.0%;-0.0%;0.0%"),                        # CG
    ("% volume ISO classé CERTAIN", 14, TEAL, "0%"),                            # CH
    ("Delta ISO estimé (milieu / tick)", 14, TEAL, "+0.0%;-0.0%;0.0%"),        # CI
    # ---- v8 : z-score du déséquilibre vs historique (feuille CONFLUENCE, colonne « Z-score déséquilibre vs réf. normale ») : colonnes CJ -> CK ----
    ("Dispersion réf. déséquilibre (MAD)", 14, PURPLE, "0.0%"),                 # CJ
    ("Z-score déséquilibre vs réf. (robuste)", 14, PURPLE, "+0.00;-0.00;0.00"), # CK
    # ---- v9 : SPEED OF TAPE : colonnes CL -> CN ----
    ("Vitesse du tape (trades / min)", 13, NAVY, "0.0"),                        # CL
    ("Réf. vitesse tape (médiane historique, même tranche)", 15, NAVY, "0.0"),  # CM
    ("Ratio vitesse tape vs réf. normale", 14, NAVY, '0.00"x"'),                # CN
    # ---- v10 : DELTA EXÉCUTÉ par taille de trade : CO -> CR (net achat − vente) puis CS -> CV (volume classé) ----
    *[(f"Delta exécuté net – trades {lbl} actions", 14, RED, "#,##0;-#,##0;0") for lbl in _trade_class_labels()],   # CO -> CR
    *[(f"Volume exécuté classé – trades {lbl} actions", 14, NAVY, "#,##0") for lbl in _trade_class_labels()],      # CS -> CV
]
(C_PERS, C_WALLB, C_WALLA, C_WALLX, C_IMBP, C_IMBGAP, C_CLS1, C_CLS2, C_CLS3, C_CLS4, C_RMB, C_CONB, C_RMA, C_CONA,
 C_PCB, C_PCA, C_PCS, C_PULL, C_TRAP, C_J1, C_J1GAP) = range(62, 83)
C_ISOBC, C_ISOSC, C_ISODC, C_ISOPC, C_ISODE = range(83, 88)
C_J1MAD, C_J1Z = 88, 89
C_TAPE, C_TAPEREF, C_TAPERAT = 90, 91, 92
C_TCN1, C_TCV1 = 93, 97          # v10 : 4 colonnes nettes (93-96) puis 4 colonnes volume (97-100)
assert len(COLS) == 100 and COLS[REF_COL - 1][0].startswith("Volume réf.")


def _num(v):
    return None if pd.isna(v) else float(v)


def _int_or_none(v):
    return None if pd.isna(v) else int(round(float(v)))


def _pull_text(v):
    """Code numérique du retrait de mur (0 / 1 bid / 2 ask / 3 les deux) -> texte ; NaN -> cellule vide."""
    if pd.isna(v):
        return None
    return {0: "", 1: "BID", 2: "ASK", 3: "BID+ASK"}.get(int(v), "")


# ----------------------------------------------------------------- formules partagées (IEX et SIP)
def _rolling_ratio(col, i):
    """Ratio de la bougie courante vs moyenne des bougies précédentes non vides."""
    if i == 2:
        return '=""'
    start = max(2, i - WHALE_WINDOW)
    ref = f"${col}${start}:{col}{i - 1}"
    return (f'=IF(COUNTIF({ref},">0")=0,"",IF({col}{i}="","",'
            f'{col}{i}/AVERAGEIF({ref},">0")))')


def _whale_score(i):
    """Score composite : concentration 30 %, taille moy. 25 %, block 25 %, volume 20 %."""
    if i < WHALE_MIN_HISTORY + 2:
        return '=""'
    start = max(2, i - WHALE_WINDOW)
    hist = f"$I${start}:I{i - 1}"
    return (
        f'=IF(OR(I{i}=0,J{i}<{WHALE_MIN_TRADES},'
        f'COUNTIF({hist},">0")<{WHALE_MIN_HISTORY},AK{i}="",AM{i}="",AN{i}=""),"",'
        f'ROUND(100*(0.30*MIN(1,AE{i}/{WHALE_CONCENTRATION_MAX})'
        f'+0.25*MIN(1,AK{i}/{WHALE_AVG_SIZE_MAX})'
        f'+0.25*MIN(1,AN{i}/{WHALE_MAX_TRADE_MAX})'
        f'+0.20*MIN(1,AM{i}/{WHALE_VOLUME_MAX})),0))'
    )


def confluence_formulas(flux_refs, whale_refs, net_cell):
    """Même logique que confluence_formulas, pour N ETF : seuils en parts d'ETF alignés (pas en 3/3, 2/3)."""
    s = CONF_SEUIL
    pos = "(" + "+".join(f"(N({r})>={s})" for r in flux_refs) + ")"
    neg = "(" + "+".join(f"(N({r})<=-{s})" for r in flux_refs) + ")"
    cnt = ",".join(flux_refs)
    avg = f'=IF(COUNT({cnt})=0,"",ROUND(AVERAGE({cnt}),0))'
    net = f"={pos}-{neg}"
    whales = "=" + "+".join(f"(N({r})>={WHALE_ALERTE})" for r in whale_refs)
    share = f"{net_cell}/COUNT({cnt})"
    verdict = (
        f'=IF(COUNT({cnt})=0,"Pas de données",'
        f'IF({share}>={COIN_CONF_FORT},"Confluence ACHAT forte",'
        f'IF({share}<=-{COIN_CONF_FORT},"Confluence VENTE forte",'
        f'IF({share}>={COIN_CONF_SIMPLE},"Confluence achat",'
        f'IF({share}<=-{COIN_CONF_SIMPLE},"Confluence vente",'
        f'IF(AND({pos}>=1,{neg}>=1),"Divergence","Neutre / mixte"))))))'
    )
    return avg, net, whales, verdict


def delta_carnet_formulas(delta_refs, carnet_refs, vol_refs, nd_cell, nc_cell):
    """Delta / Carnet alignés (net) PONDÉRÉS PAR LE VOLUME, seuil de 10 %, et verdict CONFLUENCE DELTA + CARNET.

    Pour chaque ETF : vote +1 si valeur >= seuil, -1 si <= -seuil, 0 sinon ; puis Σ(vote × volume) / Σ volume
    (de -100 % à +100 %). Le verdict s'appuie sur (delta aligné + carnet aligné) / 2."""
    d, c = SEUIL_DELTA, SEUIL_CARNET

    def weighted(refs, thr):
        num = "+".join(f"IF(ISNUMBER({r}),IF({r}>={thr},1,IF({r}<=-{thr},-1,0))*{v},0)" for r, v in zip(refs, vol_refs))
        den = "+".join(f"IF(ISNUMBER({r}),{v},0)" for r, v in zip(refs, vol_refs))
        return f'=IF(SUM({den})=0,"",SUM({num})/SUM({den}))'

    nd, nc = weighted(delta_refs, d), weighted(carnet_refs, c)
    sd, sc = f"N({nd_cell})", f"N({nc_cell})"
    share = f"({sd}+{sc})/2"
    verdict = (
        f'=IF(AND({nd_cell}="",{nc_cell}=""),"Pas de données",'
        f'IF({share}>={DC_FORT},"ACHAT fort : delta + carnet",'
        f'IF({share}>={DC_SIMPLE},"Achat : delta + carnet",'
        f'IF({share}<=-{DC_FORT},"VENTE forte : delta + carnet",'
        f'IF({share}<=-{DC_SIMPLE},"Vente : delta + carnet",'
        f'IF(OR(AND({sd}>=0.5,{sc}<=-0.5),AND({sd}<=-0.5,{sc}>=0.5)),'
        f'"Divergence exécution / carnet","Neutre / mixte"))))))'
    )
    return nd, nc, verdict


def _add_confluence_cf(ws, avg_rng, verdict_rng):
    for txt, fill, font in [('"Confluence ACHAT forte"', FILL_G, GREEN_FONT), ('"Confluence achat"', FILL_G, GREEN_FONT),
                            ('"Confluence VENTE forte"', FILL_R, RED_FONT), ('"Confluence vente"', FILL_R, RED_FONT)]:
        ws.conditional_formatting.add(verdict_rng, CellIsRule(operator="equal", formula=[txt], fill=fill, font=font))
    ws.conditional_formatting.add(
        avg_rng, ColorScaleRule(start_type="num", start_value=-100, start_color="F8696B", mid_type="num", mid_value=0,
                                mid_color="FFFFFF", end_type="num", end_value=100, end_color="63BE7B"))


def _add_dc_cf(ws, verdict_rng):
    for txt, fill, font in [
        ('"ACHAT fort : delta + carnet"', FILL_G, GREEN_FONT),
        ('"Achat : delta + carnet"', FILL_G, GREEN_FONT),
        ('"VENTE forte : delta + carnet"', FILL_R, RED_FONT),
        ('"Vente : delta + carnet"', FILL_R, RED_FONT),
        ('"Divergence exécution / carnet"', FILL_Y, FONT_Y),
    ]:
        ws.conditional_formatting.add(verdict_rng, CellIsRule(operator="equal", formula=[txt], fill=fill, font=font))


def _add_sign_cf(ws, rng):
    """Police verte si > 0, rouge si < 0."""
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN_FONT))
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED_FONT))


def _add_rwg_scale(ws, rng):
    """Nuancier rouge -> blanc -> vert sur une valeur de 0 à 1 (0 % -> 100 %)."""
    ws.conditional_formatting.add(
        rng, ColorScaleRule(start_type="num", start_value=0, start_color="F8696B", mid_type="num", mid_value=0.5,
                            mid_color="FFFFFF", end_type="num", end_value=1, end_color="63BE7B"))


def _add_whale_scale(ws, rng):
    """Nuancier blanc -> jaune -> rouge sur 0 -> 100 (comme le score whale)."""
    ws.conditional_formatting.add(
        rng, ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", mid_type="num", mid_value=50,
                            mid_color="FFD966", end_type="num", end_value=100, end_color="F8696B"))


def write_sheet(ws, bars, feed, ref_avg=None, ref_date=None, ref_slots=None):
    ncols = len(COLS)               # colonnes A -> BI
    ref_row = len(bars) + 4          # ligne de la moyenne de référence (sous la ligne TOTAL)
    for c, (title, w, fill, _) in enumerate(COLS[:ncols], 1):
        cell = ws.cell(1, c, title)
        cell.font, cell.fill, cell.border = F_HEAD, fill, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(c)].width = w
    ws.row_dimensions[1].height = 52
    comments = {
        32: "Volume de la bougie / volume « normal » à la même heure = médiane des séances de l'historique local (colonne BA : lissée sur ±1 bougie, plancher 10 % du volume moyen).",
        53: "Référence « normale » par tranche horaire : médiane des dernières séances lues dans le dossier d'historique (même heure New York), lissée et plafonnée par le bas. Une ou deux journées anormales ne la déplacent pas.",
        35: "Volume / nombre de trades. Une hausse peut signaler des ordres plus gros, mais aussi une simple baisse de la fragmentation.",
        36: "Nombre de trades pour 1 000 titres. Valeur faible = tickets plus gros ; valeur haute = exécution plus fragmentée.",
        37: f"Taille moyenne actuelle / moyenne des {WHALE_WINDOW} bougies précédentes, hors bougies sans activité. La bougie courante est exclue de la référence.",
        41: "Proxy de 0 à 100 : concentration du plus gros trade (30 %), taille moyenne relative (25 %), plus gros trade relatif (25 %) et volume relatif (20 %). Ne prouve pas la présence d'une whale.",
        42: "Forte : score >= 75 ; modérée : >= 50 ; faible : >= 25. Historique minimum : 3 bougies avec volume et 3 trades sur la bougie courante.",
        43: "Accumulation/distribution seulement si le score whale est >= 50 et si le sens du plus gros trade est confirmé par le delta de la bougie.",
        45: "Volume des trades ISO (Intermarket Sweep Order, condition F) / volume total de la bougie.",
        46: "(Volume ISO acheteur estimé − volume ISO vendeur estimé) / volume total de la bougie. Colonne séparée : le DELTA et le Delta % existants ne sont pas modifiés.",
        51: "Score directionnel : delta 40 %, carnet 20 %, plus gros trade signé 20 %, close vs VWAP 20 %. Positif = biais acheteur ; négatif = biais vendeur.",
        52: "Accumulation/distribution forte exige aussi un volume élevé : >= 1,5x la moyenne glissante ou >= 2x la moyenne J-1.",
        # ---- v2 « qualité du carnet » ----
        54: (f"Médiane du spread moyen (colonne AA) des {BOOK_QUALITY_WINDOW} bougies précédentes de CET ETF ÷ spread moyen de la "
             "bougie, plafonné à 100 %. 100 % = carnet aussi serré que d'habitude (ou mieux) ; plus bas = carnet plus large que "
             f"d'habitude. Vide s'il y a moins de {BOOK_MIN_HISTORY} bougies d'historique. Calculé séparément pour IEX et SIP : sur IEX, "
             "le spread peut être artificiellement large quand IEX n'est pas le meilleur marché (quotes IEX uniquement)."),
        55: "Nombre de mises à jour du NBBO (quotes) par minute sur la bougie. Élevé = market makers actifs ; faible = carnet dormant, plus facile à manipuler.",
        56: ("Rapport max / min de (taille bid + taille ask) parmi les quotes de la bougie. Une taille affichée qui oscille fortement "
             "sans trades correspondants est un proxy de spoofing / quote stuffing. 1,00x = tailles stables. Vide s'il y a moins de 2 quotes."),
        57: (f"Déséquilibre bid/ask (taille bid − taille ask) / (somme) pondéré par la durée d'affichage de chaque état du carnet "
             f"(état déjà affiché au début de la bougie inclus, durée plafonnée à {TW_CAP_SECONDS // 60} min). Positif = carnet acheteur "
             "la plupart du temps. Remplace avantageusement la moyenne simple de la colonne AB."),
        58: (f"Déplacement du mid dans les {IMPACT_HORIZON_S} s après chaque trade, divisé par le spread au moment du trade, moyenné "
             "pondéré par la taille des trades de la bougie. Positif = le mid a monté, négatif = il a baissé. Plus l'impact est "
             "faible pour de gros trades, plus le carnet derrière le best bid/ask est profond."),
        59: (f"Exécution À CONTRE du carnet affiché : max(0, −signe(déséquilibre pondéré temps) × Delta %) ÷ {ALIGN_SCALE:.0%} (borné à 1) "
             f"× |déséquilibre| ÷ {ALIGN_SCALE:.0%} (borné à 1) × coef qualité (BB), sur 100. 0 si le déséquilibre est < {SEUIL_CARNET:.0%} "
             f"; vide si moins de {WHALE_MIN_TRADES} trades ou si la qualité n'est pas calculable. Proxy d'un gros acteur qui absorbe la liquidité, "
             "pas une identification."),
        60: (f"Sur les {CONFIRM_WINDOW} dernières bougies dont le résultat est connu : parmi celles où le carnet affichait un déséquilibre "
             f">= {SEUIL_CARNET:.0%}, part des cas où le delta de la bougie SUIVANTE est du même signe. 50 % = le carnet ne prédit rien ; "
             f"> 60 % = carnet fiable ; < 40 % = le carnet ment souvent. Vide s'il y a moins de {CONFIRM_MIN_SAMPLES} cas."),
        61: (f"Fiable : qualité (BB) >= {VERDICT_FIABLE:.0%} ET confirmation (BH) >= {VERDICT_FIABLE:.0%}. Peu fiable : qualité OU "
             f"confirmation < {VERDICT_PEU_FIABLE:.0%}. Sinon Moyen. « Historique insuffisant » si BB ou BH est vide."),
        # ---- v3 « taille affichée crédible » ----
        C_PERS: (f"Pour chaque état du carnet (quote -> quote suivante) : min(1, âge de l'état ÷ médiane des durées d'état des {REF_WINDOW_BARS} "
                 "bougies précédentes), moyenné par COMPTAGE d'états (chaque état pèse pareil : une pluie de flashs fait baisser le score, alors qu'une pondération par la durée saturait à 100 %). L'âge compte depuis l'apparition de l'état "
                 "(avant-bougie inclus), sans information future. Bas = carnet de flashs ; haut = carnet installé."),
        C_WALLB: (f"Part de la bougie (en temps) où la taille bid dépasse {WALL_MULT:g} x la médiane des tailles bid des {REF_WINDOW_BARS} "
                  f"bougies précédentes (mur anormal). Vide si moins de {REF_MIN_STATES} quotes d'historique."),
        C_WALLA: (f"Part de la bougie (en temps) où la taille ask dépasse {WALL_MULT:g} x la médiane des tailles ask des {REF_WINDOW_BARS} "
                  f"bougies précédentes (mur anormal). Vide si moins de {REF_MIN_STATES} quotes d'historique."),
        C_WALLX: "Plus grande taille affichée (bid ou ask) de la bougie, en multiple de la médiane de son côté. Identifie un événement (pose inhabituelle), pas un état.",
        C_IMBP: (f"Même calcul que BE (déséquilibre bid/ask pondéré par la durée), mais chaque état pèse durée x persistance (colonne {L(C_PERS)}) "
                 f"au lieu de la durée seule (même plafond de {TW_CAP_SECONDS // 60} min). Version « crédibilisée » ; BE est conservée en parallèle."),
        C_IMBGAP: f"{L(C_IMBP)} − BE. BE fort avec {L(C_IMBP)} faible = carnet rempli de flashs. L'écart lui-même est un signal.",
        C_CLS1: "Déséquilibre bid/ask pondéré par la durée, restreint aux états dont la plus grande taille (bid ou ask) est dans la classe. Vide si aucun état dans la classe. Petites tailles = plutôt retail.",
        C_CLS2: "Idem, classe de taille intermédiaire basse. Bornes calibrables en tête de script (SIZE_CLASSES).",
        C_CLS3: "Idem, classe de taille intermédiaire haute. Bornes calibrables en tête de script (SIZE_CLASSES).",
        C_CLS4: "Idem, plus grosses tailles : intention affichée (institutionnelle ou spoof), à croiser avec la persistance et la consommation.",
        C_RMB: "Actions retirées du bid pendant la bougie : baisse de taille au même prix, ou taille entière du niveau quand le bid recule. L'instant du changement donne la bougie.",
        C_CONB: "Part des retraits du bid consommée par des trades : min(retrait, volume vendeur exécuté au bid ou en dessous pendant que l'état était affiché). Le reste est considéré comme annulé.",
        C_RMA: "Actions retirées de l'ask pendant la bougie : baisse de taille au même prix, ou taille entière du niveau quand l'ask monte.",
        C_CONA: "Part des retraits de l'ask consommée par des trades : min(retrait, volume acheteur exécuté à l'ask ou au-dessus pendant que l'état était affiché). Le reste est considéré comme annulé.",
        C_PCB: f"{L(C_CONB)} ÷ {L(C_RMB)}. ~100 % = retraits réels (liquidité consommée, absorption) ; bas = retraits massifs sans trades (carnet qui ment). Sur IEX seul, les trades sont rares face aux changements de quotes : ce % est naturellement très bas (la quasi-totalité des retraits sont des annulations) ; lis-le en relatif d'une bougie à l'autre.",
        C_PCA: f"{L(C_CONA)} ÷ {L(C_RMA)}. Même lecture côté ask.",
        C_PCS: "% consommé bid − % consommé ask. Positif = les retraits côté bid sont plus « réels » que côté ask.",
        C_PULL: (f"Mur anormal (> {WALL_MULT:g} x médiane) présent >= {PULL_PRESENCE:.0%} de la bougie, retiré dans les {PULL_WINDOW_S} dernières secondes "
                 f"(taille revenue sous {WALL_GONE_MULT:g} x médiane ou niveau disparu), et retrait annulé à >= {PULL_CANCEL_FRAC:.0%} (non consommé par des trades). BID, ASK ou BID+ASK."),
        C_TRAP: (f"Si {L(C_PULL)} est rempli : « Piège retiré → sweep » si le delta % de la bougie SUIVANTE est >= {SEUIL_DELTA:.0%} dans le sens opposé au mur "
                 "(mur bid retiré -> delta vendeur ; mur ask retiré -> delta acheteur) ; « Retrait de mur – sans sweep » sinon ; « en attente » tant que la bougie suivante n'existe pas."),
        C_J1: (f"Médiane, sur les séances de l'historique local et les {2 * REF_SMOOTH + 1} tranches voisines (même heure New York), du déséquilibre pondéré temps (calcul identique à BE). "
               "Vide si BOOK_J1_BASELINE = False ou sans historique ni quotes J-1."),
        C_ISOBC: ("Volume des trades ISO dont la direction est CERTAINE et ACHAT : prix >= ask de la dernière cotation strictement antérieure (règle bid/ask, pas le milieu), "
                  f"ou balayage ISO (rafale < {SWEEP_GAP_MS} ms, >= {SWEEP_MIN_LEVELS} prix distincts toujours croissants). Condition non régulière, marché locked/croisé ou règles contradictoires = non classé."),
        C_ISOSC: ("Volume des trades ISO dont la direction est CERTAINE et VENTE : prix <= bid de la dernière cotation strictement antérieure, "
                  f"ou balayage ISO à prix toujours décroissants (rafale < {SWEEP_GAP_MS} ms, >= {SWEEP_MIN_LEVELS} niveaux)."),
        C_ISODC: "(Volume ISO acheteur certain − vendeur certain) ÷ volume de la bougie. Ne contient AUCUN trade classé par le milieu ou la tick rule. Colonne séparée : le Delta ISO historique (AT) n'est pas modifié.",
        C_ISOPC: "Part du volume ISO dont la direction est certaine (bid/ask ou sweep). Haute = le delta ISO certain est représentatif ; basse = beaucoup de trades au milieu, lire avec prudence.",
        C_ISODE: "(ISO acheteur − vendeur) ÷ volume, calculé UNIQUEMENT sur les trades ISO non certains (classés par le milieu / tick rule : sens historique). Delta ISO certain + estimé = delta ISO de toute la bougie avec la règle hybride.",
        C_J1MAD: "Dispersion historique du déséquilibre à cette heure : médiane des écarts absolus (MAD) des séances de l'historique, lissée ±1 tranche. Sert à normaliser l'écart (colonne suivante).",
        C_J1Z: (f"(BE − réf. normale) ÷ max({Z_MAD_FLOOR:.0%} ; 1,4826 × MAD). z = 2 : écart de 2 écarts-types habituels à cette heure. Un écart de +5 points est une anomalie là où "
                "l'historique varie de ±2 points, et un non-événement là où il varie de ±8 points. Vide sans historique (< HIST_MIN_DAYS séances)."),
        **{C_TCN1 + k: f"Delta EXÉCUTÉ des trades de taille {lbl} actions : volume acheteur − volume vendeur (sens du trade, comme la colonne N). Volume réellement échangé, contrairement aux ordres affichés qui peuvent être retirés. Bornes : TRADE_SIZE_CLASSES."
           for k, lbl in enumerate(_trade_class_labels())},
        **{C_TCV1 + k: f"Volume acheteur + vendeur des trades de taille {lbl} actions (dénominateur du delta % exécuté sur CONFLUENCE). Les trades au sens indéterminé n'y figurent pas."
           for k, lbl in enumerate(_trade_class_labels())},
        C_TAPE: f"Speed of tape = nombre de trades (colonne J) ÷ {BAR_MINUTES} min : vitesse d'impression du ruban, en trades par minute. Mesure l'AGITATION de l'exécution, pas le sens (lire avec le delta).",
        C_TAPEREF: f"Vitesse « normale » à cette heure : médiane (trades / min) des séances de l'historique local, même heure New York, lissée ±{REF_SMOOTH} bougie, plancher {TAPE_REF_FLOOR_FRAC:.0%} de la vitesse moyenne. Vide sans historique (< HIST_MIN_DAYS séances) ou si les anciens fichiers n'ont pas la colonne Trades.",
        C_TAPERAT: f"Vitesse du tape ÷ vitesse normale à la même heure. 1x = normal ; >= {TAPE_RATIO_ALERTE}x alerte ; >= {TAPE_RATIO_FORT}x fort. Un ratio élevé avec un volume normal = beaucoup de petits trades (retail / algos) ; un volume élevé avec un ratio normal = gros prints.",
        C_J1GAP: "BE − réf. normale (médiane historique), en points. Positif = le carnet penche plus acheteur que d'habitude à cette heure ; normalise l'effet horaire (ouverture, midi, clôture).",
    }
    for col, message in comments.items():
        if col > ncols:
            continue
        ws.cell(1, col).comment = Comment(message, "Script Alpaca")

    for i, (ts, r) in enumerate(bars.iterrows(), start=2):
        ny = ts.tz_convert(NY)
        n_bar = i - 1   # position de la bougie dans la fenêtre (1 = première)
        vals = {
            1: ny.date(), 2: n_bar, 3: ts.time(), 4: ny.time(),
            5: _num(r.open), 6: _num(r.high), 7: _num(r.low), 8: _num(r.close),
            9: int(r.volume), 10: int(r.trades), 11: _num(r.vwap),
            12: int(r.buy_vol), 13: int(r.sell_vol),
            14: f"=L{i}-M{i}",
            15: f"=IF(I{i}=0,0,N{i}/I{i})",
            16: f"=SUM($N$2:N{i})",
            17: _num(r.bid_last), 18: _num(r.bid_size_last),
            19: _num(r.ask_last), 20: _num(r.ask_size_last),
            21: f'=IF(OR(Q{i}="",S{i}=""),"",S{i}-Q{i})',
            22: f'=IF(OR(Q{i}="",S{i}=""),"",(Q{i}+S{i})/2)',
            23: _num(r.bid_avg), 24: _num(r.ask_avg),
            25: _num(r.bid_size_avg), 26: _num(r.ask_size_avg),
            27: _num(r.spread_avg),
            28: f'=IF(OR(Y{i}="",Z{i}=""),"",IF(Y{i}+Z{i}=0,0,(Y{i}-Z{i})/(Y{i}+Z{i})))',
            29: int(r.n_quotes),
            30: _num(r.max_trade),
            31: f"=IF(I{i}=0,0,AD{i}/I{i})",
            32: f'=IF(OR(BA{i}="",N(BA{i})<=0),"",I{i}/BA{i})',
            33: ("ACHAT" if r.max_trade_side > 0 else "VENTE" if r.max_trade_side < 0 else "")
                if pd.notna(r.max_trade_side) else "",
            34: (f'=IF(IF(O{i}>={SEUIL_DELTA},1,IF(O{i}<=-{SEUIL_DELTA},-1,0))'
                 f'+IF(AB{i}="",0,IF(AB{i}>={SEUIL_CARNET},1,IF(AB{i}<=-{SEUIL_CARNET},-1,0)))>0,"Pression achat",'
                 f'IF(IF(O{i}>={SEUIL_DELTA},1,IF(O{i}<=-{SEUIL_DELTA},-1,0))'
                 f'+IF(AB{i}="",0,IF(AB{i}>={SEUIL_CARNET},1,IF(AB{i}<=-{SEUIL_CARNET},-1,0)))<0,"Pression vente","Neutre"))'),
            # ---- à partir de AI ----
            35: f'=IF(J{i}=0,"",I{i}/J{i})',
            36: f'=IF(I{i}=0,"",1000*J{i}/I{i})',
            37: _rolling_ratio("AI", i),
            38: _rolling_ratio("J", i),
            39: _rolling_ratio("I", i),
            40: _rolling_ratio("AD", i),
            41: _whale_score(i),
            42: (f'=IF(I{i}=0,"Pas d’activité",IF(AO{i}="","Historique insuffisant",IF(AO{i}>={WHALE_FORT},"Forte",'
                 f'IF(AO{i}>={WHALE_ALERTE},"Modérée",IF(AO{i}>=25,"Faible","Non détectée")))))'),
            43: (f'=IF(AO{i}="","Non évalué",IF(AO{i}<{WHALE_ALERTE},"Non confirmé",'
                 f'IF(AR{i}="","Non confirmé",IF(AND(AR{i}>0,O{i}>={SEUIL_DELTA}),"Accumulation probable",'
                 f'IF(AND(AR{i}<0,O{i}<=-{SEUIL_DELTA}),"Distribution probable","Mixte / à confirmer")))))'),
            44: f'=IF(OR(AD{i}="",AG{i}=""),"",IF(AG{i}="ACHAT",AD{i},-AD{i}))',
            # ISO (nouveau) : valeurs issues des trades, aucune formule existante n'est touchée
            45: (float(r.iso_vol) / float(r.volume)) if r.volume > 0 else "",
            46: (float(r.iso_buy - r.iso_sell) / float(r.volume)) if r.volume > 0 else "",
            C_ISOBC: float(r.iso_buy_sure),
            C_ISOSC: float(r.iso_sell_sure),
            C_ISODC: (float(r.iso_buy_sure - r.iso_sell_sure) / float(r.volume)) if r.volume > 0 else "",
            C_ISOPC: (float(r.iso_buy_sure + r.iso_sell_sure) / float(r.iso_vol)) if r.iso_vol > 0 else "",
            C_ISODE: (float(r.iso_buy_est - r.iso_sell_est) / float(r.volume)) if r.volume > 0 else "",
            47: f'=IF(OR(H{i}="",K{i}="",K{i}=0),"",H{i}/K{i}-1)',
            48: f'=IF(OR(I{i}=0,E{i}="",H{i}=""),"",ABS(H{i}-E{i})/I{i}*10000)',
            49: f'=IF(OR(F{i}="",G{i}="",K{i}="",K{i}=0),"",(F{i}-G{i})/K{i})',
            50: f'=IF(I{i}=0,"",L{i}/I{i})',
            51: (f'=IF(I{i}=0,"",ROUND(100*(0.40*O{i}+0.20*IF(AB{i}="",0,AB{i})'
                 f'+0.20*IF(AR{i}="",0,AR{i}/I{i})'
                 f'+0.20*IF(AU{i}="",0,MAX(-1,MIN(1,AU{i}/{FLUX_VWAP_SCALE})))),0))'),
            53: (ref_slots[i - 2] if (ref_slots and (i - 2) < len(ref_slots)) else None),
            52: (f'=IF(AY{i}="","Pas de données",IF(AND(AY{i}>=40,OR(IF(AM{i}="",0,AM{i})>=1.5,IF(AF{i}="",0,AF{i})>={SIGNAL_RATIO_J1})),"Accumulation forte",'
                 f'IF(AND(AY{i}<=-40,OR(IF(AM{i}="",0,AM{i})>=1.5,IF(AF{i}="",0,AF{i})>={SIGNAL_RATIO_J1})),"Distribution forte",'
                 f'IF(AY{i}>=20,"Biais acheteur",IF(AY{i}<=-20,"Biais vendeur","Neutre / mixte")))))'),
            # ---- v2 « qualité du carnet » : BB -> BH = valeurs (seconde passe sur quotes/trades), BI = formule ----
            54: _num(r.q_coef),
            55: _num(r.churn),
            56: _num(r.size_rot),
            57: _num(r.imb_tw),
            58: _num(r.impact),
            59: None if pd.isna(r.absorb) else int(r.absorb),
            60: _num(r.confirm),
            61: (f'=IF(OR(BB{i}="",BH{i}=""),"Historique insuffisant",'
                 f'IF(AND(BB{i}>={VERDICT_FIABLE},BH{i}>={VERDICT_FIABLE}),"Fiable",'
                 f'IF(OR(BB{i}<{VERDICT_PEU_FIABLE},BH{i}<{VERDICT_PEU_FIABLE}),"Peu fiable","Moyen")))'),
            # ---- v3 « taille affichée crédible » : BJ -> CD ----
            C_PERS: _num(r["pers"]),
            C_WALLB: _num(r["wall_b"]),
            C_WALLA: _num(r["wall_a"]),
            C_WALLX: _num(r["wall_x"]),
            C_IMBP: _num(r["imb_p"]),
            C_IMBGAP: f'=IF(OR({L(C_IMBP)}{i}="",BE{i}=""),"",{L(C_IMBP)}{i}-BE{i})',
            C_CLS1: _num(r["cls1"]),
            C_CLS2: _num(r["cls2"]),
            C_CLS3: _num(r["cls3"]),
            C_CLS4: _num(r["cls4"]),
            C_RMB: _int_or_none(r["rm_b"]),
            C_CONB: _int_or_none(r["con_b"]),
            C_RMA: _int_or_none(r["rm_a"]),
            C_CONA: _int_or_none(r["con_a"]),
            C_PCB: f'=IF(OR({L(C_RMB)}{i}="",N({L(C_RMB)}{i})<=0),"",MIN(1,{L(C_CONB)}{i}/{L(C_RMB)}{i}))',
            C_PCA: f'=IF(OR({L(C_RMA)}{i}="",N({L(C_RMA)}{i})<=0),"",MIN(1,{L(C_CONA)}{i}/{L(C_RMA)}{i}))',
            C_PCS: f'=IF(OR({L(C_PCB)}{i}="",{L(C_PCA)}{i}=""),"",{L(C_PCB)}{i}-{L(C_PCA)}{i})',
            C_PULL: _pull_text(r["pull"]),
            C_TRAP: (f'=IF({L(C_PULL)}{i}="","",IF(NOT(ISNUMBER(O{i + 1})),"Retrait de mur – en attente",'
                     f'IF(OR(AND(ISNUMBER(SEARCH("BID",{L(C_PULL)}{i})),O{i + 1}<=-{SEUIL_DELTA}),'
                     f'AND(ISNUMBER(SEARCH("ASK",{L(C_PULL)}{i})),O{i + 1}>={SEUIL_DELTA})),'
                     f'"Piège retiré → sweep","Retrait de mur – sans sweep")))'),
            C_J1: _num(r["imb_ref"]),
            C_J1GAP: f'=IF(OR(BE{i}="",{L(C_J1)}{i}=""),"",BE{i}-{L(C_J1)}{i})',
            C_J1MAD: _num(r.get("imb_mad", np.nan)),
            C_J1Z: (f'=IF(OR({L(C_J1GAP)}{i}="",{L(C_J1MAD)}{i}=""),"",'
                    f'{L(C_J1GAP)}{i}/MAX({Z_MAD_FLOOR},1.4826*{L(C_J1MAD)}{i}))'),
            **{C_TCN1 + k: int(r[f"tcb{k + 1}"] - r[f"tcs{k + 1}"]) for k in range(4)},
            **{C_TCV1 + k: int(r[f"tcb{k + 1}"] + r[f"tcs{k + 1}"]) for k in range(4)},
            C_TAPE: f'=J{i}/{BAR_MINUTES}',
            C_TAPEREF: _num(r.get("tape_ref", np.nan)),
            C_TAPERAT: f'=IF(OR({L(C_TAPEREF)}{i}="",N({L(C_TAPEREF)}{i})<=0),"",{L(C_TAPE)}{i}/{L(C_TAPEREF)}{i})',
        }
        for c, v in vals.items():
            cell = ws.cell(i, c, v)
            cell.number_format = COLS[c - 1][3]
            cell.font, cell.border = F, BORDER
            cell.alignment = Alignment(horizontal="center" if (c <= 4 or c in (33, 34, 42, 43, 52, 61, C_PULL, C_TRAP)) else "right")

    last = len(bars) + 1
    t = last + 2
    ws.cell(t, 1, "TOTAL").font = FB
    for c, col in [(9, "I"), (10, "J"), (12, "L"), (13, "M"), (14, "N")]:
        cell = ws.cell(t, c, f"=SUM({col}2:{col}{last})")
        cell.number_format, cell.font = COLS[c - 1][3], FB
    cell = ws.cell(t, 15, f"=IF(I{t}=0,0,N{t}/I{t})")
    cell.number_format, cell.font = "0.0%", FB
    for c, formula, fmt in [
        (35, f'=IF(J{t}=0,"",I{t}/J{t})', "#,##0.0"),
        (36, f'=IF(I{t}=0,"",1000*J{t}/I{t})', "0.00"),
        (44, f'=SUM(AR2:AR{last})', "#,##0;-#,##0;0"),
        (50, f'=IF(I{t}=0,"",L{t}/I{t})', "0.0%"),
    ]:
        cell = ws.cell(t, c, formula)
        cell.number_format, cell.font = fmt, FB

    for rng in [f"N2:P{last}", f"N{t}:O{t}", f"AB2:AB{last}", f"AT2:AT{last}"]:
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN_FONT))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED_FONT))

    for rng, txt, fill, col in [
        (f"AG2:AG{last}", '"ACHAT"', FILL_G, "007A33"), (f"AG2:AG{last}", '"VENTE"', FILL_R, "C00000"),
        (f"AH2:AH{last}", '"Pression achat"', FILL_G, "007A33"), (f"AH2:AH{last}", '"Pression vente"', FILL_R, "C00000"),
        (f"AP2:AP{last}", '"Forte"', FILL_R, "C00000"),
        (f"AQ2:AQ{last}", '"Accumulation probable"', FILL_G, "007A33"),
        (f"AQ2:AQ{last}", '"Distribution probable"', FILL_R, "C00000"),
        (f"AZ2:AZ{last}", '"Accumulation forte"', FILL_G, "007A33"),
        (f"AZ2:AZ{last}", '"Distribution forte"', FILL_R, "C00000"),
    ]:
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[txt], fill=fill,
                                                      font=Font(name="Arial", size=10, bold=True, color=col)))

    for seuil, color in [(RATIO_FORT, "F4B183"), (RATIO_ALERTE, "FCE4D6")]:
        ws.conditional_formatting.add(
            f"AF2:AF{last}",
            CellIsRule(operator="greaterThanOrEqual", formula=[str(seuil)], stopIfTrue=True,
                       fill=PatternFill("solid", bgColor=color, fgColor=color),
                       font=Font(name="Arial", size=10, bold=True, color="C55A11")))

    ws.conditional_formatting.add(
        f"AO2:AO{last}",
        ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                       mid_type="num", mid_value=50, mid_color="FFD966",
                       end_type="num", end_value=100, end_color="F8696B"))
    ws.conditional_formatting.add(
        f"AY2:AY{last}",
        ColorScaleRule(start_type="num", start_value=-100, start_color="F8696B",
                       mid_type="num", mid_value=0, mid_color="FFFFFF",
                       end_type="num", end_value=100, end_color="63BE7B"))

    # ---- v2 : mises en forme des colonnes BB -> BI ----
    _add_rwg_scale(ws, f"BB2:BB{last}")                       # BB : rouge -> blanc -> vert (0 -> 100 %)
    ws.conditional_formatting.add(                            # BC : blanc -> bleu
        f"BC2:BC{last}",
        ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                       end_type="percentile", end_value=95, end_color="5B9BD5"))
    _add_sign_cf(ws, f"BE2:BE{last}")                         # BE : vert / rouge selon le signe
    _add_sign_cf(ws, f"BF2:BF{last}")                         # BF : vert / rouge selon le signe
    _add_whale_scale(ws, f"BG2:BG{last}")                     # BG : blanc -> jaune -> rouge (comme le score whale)
    _add_rwg_scale(ws, f"BH2:BH{last}")                       # BH : rouge -> blanc -> vert
    for txt, fill, font in [('"Fiable"', FILL_G, GREEN_FONT), ('"Moyen"', FILL_Y, FONT_Y), ('"Peu fiable"', FILL_R, RED_FONT)]:
        ws.conditional_formatting.add(f"BI2:BI{last}", CellIsRule(operator="equal", formula=[txt], fill=fill, font=font))

    # ---- v3 : mises en forme des colonnes BJ -> CD ----
    _add_rwg_scale(ws, f"{L(C_PERS)}2:{L(C_PERS)}{last}")
    for c in (C_WALLB, C_WALLA):
        ws.conditional_formatting.add(
            f"{L(c)}2:{L(c)}{last}",
            ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                           end_type="num", end_value=1, end_color="5B9BD5"))
    for c in (C_IMBP, C_IMBGAP, C_CLS1, C_CLS2, C_CLS3, C_CLS4, C_PCS, C_J1GAP, C_J1Z, C_ISODC, C_ISODE):
        _add_sign_cf(ws, f"{L(c)}2:{L(c)}{last}")
    for seuil, color in [(TAPE_RATIO_FORT, "F4B183"), (TAPE_RATIO_ALERTE, "FCE4D6")]:      # v9 : speed of tape
        ws.conditional_formatting.add(f"{L(C_TAPERAT)}2:{L(C_TAPERAT)}{last}", CellIsRule(
            operator="greaterThanOrEqual", formula=[str(seuil)], stopIfTrue=True,
            fill=PatternFill("solid", bgColor=color, fgColor=color),
            font=Font(name="Arial", size=10, bold=True, color="C55A11")))
    _add_rwg_scale(ws, f"{L(C_ISOPC)}2:{L(C_ISOPC)}{last}")          # % ISO certain : rouge -> blanc -> vert
    for c in (C_PCB, C_PCA):
        _add_rwg_scale(ws, f"{L(c)}2:{L(c)}{last}")
    for txt, fill, font in [('"Piège retiré → sweep"', FILL_R, RED_FONT), ('"Retrait de mur – sans sweep"', FILL_Y, FONT_Y),
                            ('"Retrait de mur – en attente"', FILL_O, FONT_O)]:
        ws.conditional_formatting.add(f"{L(C_TRAP)}2:{L(C_TRAP)}{last}",
                                      CellIsRule(operator="equal", formula=[txt], fill=fill, font=font))

    _hd = HIST_INFO.get(feed)
    lab = ws.cell(ref_row, 1, ("Moy. volume par bougie – médiane des " + str(len(_hd)) + " séances du "
                               + _hd[-1].strftime("%d/%m") + " au " + _hd[0].strftime("%d/%m/%Y") + " (référence normale"
                               + (f", dont {HIST_RETRO[feed]} postérieure(s) à la date analysée" if HIST_RETRO.get(feed) else "") + ")")
                  if _hd else ("Moy. volume par bougie – séance du "
                               + (ref_date.strftime("%d/%m/%Y") if ref_date else "n/a") + " (référence J-1)"))
    lab.font = FB
    ref = ws.cell(ref_row, 9, ref_avg)
    ref.number_format, ref.font = "#,##0", FB

    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:{L(ncols)}{last}"


# ----------------------------------------------------------------- feuille CONFLUENCE (12 ETF BTC côte à côte)
def _px_formulas(i, sheets_etf, wcr, ra, dyn_ref, prev_cvd_ref, ret_col, amp_col):
    """v12 : formules de la bougie i (ligne Excel) pour « CVD carnet prix ».

    Retourne (_px_ret, _px_amp, CVD carnet prix) :
      _px_ret = somme sur les ETF (poids = _wc_, ceux du carnet dynamique) de poids x rendement du close sur la bougie ;
      _px_amp = somme des mêmes poids x amplitude moyenne de l'ETF (colonne AW, fenêtre ra..i) ; seuls les ETF avec close courant ET précédent comptent ;
      réaction du prix = somme sur PX_BARS bougies de _px_ret / _px_amp (rendement en unités d'amplitude moyenne) ;
      k = bornée(signe(carnet) x réaction / PX_TAU, -1, +1) ; CVD(i) = CVD(i-1) x decay + carnet x k.
    """
    ret_t, amp_t = [], []
    for sh, w in zip(sheets_etf, wcr):
        ca, cb = f"'{sh}'!H{i}", f"'{sh}'!H{i - 1}"
        ok = f"AND(ISNUMBER({ca}),ISNUMBER({cb}),N({w})>0)"
        ret_t.append(f"IF({ok},IFERROR({w}*({ca}/{cb}-1),0),0)")
        amp_t.append(f"IF({ok},IFERROR({w}*AVERAGE('{sh}'!AW{ra}:AW{i}),0),0)")
    f_ret = "=" + "+".join(ret_t)
    f_amp = "=" + "+".join(amp_t)
    react = "+".join(f"IF(N({amp_col}{j})>0,N({ret_col}{j})/{amp_col}{j},0)" for j in range(max(2, i - PX_BARS + 1), i + 1))
    decay = 0.5 ** (BAR_MINUTES / PX_HALF_LIFE_MIN)
    f_cvd = (f'=IF(C{i}<TIME(9,30,0),"",{decay:.6f}*IF(ISNUMBER({prev_cvd_ref}),{prev_cvd_ref},0)'
             f'+IF(ISNUMBER({dyn_ref}),{dyn_ref}*MAX(-1,MIN(1,SIGN({dyn_ref})*({react})/{PX_TAU})),0))')
    return f_ret, f_amp, f_cvd


def write_confluence(ws, bars_by_sym, feed, oi=None, sip_cvd=None, auctions=None, sip_imb=None):
    """Feuille CONFLUENCE_<flux> : une ligne par bougie, les ETF BTC côte à côte (delta ET carnet de chacun),
    puis les verdicts de confluence, la colonne OI BTC et le bloc « moyenne des ETF » (indicateurs AF -> AX).

    `oi` (dict, même vide) ajoute la colonne OI BTC ; None la supprime.
    `auctions` ({sym: {...}}) ajoute les colonnes d'enchères ; vide / None = colonnes absentes.
    """
    syms = [s for s in ETFS if s in bars_by_sym]
    n = max((len(bars_by_sym[s]) for s in syms), default=0)
    if n == 0:
        ws["A1"] = f"Pas de données pour {feed}"
        return
    ref_bars = bars_by_sym[syms[0]]
    N = len(syms)
    lab = f"{N} ETF BTC"
    with_oi = oi is not None
    per = [("Close", "H", "0.0000"), ("Delta %", "O", "0.0%"), ("Carnet %", "AB", "0.0%"),
           ("Score whale", "AO", "0"), ("Score flux", "AY", "0;-0;0"), ("Signal flux", "AZ", "General")]
    W = len(per)   # 6 colonnes par ETF

    # ================= ORGANISATION DES COLONNES (feuilles CONFLUENCE uniquement) =================
    # 1) on DÉFINIT toutes les colonnes (hors N° bougie / heures / blocs ETF) dans l'ordre historique ;
    # 2) CONFLUENCE_FRONT dit lesquelles sont placées juste après « Heure New York » (et dans quel ordre) ;
    # 3) toutes les autres sont décalées à droite après les blocs ETF, dans leur ordre d'origine.
    # Toutes les formules / mises en forme / commentaires utilisent les positions calculées ici (pos[...]).
    # enchères : colonnes présentes seulement si Alpaca a renvoyé des volumes de croix pour au moins un ETF
    auc = auctions or {}
    auc_open = sum(float(auc.get(s, {}).get("open_size") or 0) for s in syms)
    auc_close = sum(float(auc.get(s, {}).get("close_size") or 0) for s in syms)
    has_auc = auc_open > 0 or auc_close > 0

    defs = {}                                     # clé -> (titre, largeur, couleur, format), ordre historique

    def add(key, title, w, fill, fmt):
        defs[key] = (title, w, fill, fmt)

    add("ratio", COLS[31][0] + f" (agrégé {N} ETF)", 13, PURPLE, COLS[31][3])
    # Delta / Carnet aligné pondérés (proportionnels) à la place des CVD Delta ; les colonnes « aligné (net) » restent.
    mid = [("ndw", "Delta aligné pondéré volume", 13, PURPLE, "+0%;-0%;0%"),
           ("cvd_d", "CVD cumulé delta pondéré (depuis 9h30 NY)", 16, PURPLE, "0.0%"),
           ("ncw", "Carnet aligné pondéré volume", 13, PURPLE, "+0%;-0%;0%"),
           ("cvd_c", "CVD moyen Carnet (depuis 9h30 NY)", 16, PURPLE, "0.0%")]
    if sip_cvd is not None:
        mid.append(("cvd_sip", "CVD carnet SIP (depuis 9h30 NY)", 16, PURPLE, "0.0%"))
    mid += [("iso_p", "% volume ISO moyen", 13, PURPLE, "0.0%"),
            ("iso_d", "Delta ISO moyen (signé)", 14, PURPLE, "0.0%;-0.0%;0.0%")]
    if with_oi:
        mid.append(("oi", "OI BTC agrégé (var. % / bougie)", 14, PURPLE, "0.00%"))
    mid += [("nd", "Delta aligné (net, pondéré volume)", 13, PURPLE, "+0%;-0%;0%"),
            ("nc", "Carnet aligné (net, pondéré volume)", 13, PURPLE, "+0%;-0%;0%"),
            ("vdc", "CONFLUENCE DELTA + CARNET", 30, PURPLE, "General"),
            ("vdcw", "CONFLUENCE DELTA + CARNET (pondéré volume)", 30, PURPLE, "General"),
            ("avg", f"Score flux moyen {lab}", 14, PURPLE, "0;-0;0"),
            ("net", "ETF alignés flux (net)", 13, PURPLE, "+0;-0;0"),
            ("wh", f"Whales simultanées (sur {N})", 13, PURPLE, "0"),
            ("ver", "CONFLUENCE FLUX GLOBAL", 28, PURPLE, "General")]
    for m in mid:
        add(*m)

    # ---- bloc « moyenne des ETF » : indicateurs des feuilles ETF (AG -> AR, AU -> AZ) ----
    # AF (ratio volume) = colonne « ratio » ; AS/AT (ISO) ont leurs propres colonnes agrégées pondérées.
    SRC = [32] + list(range(33, 45)) + list(range(47, 53))
    TEXT_SRC = {33, 34, 42, 43, 52}              # AG, AH, AP, AQ, AZ : colonnes texte -> verdict recalculé
    for sc in SRC:
        if sc == 32:
            continue
        title, w, _, fmt = COLS[sc - 1]
        suffix = f" ({lab})" if sc in TEXT_SRC else f" (moy. {N})"
        add(f"a{sc}", title + suffix, max(w, 18) if sc in TEXT_SRC else w, PURPLE, fmt)
    # Delta % et Carnet % moyens pondérés
    add("dm", f"Delta % moyen pondéré {lab}", 15, PURPLE, "0.0%")
    add("cm", f"Carnet % moyen pondéré {lab}", 15, PURPLE, "0.0%")

    # ---- v2 : bloc final (divergence, carnet qualité, absorption, confirmation, enchères) ----
    add("div_etf", "ETF divergent (signe opposé au consensus)", 22, PURPLE, "General")
    add("div_gap", "Écart max de divergence (pts, signé)", 15, PURPLE, "+0.0%;-0.0%;0.0%")
    add("div_ver", "VERDICT DIVERGENCE", 22, PURPLE, "General")
    add("qm", "Carnet qualité moyen (pondéré volume)", 14, PURPLE, "0%")
    add("qc", "Carnet % moyen qualité-pondéré", 15, PURPLE, "0.0%")
    add("abs", f"Absorption pondérée volume ({lab})", 13, PURPLE, "0")
    add("conf", f"Confirmation carnet pondérée volume ({lab})", 14, PURPLE, "0%")
    if has_auc:
        add("auc_o", f"Croix 9h30 – volume (Σ {N} ETF)", 15, PURPLE, "#,##0")
        add("auc_c", f"Croix 16h00 – volume (Σ {N} ETF)", 15, PURPLE, "#,##0")
        add("auc_r", "Croix clôture / volume du jour", 15, PURPLE, "0.0%")
    # colonnes helper masquées _gap_<ETF> : écart (en valeur absolue) de chaque ETF divergent, 0 sinon
    for k, sym in enumerate(syms):
        add(f"gap{k}", f"_gap_{sym}", 9, PURPLE, "0.0%")
    # v2.1 : indicateurs carnet BC -> BF + verdict BI agrégés (pondérés volume)
    add("chn", f"Churn carnet pondéré volume ({lab})", 13, PURPLE, "0.0")
    add("chn_ratio", f"Ratio churn carnet vs réf. normale (médiane {HIST_LABEL}, même tranche) (agrégé {N} ETF)", 13, PURPLE, '0.00"x"')
    add("chn_ref", "_réf_churn_J-1", 9, PURPLE, "0.0")
    add("tape_ratio", f"Speed of tape vs réf. normale (médiane {HIST_LABEL}, même tranche) (agrégé {N} ETF)", 13, NAVY, '0.00"x"')
    add("tape_ref", "_réf_tape", 9, NAVY, "0.0")
    add("rot", f"Rotation tailles pondérée volume ({lab})", 14, PURPLE, '0.00"x"')
    add("imb", f"Déséquilibre pondéré temps – moy. pondérée volume ({lab})", 16, PURPLE, "0.0%;-0.0%;0.0%")
    # v3 : colonnes à côté de « Déséquilibre pondéré temps » (M) ; M n'est pas modifiée
    add("imb_pers", f"Déséquilibre temps × persistance – moy. pondérée volume ({lab})", 16, PURPLE, "0.0%;-0.0%;0.0%")
    add("imb_gap", "Écart persistance − M (flashs)", 13, PURPLE, "+0.0%;-0.0%;0.0%")
    add("imb_hyb", f"Déséquilibre pondéré temps – poids volume × fiabilité ({lab})", 16, PURPLE, "0.0%;-0.0%;0.0%")
    if sip_imb is not None:
        add("imb_sip", "Écart IEX vs SIP – déséquilibre pondéré temps (abs.)", 14, PURPLE, "0.0%")
    for j, lbl in enumerate(_class_labels()):
        add(f"cls{j + 1}", f"Déséquilibre {lbl} actions – moy. pondérée volume", 14, PURPLE, "0.0%;-0.0%;0.0%")
    add("wall_d", f"Mur bid − mur ask (% bougie) – moy. pondérée volume ({lab})", 15, PURPLE, "+0%;-0%;0%")
    add("pers", f"Persistance affichage – moy. pondérée volume ({lab})", 13, PURPLE, "0%")
    add("cons_b", f"% consommé bid – Σ {N} ETF", 12, PURPLE, "0%")
    add("cons_a", f"% consommé ask – Σ {N} ETF", 12, PURPLE, "0%")
    add("cons_s", "Consommation bid − ask (Σ ETF)", 13, PURPLE, "+0%;-0%;0%")
    add("trap", "Alerte piège (retrait de mur → sweep)", 26, PURPLE, "General")
    add("imb_j1", f"Z-score déséquilibre vs réf. normale – poids carnet dynamique ({lab})", 16, PURPLE, "+0.00;-0.00;0.00")
    add("imp", f"Impact post-trade pondéré volume ({lab})", 14, PURPLE, "0.000;-0.000;0.000")
    add("vdq", "Verdict carnet global", 19, PURPLE, "General")
    # v2.2 : Carnet % moyen pondéré par un score de QUALITÉ COMPOSITE (spread x stabilité des tailles x track record BH).
    # Colonnes d'aide masquées : _imb_<ETF> (déséquilibre retenu) et _w_<ETF> (poids composite).
    add("qcc", f"Carnet % moyen qualité composite ({lab})", 16, PURPLE, "0.0%")
    for k, sym in enumerate(syms):
        add(f"hi{k}", f"_imb_{sym}", 9, PURPLE, "0.0%")
    for k, sym in enumerate(syms):
        add(f"hw{k}", f"_w_{sym}", 9, PURPLE, "0.00")
    # v4 : carnet dynamique (fiabilité mesurée sur la réaction du prix) + colonnes d'aide masquées par ETF
    add("dyn_c", f"Carnet % dynamique – pondéré fiabilité prix ({lab})", 16, PURPLE, "+0.0%;-0.0%;0.0%")
    add("dyn_cvd", f"CVD carnet à demi-vie ({REACT_HALF_LIFE_MIN} min)", 15, PURPLE, "0.0%")
    add("dyn_cvd_px", f"CVD carnet prix (demi-vie {PX_HALF_LIFE_MIN} min, validé par la réaction du prix)", 15, PURPLE, "0.0%")
    add("px_ret", "_px_ret", 9, PURPLE, "0.0000%")
    add("px_amp", "_px_amp", 9, PURPLE, "0.0000%")
    add("dyn_top", "Carnet le plus fiable (poids)", 18, PURPLE, "General")
    add("dyn_conf", "Taux de validation prix agrégé (TVP)", 14, PURPLE, "0%")
    add("dyn_gate", "Validation prix du carnet", 24, PURPLE, "General")
    add("dyn_ver", "Verdict carnet final (validé prix)", 26, PURPLE, "General")
    add("iso_dc", f"Delta ISO CERTAIN – moy. pondérée volume ({lab})", 15, PURPLE, "+0.0%;-0.0%;0.0%")
    add("iso_pc", f"% volume ISO classé CERTAIN ({lab})", 14, PURPLE, "0%")
    add("dtop_val", "Delta % de l'ETF au carnet le plus fiable", 15, PURPLE, "+0.0%;-0.0%;0.0%")
    add("dd_c", (f"Delta % dynamique – mêmes ETF et mêmes poids que le carnet ({lab})" if DELTA_SOURCE == "carnet" else
                 f"Delta % de l'ETF au carnet le plus fiable ({lab})" if DELTA_SOURCE == "top" else
                 f"Delta % dynamique – pondéré fiabilité prix ({lab})"), 16, PURPLE, "+0.0%;-0.0%;0.0%")
    add("dd_cvd", f"CVD delta à demi-vie ({REACT_HALF_LIFE_MIN} min)", 15, PURPLE, "0.0%")
    add("dd_top", "Delta le plus fiable (poids)", 18, PURPLE, "General")
    add("dd_conf", "TVP delta agrégé", 12, PURPLE, "0%")
    add("dd_gate", "Validation prix du delta", 24, PURPLE, "General")
    add("dc_sc", "Score DELTA + CARNET (validé prix)", 14, PURPLE, "+0.0%;-0.0%;0.0%")
    add("dc_ver", "CONFLUENCE DELTA + CARNET (validé prix)", 32, PURPLE, "General")
    add("imb_dyn", f"Déséquilibre pondéré temps – carnet le plus fiable, dynamique TVP ({lab})", 16, PURPLE, "0.0%;-0.0%;0.0%")
    for j, lbl in enumerate(_class_labels()):
        add(f"cls_dyn{j + 1}", f"Déséquilibre {lbl} actions – carnet le plus fiable, dynamique TVP", 14, PURPLE, "0.0%;-0.0%;0.0%")
    for j, lbl in enumerate(_trade_class_labels()):
        add(f"tx_cls{j + 1}", f"Delta % EXÉCUTÉ – trades {lbl} actions (Σ {N} ETF)", 14, PURPLE, "+0.0%;-0.0%;0.0%")
    for k, sym in enumerate(syms):
        add(f"wc{k}", f"_wc_{sym}", 9, PURPLE, "0.00")
    for q in (1, 2, 3):
        add(f"capS{q}", f"_cap{q}", 9, PURPLE, "0.00")
    for q in (1, 2, 3):
        add(f"dcapS{q}", f"_dcap{q}", 9, PURPLE, "0.00")
    for k, sym in enumerate(syms):
        add(f"drx{k}", f"_drx_{sym}", 9, PURPLE, "0")
    for k, sym in enumerate(syms):
        add(f"drp{k}", f"_dp_{sym}", 9, PURPLE, "0%")
    for k, sym in enumerate(syms):
        add(f"dwd{k}", f"_dwd_{sym}", 9, PURPLE, "0.00")
    for k, sym in enumerate(syms):
        add(f"sg{k}", f"_sg_{sym}", 9, PURPLE, "0.0%")
    for k, sym in enumerate(syms):
        add(f"rx{k}", f"_rx_{sym}", 9, PURPLE, "0")
    for k, sym in enumerate(syms):
        add(f"rp{k}", f"_p_{sym}", 9, PURPLE, "0%")
    for k, sym in enumerate(syms):
        add(f"wd{k}", f"_wd_{sym}", 9, PURPLE, "0.00")
    # v8 : colonnes d'appui du z-score carnet, AJOUTÉES TOUT À DROITE (aucune autre colonne ne bouge ; pour les avancer en tête de feuille, ajouter leur clé à CONFLUENCE_FRONT)
    add("imb_raw", f"Écart déséquilibre brut vs réf. normale – moy. pondérée volume (ancienne F) ({lab})", 16, PURPLE, "+0.0%;-0.0%;0.0%")
    add("imb_ema", "Z-score déséquilibre – lissé EMA (mémoire)", 15, PURPLE, "+0.00;-0.00;0.00")
    add("imb_ver", "Carnet vs historique (persistance 2 bougies)", 28, PURPLE, "General")

    # ---- placement ----
    optional = {"cvd_sip", "oi", "auc_o", "auc_c", "auc_r", "imb_sip"}          # absentes selon les options du run : ignorées
    for key in CONFLUENCE_FRONT:
        if re.fullmatch(r"(gap|hi|hw|sg|rx|rp|wd|wc|drx|drp|dwd)\d+", key):
            raise ValueError(f"CONFLUENCE_FRONT : « {key} » est une colonne d'aide masquée, non déplaçable")
        if key not in defs and key not in optional:
            raise ValueError(f"CONFLUENCE_FRONT : clé inconnue « {key} »")
    _tl = _trade_class_labels()
    for _key, _t, _w in CONFLUENCE_LAYOUT:       # v11 : en-têtes et largeurs de l'organisation demandée (la couleur violette est conservée)
        if _key in defs:
            _fmt = defs[_key][3]
            defs[_key] = (_t.format(N=N, T3=_tl[2], T4=_tl[3]), _w, PURPLE, _fmt)
    front = []
    for key in CONFLUENCE_FRONT:
        if key in defs and key not in front:
            front.append(key)
    rest = [k for k in defs if k not in front]
    OFF = 3 + len(front)       # colonnes avant les blocs ETF : N° bougie | Heure UTC | Heure NY | colonnes « front »
    pos = {}
    for j, key in enumerate(front):
        pos[key] = 4 + j
    for j, key in enumerate(rest):
        pos[key] = OFF + W * N + 1 + j
    heads = [("N° bougie", CONF_ABC_WIDTHS[0], NAVY, "0"), ("Heure UTC", CONF_ABC_WIDTHS[1], NAVY, "HH:MM"),
             ("Heure New York", CONF_ABC_WIDTHS[2], NAVY, "HH:MM")]
    heads += [defs[key] for key in front]
    for k, sym in enumerate(syms):
        for title, _, fmt in per:
            heads.append((f"{sym} {title}", 18 if title == "Signal flux" else 12, [NAVY, TEAL, RED][k % 3], fmt))
    heads += [defs[key] for key in rest]
    assert len(heads) == max(pos.values()) == len(defs) + 3 + W * N

    ratio_col = pos["ratio"]
    (c_ndw, c_cvd_d, c_ncw, c_cvd_c, c_iso_p, c_iso_d, c_nd, c_nc, c_vdc,
     c_vdcw, c_avg, c_net, c_wh, c_ver) = (
        pos[k] for k in ("ndw", "cvd_d", "ncw", "cvd_c", "iso_p", "iso_d", "nd", "nc", "vdc",
                         "vdcw", "avg", "net", "wh", "ver"))
    c_oi = pos.get("oi")
    c_cvd_sip = pos.get("cvd_sip")
    newcol = {32: ratio_col}
    newcol.update({sc: pos[f"a{sc}"] for sc in SRC if sc != 32})
    c_dm, c_cm = pos["dm"], pos["cm"]
    c_div_etf, c_div_gap, c_div_ver = pos["div_etf"], pos["div_gap"], pos["div_ver"]
    c_qm, c_qc, c_abs, c_conf = pos["qm"], pos["qc"], pos["abs"], pos["conf"]
    c_auc_o, c_auc_c, c_auc_r = pos.get("auc_o"), pos.get("auc_c"), pos.get("auc_r")
    c_gap0 = pos["gap0"]
    c_chn, c_rot, c_imb, c_imp, c_vdq, c_qcc = (pos[k] for k in ("chn", "rot", "imb", "imp", "vdq", "qcc"))
    c_imb0, c_w0 = pos["hi0"], pos["hw0"]
    c_dyn, c_dyn_cvd, c_dyn_top, c_dyn_conf, c_dyn_gate, c_dyn_ver = (pos[k] for k in ("dyn_c", "dyn_cvd", "dyn_top", "dyn_conf", "dyn_gate", "dyn_ver"))
    c_sg0, c_rx0, c_rp0, c_wd0 = pos["sg0"], pos["rx0"], pos["rp0"], pos["wd0"]
    c_px, c_px_ret, c_px_amp = pos["dyn_cvd_px"], pos["px_ret"], pos["px_amp"]
    c_cap = [pos[f"capS{q}"] for q in (1, 2, 3)]
    c_imb_dyn, c_wc0 = pos["imb_dyn"], pos["wc0"]
    c_iso_dc, c_iso_pc = pos["iso_dc"], pos["iso_pc"]
    c_cls_dyn = [pos[f"cls_dyn{j + 1}"] for j in range(4)]
    c_tx = [pos[f"tx_cls{j + 1}"] for j in range(4)]
    c_dcap = [pos[f"dcapS{q}"] for q in (1, 2, 3)]
    c_dtop_val, c_dd, c_dd_cvd, c_dd_top, c_dd_conf, c_dd_gate, c_dc_sc, c_dc_ver = (
        pos[k] for k in ("dtop_val", "dd_c", "dd_cvd", "dd_top", "dd_conf", "dd_gate", "dc_sc", "dc_ver"))
    c_drx0, c_drp0, c_dwd0 = pos["drx0"], pos["drp0"], pos["dwd0"]
    c_imbp, c_imbg, c_hyb = pos["imb_pers"], pos["imb_gap"], pos["imb_hyb"]
    c_imbs = pos.get("imb_sip")
    c_cls = [pos[f"cls{j + 1}"] for j in range(4)]
    c_walld = pos["wall_d"]
    c_pers, c_conb, c_cona, c_cons, c_trap, c_j1 = (pos[k] for k in ("pers", "cons_b", "cons_a", "cons_s", "trap", "imb_j1"))
    c_imb_raw, c_imb_ema, c_imb_ver = pos["imb_raw"], pos["imb_ema"], pos["imb_ver"]
    _n_hist = len(HIST_INFO.get(feed) or [])
    _z_conf = min(1.0, _n_hist / float(HIST_CONF_FULL)) if HIST_CONF_FULL else 1.0     # confiance dans la référence : n séances / HIST_CONF_FULL (réduit F les premiers jours)
    sheets_etf = [f"{sy}_{feed}" for sy in syms]

    for c, (title, w, fill, _) in enumerate(heads, 1):
        cell = ws.cell(1, c, title)
        cell.font, cell.fill, cell.border = F_HEAD, fill, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(c)].width = w
    ws.row_dimensions[1].height = CONF_HEADER_HEIGHT
    for k in range(N):
        ws.column_dimensions[L(c_gap0 + k)].hidden = True
        ws.column_dimensions[L(c_imb0 + k)].hidden = True
        ws.column_dimensions[L(c_w0 + k)].hidden = True
        for c_h in (c_sg0, c_rx0, c_rp0, c_wd0, c_wc0, c_drx0, c_drp0, c_dwd0):
            ws.column_dimensions[L(c_h + k)].hidden = True
    for c_h in c_cap + c_dcap:
        ws.column_dimensions[L(c_h)].hidden = True
    ws.column_dimensions[L(c_px_ret)].hidden = True            # v12 : colonnes d'aide du CVD carnet prix
    ws.column_dimensions[L(c_px_amp)].hidden = True
    ws.column_dimensions[L(pos["chn_ref"])].hidden = True      # réf. churn J-1 (Σ ETF) : colonne d'aide du ratio churn
    ws.column_dimensions[L(pos["tape_ref"])].hidden = True     # v9 : réf. speed of tape (Σ ETF) : colonne d'aide du ratio tape
    if with_oi:
        ws.cell(1, c_oi).comment = Comment(
            "Variation de l'Open Interest BTC agrégé (perp + futures à échéance des principaux exchanges, en USD) sur la "
            "bougie : OI de fin ÷ OI de début − 1. Source Coinalyze, valeur fixe (pas une formule), identique sur IEX et SIP "
            "et pour les ETF (proxy BTC). Hausse = nouvelles positions ouvertes ; baisse = positions fermées. "
            "N'entre dans aucun score ni verdict.", "Script Alpaca")
    ws.cell(1, c_dm).comment = Comment(
        "Moyenne du Delta % de chaque ETF pondérée par son volume de la bougie (Σ(Delta%×Volume)/ΣVolume), "
        "et non une simple moyenne arithmétique : un ETF illiquide (ex. MSBT, BTCO) ne peut plus, à lui seul, "
        "faire basculer le signal face aux gros ETF (IBIT, FBTC, GBTC...).", "Script Alpaca")
    ws.cell(1, c_cm).comment = Comment(
        "Moyenne du Carnet % de chaque ETF pondérée par le même volume que Delta % moyen pondéré, pour rester "
        "cohérente avec elle (le carnet n'a pas de volume propre).", "Script Alpaca")
    ws.cell(1, c_iso_p).comment = Comment(
        "Part du volume exécutée en ordres ISO (Intermarket Sweep Order, condition F), moyenne des ETF pondérée par leur "
        "volume de la bougie. Ne remplace aucun delta.", "Script Alpaca")
    ws.cell(1, c_iso_d).comment = Comment(
        "(Volume ISO acheteur − volume ISO vendeur) ÷ volume total, moyenne des ETF pondérée par leur volume. "
        "Colonne séparée : Delta % moyen pondéré n'est pas modifié.", "Script Alpaca")
    ws.cell(1, c_ndw).comment = Comment(
        f"Pour chaque ETF : valeur = Delta % ÷ {ALIGN_SCALE:.0%}, limitée entre -1 et +1 (donc proportionnelle, sans effet de seuil). "
        "On pondère par le volume de la bougie : Σ(valeur × volume) ÷ Σ volume. De -100 % à +100 %. "
        "Les ETF sans trade ont un poids nul. « Delta aligné (net) » reste affiché (comptage ±1 par ETF).", "Script Alpaca")
    ws.cell(1, c_ncw).comment = Comment(
        f"Même calcul que « Delta aligné pondéré volume », appliqué au Carnet % (÷ {ALIGN_SCALE:.0%}, limité à ±1, pondéré volume).",
        "Script Alpaca")
    ws.cell(1, c_vdcw).comment = Comment(
        f"Moyenne de Delta aligné pondéré et Carnet aligné pondéré : >= {DCW_FORT:.0%} achat fort, >= {DCW_SIMPLE:.0%} achat, "
        f"<= -{DCW_SIMPLE:.0%} vente, <= -{DCW_FORT:.0%} vente forte. Divergence si l'un est >= +{DCW_SIMPLE:.0%} et l'autre <= -{DCW_SIMPLE:.0%}. "
        "Le verdict non pondéré reste affiché.", "Script Alpaca")
    ws.cell(1, c_cvd_d).comment = Comment(
        "Cumul, depuis 9h30 heure de New York, d'un delta moyen par bougie où chaque ETF pèse volume × |Delta %| : "
        "Σ(delta × volume × |delta|) ÷ Σ(volume × |delta|). Les ETF à la fois très directionnels et très échangés dominent. "
        "Vide avant 9h30 NY.", "Script Alpaca")
    ws.cell(1, c_cvd_c).comment = Comment(
        "Cumul, depuis 9h30 heure de New York, de Carnet % moyen pondéré, bougie après bougie. Vide avant 9h30 NY.",
        "Script Alpaca")
    ws.cell(1, c_net).comment = Comment(
        f"Nombre d'ETF acheteurs (score >= {CONF_SEUIL}) moins nombre d'ETF vendeurs (score <= -{CONF_SEUIL}). "
        f"Confluence forte si l'écart représente >= {COIN_CONF_FORT:.0%} des ETF ayant une donnée, simple si >= {COIN_CONF_SIMPLE:.0%}. "
        "Avec un seul ETF, la confluence reprend simplement son signal.", "Script Alpaca")
    ws.cell(1, c_vdc).comment = Comment(
        f"Moyenne de Delta aligné et Carnet aligné (net, pondérés volume) : >= {DC_FORT:.0%} fort, >= {DC_SIMPLE:.0%} simple. "
        f"Divergence si l'un est >= +50 % et l'autre <= -50 %.", "Script Alpaca")
    ws.cell(1, c_nd).comment = Comment(
        f"Chaque ETF vote +1 si Delta % >= {SEUIL_DELTA:.0%}, -1 si <= -{SEUIL_DELTA:.0%}, 0 sinon ; les votes sont pondérés par le "
        "volume de la bougie de chaque ETF : Σ(vote × volume) ÷ Σ volume. De -100 % à +100 %.", "Script Alpaca")
    ws.cell(1, c_nc).comment = Comment(
        f"Même calcul que « Delta aligné (net, pondéré volume) », avec Carnet % et le seuil de {SEUIL_CARNET:.0%}.", "Script Alpaca")
    if c_cvd_sip:
        ws.cell(1, c_cvd_sip).comment = Comment(
            "Même CVD Carnet que la colonne voisine, mais lu dans la feuille CONFLUENCE_SIP (carnet consolidé SIP) pour comparer "
            "avec le carnet IEX. Vide avant 9h30 NY ou si la bougie SIP n'est pas encore disponible.", "Script Alpaca")
    # commentaires du bloc v2
    ws.cell(1, c_div_etf).comment = Comment(
        f"Consensus = « Delta % moyen pondéré ». Un ETF est divergent si son Delta % est de signe OPPOSÉ au consensus, si "
        f"|consensus| >= {DIV_CONSENSUS_MIN:.0%} et si l'écart |Delta % ETF − consensus| >= {DIV_GAP:.0%} ({DIV_GAP * 100:.0f} pts). "
        "Affiche l'ETF avec le plus grand écart et son ampleur signée. Calculé via les colonnes masquées _gap_<ETF> (à droite).",
        "Script Alpaca")
    ws.cell(1, c_div_gap).comment = Comment(
        "Plus grand écart de divergence, signé comme le sens de l'ETF divergent : vert = l'ETF est acheteur alors que le consensus "
        "est vendeur ; rouge = l'ETF est vendeur alors que le consensus est acheteur. 0 = pas de divergence.", "Script Alpaca")
    ws.cell(1, c_div_ver).comment = Comment(
        "Divergence persistante (jaune) : au moins un ETF divergent sur cette bougie ET sur la précédente. Divergence récente "
        "(orange) : divergence sur cette bougie seulement. Sinon : Pas de divergence.", "Script Alpaca")
    ws.cell(1, c_qm).comment = Comment(
        "Moyenne des coefficients qualité carnet (colonne BB des feuilles ETF) pondérée par le volume de la bougie de chaque ETF.",
        "Script Alpaca")
    ws.cell(1, c_qc).comment = Comment(
        "Moyenne du Carnet % de chaque ETF pondérée par son coefficient qualité (BB) : les carnets serrés (spread proche de "
        "d'habitude) comptent plus que les carnets bruités. C'est la pondération « inverse du spread ».", "Script Alpaca")
    ws.cell(1, c_abs).comment = Comment(
        "Moyenne du score d'absorption (colonne BG des feuilles ETF) PONDÉRÉE par le volume de la bougie de chaque ETF : "
        "Σ(BG × volume) ÷ Σ(volume), ETF sans valeur ignorés.", "Script Alpaca")
    ws.cell(1, c_conf).comment = Comment(
        "Moyenne de la confirmation carnet -> exécution (colonne BH des feuilles ETF) PONDÉRÉE par le volume de la bougie de "
        "chaque ETF : Σ(BH × volume) ÷ Σ(volume), ETF sans valeur ignorés.", "Script Alpaca")
    ws.cell(1, pos["chn_ratio"]).comment = Comment(
        f"Activité du carnet vs normale : Σ churn des ETF (MàJ NBBO / min, colonne BC) ÷ Σ churn « normal » à la même heure "
        f"(colonne masquée _réf_churn_J-1 = médiane des séances de l'historique local, lissée ±{REF_SMOOTH} bougie, plancher {CHURN_REF_FLOOR_FRAC:.0%} du churn moyen). "
        f"1x = normal pour l'heure ; >= {CHURN_RATIO_ALERTE}x alerte ; >= {CHURN_RATIO_FORT}x fort. Mesure l'agitation des cotations, pas le sens. "
        f"Seuls les ETF avec un churn du jour et une réf. J-1 valides entrent dans le calcul.", "Script Alpaca")
    ws.cell(1, pos["tape_ratio"]).comment = Comment(
        f"SPEED OF TAPE : Σ vitesse des ETF (trades / min, colonne CL des feuilles ETF) ÷ Σ vitesse « normale » à la même heure "
        f"(colonne masquée _réf_tape = médiane des séances de l'historique local, lissée ±{REF_SMOOTH} bougie, plancher {TAPE_REF_FLOOR_FRAC:.0%} de la vitesse moyenne). "
        f"1x = normal ; >= {TAPE_RATIO_ALERTE}x alerte ; >= {TAPE_RATIO_FORT}x fort. Mesure la cadence des exécutions, pas le sens : à lire avec le delta et le ratio volume. "
        f"Seuls les ETF avec une réf. valide entrent dans le calcul.", "Script Alpaca")
    ws.cell(1, c_chn).comment = Comment("Churn (MàJ NBBO / min, colonne BC des feuilles ETF), moyenne pondérée par le volume de chaque ETF.", "Script Alpaca")
    ws.cell(1, c_rot).comment = Comment("Rotation des tailles max/min (colonne BD des feuilles ETF), moyenne pondérée par le volume de chaque ETF.", "Script Alpaca")
    ws.cell(1, c_imb).comment = Comment("Déséquilibre bid/ask pondéré par la durée d'affichage (colonne BE des feuilles ETF), moyenne pondérée par le volume de chaque ETF. Positif = carnets acheteurs.", "Script Alpaca")
    ws.cell(1, c_imp).comment = Comment("Impact post-trade (× spread, 2 s, colonne BF des feuilles ETF), moyenne pondérée par le volume de chaque ETF.", "Script Alpaca")
    ws.cell(1, c_qcc).comment = Comment(
        "Moyenne des déséquilibres de carnet des ETF pondérée par un score de qualité composite = coef qualité BB (spread) "
        "x stabilité des tailles (médiane de la rotation BD des 12 bougies précédentes ÷ rotation actuelle, plafonné à 1) "
        f"x track record du carnet (BH : {VERDICT_PEU_FIABLE:.0%} ou moins -> {TRACK_FLOOR:.0%}, {VERDICT_FIABLE:.0%} ou plus -> 100 %, linéaire entre les deux ; "
        f"BH vide -> {TRACK_NEUTRAL:.0%}). Déséquilibre retenu par ETF : BE (pondéré temps), à défaut Carnet %. "
        "Le volume n'intervient PAS. Voir les colonnes masquées _imb_<ETF> et _w_<ETF>.", "Script Alpaca")
    ws.cell(1, c_imbp).comment = Comment(
        "Moyenne des « Déséquilibre temps × persistance » (colonne BN des feuilles ETF) pondérée par le volume de chaque ETF. Chaque état du "
        "carnet pèse durée x persistance (flashs minorés, murs installés conservés). À lire avec M : M reste inchangée.", "Script Alpaca")
    ws.cell(1, c_imbg).comment = Comment(
        "Moyenne pondérée volume de (BN − BE) par ETF. M fort avec cet écart très négatif/positif = carnet rempli de flashs : M est moins crédible.", "Script Alpaca")
    ws.cell(1, c_hyb).comment = Comment(
        "Même BE que M, mais chaque ETF pèse volume x fiabilité : Σ(BE x volume x _w_) ÷ Σ(volume x _w_), avec _w_ = qualité BB x stabilité des tailles BD "
        "x track record BH (colonnes masquées _w_<ETF>). Un gros ETF au carnet bruité n'écrase plus un petit ETF au carnet propre.", "Script Alpaca")
    if c_imbs:
        ws.cell(1, c_imbs).comment = Comment(
            "|M IEX − M SIP| sur la même bougie (M = Déséquilibre pondéré temps). Écart large = le carnet IEX est trompeur à ce moment-là : "
            "ne pas se fier à M en IEX. Colonne présente seulement dans CONFLUENCE_IEX et quand le SIP est disponible.", "Script Alpaca")
    for j, c_ in enumerate(c_cls):
        ws.cell(1, c_).comment = Comment(
            f"Moyenne pondérée volume du déséquilibre bid/ask (pondéré durée) restreint aux états dont la plus grande taille est dans la classe {_class_labels()[j]} actions. "
            "Retail (petites tailles) vs institutionnel (grosses tailles) ne racontent pas la même histoire. Bornes : SIZE_CLASSES.", "Script Alpaca")
    ws.cell(1, c_walld).comment = Comment(
        "Moyenne pondérée volume de (mur bid − mur ask), soit BK − BL des feuilles ETF (part de la bougie où la taille dépasse 3 x sa médiane). "
        "Positif = murs bid dominants (support affiché) ; négatif = murs ask dominants (résistance affichée) ; ~0 = pas de mur ou murs des deux côtés "
        "qui s'annulent. À lire avec la persistance et le % consommé pour juger si le mur est réel.", "Script Alpaca")
    ws.cell(1, c_pers).comment = Comment(
        "Moyenne pondérée volume de la persistance d'affichage (colonne BJ). Bas = carnets de flashs ; haut = carnets installés.", "Script Alpaca")
    ws.cell(1, c_conb).comment = Comment(
        "Σ des actions retirées du bid et consommées par des trades ÷ Σ des actions retirées du bid, sur les ETF (et non une moyenne de pourcentages). "
        "~100 % = liquidité réelle ; bas = retraits massifs sans trades, carnet qui ment.", "Script Alpaca")
    ws.cell(1, c_cona).comment = Comment("Même calcul côté ask.", "Script Alpaca")
    ws.cell(1, c_cons).comment = Comment("% consommé bid − % consommé ask.", "Script Alpaca")
    ws.cell(1, c_trap).comment = Comment(
        "Nombre d'ETF en alerte « retrait de mur pré-impact » (colonne CB des feuilles ETF) : mur anormal présent >= 60 % de la bougie, retiré dans les 30 "
        "dernières secondes, surtout annulé ; « → sweep » si le delta de la bougie suivante est fort dans le sens opposé au mur.", "Script Alpaca")
    ws.cell(1, c_j1).comment = Comment(
        "v8 : z-score du carnet vs l'historique à la MÊME heure. Par ETF : z = (BE − médiane historique) ÷ (1,4826 × MAD historique, plancher "
        f"{Z_MAD_FLOOR:.0%}) (colonne CK des feuilles ETF), puis moyenne pondérée par le poids du carnet dynamique (_wc_ : qualité x stabilité x fiabilité prix, plafonné) "
        "et non plus par le volume (un seul ETF dominait). 0 = carnet habituel pour l'heure ; +2 = nettement plus acheteur que d'habitude ; -2 = plus vendeur. "
        f"Multiplié par la confiance dans la référence (séances utilisées ÷ {HIST_CONF_FULL}, plafonné à 1). Vide sans historique. L'ancien écart brut est conservé tout à droite.", "Script Alpaca")
    ws.cell(1, c_imb_raw).comment = Comment("Ancienne colonne F : Σ(CD × volume) ÷ Σ(volume), écart brut en points de % vs la médiane historique, pondéré par le volume.", "Script Alpaca")
    ws.cell(1, c_imb_ema).comment = Comment(f"EMA du z-score de F (coefficient {Z_EMA_ALPHA}) : une valeur par bougie de 5 min, avec mémoire des bougies précédentes.", "Script Alpaca")
    ws.cell(1, c_imb_ver).comment = Comment(
        f"Verdict sur le z lissé, exigeant 2 bougies consécutives du même côté : |z| >= {Z_ANOMALIE} « anormalement », >= {Z_FORT} « très ». "
        "Seuils posés sur les percentiles ~90 / ~97 de 8 jours d'IEX : à recalibrer. Aucun pouvoir prédictif démontré à ce jour (hit-rate 50-54 % sur 8 jours) : indicateur de contexte.", "Script Alpaca")
    ws.cell(1, c_vdq).comment = Comment(
        f"Mêmes règles que la colonne BI des feuilles ETF, appliquées aux valeurs agrégées : Fiable si qualité (Carnet qualité moyen) >= {VERDICT_FIABLE:.0%} "
        f"ET confirmation pondérée >= {VERDICT_FIABLE:.0%} ; Peu fiable si l'une des deux < {VERDICT_PEU_FIABLE:.0%} ; sinon Moyen ; "
        "Historique insuffisant si l'une des deux est vide.", "Script Alpaca")
    if has_auc:
        ws.cell(1, c_auc_o).comment = Comment(
            "Volume de la croix d'ouverture (9h30 NY), somme des ETF, d'après l'endpoint d'enchères d'Alpaca. Valeur fixe, "
            "affichée sur la bougie qui contient 9h30. Prix/volume de la croix = print de plus grande taille.", "Script Alpaca")
        ws.cell(1, c_auc_c).comment = Comment(
            "Volume de la croix de clôture (16h00 NY), somme des ETF. Valeur fixe, affichée sur la bougie qui contient 16h00.",
            "Script Alpaca")
        ws.cell(1, c_auc_r).comment = Comment(
            "Volume de la croix de clôture ÷ volume total du jour des ETF (feuilles ETF, ligne TOTAL). Poids de la croix de clôture ; "
            "ce n'est pas un imbalance officiel.", "Script Alpaca")

    ws.cell(1, c_dyn).comment = Comment(
        "Signal carnet de chaque ETF (BN = déséquilibre x persistance, à défaut BE, à défaut Carnet %) pondéré par son TAUX DE VALIDATION PRIX "
        f"(TVP) : le carnet de la bougie j (>= {REACT_TRIGGER:.0%}) a-t-il annoncé le sens du close à j+{REACT_H} (mouvement >= {REACT_THR_AMP:g} x amplitude moyenne) ? "
        f"TVP sur {REACT_WINDOW} bougies, lissé ({REACT_PRIOR} cas fictifs à 50 %), neutre sous {REACT_MIN_CASES} cas. Fiabilité 0 à 50 %, 100 % à {REACT_FULL:.0%}, plancher {REACT_FLOOR:.0%}. "
        f"Poids = qualité (spread BB x stabilité des tailles x persistance) x fiabilité^{REACT_GAMMA}, SANS volume exécuté, plafonné à {REACT_CAP:.0%} par ETF. "
        f"ETF exclu s'il a moins de {REACT_MIN_TRADES} trades sur la fenêtre. Voir colonnes masquées _sg_, _rx_, _p_, _wd_.", "Script Alpaca")
    ws.cell(1, c_dyn_cvd).comment = Comment(
        f"Carnet % dynamique cumulé à mémoire courte : CVD(i) = CVD(i-1) x 0,5^({BAR_MINUTES}/{REACT_HALF_LIFE_MIN}) + carnet dynamique(i). "
        "Une bougie perd la moitié de son poids après la demi-vie. Vide avant 9h30 NY.", "Script Alpaca")
    ws.cell(1, c_px).comment = Comment(
        "Même carnet dynamique que « CVD carnet », mais chaque bougie n'est créditée que dans la mesure où le PRIX a réagi dans son sens : "
        "CVD(i) = CVD(i-1) x 0,5^(" + f"{BAR_MINUTES}/{PX_HALF_LIFE_MIN}" + ") + carnet dynamique(i) x k(i). "
        f"k = réaction du prix (rendement pondéré des ETF sur {PX_BARS} bougies ÷ amplitude moyenne) dans le sens du carnet ÷ {PX_TAU:g}, bornée entre -1 et +1 : "
        "+1 = carnet confirmé par le prix, 0 = carnet sans effet sur le prix (ignoré), -1 = le prix va à contre-sens du carnet (absorption : la contribution s'inverse). "
        "Se retourne donc quand le prix casse alors que le carnet reste acheteur. « CVD carnet » est conservée à côté pour comparer. Colonnes masquées _px_ret et _px_amp.", "Script Alpaca")
    ws.cell(1, c_dyn_top).comment = Comment(
        "ETF qui porte le poids le plus élevé dans le carnet dynamique sur la bougie (après plafond), et son poids en % du total.", "Script Alpaca")
    ws.cell(1, c_dyn_conf).comment = Comment(
        "TVP (carnet -> sens du prix) moyen des ETF, pondéré par les mêmes poids dynamiques. 50 % = le carnet agrégé ne prédit rien.", "Script Alpaca")
    ws.cell(1, c_dyn_gate).comment = Comment(
        f"« Carnet validé par le prix » si le TVP agrégé >= {REACT_GATE:.0%} ; sinon « Carnet non fiable – attendre » : on ne valide pas un carnet "
        "avant d'avoir vu la réaction du prix.", "Script Alpaca")
    ws.cell(1, c_dyn_ver).comment = Comment(
        f"Carnet dynamique >= +{REACT_VERDICT:.0%} : ACHAT ; <= -{REACT_VERDICT:.0%} : VENTE, uniquement si le TVP agrégé >= {REACT_GATE:.0%} ; sinon Neutre / carnet non fiable.", "Script Alpaca")
    ws.cell(1, c_imb_dyn).comment = Comment(
        "Même logique que « Déséquilibre pondéré temps » (BE de chaque ETF, pondéré durée), mais chaque ETF est pondéré par son poids "
        "DYNAMIQUE TVP (qualité du carnet x fiabilité prix^2, plafonné) : le carnet le plus fiable pèse le plus, sans volume exécuté. "
        "Remplace « poids volume x fiabilité » en tête de feuille (l'ancienne colonne reste à droite des blocs ETF). Poids visibles dans _wc_<ETF>.", "Script Alpaca")
    for j, c_ in enumerate(c_tx):
        ws.cell(1, c_).comment = Comment(
            f"Delta % EXÉCUTÉ des trades de taille {_trade_class_labels()[j]} actions : Σ (achat − vente) ÷ Σ (achat + vente) de ces trades, sur tous les ETF "
            "(colonnes CO -> CV des feuilles ETF). Contrairement aux colonnes « ordres affichés » (déplacées à droite), il s'agit de volume réellement exécuté : "
            "petits trades = plutôt retail, gros trades = plutôt institutionnel. Vide si aucun trade de cette taille sur la bougie. Bornes : TRADE_SIZE_CLASSES.", "Script Alpaca")
    for j, c_ in enumerate(c_cls_dyn):
        ws.cell(1, c_).comment = Comment(
            f"Même logique que « Déséquilibre {_class_labels()[j]} actions » (déséquilibre pondéré durée restreint aux états dont la plus grande taille "
            "est dans la classe), mais pondéré par le poids DYNAMIQUE TVP de chaque ETF (carnet le plus fiable = poids le plus élevé, sans volume). "
            "Remplace la version pondérée volume en tête de feuille (elle reste à droite des blocs ETF).", "Script Alpaca")
    ws.cell(1, c_iso_dc).comment = Comment(
        "Delta ISO CERTAIN (colonne CG des feuilles ETF : trades ISO classés par la règle bid/ask ou par balayage, jamais par le milieu / tick rule), "
        "moyenne des ETF pondérée par le volume de la bougie. Colonne séparée : « Delta ISO » historique inchangé.", "Script Alpaca")
    ws.cell(1, c_iso_pc).comment = Comment(
        "Part du volume ISO (Σ ETF) dont la direction est certaine. Haute = le Delta ISO certain représente bien le flux ISO ; basse = beaucoup de trades au milieu du spread.", "Script Alpaca")
    ws.cell(1, c_dtop_val).comment = Comment(
        "Delta % (delta ÷ volume de la bougie) de l'ETF qui porte le plus gros poids dans le carnet dynamique, c'est-à-dire l'ETF "
        "indiqué dans « Carnet le plus fiable (poids) ».", "Script Alpaca")
    ws.cell(1, c_dd).comment = Comment(
        f"Même exercice que le carnet, appliqué au DELTA : TVP delta de chaque ETF = part des cas où le Delta % de la bougie j (>= {SEUIL_DELTA:.0%}) "
        f"a annoncé le sens du close à j+{REACT_H} (mouvement >= {REACT_THR_AMP:g} x amplitude moyenne). Lissé, neutre sous {REACT_MIN_CASES} cas. "
        f"Poids = racine(volume sur {REACT_WINDOW} bougies) x fiabilité^{REACT_GAMMA}, plafonné à {REACT_CAP:.0%} par ETF. "
        "Le volume garde un rôle modéré car il décrit l'exécution. Colonnes masquées _drx_, _dp_, _dwd_.", "Script Alpaca")
    ws.cell(1, c_dd_cvd).comment = Comment(
        f"Delta % dynamique cumulé à mémoire courte (demi-vie {REACT_HALF_LIFE_MIN} min). Vide avant 9h30 NY.", "Script Alpaca")
    ws.cell(1, c_dd_top).comment = Comment("ETF qui porte le plus gros poids dans le delta dynamique, et son poids en %.", "Script Alpaca")
    ws.cell(1, c_dd_conf).comment = Comment("TVP delta moyen des ETF, pondéré par les poids du delta dynamique. 50 % = le delta ne prédit rien.", "Script Alpaca")
    ws.cell(1, c_dd_gate).comment = Comment(f"« Delta validé par le prix » si le TVP delta agrégé >= {REACT_GATE:.0%}.", "Script Alpaca")
    if DELTA_SOURCE == "carnet":
        ws.cell(1, c_dd).comment = Comment(
            "Delta % pris sur les MÊMES ETF et avec les MÊMES poids que le carnet dynamique (_wc_ : qualité du carnet x stabilité x persistance x fiabilité prix, "
            f"plafonné à {REACT_CAP:.0%} par ETF). Un ETF sans delta valide sort du calcul. Le TVP delta (le Delta % de la bougie j a-t-il annoncé le sens du close à j+{REACT_H} ?) "
            "sert uniquement à valider le delta agrégé, il ne change plus les poids.", "Script Alpaca")
        ws.cell(1, c_dd_top).comment = Comment("Identique à « Carnet le plus fiable » : le delta suit les mêmes ETF et les mêmes poids que le carnet.", "Script Alpaca")
        ws.cell(1, c_dd_conf).comment = Comment("TVP delta moyen des ETF, pondéré par les poids du carnet. 50 % = le delta ne prédit rien.", "Script Alpaca")
    elif DELTA_SOURCE == "top":
        ws.cell(1, c_dd).comment = Comment("Delta % du seul ETF dont le carnet est le plus fiable (celui de « Carnet le plus fiable (poids) »).", "Script Alpaca")
        ws.cell(1, c_dd_top).comment = Comment("Identique à « Carnet le plus fiable » : le delta vient de cet ETF.", "Script Alpaca")
        ws.cell(1, c_dd_conf).comment = Comment("TVP delta de cet ETF seul. 50 % = son delta ne prédit rien.", "Script Alpaca")
    ws.cell(1, c_dc_sc).comment = Comment(
        "Moyenne du delta dynamique et du carnet dynamique, en ne retenant que les composantes VALIDÉES par le prix (TVP >= porte). "
        "Une seule composante validée = elle seule ; aucune = vide.", "Script Alpaca")
    ws.cell(1, c_dc_ver).comment = Comment(
        f"Score >= {DCW_FORT:.0%} fort, >= {DCW_SIMPLE:.0%} simple (achat), <= -{DCW_FORT:.0%} / -{DCW_SIMPLE:.0%} vente. "
        "Précise quelles composantes sont validées. « Neutre / non validé par le prix » si ni le delta ni le carnet n'ont de TVP suffisant.", "Script Alpaca")
    delta_col = [L(OFF + 1 + W * k + 1) for k in range(N)]
    carnet_col = [L(OFF + 1 + W * k + 2) for k in range(N)]
    whale_col = [L(OFF + 1 + W * k + 3) for k in range(N)]
    flux_col = [L(OFF + 1 + W * k + 4) for k in range(N)]
    gap_col = [L(c_gap0 + k) for k in range(N)]
    gap_head = f"${gap_col[0]}$1:${gap_col[-1]}$1"

    # lignes des enchères (bougie contenant 9h30 / 16h00 NY du jour affiché)
    auc_rows = {}
    if has_auc:
        for k in range(len(ref_bars)):
            ny0 = ref_bars.index[k].tz_convert(NY)
            m0 = ny0.hour * 60 + ny0.minute
            if m0 <= 9 * 60 + 30 < m0 + BAR_MINUTES:
                auc_rows["open"] = k + 2
            if m0 <= 16 * 60 < m0 + BAR_MINUTES:
                auc_rows["close"] = k + 2

    c_chr, c_chrat = pos["chn_ref"], pos["chn_ratio"]
    c_tpr, c_tprat = pos["tape_ref"], pos["tape_ratio"]
    _tp_ref = {sy: (bars_by_sym[sy]["tape_ref"].to_numpy(dtype=float) if "tape_ref" in bars_by_sym[sy].columns
                    else np.full(len(bars_by_sym[sy]), np.nan)) for sy in syms}
    _chn_cur, _chn_ref = {}, {}
    for sy in syms:
        b_ = bars_by_sym[sy]
        _chn_cur[sy] = b_["churn"].to_numpy(dtype=float) if "churn" in b_.columns else np.full(len(b_), np.nan)
        _chn_ref[sy] = b_["churn_ref"].to_numpy(dtype=float) if "churn_ref" in b_.columns else np.full(len(b_), np.nan)

    for i in range(2, n + 2):
        ts = ref_bars.index[i - 2] if (i - 2) < len(ref_bars) else None
        ny = ts.tz_convert(NY) if ts is not None else None
        row = {1: (i - 1) if ny is not None else None,
               2: ts.time() if ts is not None else None,
               3: ny.time() if ny is not None else None}
        for k, sym in enumerate(syms):
            for j, (_, col, _) in enumerate(per):
                src = f"'{sym}_{feed}'!{col}{i}"
                row[OFF + 1 + W * k + j] = f'=IF({src}="","",{src})'
        d_refs = [f"{c}{i}" for c in delta_col]
        c_refs = [f"{c}{i}" for c in carnet_col]
        vol_refs = [f"'{sh}'!I{i}" for sh in sheets_etf]           # volume brut de la bougie, par ETF
        avg_f, net_f, whales_f, verdict_f = confluence_formulas(
            [f"{c}{i}" for c in flux_col], [f"{c}{i}" for c in whale_col], f"{L(c_net)}{i}")
        nd_f, nc_f, vdc_f = delta_carnet_formulas(d_refs, c_refs, vol_refs, f"{L(c_nd)}{i}", f"{L(c_nc)}{i}")

        # moyennes pondérées par le volume (au lieu d'une simple moyenne arithmétique)
        num_d = "+".join(f'IF(ISNUMBER({d}),{d}*{v},0)' for d, v in zip(d_refs, vol_refs))
        den_d = "+".join(f'IF(ISNUMBER({d}),{v},0)' for d, v in zip(d_refs, vol_refs))
        row[c_dm] = f'=IF(SUM({den_d})=0,"",SUM({num_d})/SUM({den_d}))'
        num_c = "+".join(f'IF(ISNUMBER({c}),{c}*{v},0)' for c, v in zip(c_refs, vol_refs))
        den_c = "+".join(f'IF(ISNUMBER({c}),{v},0)' for c, v in zip(c_refs, vol_refs))
        row[c_cm] = f'=IF(SUM({den_c})=0,"",SUM({num_c})/SUM({den_c}))'

        # ---- nouvelles colonnes (les existantes ne sont pas modifiées) ----
        # ISO agrégé, pondéré par le volume de chaque ETF
        for cc, colx in ((c_iso_p, "AS"), (c_iso_d, "AT")):
            refs_x = [f"'{sh}'!{colx}{i}" for sh in sheets_etf]
            nx = "+".join(f'IF(ISNUMBER({a}),{a}*{v},0)' for a, v in zip(refs_x, vol_refs))
            dx = "+".join(f'IF(ISNUMBER({a}),{v},0)' for a, v in zip(refs_x, vol_refs))
            row[cc] = f'=IF(SUM({dx})=0,"",SUM({nx})/SUM({dx}))'
        # Delta / Carnet alignés pondérés : valeur ETF = % ÷ ALIGN_SCALE limitée à ±1, × volume, ÷ volume
        sgn_d = "+".join(f'IF(ISNUMBER({d}),MAX(-1,MIN(1,{d}/{ALIGN_SCALE}))*{v},0)' for d, v in zip(d_refs, vol_refs))
        sgn_c = "+".join(f'IF(ISNUMBER({c}),MAX(-1,MIN(1,{c}/{ALIGN_SCALE}))*{v},0)' for c, v in zip(c_refs, vol_refs))
        row[c_ndw] = f'=IF(SUM({den_d})=0,"",SUM({sgn_d})/SUM({den_d}))'
        row[c_ncw] = f'=IF(SUM({den_c})=0,"",SUM({sgn_c})/SUM({den_c}))'
        ndw, ncw = f"{L(c_ndw)}{i}", f"{L(c_ncw)}{i}"
        shw = f"(N({ndw})+N({ncw}))/2"
        row[c_vdcw] = (
            f'=IF(AND({ndw}="",{ncw}=""),"Pas de données",'
            f'IF({shw}>={DCW_FORT},"ACHAT fort : delta + carnet",'
            f'IF({shw}>={DCW_SIMPLE},"Achat : delta + carnet",'
            f'IF({shw}<=-{DCW_FORT},"VENTE forte : delta + carnet",'
            f'IF({shw}<=-{DCW_SIMPLE},"Vente : delta + carnet",'
            f'IF(OR(AND(N({ndw})>={DCW_SIMPLE},N({ncw})<=-{DCW_SIMPLE}),AND(N({ndw})<=-{DCW_SIMPLE},N({ncw})>={DCW_SIMPLE})),'
            f'"Divergence exécution / carnet","Neutre / mixte"))))))')

        # CVD (cumul depuis l'ouverture US 9h30 NY) : vide avant, puis cumul = ligne précédente + apport de la bougie
        prev_cvd_c = f"{L(c_cvd_c)}{i - 1}"
        cm_cell = f"{L(c_cm)}{i}"
        prev_cvd_d = f"{L(c_cvd_d)}{i - 1}"
        num_w = "+".join(f"IF(ISNUMBER({d}),{d}*{v}*ABS({d}),0)" for d, v in zip(d_refs, vol_refs))
        den_w = "+".join(f"IF(ISNUMBER({d}),{v}*ABS({d}),0)" for d, v in zip(d_refs, vol_refs))
        row[c_cvd_d] = (f'=IF(C{i}<TIME(9,30,0),"",IF(ISNUMBER({prev_cvd_d}),{prev_cvd_d},0)'
                        f'+IF(SUM({den_w})=0,0,SUM({num_w})/SUM({den_w})))')
        row[c_cvd_c] = (f'=IF(C{i}<TIME(9,30,0),"",IF(ISNUMBER({prev_cvd_c}),{prev_cvd_c},0)'
                        f'+IF(ISNUMBER({cm_cell}),{cm_cell},0))')
        if c_cvd_sip:
            sref = f"'{sip_cvd[0]}'!{sip_cvd[1]}{i}"
            row[c_cvd_sip] = f'=IF(ISNUMBER({sref}),{sref},"")'

        if with_oi:
            row[c_oi] = oi.get(int(ts.timestamp())) if (oi and ts is not None) else None
        row[c_nd], row[c_nc], row[c_vdc] = nd_f, nc_f, vdc_f
        row[c_avg], row[c_net], row[c_wh], row[c_ver] = avg_f, net_f, whales_f, verdict_f

        # moyenne des ETF pour AF -> AX (même ligne = même bougie dans toutes les feuilles ETF)
        def mk(sc, i=i):
            return f"{L(newcol[sc])}{i}"
        for sc in SRC:
            if sc in TEXT_SRC:
                continue
            if sc == 32:    # BR : ratio agrégé = Σ volumes / Σ références J-1 (ETF avec ratio valide), pas une moyenne de ratios
                bcol = L(REF_COL)                              # colonne BA des feuilles ETF (référence J-1)
                afs = [f"'{sh}'!AF{i}" for sh in sheets_etf]
                bas = [f"'{sh}'!{bcol}{i}" for sh in sheets_etf]
                nn = "+".join(f"IF(ISNUMBER({a}),{v},0)" for a, v in zip(afs, vol_refs))
                dd = "+".join(f"IF(ISNUMBER({a}),{b},0)" for a, b in zip(afs, bas))
                row[newcol[32]] = f'=IF(SUM({dd})=0,"",SUM({nn})/SUM({dd}))'
                continue
            refs = ",".join(f"'{sh}'!{L(sc)}{i}" for sh in sheets_etf)
            if sc == 51:    # score flux global : même arrondi que « Score flux moyen »
                row[newcol[sc]] = f'=IF(COUNT({refs})=0,"",ROUND(AVERAGE({refs}),0))'
            else:
                row[newcol[sc]] = f'=IF(COUNT({refs})=0,"",AVERAGE({refs}))'
        # AG : sens du plus gros trade = vote majoritaire des ETF
        nb_a = "+".join(f"('{sh}'!AG{i}=\"ACHAT\")" for sh in sheets_etf)
        nb_v = "+".join(f"('{sh}'!AG{i}=\"VENTE\")" for sh in sheets_etf)
        row[newcol[33]] = (f'=IF(({nb_a})+({nb_v})=0,"",IF(({nb_a})>({nb_v}),"ACHAT",'
                           f'IF(({nb_v})>({nb_a}),"VENTE","MIXTE")))')
        # AH : pression = même règle que la colonne AH, appliquée au delta % moyen et au carnet % moyen
        dm, cm = f"{L(c_dm)}{i}", f"{L(c_cm)}{i}"
        sd = f"IF({dm}>={SEUIL_DELTA},1,IF({dm}<=-{SEUIL_DELTA},-1,0))"
        sk = f"IF({cm}=\"\",0,IF({cm}>={SEUIL_CARNET},1,IF({cm}<=-{SEUIL_CARNET},-1,0)))"
        row[newcol[34]] = (f'=IF({dm}="","",IF({sd}+{sk}>0,"Pression achat",'
                           f'IF({sd}+{sk}<0,"Pression vente","Neutre")))')
        # AP / AQ : mêmes règles que les colonnes AP / AQ, appliquées aux valeurs moyennes
        volN = "SUM(" + ",".join(f"'{sh}'!I{i}" for sh in sheets_etf) + ")"
        ao, ar = mk(41), mk(44)
        row[newcol[42]] = (f'=IF({volN}=0,"Pas d’activité",IF({ao}="","Historique insuffisant",IF({ao}>={WHALE_FORT},"Forte",'
                           f'IF({ao}>={WHALE_ALERTE},"Modérée",IF({ao}>=25,"Faible","Non détectée")))))')
        row[newcol[43]] = (f'=IF({ao}="","Non évalué",IF({ao}<{WHALE_ALERTE},"Non confirmé",'
                           f'IF(OR({ar}="",{dm}=""),"Non confirmé",IF(AND({ar}>0,{dm}>={SEUIL_DELTA}),"Accumulation probable",'
                           f'IF(AND({ar}<0,{dm}<=-{SEUIL_DELTA}),"Distribution probable","Mixte / à confirmer")))))')
        # AX : même règle que la colonne AX, appliquée au score flux moyen, au volume vs 12 bougies moyen et au ratio J-1 moyen
        aw, am, af = mk(51), mk(39), mk(32)
        row[newcol[52]] = (f'=IF({aw}="","Pas de données",IF(AND({aw}>=40,OR(IF({am}="",0,{am})>=1.5,IF({af}="",0,{af})>={SIGNAL_RATIO_J1})),"Accumulation forte",'
                           f'IF(AND({aw}<=-40,OR(IF({am}="",0,{am})>=1.5,IF({af}="",0,{af})>={SIGNAL_RATIO_J1})),"Distribution forte",'
                           f'IF({aw}>=20,"Biais acheteur",IF({aw}<=-20,"Biais vendeur","Neutre / mixte")))))')

        # ---- v2 : divergence (via les colonnes helper _gap_<ETF>, une par ETF) ----
        cons = f"{L(c_dm)}{i}"                                  # consensus = Delta % moyen pondéré
        for k in range(N):
            d = d_refs[k]
            row[c_gap0 + k] = (
                f'=IFERROR(IF(AND(ISNUMBER({d}),ISNUMBER({cons}),ABS({cons})>={DIV_CONSENSUS_MIN},'
                f'{d}*{cons}<0,ABS({d}-{cons})>={DIV_GAP}),ABS({d}-{cons}),0),0)')
        hrng = f"${gap_col[0]}{i}:${gap_col[-1]}{i}"
        hprev = f"${gap_col[0]}{i - 1}:${gap_col[-1]}{i - 1}"
        mx = f"MAX({hrng})"
        row[c_div_etf] = (f'=IF({mx}<=0,"",MID(INDEX({gap_head},MATCH({mx},{hrng},0)),6,12)'
                          f'&" ("&TEXT(-SIGN({cons})*{mx}*100,"+0;-0")&" pts)")')
        row[c_div_gap] = f'=IF({mx}<=0,0,-SIGN({cons})*{mx})'
        row[c_div_ver] = (f'=IF({mx}<=0,"Pas de divergence",'
                          f'IF(MAX({hprev})>0,"Divergence persistante","Divergence récente"))')

        # ---- v2 : carnet qualité, absorption, confirmation ----
        bb_refs = [f"'{sh}'!BB{i}" for sh in sheets_etf]
        qn = "+".join(f"IF(ISNUMBER({b}),{b}*{v},0)" for b, v in zip(bb_refs, vol_refs))
        qd = "+".join(f"IF(ISNUMBER({b}),{v},0)" for b, v in zip(bb_refs, vol_refs))
        row[c_qm] = f'=IF(SUM({qd})=0,"",SUM({qn})/SUM({qd}))'
        qcn = "+".join(f"IF(AND(ISNUMBER({c}),ISNUMBER({b})),{c}*{b},0)" for c, b in zip(c_refs, bb_refs))
        qcd = "+".join(f"IF(AND(ISNUMBER({c}),ISNUMBER({b})),{b},0)" for c, b in zip(c_refs, bb_refs))
        row[c_qc] = f'=IF(SUM({qcd})=0,"",SUM({qcn})/SUM({qcd}))'
        def wavg(col, rnd=None, i=i):
            """Moyenne pondérée par le volume brut de la bougie de chaque ETF ; ETF sans valeur = poids nul."""
            refs = [f"'{sh}'!{col}{i}" for sh in sheets_etf]
            nu = "+".join(f"IF(ISNUMBER({r}),{r}*{v},0)" for r, v in zip(refs, vol_refs))
            de = "+".join(f"IF(ISNUMBER({r}),{v},0)" for r, v in zip(refs, vol_refs))
            core = f"SUM({nu})/SUM({de})"
            if rnd is not None:
                core = f"ROUND({core},{rnd})"
            return f'=IF(SUM({de})=0,"",{core})'
        row[c_abs] = wavg("BG", 0)          # absorption pondérée volume (remplace la moyenne simple)
        row[c_conf] = wavg("BH")            # confirmation pondérée volume (remplace la moyenne simple)
        row[c_chn] = wavg("BC")
        # ratio churn vs J-1 : Σ churn des ETF / Σ réf. J-1, sur les ETF où les deux valeurs existent (même principe que le ratio volume)
        _k = i - 2
        _use = [(sh, sy) for sh, sy in zip(sheets_etf, syms)
                if _k < len(_chn_cur[sy]) and _k < len(_chn_ref[sy])
                and np.isfinite(_chn_cur[sy][_k]) and np.isfinite(_chn_ref[sy][_k]) and _chn_ref[sy][_k] > 0]
        if _use:
            row[c_chr] = float(sum(_chn_ref[sy][_k] for _, sy in _use))
            _num = "+".join(f"'{sh}'!BC{i}" for sh, _ in _use)
            row[c_chrat] = f'=IF(N({L(c_chr)}{i})=0,"",({_num})/{L(c_chr)}{i})'
        # v9 : speed of tape agrégé = Σ vitesses (CL) ÷ Σ réf. (CM), sur les ETF où la réf. historique existe
        _use_t = [sh for sh, sy in zip(sheets_etf, syms) if _k < len(_tp_ref[sy]) and np.isfinite(_tp_ref[sy][_k]) and _tp_ref[sy][_k] > 0]
        if _use_t:
            row[c_tpr] = float(sum(_tp_ref[sy][_k] for sh, sy in zip(sheets_etf, syms) if sh in _use_t))
            _tnum = "+".join(f"'{sh}'!{L(C_TAPE)}{i}" for sh in _use_t)
            row[c_tprat] = f'=IF(N({L(c_tpr)}{i})=0,"",({_tnum})/{L(c_tpr)}{i})'
        row[c_rot] = wavg("BD")
        row[c_imb] = wavg("BE")
        row[c_imp] = wavg("BF")
        # v2.2 : qualité composite (poids par ETF dans les colonnes masquées _w_<ETF>)
        a0 = max(2, i - BOOK_QUALITY_WINDOW)
        for k, sh in enumerate(sheets_etf):
            bd, bb, bh, be = (f"'{sh}'!BD{i}", f"'{sh}'!BB{i}", f"'{sh}'!BH{i}", f"'{sh}'!BE{i}")
            imb = f'IF(ISNUMBER({be}),{be},IF(ISNUMBER({c_refs[k]}),{c_refs[k]},""))'
            row[c_imb0 + k] = "=" + imb
            if i > 2:
                rng = f"'{sh}'!BD{a0}:BD{i - 1}"
                stab = f"IF(AND(ISNUMBER({bd}),COUNT({rng})>={BOOK_MIN_HISTORY}),MIN(1,MEDIAN({rng})/{bd}),1)"
            else:
                stab = "1"
            track = (f"IF(ISNUMBER({bh}),MIN(1,MAX({TRACK_FLOOR},({bh}-{VERDICT_PEU_FIABLE})/({VERDICT_FIABLE}-{VERDICT_PEU_FIABLE}))),{TRACK_NEUTRAL})")
            row[c_w0 + k] = f"=IF(AND(ISNUMBER({L(c_imb0 + k)}{i}),ISNUMBER({bb})),{bb}*{stab}*{track},0)"
        wr = f"{L(c_w0)}{i}:{L(c_w0 + N - 1)}{i}"
        ir = f"{L(c_imb0)}{i}:{L(c_imb0 + N - 1)}{i}"
        row[c_qcc] = f'=IF(SUM({wr})=0,"",SUMPRODUCT({wr},{ir})/SUM({wr}))'
        # ---- v5 : carnet DYNAMIQUE validé par le PRIX ----
        # _sg_ : signal carnet de l'ETF = BN (déséquilibre x persistance), à défaut BE, à défaut Carnet %.
        # _rx_ (ligne i) : jugement du signal de la bougie j = i - H : 1 si le close a bougé d'au moins THR_AMP x amplitude moyenne
        #                  DANS le sens du carnet (j -> i), 0 s'il a bougé en sens CONTRAIRE (absorption comprise), vide si le prix n'a pas bougé
        #                  ou si le carnet ne parlait pas (|signal| < TRIGGER). Connu à la clôture de la bougie i : aucune information future.
        # _p_  : TVP lissé = (réussites + 0,5 x PRIOR) / (cas + PRIOR) sur la fenêtre ; 50 % tant que cas < MIN_CASES.
        # _wd_ : poids = qualité BB x stabilité des tailles x persistance x fiabilité_prix^GAMMA (SANS volume exécuté) ; 0 si ETF inactif.
        ra = max(2, i - REACT_WINDOW + 1)
        for k, sh in enumerate(sheets_etf):
            imb_k = f"{L(c_imb0 + k)}{i}"
            bn = f"'{sh}'!{L(C_IMBP)}{i}"
            row[c_sg0 + k] = f'=IF(ISNUMBER({bn}),{bn},IF(ISNUMBER({imb_k}),{imb_k},""))'
            j = i - REACT_H
            if j >= 2:
                sg_j = f"{L(c_sg0 + k)}{j}"
                ca, cb = f"'{sh}'!H{i}", f"'{sh}'!H{j}"
                thr = f"{REACT_THR_AMP}*AVERAGE('{sh}'!AW{ra}:AW{i})"
                row[c_rx0 + k] = (f'=IFERROR(IF(AND(ISNUMBER({sg_j}),ISNUMBER({ca}),ISNUMBER({cb})),'
                                  f'IF(AND(ABS({sg_j})>={REACT_TRIGGER},{cb}>0),'
                                  f'IF(ABS({ca}/{cb}-1)>={thr},IF(SIGN({sg_j})=SIGN({ca}-{cb}),1,0),""),""),""),"")')
            rrng = f"{L(c_rx0 + k)}{ra}:{L(c_rx0 + k)}{i}"
            p_k = f"{L(c_rp0 + k)}{i}"
            row[c_rp0 + k] = (f"=IF(COUNT({rrng})<{REACT_MIN_CASES},0.5,"
                              f"(COUNTIF({rrng},1)+0.5*{REACT_PRIOR})/(COUNT({rrng})+{REACT_PRIOR}))")
            rel = f"({REACT_FLOOR}+{1 - REACT_FLOOR}*MAX(0,MIN(1,({p_k}-0.5)/{REACT_FULL - 0.5})))"
            bb_, bd_ = f"'{sh}'!BB{i}", f"'{sh}'!BD{i}"
            pe_ = f"'{sh}'!{L(C_PERS)}{i}"
            if i > 2:
                srng = f"'{sh}'!BD{max(2, i - BOOK_QUALITY_WINDOW)}:BD{i - 1}"
                stab = f"IF(AND(ISNUMBER({bd_}),COUNT({srng})>={BOOK_MIN_HISTORY}),MIN(1,MEDIAN({srng})/{bd_}),1)"
            else:
                stab = "1"
            qual = f"IF(ISNUMBER({bb_}),{bb_},0.5)*{stab}*IF(ISNUMBER({pe_}),{pe_},0.5)"
            row[c_wd0 + k] = (f"=IF(AND(ISNUMBER({L(c_sg0 + k)}{i}),SUM('{sh}'!J{ra}:J{i})>={REACT_MIN_TRADES}),"
                              f"{qual}*{rel}^{REACT_GAMMA},0)")
        wd_refs = [f"{L(c_wd0 + k)}{i}" for k in range(N)]
        sg_refs = [f"{L(c_sg0 + k)}{i}" for k in range(N)]
        rp_refs = [f"{L(c_rp0 + k)}{i}" for k in range(N)]
        wd_rng = f"{L(c_wd0)}{i}:{L(c_wd0 + N - 1)}{i}"
        wd_head = f"${L(c_wd0)}$1:${L(c_wd0 + N - 1)}$1"
        # plafond par ETF itéré 3 fois (_cap1.._cap3) : le plafond est calculé sur la somme DÉJÀ plafonnée -> part réelle <= REACT_CAP
        prev_s = f"SUM({wd_rng})"
        for q in range(3):
            row[c_cap[q]] = "=" + "+".join(f"MIN({w_},{REACT_CAP}*{prev_s})" for w_ in wd_refs)
            prev_s = f"{L(c_cap[q])}{i}"
        cap = f"{REACT_CAP}*{prev_s}"
        wc = [f"MIN({w_},{cap})" for w_ in wd_refs]
        swc = "+".join(wc)
        row[c_dyn] = (f'=IF(SUM({wd_rng})=0,"",(' + "+".join(f"IF(ISNUMBER({g_}),{g_}*{w_},0)" for g_, w_ in zip(sg_refs, wc))
                      + f")/({swc}))")
        row[c_dyn_conf] = (f'=IF(SUM({wd_rng})=0,"",(' + "+".join(f"{p_}*{w_}" for p_, w_ in zip(rp_refs, wc)) + f")/({swc}))")
        row[c_dyn_top] = (f'=IF(SUM({wd_rng})=0,"",MID(INDEX({wd_head},MATCH(MAX({wd_rng}),{wd_rng},0)),5,12)'
                          f'&" ("&TEXT(MIN(MAX({wd_rng}),{cap})/({swc}),"0%")&")")')
        dyn, conf = f"{L(c_dyn)}{i}", f"{L(c_dyn_conf)}{i}"
        row[c_dyn_gate] = (f'=IF({dyn}="","Pas de données",IF({conf}>={REACT_GATE},"Carnet validé par le prix",'
                           f'"Carnet non fiable – attendre"))')
        row[c_dyn_ver] = (f'=IF({dyn}="","Pas de données",IF({conf}<{REACT_GATE},"Neutre / carnet non fiable",'
                          f'IF({dyn}>={REACT_VERDICT},"Carnet ACHAT (validé prix)",IF({dyn}<=-{REACT_VERDICT},'
                          f'"Carnet VENTE (validé prix)","Neutre / mixte"))))')
        decay = 0.5 ** (BAR_MINUTES / REACT_HALF_LIFE_MIN)
        prev_cvd_dyn = f"{L(c_dyn_cvd)}{i - 1}"
        row[c_dyn_cvd] = (f'=IF(C{i}<TIME(9,30,0),"",{decay:.6f}*IF(ISNUMBER({prev_cvd_dyn}),{prev_cvd_dyn},0)'
                          f'+IF(ISNUMBER({dyn}),{dyn},0))')
        # ---- v12 : CVD carnet prix (le CVD carnet ci-dessus n'est pas modifié) ----
        _wcr = [f"{L(c_wc0 + k)}{i}" for k in range(N)]
        row[c_px_ret], row[c_px_amp], row[c_px] = _px_formulas(
            i, sheets_etf, _wcr, ra, dyn, f"{L(c_px)}{i - 1}", L(c_px_ret), L(c_px_amp))
        # ---- v6 : delta ISO certain agrégé (pondéré volume) et part ISO classée certaine (pondérée par le volume ISO) ----
        row[c_iso_dc] = wavg(L(C_ISODC))
        _n = "+".join(f"N('{sh}'!{L(C_ISOBC)}{i})+N('{sh}'!{L(C_ISOSC)}{i})" for sh in sheets_etf)
        _d = "+".join(f"N('{sh}'!AS{i})*N('{sh}'!I{i})" for sh in sheets_etf)
        row[c_iso_pc] = f'=IF(({_d})=0,"",({_n})/({_d}))'
        # ---- v5 ter : poids FINAL du carnet dynamique (_wc_ = poids TVP plafonné) et indicateurs de déséquilibre pondérés par ces poids ----
        # Remplacent (à leur place, en tête de feuille) « poids volume x fiabilité » et les classes de taille pondérées volume :
        # même principe / même logique (BE, puis BE restreint à une classe de taille), mais chaque ETF pèse son poids dynamique TVP.
        for k in range(N):
            row[c_wc0 + k] = f"=MIN({wd_refs[k]},{cap})"
        wcr = [f"{L(c_wc0 + k)}{i}" for k in range(N)]

        def wdyn(col_letter):
            refs = [f"'{sh}'!{col_letter}{i}" for sh in sheets_etf]
            nu = "+".join(f"IF(ISNUMBER({r_}),{r_}*{w_},0)" for r_, w_ in zip(refs, wcr))
            de = "+".join(f"IF(ISNUMBER({r_}),{w_},0)" for r_, w_ in zip(refs, wcr))
            return f'=IF(({de})=0,"",({nu})/({de}))'
        row[c_imb_dyn] = wdyn("BE")
        for j in range(4):      # v10 : delta exécuté par taille de trade = Σ net ÷ Σ volume classé (les gros ETF pèsent naturellement plus)
            _tn = "+".join(f"N('{sh}'!{L(C_TCN1 + j)}{i})" for sh in sheets_etf)
            _tv = "+".join(f"N('{sh}'!{L(C_TCV1 + j)}{i})" for sh in sheets_etf)
            row[c_tx[j]] = f'=IF(({_tv})=0,"",({_tn})/({_tv}))'
        for j in range(4):
            row[c_cls_dyn[j]] = wdyn(L(C_CLS1 + j))
        # ---- v5 bis : DELTA dynamique validé par le prix (même exercice que le carnet) ----
        # _drx_ : le Delta % de la bougie j = i - H (|Delta %| >= SEUIL_DELTA) a-t-il annoncé le sens du close à i ? 1 / 0 / vide.
        # _dp_  : TVP delta lissé (même lissage que le carnet). _dwd_ : poids = racine(volume fenêtre) x fiabilité^GAMMA, 0 si pas de volume.
        for k, sh in enumerate(sheets_etf):
            j = i - REACT_H
            if j >= 2:
                dg_j = d_refs[k].rstrip("0123456789") + str(j)
                ca, cb = f"'{sh}'!H{i}", f"'{sh}'!H{j}"
                thr = f"{REACT_THR_AMP}*AVERAGE('{sh}'!AW{ra}:AW{i})"
                row[c_drx0 + k] = (f'=IFERROR(IF(AND(ISNUMBER({dg_j}),ISNUMBER({ca}),ISNUMBER({cb})),'
                                   f'IF(AND(ABS({dg_j})>={SEUIL_DELTA},{cb}>0),'
                                   f'IF(ABS({ca}/{cb}-1)>={thr},IF(SIGN({dg_j})=SIGN({ca}-{cb}),1,0),""),""),""),"")')
            drng = f"{L(c_drx0 + k)}{ra}:{L(c_drx0 + k)}{i}"
            dp_k = f"{L(c_drp0 + k)}{i}"
            row[c_drp0 + k] = (f"=IF(COUNT({drng})<{REACT_MIN_CASES},0.5,"
                               f"(COUNTIF({drng},1)+0.5*{REACT_PRIOR})/(COUNT({drng})+{REACT_PRIOR}))")
            drel = f"({REACT_FLOOR}+{1 - REACT_FLOOR}*MAX(0,MIN(1,({dp_k}-0.5)/{REACT_FULL - 0.5})))"
            row[c_dwd0 + k] = (f"=IF(AND(ISNUMBER({d_refs[k]}),SUM('{sh}'!I{ra}:I{i})>0),"
                               f"SQRT(SUM('{sh}'!I{ra}:I{i}))*{drel}^{REACT_GAMMA},0)")
        dw_refs = [f"{L(c_dwd0 + k)}{i}" for k in range(N)]
        dp_refs = [f"{L(c_drp0 + k)}{i}" for k in range(N)]
        dw_rng = f"{L(c_dwd0)}{i}:{L(c_dwd0 + N - 1)}{i}"
        dw_head = f"${L(c_dwd0)}$1:${L(c_dwd0 + N - 1)}$1"
        prev_s = f"SUM({dw_rng})"
        for q in range(3):
            row[c_dcap[q]] = "=" + "+".join(f"MIN({w_},{REACT_CAP}*{prev_s})" for w_ in dw_refs)
            prev_s = f"{L(c_dcap[q])}{i}"
        dcap = f"{REACT_CAP}*{prev_s}"
        dwc = [f"MIN({w_},{dcap})" for w_ in dw_refs]
        dswc = "+".join(dwc)
        if DELTA_SOURCE == "own":
            row[c_dd] = (f'=IF(SUM({dw_rng})=0,"",(' + "+".join(f"IF(ISNUMBER({d_}),{d_}*{w_},0)" for d_, w_ in zip(d_refs, dwc))
                         + f")/({dswc}))")
            row[c_dd_conf] = (f'=IF(SUM({dw_rng})=0,"",(' + "+".join(f"{p_}*{w_}" for p_, w_ in zip(dp_refs, dwc)) + f")/({dswc}))")
            row[c_dd_top] = (f'=IF(SUM({dw_rng})=0,"",MID(INDEX({dw_head},MATCH(MAX({dw_rng}),{dw_rng},0)),6,12)'
                             f'&" ("&TEXT(MIN(MAX({dw_rng}),{dcap})/({dswc}),"0%")&")")')
        elif DELTA_SOURCE == "top":
            # delta (et TVP delta) du seul ETF dont le carnet est le plus fiable : même ETF que « Carnet le plus fiable (poids) »
            _m = f"MATCH(MAX({wd_rng}),{wd_rng},0)"
            row[c_dd] = f'=IF(SUM({wd_rng})=0,"",IFERROR(1*CHOOSE({_m},' + ",".join(d_refs) + '),""))'
            row[c_dd_conf] = f'=IF(SUM({wd_rng})=0,"",IFERROR(1*CHOOSE({_m},' + ",".join(dp_refs) + '),""))'
            row[c_dd_top] = row[c_dyn_top]
        else:
            # "carnet" : mêmes ETF, mêmes poids (_wc_) que le carnet dynamique ; un ETF sans delta valide sort du dénominateur
            _num = "+".join(f"IF(ISNUMBER({d_}),{d_}*{w_},0)" for d_, w_ in zip(d_refs, wc))
            _den = "+".join(f"IF(ISNUMBER({d_}),{w_},0)" for d_, w_ in zip(d_refs, wc))
            _pn = "+".join(f"IF(ISNUMBER({d_}),{p_}*{w_},0)" for d_, p_, w_ in zip(d_refs, dp_refs, wc))
            row[c_dd] = f'=IF(({_den})=0,"",({_num})/({_den}))'
            row[c_dd_conf] = f'=IF(({_den})=0,"",({_pn})/({_den}))'
            row[c_dd_top] = row[c_dyn_top]
        ddv, ddc = f"{L(c_dd)}{i}", f"{L(c_dd_conf)}{i}"
        row[c_dd_gate] = (f'=IF({ddv}="","Pas de données",IF({ddc}>={REACT_GATE},"Delta validé par le prix",'
                          f'"Delta non fiable – attendre"))')
        prev_dcvd = f"{L(c_dd_cvd)}{i - 1}"
        row[c_dd_cvd] = (f'=IF(C{i}<TIME(9,30,0),"",{decay:.6f}*IF(ISNUMBER({prev_dcvd}),{prev_dcvd},0)'
                         f'+IF(ISNUMBER({ddv}),{ddv},0))')
        # delta de l'ETF dont le carnet est le plus fiable (celui de « Carnet le plus fiable (poids) »)
        row[c_dtop_val] = (f'=IF(SUM({wd_rng})=0,"",IFERROR(1*CHOOSE(MATCH(MAX({wd_rng}),{wd_rng},0),'
                           + ",".join(d_refs) + '),""))')
        # confluence DELTA + CARNET validée par le prix : seules les composantes validées (TVP >= porte) entrent dans le score
        gd, gc = f'AND(ISNUMBER({ddc}),{ddc}>={REACT_GATE})', f'AND(ISNUMBER({conf}),{conf}>={REACT_GATE})'
        sc = f"{L(c_dc_sc)}{i}"
        row[c_dc_sc] = (f'=IF(AND({gd},{gc}),(N({ddv})+N({dyn}))/2,IF({gd},N({ddv}),IF({gc},N({dyn}),"")))')
        src_ = f'IF(AND({gd},{gc}),"delta + carnet validés",IF({gd},"delta seul validé","carnet seul validé"))'
        row[c_dc_ver] = (f'=IF({sc}="","Neutre / non validé par le prix",'
                         f'IF({sc}>={DCW_FORT},"ACHAT fort ("&{src_}&")",'
                         f'IF({sc}>={DCW_SIMPLE},"Achat ("&{src_}&")",'
                         f'IF({sc}<=-{DCW_FORT},"VENTE forte ("&{src_}&")",'
                         f'IF({sc}<=-{DCW_SIMPLE},"Vente ("&{src_}&")","Neutre / mixte")))))')
        # ---- v3 : agrégats des colonnes BJ -> CD des feuilles ETF (M n'est pas modifiée) ----
        row[c_imbp] = wavg(L(C_IMBP))                     # déséquilibre temps x persistance, pondéré volume
        row[c_imbg] = wavg(L(C_IMBGAP))                   # écart persistance - BE, pondéré volume
        row[c_pers] = wavg(L(C_PERS))
        # delta murs : moyenne pondérée volume de (mur bid − mur ask) ; ETF sans les deux valeurs = poids nul
        wb_r = [f"'{sh}'!{L(C_WALLB)}{i}" for sh in sheets_etf]
        wa_r = [f"'{sh}'!{L(C_WALLA)}{i}" for sh in sheets_etf]
        wn = "+".join(f"IF(AND(ISNUMBER({b_}),ISNUMBER({a_})),({b_}-{a_})*{v},0)" for b_, a_, v in zip(wb_r, wa_r, vol_refs))
        wd = "+".join(f"IF(AND(ISNUMBER({b_}),ISNUMBER({a_})),{v},0)" for b_, a_, v in zip(wb_r, wa_r, vol_refs))
        row[c_walld] = f'=IF(SUM({wd})=0,"",SUM({wn})/SUM({wd}))'
        for j in range(4):
            row[c_cls[j]] = wavg(L(C_CLS1 + j))
        # v8 : F = z-score agrégé. Chaque ETF : z = (BE − réf.) ÷ dispersion historique (colonne CK des feuilles ETF), pondéré par le poids
        # du carnet dynamique (_wc_ : qualité x stabilité x fiabilité prix, plafonné) au lieu du volume (IBIT = ~81 % du volume IEX).
        # Si aucun ETF n'a de poids (< REACT_MIN_TRADES trades), repli sur la moyenne simple des z. Le tout x confiance (séances / HIST_CONF_FULL).
        zr = [f"'{sh}'!{L(C_J1Z)}{i}" for sh in sheets_etf]
        z_num = "+".join(f"IF(ISNUMBER({r_}),{r_}*{w_},0)" for r_, w_ in zip(zr, wcr))
        z_den = "+".join(f"IF(ISNUMBER({r_}),{w_},0)" for r_, w_ in zip(zr, wcr))
        z_sum = "+".join(f"IF(ISNUMBER({r_}),{r_},0)" for r_ in zr)
        z_cnt = "+".join(f"IF(ISNUMBER({r_}),1,0)" for r_ in zr)
        row[c_j1] = (f'=IF(({z_cnt})=0,"",{_z_conf:.4f}*IF(({z_den})>0,({z_num})/({z_den}),({z_sum})/({z_cnt})))')
        row[c_imb_raw] = wavg(L(C_J1GAP))                # ancienne colonne F (écart brut en points de %, pondéré volume), conservée
        _fz, _fp = f"{L(c_j1)}{i}", f"{L(c_imb_ema)}{i - 1}"
        if i == 2:
            row[c_imb_ema] = f'=IF({_fz}="","",{_fz})'
        else:
            row[c_imb_ema] = (f'=IF({_fz}="",IF({_fp}="","",{_fp}),IF({_fp}="",{_fz},{Z_EMA_ALPHA}*{_fz}+{1 - Z_EMA_ALPHA}*{_fp}))')
            _e, _q = f"{L(c_imb_ema)}{i}", f"{L(c_imb_ema)}{i - 1}"
            row[c_imb_ver] = (f'=IF(OR({_e}="",{_q}=""),"",'
                              f'IF(AND({_e}>={Z_FORT},{_q}>={Z_FORT}),"Carnet très ACHETEUR vs historique",'
                              f'IF(AND({_e}>={Z_ANOMALIE},{_q}>={Z_ANOMALIE}),"Carnet anormalement ACHETEUR",'
                              f'IF(AND({_e}<=-{Z_FORT},{_q}<=-{Z_FORT}),"Carnet très VENDEUR vs historique",'
                              f'IF(AND({_e}<=-{Z_ANOMALIE},{_q}<=-{Z_ANOMALIE}),"Carnet anormalement VENDEUR","")))))')
        # M bis : moyenne des BE pondérée par volume x fiabilité (_w_ = qualité BB x stabilité BD x track record BH)
        be_refs = [f"'{sh}'!BE{i}" for sh in sheets_etf]
        w_refs = [f"{L(c_w0 + k)}{i}" for k in range(N)]
        hn = "+".join(f"IF(AND(ISNUMBER({b_}),ISNUMBER({w_})),{b_}*{v}*{w_},0)" for b_, w_, v in zip(be_refs, w_refs, vol_refs))
        hd = "+".join(f"IF(AND(ISNUMBER({b_}),ISNUMBER({w_})),{v}*{w_},0)" for b_, w_, v in zip(be_refs, w_refs, vol_refs))
        row[c_hyb] = f'=IF(SUM({hd})=0,"",SUM({hn})/SUM({hd}))'
        # écart IEX vs SIP sur M (en valeur absolue), uniquement dans la feuille IEX quand le SIP est disponible
        if c_imbs and sip_imb:
            sref = f"'{sip_imb[0]}'!{sip_imb[1]}{i}"
            m_ = f"{L(c_imb)}{i}"
            row[c_imbs] = f'=IF(AND(ISNUMBER({m_}),ISNUMBER({sref})),ABS({m_}-{sref}),"")'
        # consommation des retraits de taille : Σ consommé ÷ Σ retiré sur les ETF (et non une moyenne de pourcentages)
        for c_dst, c_rm_, c_con_ in ((c_conb, C_RMB, C_CONB), (c_cona, C_RMA, C_CONA)):
            rms = ",".join(f"'{sh}'!{L(c_rm_)}{i}" for sh in sheets_etf)
            cns = ",".join(f"'{sh}'!{L(c_con_)}{i}" for sh in sheets_etf)
            row[c_dst] = f'=IF(N(SUM({rms}))<=0,"",MIN(1,SUM({cns})/SUM({rms})))'
        row[c_cons] = f'=IF(OR({L(c_conb)}{i}="",{L(c_cona)}{i}=""),"",{L(c_conb)}{i}-{L(c_cona)}{i})'
        # alerte piège : nombre d'ETF en alerte (le delta de la bougie suivante est évalué ETF par ETF dans CB)
        trp = [f"'{sh}'!{L(C_TRAP)}{i}" for sh in sheets_etf]
        n_sw = "+".join(f'({t_}="Piège retiré → sweep")' for t_ in trp)
        n_ns = "+".join(f'({t_}="Retrait de mur – sans sweep")' for t_ in trp)
        n_wt = "+".join(f'({t_}="Retrait de mur – en attente")' for t_ in trp)
        row[c_trap] = (f'=IF(({n_sw})>0,"Piège retiré → sweep ("&({n_sw})&" ETF)",'
                       f'IF(({n_wt})>0,"Retrait de mur – en attente ("&({n_wt})&" ETF)",'
                       f'IF(({n_ns})>0,"Retrait de mur – sans sweep ("&({n_ns})&" ETF)","")))')
        qm_, cf_ = f"{L(c_qm)}{i}", f"{L(c_conf)}{i}"
        row[c_vdq] = (f'=IF(OR({qm_}="",{cf_}=""),"Historique insuffisant",'
                      f'IF(AND({qm_}>={VERDICT_FIABLE},{cf_}>={VERDICT_FIABLE}),"Fiable",'
                      f'IF(OR({qm_}<{VERDICT_PEU_FIABLE},{cf_}<{VERDICT_PEU_FIABLE}),"Peu fiable","Moyen")))')

        # ---- v2 : enchères (valeurs fixes sur la bougie qui contient 9h30 / 16h00) ----
        if has_auc:
            if auc_rows.get("open") == i and auc_open > 0:
                row[c_auc_o] = auc_open
            if auc_rows.get("close") == i and auc_close > 0:
                row[c_auc_c] = auc_close
                tots = ",".join(f"'{sh}'!I{len(bars_by_sym[sy]) + 3}" for sh, sy in zip(sheets_etf, syms))
                row[c_auc_r] = f'=IF(SUM({tots})=0,"",{L(c_auc_c)}{i}/SUM({tots}))'

        for c, v in row.items():
            cell = ws.cell(i, c, v)
            cell.number_format = heads[c - 1][3]
            cell.font, cell.border = F, BORDER
            cell.alignment = Alignment(horizontal="center" if (c <= 3 or heads[c - 1][3] == "General") else "right")

    last = n + 1
    signed = [f"{L(c_dm)}2:{L(c_dm)}{last}", f"{L(c_iso_d)}2:{L(c_iso_d)}{last}",
              f"{L(c_ndw)}2:{L(c_ndw)}{last}", f"{L(c_ncw)}2:{L(c_ncw)}{last}",
              f"{L(c_cm)}2:{L(c_cm)}{last}", f"{L(c_cvd_c)}2:{L(c_cvd_c)}{last}",
              f"{L(c_cvd_d)}2:{L(c_cvd_d)}{last}",
              f"{L(c_nd)}2:{L(c_nd)}{last}", f"{L(c_nc)}2:{L(c_nc)}{last}", f"{L(c_net)}2:{L(c_net)}{last}",
              f"{L(c_div_gap)}2:{L(c_div_gap)}{last}", f"{L(c_qc)}2:{L(c_qc)}{last}",
              f"{L(c_dyn)}2:{L(c_dyn)}{last}", f"{L(c_dyn_cvd)}2:{L(c_dyn_cvd)}{last}", f"{L(c_px)}2:{L(c_px)}{last}",
              f"{L(c_dd)}2:{L(c_dd)}{last}", f"{L(c_dd_cvd)}2:{L(c_dd_cvd)}{last}",
              f"{L(c_dtop_val)}2:{L(c_dtop_val)}{last}", f"{L(c_dc_sc)}2:{L(c_dc_sc)}{last}",
              f"{L(c_iso_dc)}2:{L(c_iso_dc)}{last}"]
    if with_oi:
        signed.append(f"{L(c_oi)}2:{L(c_oi)}{last}")
    if c_cvd_sip:
        signed.append(f"{L(c_cvd_sip)}2:{L(c_cvd_sip)}{last}")
    signed += [f"{c}2:{c}{last}" for c in delta_col + carnet_col + flux_col]
    for rng in signed:
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN_FONT))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED_FONT))
    for k in range(N):
        sig = L(OFF + W * k + 5)
        for txt, fill, col in [('"Accumulation forte"', FILL_G, "007A33"), ('"Distribution forte"', FILL_R, "C00000")]:
            ws.conditional_formatting.add(f"{sig}2:{sig}{last}", CellIsRule(
                operator="equal", formula=[txt], fill=fill, font=Font(name="Arial", size=10, bold=True, color=col)))
    _add_confluence_cf(ws, f"{L(c_avg)}2:{L(c_avg)}{last}", f"{L(c_ver)}2:{L(c_ver)}{last}")
    _add_dc_cf(ws, f"{L(c_vdc)}2:{L(c_vdc)}{last}")
    _add_dc_cf(ws, f"{L(c_vdcw)}2:{L(c_vdcw)}{last}")
    # v5 : codes couleurs du carnet / delta dynamiques (mêmes codes que le reste de la feuille : vert = achat / validé, rouge = vente / non fiable, jaune = neutre non validé)
    _add_rwg_scale(ws, f"{L(c_iso_pc)}2:{L(c_iso_pc)}{last}")
    for c_ in (c_dyn_conf, c_dd_conf):
        _add_rwg_scale(ws, f"{L(c_)}2:{L(c_)}{last}")                              # TVP : rouge -> blanc -> vert
    for c_, txt_ok, txt_ko in ((c_dyn_gate, '"Carnet validé par le prix"', '"Carnet non fiable – attendre"'),
                               (c_dd_gate, '"Delta validé par le prix"', '"Delta non fiable – attendre"')):
        rg_ = f"{L(c_)}2:{L(c_)}{last}"
        ws.conditional_formatting.add(rg_, CellIsRule(operator="equal", formula=[txt_ok], fill=FILL_G, font=GREEN_FONT))
        ws.conditional_formatting.add(rg_, CellIsRule(operator="equal", formula=[txt_ko], fill=FILL_R, font=RED_FONT))
    rg_ = f"{L(c_dyn_ver)}2:{L(c_dyn_ver)}{last}"
    for txt, fill, font in [('"Carnet ACHAT (validé prix)"', FILL_G, GREEN_FONT), ('"Carnet VENTE (validé prix)"', FILL_R, RED_FONT),
                            ('"Neutre / carnet non fiable"', FILL_Y, FONT_Y)]:
        ws.conditional_formatting.add(rg_, CellIsRule(operator="equal", formula=[txt], fill=fill, font=font))
    rg_ = f"{L(c_dc_ver)}2:{L(c_dc_ver)}{last}"
    cv0 = f"{L(c_dc_ver)}2"
    ws.conditional_formatting.add(rg_, FormulaRule(formula=[f'LEFT({cv0},5)="Achat"'], fill=FILL_G, font=GREEN_FONT))
    ws.conditional_formatting.add(rg_, FormulaRule(formula=[f'LEFT({cv0},5)="Vente"'], fill=FILL_R, font=RED_FONT))
    ws.conditional_formatting.add(rg_, FormulaRule(formula=[f'LEFT({cv0},14)="Neutre / non v"'], fill=FILL_Y, font=FONT_Y))

    # v2 : mises en forme du bloc final (mêmes nuanciers que les feuilles ETF)
    ws.conditional_formatting.add(f"{L(c_div_ver)}2:{L(c_div_ver)}{last}", CellIsRule(
        operator="equal", formula=['"Divergence persistante"'], fill=FILL_Y, font=FONT_Y))
    ws.conditional_formatting.add(f"{L(c_div_ver)}2:{L(c_div_ver)}{last}", CellIsRule(
        operator="equal", formula=['"Divergence récente"'], fill=FILL_O, font=FONT_O))
    _add_rwg_scale(ws, f"{L(c_qm)}2:{L(c_qm)}{last}")
    _add_whale_scale(ws, f"{L(c_abs)}2:{L(c_abs)}{last}")
    _add_rwg_scale(ws, f"{L(c_conf)}2:{L(c_conf)}{last}")
    ws.conditional_formatting.add(f"{L(c_chn)}2:{L(c_chn)}{last}", ColorScaleRule(
        start_type="num", start_value=0, start_color="FFFFFF", end_type="percentile", end_value=95, end_color="5B9BD5"))
    for seuil, color in [(CHURN_RATIO_FORT, "F4B183"), (CHURN_RATIO_ALERTE, "FCE4D6")]:
        ws.conditional_formatting.add(f"{L(c_chrat)}2:{L(c_chrat)}{last}", CellIsRule(
            operator="greaterThanOrEqual", formula=[str(seuil)], stopIfTrue=True,
            fill=PatternFill("solid", bgColor=color, fgColor=color),
            font=Font(name="Arial", size=10, bold=True, color="C55A11")))
    for seuil, color in [(TAPE_RATIO_FORT, "F4B183"), (TAPE_RATIO_ALERTE, "FCE4D6")]:      # v9 : speed of tape
        ws.conditional_formatting.add(f"{L(c_tprat)}2:{L(c_tprat)}{last}", CellIsRule(
            operator="greaterThanOrEqual", formula=[str(seuil)], stopIfTrue=True,
            fill=PatternFill("solid", bgColor=color, fgColor=color),
            font=Font(name="Arial", size=10, bold=True, color="C55A11")))
    _add_sign_cf(ws, f"{L(c_imb)}2:{L(c_imb)}{last}")
    _add_sign_cf(ws, f"{L(c_imp)}2:{L(c_imp)}{last}")
    _add_sign_cf(ws, f"{L(c_qcc)}2:{L(c_qcc)}{last}")
    # v3
    for c_ in [c_imbp, c_imbg, c_hyb, c_cons, c_j1, c_walld, c_imb_dyn, c_imb_raw, c_imb_ema] + c_cls + c_cls_dyn + c_tx:
        _add_sign_cf(ws, f"{L(c_)}2:{L(c_)}{last}")
    for c_ in (c_pers, c_conb, c_cona):
        _add_rwg_scale(ws, f"{L(c_)}2:{L(c_)}{last}")
    if c_imbs:
        ws.conditional_formatting.add(f"{L(c_imbs)}2:{L(c_imbs)}{last}", ColorScaleRule(
            start_type="num", start_value=0, start_color="FFFFFF", end_type="percentile", end_value=95, end_color="5B9BD5"))
    for txt, fill, font in [('"Carnet très ACHETEUR vs historique"', FILL_G, GREEN_FONT), ('"Carnet anormalement ACHETEUR"', FILL_G, GREEN_FONT),
                            ('"Carnet très VENDEUR vs historique"', FILL_R, RED_FONT), ('"Carnet anormalement VENDEUR"', FILL_R, RED_FONT)]:
        ws.conditional_formatting.add(f"{L(c_imb_ver)}2:{L(c_imb_ver)}{last}", CellIsRule(operator="equal", formula=[txt], fill=fill, font=font))
    r_trap = f"{L(c_trap)}2:{L(c_trap)}{last}"
    ws.conditional_formatting.add(r_trap, FormulaRule(formula=[f'LEFT({L(c_trap)}2,5)="Piège"'], fill=FILL_R, font=RED_FONT, stopIfTrue=True))
    ws.conditional_formatting.add(r_trap, FormulaRule(formula=[f'LEFT({L(c_trap)}2,7)="Retrait"'], fill=FILL_Y, font=FONT_Y))
    for txt, fill, font in [('"Fiable"', FILL_G, GREEN_FONT), ('"Moyen"', FILL_Y, FONT_Y), ('"Peu fiable"', FILL_R, RED_FONT)]:
        ws.conditional_formatting.add(f"{L(c_vdq)}2:{L(c_vdq)}{last}", CellIsRule(operator="equal", formula=[txt], fill=fill, font=font))

    # bloc « moyenne des ETF » : mêmes mises en forme que les feuilles ETF
    rg = lambda sc: f"{L(newcol[sc])}2:{L(newcol[sc])}{last}"
    for sc in (44, 47):                                    # AR, AU : valeurs signées
        ws.conditional_formatting.add(rg(sc), CellIsRule(operator="greaterThan", formula=["0"], font=GREEN_FONT))
        ws.conditional_formatting.add(rg(sc), CellIsRule(operator="lessThan", formula=["0"], font=RED_FONT))
    for sc, txt, fill, col in [
        (33, '"ACHAT"', FILL_G, "007A33"), (33, '"VENTE"', FILL_R, "C00000"),
        (34, '"Pression achat"', FILL_G, "007A33"), (34, '"Pression vente"', FILL_R, "C00000"),
        (42, '"Forte"', FILL_R, "C00000"),
        (43, '"Accumulation probable"', FILL_G, "007A33"), (43, '"Distribution probable"', FILL_R, "C00000"),
        (52, '"Accumulation forte"', FILL_G, "007A33"), (52, '"Distribution forte"', FILL_R, "C00000"),
    ]:
        ws.conditional_formatting.add(rg(sc), CellIsRule(operator="equal", formula=[txt], fill=fill,
                                                         font=Font(name="Arial", size=10, bold=True, color=col)))
    for seuil, color in [(RATIO_FORT, "F4B183"), (RATIO_ALERTE, "FCE4D6")]:
        ws.conditional_formatting.add(rg(32), CellIsRule(
            operator="greaterThanOrEqual", formula=[str(seuil)], stopIfTrue=True,
            fill=PatternFill("solid", bgColor=color, fgColor=color),
            font=Font(name="Arial", size=10, bold=True, color="C55A11")))
    ws.conditional_formatting.add(rg(41), ColorScaleRule(
        start_type="num", start_value=0, start_color="FFFFFF", mid_type="num", mid_value=50,
        mid_color="FFD966", end_type="num", end_value=100, end_color="F8696B"))
    ws.conditional_formatting.add(rg(51), ColorScaleRule(
        start_type="num", start_value=-100, start_color="F8696B", mid_type="num", mid_value=0,
        mid_color="FFFFFF", end_type="num", end_value=100, end_color="63BE7B"))
    ws.cell(1, newcol[33]).comment = Comment(
        f"Colonnes numériques : moyenne des {N} ETF BTC ({', '.join(syms)}) pour la même bougie, ETF sans valeur ignorés. "
        "Colonnes texte : mêmes règles que les feuilles ETF, appliquées aux moyennes (pression = delta % et carnet % moyens ; "
        "whale / biais whale = score whale moyen ; signal flux = score flux moyen). Sens du plus gros trade = vote majoritaire.",
        "Script Alpaca")

    # récapitulatif : nombre de bougies par type de confluence
    r = last + 2
    ws.cell(r, 1, "Récapitulatif – CONFLUENCE DELTA + CARNET").font = FB
    rng = f"${L(c_vdc)}$2:${L(c_vdc)}${last}"
    for k, label in enumerate(["ACHAT fort : delta + carnet", "Achat : delta + carnet", "Neutre / mixte",
                               "Divergence exécution / carnet", "Vente : delta + carnet",
                               "VENTE forte : delta + carnet"], start=1):
        ws.cell(r + k, 1, label).font = F
        ws.merge_cells(start_row=r + k, start_column=1, end_row=r + k, end_column=4)
        cell = ws.cell(r + k, 5, f'=COUNTIF({rng},"{label}")')
        cell.font, cell.number_format = FB, "0"
    r += 8
    ws.cell(r, 1, "Récapitulatif – CONFLUENCE DELTA + CARNET (pondéré volume)").font = FB
    rng = f"${L(c_vdcw)}$2:${L(c_vdcw)}${last}"
    for k, label in enumerate(["ACHAT fort : delta + carnet", "Achat : delta + carnet", "Neutre / mixte",
                               "Divergence exécution / carnet", "Vente : delta + carnet",
                               "VENTE forte : delta + carnet"], start=1):
        ws.cell(r + k, 1, label).font = F
        ws.merge_cells(start_row=r + k, start_column=1, end_row=r + k, end_column=4)
        cell = ws.cell(r + k, 5, f'=COUNTIF({rng},"{label}")')
        cell.font, cell.number_format = FB, "0"
    r += 8
    ws.cell(r, 1, "Récapitulatif – CONFLUENCE FLUX GLOBAL").font = FB
    rng = f"${L(c_ver)}$2:${L(c_ver)}${last}"
    for k, label in enumerate(["Confluence ACHAT forte", "Confluence achat", "Neutre / mixte", "Divergence",
                               "Confluence vente", "Confluence VENTE forte"], start=1):
        ws.cell(r + k, 1, label).font = F
        ws.merge_cells(start_row=r + k, start_column=1, end_row=r + k, end_column=4)
        cell = ws.cell(r + k, 5, f'=COUNTIF({rng},"{label}")')
        cell.font, cell.number_format = FB, "0"
    r += 8
    ws.cell(r, 1, "Récapitulatif – VERDICT DIVERGENCE").font = FB
    rng = f"${L(c_div_ver)}$2:${L(c_div_ver)}${last}"
    for k, label in enumerate(["Divergence persistante", "Divergence récente", "Pas de divergence"], start=1):
        ws.cell(r + k, 1, label).font = F
        ws.merge_cells(start_row=r + k, start_column=1, end_row=r + k, end_column=4)
        cell = ws.cell(r + k, 5, f'=COUNTIF({rng},"{label}")')
        cell.font, cell.number_format = FB, "0"

    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{L(len(heads))}{last}"
    return {"cvd_c": L(c_cvd_c), "imb": L(c_imb)}


def write_guide(wb):
    ws = wb.create_sheet("GUIDE")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:D1")
    ws["A1"] = "Lecture des indicateurs avancés"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = NAVY
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30
    for col, width in {"A": 28, "B": 88, "C": 22, "D": 22}.items():
        ws.column_dimensions[col].width = width
    rows = [
        ("Indicateur", "Interprétation", "Signal utile", "Limite principale"),
        ("Taille moyenne / trade", "Volume divisé par nombre de trades : mesure le ticket moyen exécuté.", "Hausse = tickets plus gros", "Ne mesure pas la taille du compte"),
        ("Trades / 1 000 titres", "Fragmentation des exécutions ; varie généralement à l'inverse du ticket moyen.", "Baisse = moins fragmenté", "Dépend du lieu d'exécution"),
        ("Ratios vs 12 bougies", "Compare la bougie aux 12 bougies précédentes, en excluant les bougies vides du calcul et sans s'inclure elle-même.", "> 1 = activité supérieure", "Moins fiable avec peu d'historique"),
        ("Score whale (0–100)", "30 % concentration + 25 % ticket moyen relatif + 25 % plus gros trade relatif + 20 % volume relatif.", ">= 50 : surveillance ; >= 75 : forte", "Proxy statistique, pas une identification"),
        ("Présence whale estimée", "Traduit le score en niveau et exige au moins 3 bougies d'historique et 3 trades courants.", "Forte = convergence d'anomalies", "Un découpage d'ordre peut échapper au signal"),
        ("Biais whale estimé", "Recoupe le sens estimé du plus gros trade avec le delta de la bougie.", "Accumulation ou distribution probable", "Le sens des trades reste estimé"),
        ("Plus gros trade signé", "Taille du plus gros trade, positive à l'achat estimé et négative à la vente estimée.", "Contribution au score flux", "Un seul trade ne résume pas toute la bougie"),
        ("Close vs VWAP", "Position du close par rapport au prix moyen pondéré par le volume.", "Positif = clôture au-dessus du VWAP", "Signal de très court terme"),
        ("Impact / 10k titres", "Variation absolue open-close rapportée à 10 000 titres échangés.", "Élevé = fort impact du volume", "Sensible aux spreads et au faible volume"),
        ("Score flux global", "Delta 40 % + carnet 20 % + gros trade signé 20 % + close/VWAP 20 %.", "+/-20 : biais ; +/-40 : fort si volume", "Ce n'est pas une probabilité de gain"),
        ("Signal flux global", "Synthèse directionnelle ; les signaux forts exigent aussi un volume anormal.", "Filtre les signaux sans participation", "À confirmer avec le contexte de marché"),
        ("Delta % / Carnet % moyen pondéré", "Sur les feuilles CONFLUENCE, moyenne des ETF pondérée par le volume de la bougie de chacun (Σ(valeur×Volume)/ΣVolume), au lieu d'une moyenne arithmétique simple. Un ETF peu liquide (ex. MSBT, BTCO) ne peut plus, à lui seul, faire basculer le signal face aux gros ETF (IBIT, FBTC, GBTC...).", "Reflète le poids réel de chaque ETF dans le flux", "Un ETF sans volume sur la bougie ne contribue pas"),
        ("CVD moyen Delta / Carnet", "Cumul depuis l'ouverture du marché actions US (9h30 NY) : CVD Carnet = cumul du Carnet % moyen pondéré, bougie après bougie. Vide avant 9h30 NY.", "Tendance d'accumulation/distribution de la séance US", "Repart de zéro chaque jour à 9h30 NY ; ne couvre pas la nuit"),
        ("CONFLUENCE FLUX GLOBAL", f"Compare les ETF BTC sur la même bougie. Un ETF est acheteur si son score flux >= {CONF_SEUIL}, vendeur si <= -{CONF_SEUIL}. Confluence forte si l'écart acheteurs − vendeurs représente >= {COIN_CONF_FORT:.0%} des ETF ayant une donnée, simple si >= {COIN_CONF_SIMPLE:.0%} ; mélange = divergence.", "Beaucoup d'ETF alignés = signal plus solide", "Les ETF BTC sont très corrélés : l'alignement est fréquent"),
        ("CONFLUENCE DELTA + CARNET", f"Additionne l'alignement des ETF sur le delta % (seuil {SEUIL_DELTA:.0%}) et sur le déséquilibre du carnet (seuil {SEUIL_CARNET:.0%}), en part du maximum possible : >= {DC_FORT:.0%} fort, >= {DC_SIMPLE:.0%} simple. Divergence = l'exécution et le carnet pointent en sens opposés.", "Exécution ET carnet dans le même sens", "Le carnet peut être spoofé / retiré avant exécution"),
        ("% volume ISO / Delta ISO (signé)", "Feuilles ETF (AS, AT) : % volume ISO = volume des ordres ISO (Intermarket Sweep Order, condition F) ÷ volume de la bougie ; Delta ISO = (ISO acheteur − ISO vendeur) ÷ volume de la bougie. Feuilles CONFLUENCE : moyennes pondérées par le volume des ETF. Colonnes séparées, le delta existant n'est pas modifié.", "Les ordres ISO balayent plusieurs marchés : exécution agressive et urgente", "Le sens achat/vente reste estimé ; couverture des conditions variable selon le flux (IEX/SIP)"),
        ("Delta / Carnet aligné pondéré, CONFLUENCE DELTA + CARNET (pondéré volume)", "Pour chaque ETF, valeur = Delta % (ou Carnet %) ÷ 50 %, limitée entre -1 et +1, pondérée par le volume de la bougie : de -100 % à +100 %. Pas d'effet de seuil (IBIT à +12 % pèse 0,24, pas +1). Verdict = moyenne des deux : >= 25 % fort, >= 15 % simple, divergence si l'un >= +15 % et l'autre <= -15 %. Les colonnes « aligné (net) » (comptage ±1 par ETF) restent affichées.", "Mesure où est passé le volume et avec quelle intensité", "Dominé par l'ETF le plus liquide (IBIT) ; lire avec Delta % moyen pondéré"),
        ("Direction CERTAINE des trades ISO (CE -> CI)", f"Règle bid/ask au lieu du milieu : prix >= ask de la cotation strictement antérieure = achat, prix <= bid = vente ; marché locked/croisé ou condition non régulière = non classé. Balayage ISO : rafale (< {SWEEP_GAP_MS} ms) sur >= {SWEEP_MIN_LEVELS} prix distincts toujours croissants (achat) ou décroissants (vente). Le milieu et la tick rule ne servent jamais dans le « certain » ; leurs trades vont dans « Delta ISO estimé ». % volume ISO certain = représentativité.", "Sépare ce qui est certain de ce qui est estimé", "« Certain » = quasi certain (une limite marketable posée pile à l'ask reste possible) ; sur IEX la cotation est celle d'IEX, pas le NBBO"),
        ("Carnet dynamique validé par le prix (TVP)", f"Chaque ETF est pondéré par son Taux de Validation Prix : sur {REACT_WINDOW} bougies, part des signaux carnet (|BN| >= {REACT_TRIGGER:.0%}) suivis d'un mouvement du close d'au moins {REACT_THR_AMP:g} x l'amplitude moyenne dans le même sens à j+{REACT_H} (mouvement inverse = échec, prix immobile = non compté). Lissé ({REACT_PRIOR} cas à 50 %), neutre sous {REACT_MIN_CASES} cas. Poids = qualité du carnet x fiabilité^{REACT_GAMMA}, sans volume exécuté, plafond {REACT_CAP:.0%} par ETF. CVD à demi-vie {REACT_HALF_LIFE_MIN} min, porte de validation (TVP agrégé >= {REACT_GATE:.0%}) et verdict carnet final.", "Un carnet n'influence le verdict que s'il a fait bouger le prix dans le sens annoncé", "Peu de cas en début de séance (TVP neutre, poids proches de la qualité) ; réaction mesurée sur le close de l'ETF ; pas encore de mode absorption ni de filtre de tendance"),
        ("Bloc MOY. DES ETF (AG → AZ)", "Sur les feuilles CONFLUENCE, reprend les indicateurs AG à AR et AU à AZ des feuilles ETF (le ratio volume J-1 moyen est juste après les blocs ETF) : moyenne des ETF pour les colonnes numériques ; pour les colonnes texte (sens du plus gros trade, pression, présence/biais whale, signal flux) les mêmes règles sont appliquées aux valeurs moyennes (vote majoritaire pour le sens du plus gros trade).", "Lecture globale d'une bougie en un coup d'œil", "Une moyenne peut masquer un ETF très différent des autres"),
        ("OI BTC agrégé (var. % / bougie)", "Sur les feuilles CONFLUENCE, à côté du delta % et du carnet % moyens : variation de l'Open Interest BTC agrégé (perp + futures à échéance des principaux exchanges, USD, source Coinalyze) sur la bougie.", "OI qui monte = nouvelles positions ouvertes", "Proxy BTC (pas d'OI sur les ETF) ; options non incluses ; n'entre dans aucun score"),
        ("Ratio volume vs réf. normale (AF / BR)", f"Volume de la bougie ÷ volume normal à la même heure = médiane des séances de l'historique local (colonne BA, lissée ±{REF_SMOOTH} bougie, plancher {REF_FLOOR_FRAC:.0%} du volume moyen J-1). Sur CONFLUENCE : ratio agrégé Σ volumes ÷ Σ références.", f"Alerte >= {RATIO_ALERTE}x ; fort >= {RATIO_FORT}x", "Dépend de l'historique : si plus de la moitié des séances sont atypiques, la référence l'est aussi"),
        ("Ratio churn carnet vs réf. J-1 (CONFLUENCE uniquement)", f"Σ churn des ETF (MàJ NBBO / min, colonne BC) ÷ Σ churn normal à la même heure (médiane des séances de l'historique local, colonne masquée _réf_churn_J-1, lissée ±{REF_SMOOTH} bougie, plancher {CHURN_REF_FLOOR_FRAC:.0%} du churn moyen J-1). Placé juste après le ratio volume.", f"Alerte >= {CHURN_RATIO_ALERTE}x ; fort >= {CHURN_RATIO_FORT}x", "Mesure l'agitation des cotations, pas leur sens ; sur IEX ce sont les cotations d'IEX ; à lire avec le ratio volume (un volume élevé fait monter le churn)"),
        ("Speed of tape (CL -> CN des feuilles ETF ; CONFLUENCE : colonne F)", f"Vitesse d'impression du ruban = trades ÷ {BAR_MINUTES} min (trades / minute), comparée à la médiane de la MÊME heure dans l'historique local (lissée ±{REF_SMOOTH} bougie, plancher {TAPE_REF_FLOOR_FRAC:.0%} de la vitesse moyenne). Sur CONFLUENCE : Σ vitesses ÷ Σ réf. des ETF.", f"Alerte >= {TAPE_RATIO_ALERTE}x ; fort >= {TAPE_RATIO_FORT}x", "Mesure la cadence, pas le sens ; à lire avec le ratio volume (volume normal + tape rapide = beaucoup de petits trades ; volume élevé + tape normal = gros prints)"),
        ("Delta / Carnet aligné (net, pondéré volume)", f"CONFLUENCE : chaque ETF vote +1 / -1 au seuil de {SEUIL_DELTA:.0%} (delta) ou {SEUIL_CARNET:.0%} (carnet), pondéré par son volume : de -100 % à +100 %. CONFLUENCE DELTA + CARNET en est la moyenne.", "Proche de ±100 % = ETF lourds alignés", "Effet de seuil (contrairement aux colonnes « aligné pondéré volume »)"),
        ("CVD cumulé delta pondéré", "Cumul depuis 9h30 NY d'un delta moyen par bougie où chaque ETF pèse volume × |Delta %| (les ETF lourds ET directionnels dominent). Placé à droite de « Delta aligné pondéré volume ».", "Hausse continue = flux acheteur durable porté par les gros ETF", "Un seul ETF très actif peut dominer une bougie"),
        ("CVD carnet prix (v12)", f"Placée à droite de « CVD carnet » (qui reste inchangée). Même carnet dynamique, mais chaque bougie est créditée selon la réaction du prix : k = (rendement pondéré des ETF sur {PX_BARS} bougies ÷ amplitude moyenne) dans le sens du carnet ÷ {PX_TAU:g}, borné entre -1 et +1. Carnet confirmé par le prix = compte en entier ; carnet sans effet sur le prix = ignoré ; prix à contre-sens = contribution inversée (absorption). Demi-vie {PX_HALF_LIFE_MIN} min.", "Passe négative quand le prix casse alors que le carnet reste acheteur ; comparer avec CVD carnet pour repérer un retournement", "Valeur de rattrapage, pas un signal d'entrée ; paramètres PX_BARS / PX_TAU / PX_HALF_LIFE_MIN à calibrer sur plusieurs séances"),
        ("CVD carnet SIP", "Feuille CONFLUENCE_IEX uniquement : CVD Carnet lu dans CONFLUENCE_SIP (colonne cvd_sip : à droite des blocs ETF, ou à côté du CVD Carnet IEX si tu l'ajoutes à CONFLUENCE_FRONT).", "Comparer carnet IEX et carnet consolidé", "Vide si le SIP n'est pas exécuté"),
        ("Liste des ETF", f"{len(ETFS)} ETF Bitcoin spot : " + ", ".join(ETFS) + f". {MANDATORY} est obligatoire ; un autre ETF sans données est simplement ignoré.", "Modifiable en haut du script (ETFS)", "Les petits ETF ont peu de trades sur IEX"),
        ("IEX vs SIP", "Mêmes formules sur les deux flux. IEX = échantillon d'une seule bourse ; SIP = marché consolidé. Compare les deux feuilles pour juger la fiabilité d'un signal.", "Signal présent sur les 2 flux = plus fiable", "Aucun flux ne révèle l'identité du trader"),
        # ---- v2 « qualité du carnet » : 10 nouvelles lignes ----
        ("Coef qualité carnet (BB)", f"Médiane mobile du spread moyen de l'ETF ({BOOK_QUALITY_WINDOW} bougies précédentes) ÷ spread moyen de la bougie, plafonné à 100 %. 100 % = carnet aussi serré que d'habitude ; plus bas = carnet plus large. Calculé séparément pour IEX et SIP.", "Élevé = carnet serré, donc plus fiable", "Sur IEX, le spread est artificiellement large quand IEX n'est pas le meilleur marché : compare IEX et SIP"),
        ("Churn carnet (BC)", "Mises à jour du NBBO par minute (quotes), en dehors des trades. Complète la qualité : le spread mesure la qualité statique, le churn la qualité dynamique.", "Élevé + tailles qui tournent = market makers nerveux", "Un churn faible = carnet dormant, plus facile à manipuler"),
        ("Rotation tailles (BD)", "Rapport max / min de (taille bid + taille ask) sur les quotes de la bougie : proxy de spoofing / quote stuffing (ordres posés puis retirés).", "Valeur élevée sans trades = affichage douteux", "Pas de L2 : on ne voit pas qui pose ni retire les ordres"),
        ("Déséquilibre pondéré temps (BE)", f"Déséquilibre bid/ask où chaque état du carnet pèse selon sa durée d'affichage (plafond {TW_CAP_SECONDS // 60} min). Corrige la colonne AB, qui donne le même poids à un état affiché 10 s et à un état affiché 4 min.", "Vert = carnet acheteur la plupart du temps ; rouge = vendeur", "Les quotes IEX ne reflètent que le carnet d'IEX"),
        ("Impact post-trade (BF)", f"Déplacement du mid dans les {IMPACT_HORIZON_S} s après chaque trade, divisé par le spread, pondéré par la taille des trades. Mesure indirecte de la profondeur (L2 implicite).", "Vert = mid monté après les trades ; rouge = baissé", "Le mid peut bouger pour une autre raison qu'un trade (BTC, autre ETF)"),
        ("Score absorption (BG)", f"0-100 : exécution À CONTRE du carnet affiché (delta de signe opposé au déséquilibre pondéré temps) × intensité du déséquilibre × coef qualité (BB). 0 si le déséquilibre est < {SEUIL_CARNET:.0%}.", ">= 50 : un acteur semble absorber la liquidité", "Proxy ; ne prouve ni l'identité ni l'intention"),
        ("Confirmation et Verdict carnet (BH, BI)", f"BH : sur les {CONFIRM_WINDOW} dernières bougies, part des cas où un carnet déséquilibré (>= {SEUIL_CARNET:.0%}) a été suivi d'un delta de même signe à la bougie suivante. BI : Fiable (vert) si qualité >= {VERDICT_FIABLE:.0%} ET confirmation >= {VERDICT_FIABLE:.0%} ; Peu fiable (rouge) si l'une des deux < {VERDICT_PEU_FIABLE:.0%} ; sinon Moyen (jaune).", "BH > 60 % = le carnet prédit le delta ; < 40 % = il ment souvent", f"Vide s'il y a moins de {CONFIRM_MIN_SAMPLES} cas ; le taux est rétrospectif, pas une probabilité"),
        ("Taille affichée crédible (BJ -> CD)", f"Persistance des états du carnet, murs anormaux (> {WALL_MULT:g} x médiane, delta mur bid − mur ask sur CONFLUENCE), déséquilibre temps x persistance, déséquilibre par classe de taille, retraits de taille consommés par des trades vs annulés, alerte « piège retiré → sweep », écart du déséquilibre vs médiane J-1. Sur CONFLUENCE : agrégats pondérés volume à droite de « Déséquilibre pondéré temps », dont une version volume x fiabilité et l'écart IEX vs SIP.", "Persistance haute + gros mur + consommation élevée = liquidité réelle ; BE fort mais BN faible = flashs", "Sur IEX, une annulation peut être une amélioration venue d'un autre marché : compare avec le SIP. Aucun de ces indicateurs ne prouve une intention."),
        ("ETF divergent / Écart max / VERDICT DIVERGENCE", f"Sur CONFLUENCE : un ETF est divergent si son Delta % est de signe opposé au consensus (Delta % moyen pondéré), avec |consensus| >= {DIV_CONSENSUS_MIN:.0%} et un écart >= {DIV_GAP * 100:.0f} pts. L'écart max est signé (vert = ETF acheteur contre un consensus vendeur ; rouge = l'inverse). Persistante (jaune) = divergence sur 2 bougies de suite ; récente (orange) = sur la bougie seulement. Calcul via 12 colonnes masquées _gap_<ETF>.", "Un ETF qui s'écarte des autres peut signaler une rotation", "Les ETF BTC sont corrélés : une divergence peut n'être que du bruit sur un petit ETF"),
        ("Carnet qualité moyen, Carnet % qualité-pondéré, Absorption et Confirmation moyennes", "Sur CONFLUENCE : qualité moyenne pondérée volume ; Carnet % moyen pondéré par le coef qualité (les carnets serrés comptent plus que les bruités, pondération « inverse du spread ») ; moyennes simples de l'absorption (BG) et de la confirmation (BH).", "Carnet % fiable quand la qualité moyenne est élevée", "Moyennes : peuvent masquer un ETF très différent"),
        ("Enchères 9h30 / 16h00 (optionnel)", "Sur CONFLUENCE, si Alpaca renvoie les enchères pour ton plan : volume de la croix d'ouverture (9h30) et de clôture (16h00), somme des ETF, et poids de la croix de clôture dans le volume du jour. Sinon ces colonnes sont absentes, sans erreur.", "Croix de clôture lourde = flux institutionnel (MOC)", "Prints de croix, pas un imbalance officiel ; dépend de l'endpoint et du plan Alpaca"),
    ]
    for r, values in enumerate(rows, start=3):
        for c, value in enumerate(values, start=1):
            cell = ws.cell(r, c, value)
            cell.font = F_HEAD if r == 3 else F
            cell.fill = NAVY if r == 3 else PatternFill("solid", fgColor="FFFFFF" if r % 2 else "F2F2F2")
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = 34 if r > 3 else 26
    ws.freeze_panes = "A4"
    end = 3 + len(rows) + 1
    ws[f"A{end}"] = "Important : ces indicateurs décrivent des anomalies d'exécution ; ils ne prouvent ni la présence, ni l'identité, ni l'intention d'une whale."
    ws.merge_cells(f"A{end}:D{end}")
    ws[f"A{end}"].font = Font(name="Arial", size=10, bold=True, color="C00000")
    ws[f"A{end}"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[end].height = 34


def write_workbook(data, path, oi=None, auctions=None):
    """data = {feed: {"bars": {sym: DataFrame}, "refs": {sym: (moyenne, date)}}}

    Ordre des feuilles : CONFLUENCE_IEX, CONFLUENCE_SIP, puis les ETF en IEX, puis les ETF en SIP, puis GUIDE.
    """
    wb = Workbook()
    wb.remove(wb.active)
    oi = oi or {}
    conf_ws = {feed: wb.create_sheet(f"CONFLUENCE_{feed}") for feed in data}   # ordre des onglets inchangé
    sip_cvd = None
    sip_imb = None
    if "SIP" in data:      # le SIP est écrit d'abord pour connaître la colonne de son CVD Carnet
        info = write_confluence(conf_ws["SIP"], data["SIP"]["bars"], "SIP", oi, auctions=auctions)
        if info:
            sip_cvd = ("CONFLUENCE_SIP", info["cvd_c"])
            sip_imb = ("CONFLUENCE_SIP", info["imb"])
    for feed, d in data.items():
        if feed != "SIP":
            write_confluence(conf_ws[feed], d["bars"], feed, oi, sip_cvd=sip_cvd if feed == "IEX" else None,
                             auctions=auctions, sip_imb=sip_imb if feed == "IEX" else None)
    for feed, d in data.items():
        for sym in ETFS:
            ws = wb.create_sheet(f"{sym}_{feed}")
            if sym in d["bars"]:
                write_sheet(ws, d["bars"][sym], feed, *d["refs"].get(sym, (None, None, None)))
            else:
                ws["A1"] = f"Pas de données pour {sym} ({feed})"
    write_guide(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(path)


# ----------------------------------------------------------------- aide au calibrage du ratio J-1
def _print_ratio_stats(bars_by_sym, refs, feed):
    """Distribution du ratio volume / référence normale (même tranche) pour ajuster RATIO_ALERTE / RATIO_FORT.
    Calculé comme la colonne « Ratio volume » de CONFLUENCE : Σ volumes ÷ Σ références des ETF, bougie par bougie.
    Les ETF sans référence valide à cette tranche (réf. nulle : quasi aucun trade sur IEX) sont exclus, au lieu de produire des ratios aberrants."""
    nums, dens = [], []
    for sym, b in bars_by_sym.items():
        slots = refs.get(sym, (None, None, None))[2]
        if not slots:
            continue
        v = np.nan_to_num(b["volume"].to_numpy(dtype=float))
        m = min(len(v), len(slots))
        sl = np.asarray(slots[:m], dtype=float)
        ok = np.isfinite(sl) & (sl > 0)
        nums.append(np.where(ok, v[:m], 0.0))
        dens.append(np.where(ok, sl, 0.0))
    if not nums:
        return
    M = min(len(a) for a in nums)
    num = np.sum([a[:M] for a in nums], axis=0)
    den = np.sum([a[:M] for a in dens], axis=0)
    good = (den > 0) & (num > 0)
    vals = num[good] / den[good]
    if len(vals) < 20:
        return
    p50, p90, p97 = np.percentile(vals, [50, 90, 97])
    print(f"\n[calibrage {feed}] ratio vol / réf. normale (agrégé {len(nums)} ETF) : P50={p50:.2f}  P90={p90:.2f}  P97={p97:.2f}  "
          f"(seuils actuels : alerte {RATIO_ALERTE}, fort {RATIO_FORT})")


# ----------------------------------------------------------------- collecte d'un flux
# ----------------------------------------------------------------- mode live : téléchargement incrémental + caches jour
_PREV_REF_CACHE = {}      # (sym, date, feed) -> référence volume J-1 (figée pour la journée : ne pas retélécharger à chaque tick)
_BOOK_REF_CACHE = {}      # (sym, feed, tag) -> {colonne: tableau numpy} réf. carnet J-1 (figée pour la journée)


def _read_cached_csv(path):
    try:
        df = pd.read_csv(path)
    except Exception:
        return pd.DataFrame()
    if "timestamp" not in df.columns or df.empty:
        return pd.DataFrame()
    df["timestamp"] = pd.to_datetime(df["timestamp"], utc=True)
    return df.sort_values("timestamp").reset_index(drop=True)


def _fetch_incremental(client, sym, start, end, feed, tag, kind):
    """Live : réutilise le CSV brut du jour et ne télécharge que le delta depuis le dernier timestamp connu.
    Hors mode live (ou sans cache) : téléchargement complet, comportement d'origine."""
    fetch = fetch_trades if kind == "trades" else fetch_quotes
    cache = os.path.join(OUT_DIR, f"{sym}_{tag}_{feed}_{kind}.csv.gz")
    if not LIVE_MODE or not os.path.exists(cache):
        return fetch(client, sym, start, end, feed)
    old = _read_cached_csv(cache)
    if old.empty:
        return fetch(client, sym, start, end, feed)
    last_ts = old["timestamp"].iloc[-1]
    if pd.Timestamp(last_ts) >= pd.Timestamp(end):
        new = pd.DataFrame()                       # rien de nouveau (tick sans nouvelle donnée)
    else:
        new = fetch(client, sym, pd.Timestamp(last_ts) + pd.Timedelta(microseconds=1), end, feed)
    if new is None or new.empty:
        return old
    new["timestamp"] = pd.to_datetime(new["timestamp"], utc=True)
    new = new[new["timestamp"] > last_ts]
    if new.empty:
        return old
    out = pd.concat([old, new], ignore_index=True).sort_values("timestamp").reset_index(drop=True)
    print(f"(incrémental : {len(old):,} en cache + {len(new):,} nouvelles lignes)")
    return out


def _attach_book_ref_cached(client, sym, feed, ref, bars, tag):
    """Live : la réf. carnet J-1 (colonnes imb_ref / churn_ref / imb_mad / tape_ref) est figée pour la journée :
    calculée une fois, réappliquée à chaque tick (les créneaux de la journée ne changent pas)."""
    key = (sym, feed, tag)
    cached = _BOOK_REF_CACHE.get(key)
    if cached is not None:
        n = len(bars)
        for col, arr in cached.items():
            out = np.full(n, np.nan)
            m = min(n, len(arr))
            out[:m] = arr[:m]
            bars[col] = out
        return
    _attach_book_ref(client, sym, feed, ref, bars)
    cols = [c for c in ("imb_ref", "churn_ref", "imb_mad", "tape_ref") if c in bars.columns]
    if cols:
        _BOOK_REF_CACHE[key] = {c: bars[c].to_numpy(dtype=float).copy() for c in cols}


def collect_feed(client, feed, day, tag):
    from alpaca.common.exceptions import APIError

    start, end = session_window(day, feed)
    day_date = start.astimezone(NY).date()
    print(f"\n########## {feed} | fenêtre UTC : {start} -> {end} | bougies {BAR_MINUTES} min")
    bars_by_sym, refs = {}, {}

    for sym in ETFS:
        print(f"\n=== {sym} ({feed}) ===")
        try:
            trades = _fetch_incremental(client, sym, start, end, feed, tag, "trades")
            quotes = _fetch_incremental(client, sym, start, end, feed, tag, "quotes")
            _rk = (sym, str(day_date), feed)
            if LIVE_MODE and _rk in _PREV_REF_CACHE:
                refs[sym] = _PREV_REF_CACHE[_rk]
            else:
                refs[sym] = prev_session_avg_volume(client, sym, day_date, feed)
                if LIVE_MODE:
                    _PREV_REF_CACHE[_rk] = refs[sym]
        except APIError as exc:
            message = str(exc)
            if "401" in message:
                raise SystemExit(
                    "ERREUR 401 : Alpaca refuse le couple clé API / clé secrète.\n"
                    "Vérifie que les deux valeurs viennent du même compte Alpaca, sans espace "
                    "ni guillemet parasite, puis relance. Ne partage pas ta clé secrète."
                ) from exc
            if "403" in message:
                raise FeedUnavailable(f"accès refusé au flux {feed} (abonnement ?) : {exc}") from exc
            if sym == MANDATORY:
                raise SystemExit(f"Erreur Alpaca pour {sym} ({feed}) : {exc}") from exc
            print(f"[!] {sym} ({feed}) ignoré : {exc}")      # symbole inconnu, pas coté à cette date…
            continue
        print(f"trades : {len(trades):,} | quotes : {len(quotes):,}")
        avg, ref_d, _ = refs[sym]
        print(f"moyenne volume/bougie de référence (médiane de l'historique, séance la plus récente : {ref_d}) : {avg:,.0f}" if avg is not None
              else "pas de référence trouvée (ratio vide)")
        if trades.empty or quotes.empty:
            print("Pas de données (vérifie le feed / ton plan / la date).")
            continue
        os.makedirs(OUT_DIR, exist_ok=True)   # recrée le dossier s'il a disparu pendant l'exécution
        trades.to_csv(f"{OUT_DIR}/{sym}_{tag}_{feed}_trades.csv.gz", index=False)
        quotes.to_csv(f"{OUT_DIR}/{sym}_{tag}_{feed}_quotes.csv.gz", index=False)
        try:
            cl = classify(trades, quotes)
            cl.to_csv(f"{OUT_DIR}/{sym}_{tag}_{feed}_trades_classified.csv.gz", index=False)
            bars_by_sym[sym] = build_bars(cl, quotes, pd.Timestamp(start), pd.Timestamp(end))
            _attach_book_ref_cached(client, sym, feed, refs.get(sym), bars_by_sym[sym], tag)
        except Exception as exc:
            if sym == MANDATORY:
                raise
            print(f"[!] {sym} ({feed}) ignoré (données inexploitables) : {exc}")
    _print_ratio_stats(bars_by_sym, refs, feed)
    return {"bars": bars_by_sym, "refs": refs}


# ----------------------------------------------------------------- mode LIVE
# Correspondance figée entre les colonnes de la feuille CONFLUENCE_<flux> et les lettres lues par alerts.py
# (CONFLUENCE_LAYOUT, ordre fixe après N° bougie / Heure UTC / Heure NY). Si tu modifies CONFLUENCE_LAYOUT,
# mets ce tableau à jour : le contrôle d'en-têtes ci-dessous arrêtera les alertes plutôt que d'alerter sur de mauvaises colonnes.
ALERT_COLS = {"D": "ratio", "E": "chn_ratio", "F": "tape_ratio", "G": "ndw", "I": "dyn_cvd",
              "J": "dyn_cvd_px", "K": "dtop_val", "L": "dd_c", "M": "iso_d", "O": "tx_cls3", "P": "tx_cls4"}
ALERT_HEAD_CHECK = {"D": "Ratio volume", "E": "Ratio churn", "F": "Speed of tape", "G": "Delta aligné",
                    "I": "CVD carnet", "J": "CVD carnet prix", "K": "le plus fiable", "L": "Delta % dynamique",
                    "M": "Delta ISO", "O": "EXÉCUTÉ", "P": "EXÉCUTÉ"}


def _num_or_none(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) and np.isfinite(v) else None


def _extract_last_bar_values(xlsx_path, feed, day):
    """Recalcule le classeur (LibreOffice, headless) puis lit la DERNIÈRE bougie clôturée de CONFLUENCE_<feed>.
    -> (ny_iso 'AAAA-MM-JJTHH:MM:SS', {lettre: valeur|None}) ou None si indisponible / en-têtes inattendus.
    Les valeurs lues sont exactement celles qu'Excel affiche (les formules sont recalculées, pas réinterprétées)."""
    import subprocess
    import tempfile
    from openpyxl import load_workbook

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("[!] LibreOffice introuvable : alertes désactivées pour ce tick (l'Excel est quand même produit).")
        return None
    try:
        with tempfile.TemporaryDirectory(dir=OUT_DIR) as tmp:
            r = subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "xlsx",
                                "--outdir", tmp, xlsx_path], capture_output=True, text=True, timeout=300)
            recalced = os.path.join(tmp, os.path.basename(xlsx_path))
            if r.returncode != 0 or not os.path.exists(recalced):
                print(f"[!] Recalcul LibreOffice impossible ({r.stderr.strip()[-200:]}) : tick d'alerte sauté.")
                return None
            wb = load_workbook(recalced, read_only=True, data_only=True)
            try:
                sheet = f"CONFLUENCE_{feed}"
                if sheet not in wb.sheetnames:
                    return None
                ws = wb[sheet]
                rows = list(ws.iter_rows(min_row=1, max_col=16, values_only=True))
            finally:
                wb.close()
    except Exception as exc:
        print(f"[!] Lecture des valeurs recalculées impossible ({exc}) : tick d'alerte sauté.")
        return None
    if len(rows) < 2:
        return None
    hdr = rows[0]
    for letter, frag in ALERT_HEAD_CHECK.items():
        ci = ord(letter) - 65
        title = str(hdr[ci] or "")
        if frag.lower() not in title.lower():
            print(f"[!] Colonne {letter} inattendue (« {title} », attendu « {frag} ») : alertes suspendues "
                  "(la mise en page CONFLUENCE_LAYOUT a changé -> mettre ALERT_COLS à jour).")
            return None
    last = None
    for row in rows[1:]:
        if isinstance(row[0], (int, float)) and not isinstance(row[0], bool):
            last = row                                   # récapitulatifs en bas de feuille ignorés (pas de N° de bougie)
    if last is None:
        return None
    t = last[2]                                          # colonne C : Heure New York
    if hasattr(t, "hour"):
        hh, mm = t.hour, t.minute
    else:
        try:
            hh, mm = int(str(t)[:2]), int(str(t)[3:5])
        except Exception:
            return None
    ny_iso = f"{day}T{hh:02d}:{mm:02d}:00"
    out = {letter: _num_or_none(last[ord(letter) - 65]) for letter in ALERT_COLS}
    return ny_iso, out


def _fire_alerts(ny_iso, out):
    """Branche la machine à états alerts.py (Mode 1 / Mode 2 « Vito » + Telegram) sur la dernière bougie clôturée.
    Les valeurs envoyées sont celles de l'Excel, recalculées : alerte et fichier ne peuvent plus diverger.
    (Montage hybride : désactivé par défaut, c'est le bot qui alerte à partir du JSON publié — voir _git_push_tick.)"""
    try:
        import alerts
    except Exception as exc:
        print(f"[!] alerts.py indisponible ({exc}) : alertes désactivées, l'Excel continue.")
        return
    try:
        events = alerts.on_bar(ny_iso, out)
        for e in events:
            print(f"ALERTE {e.mode} {e.kind} {e.direction} ({e.ny} NY, bougie {e.bar}) {e.info}")
    except Exception as exc:
        print(f"[!] Erreur machine à alertes ({exc}) : l'Excel continue.")


def _last_bar_payload(ny_iso, feed, out):
    """Petit JSON publié à chaque tick : la dernière bougie clôturée, valeurs recalculées de l'Excel (D..P).
    C'est le canal vers le bot : il y puise G, I, J, K, L, M, O, P et les combine à ses propres ratios D/E/F."""
    hh, mm = int(ny_iso[11:13]), int(ny_iso[14:16])
    return {"day": ny_iso[:10], "ny_iso": ny_iso, "bar": (hh * 60 + mm) // BAR_MINUTES + 1,
            "feed": feed, "updated_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "values": out}


# ---- push GitHub (le dépôt sert de passerelle : ton autre script lit l'Excel via GitHub, branche dédiée) ----
_GIT_OK = None


def _git(*args, check=True):
    import subprocess
    return subprocess.run(["git", "-C", GIT_DIR, *args], capture_output=True, text=True, check=check)


def _git_setup():
    global _GIT_OK
    if _GIT_OK is not None:
        return _GIT_OK
    _GIT_OK = False
    if not GITHUB_REPO or not GITHUB_TOKEN:
        print("[i] GITHUB_REPO / GITHUB_TOKEN absents : pas de push GitHub (l'Excel reste sur le Volume).")
        return False
    if not shutil.which("git"):
        print("[!] git introuvable dans l'image : pas de push GitHub.")
        return False
    import subprocess
    url = f"https://x-access-token:{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git"
    try:
        if not os.path.isdir(os.path.join(GIT_DIR, ".git")):
            os.makedirs(os.path.dirname(GIT_DIR) or ".", exist_ok=True)
            subprocess.run(["git", "clone", url, GIT_DIR], capture_output=True, text=True, check=True)
        remote = _git("ls-remote", "--heads", "origin", GITHUB_BRANCH).stdout.strip()
        if remote:
            _git("fetch", "origin", GITHUB_BRANCH)
            _git("checkout", "-B", GITHUB_BRANCH, f"origin/{GITHUB_BRANCH}")
        else:
            _git("checkout", "-B", GITHUB_BRANCH)        # première utilisation : la branche est créée au premier push
        _git("config", "user.email", "railway-live@local")
        _git("config", "user.name", "Railway live")
        _GIT_OK = True
        print(f"Push GitHub actif : {GITHUB_REPO} (branche {GITHUB_BRANCH}).")
    except subprocess.CalledProcessError as exc:
        print(f"[!] GitHub injoignable au démarrage ({(exc.stderr or '').strip()[-200:]}) : nouvel essai à chaque tick.")
    return _GIT_OK


def _git_commit_push(msg):
    """Commit + push sur la branche de données : un seul commit « vivant » (amend + force push).
    Sans ça, chaque tick figerait un classeur complet dans l'historique git et le dépôt gonflerait
    d'un fichier Excel entier toutes les 5 minutes. La branche live-data est dédiée : personne d'autre n'y pousse."""
    import subprocess
    try:
        _git("add", "-f", "-A")
        st = _git("status", "--porcelain").stdout.strip()
        if not st:
            return
        if _git("rev-parse", "--verify", "HEAD", check=False).returncode == 0:
            _git("commit", "--amend", "-m", msg)
        else:
            _git("commit", "-m", msg)
        try:
            _git("push", "--force-with-lease", "origin", GITHUB_BRANCH)
        except subprocess.CalledProcessError:
            _git("push", "--force", "origin", GITHUB_BRANCH)
    except subprocess.CalledProcessError as exc:
        print(f"[!] Push GitHub échoué ({(exc.stderr or '').strip()[-200:]}) : réessaie au prochain tick.")


def _git_push_tick(xlsx_path, day, last_bar=None):
    if not _git_setup():
        return
    os.makedirs(os.path.join(GIT_DIR, "live"), exist_ok=True)
    # nom stable pour ton autre script : toujours le même chemin, branche dédiée
    shutil.copy2(xlsx_path, os.path.join(GIT_DIR, "live", "flux_confluence_live.xlsx"))
    # JSON de la dernière bougie clôturée : c'est lui que le bot lit pour alimenter les alertes (G, I, J, K, L, M, O, P)
    if last_bar:
        ny_iso, feed, out = last_bar
        try:
            with open(os.path.join(GIT_DIR, "live", "confluence_derniere_bougie.json"), "w", encoding="utf-8") as fh:
                json.dump(_last_bar_payload(ny_iso, feed, out), fh, ensure_ascii=False)
        except Exception as exc:
            print(f"[!] Écriture du JSON de bougie impossible ({exc}).")
    # copie datée dans history/ (alimente replay.py et l'historique partagé)
    os.makedirs(os.path.join(GIT_DIR, "history"), exist_ok=True)
    shutil.copy2(xlsx_path, os.path.join(GIT_DIR, "history", os.path.basename(xlsx_path)))
    _git_commit_push(f"live {day} {datetime.now(NY).strftime('%H:%M')} NY")


# ----------------------------------------------------------------- boucle live
def _bootstrap_history():
    """Premier démarrage sur Railway : si l'historique local est vide ou presque, on l'amorce avec les Excel
    quotidiens déposés dans le dépôt (dossier history/) -> les ratios D / E / F ont leurs références dès le
    premier tick. Ensuite le système s'auto-alimente (chaque journée vécue rejoint l'historique)."""
    if not HIST_DIR:
        return
    try:
        os.makedirs(HIST_DIR, exist_ok=True)
        have = [n for n in os.listdir(HIST_DIR)
                if re.match(r"flux_BTC_ETF_\d{8}_.*\.xlsx$", n, re.I) and not n.startswith("~$")]
        if len(have) >= HIST_MIN_DAYS:
            return                                        # historique déjà présent : rien à faire
        src = os.path.join(GIT_DIR, "history")
        if not os.path.isdir(src):
            print(f"[i] Pas de dossier history/ dans le dépôt : l'historique se construira jour après jour "
                  f"(repli J-1 en dessous de {HIST_MIN_DAYS} séances).")
            return
        n = 0
        for name in os.listdir(src):
            if re.match(r"flux_BTC_ETF_\d{8}_.*\.xlsx$", name, re.I) and not name.startswith("~$"):
                shutil.copy2(os.path.join(src, name), os.path.join(HIST_DIR, name))
                n += 1
        if n:
            print(f"Historique amorcé depuis le dépôt : {n} fichier(s) Excel copié(s) dans {HIST_DIR}.")
        else:
            print("[i] Aucun Excel quotidien trouvé dans history/ : l'historique se construira jour après jour.")
    except Exception as exc:
        print(f"[!] Amorçage de l'historique impossible ({exc}) : le live continue, références limitées les premiers jours.")


def run_day(day, feeds):
    """Une passe complète (batch ou tick live) -> (chemin Excel, data). Corps commun, extrait de main()."""
    client = _client()
    os.makedirs(OUT_DIR, exist_ok=True)
    d0 = datetime.strptime(day, "%Y-%m-%d")
    tag = d0.strftime("%Y%m%d")

    data = {}
    for feed in feeds:
        try:
            data[feed] = collect_feed(client, feed, day, tag)
        except (FeedUnavailable, ValueError) as exc:
            print(f"\n[!] Flux {feed} ignoré : {exc}")
    if not data:
        raise SystemExit("Aucun flux exploitable.")

    out = f"flux_BTC_ETF_{tag}_{'_'.join(data)}_confluence{'_24h' if FULL_DAY else ''}.xlsx"
    oi = {}
    idx = [b.index for d in data.values() for b in d["bars"].values() if len(b)]
    if idx:
        t0 = min(ix[0] for ix in idx)
        t1 = min(max(ix[-1] for ix in idx) + pd.Timedelta(minutes=BAR_MINUTES), pd.Timestamp.now(tz="UTC"))
        oi = fetch_btc_oi(t0, t1)
    auctions = {}
    syms_all = [s for s in ETFS if any(s in d["bars"] for d in data.values())]
    if syms_all:
        try:
            auctions = fetch_auctions(syms_all, d0.date())
        except Exception as exc:
            print(f"\n[!] Enchères ignorées : {exc}")
            auctions = {}
    write_workbook(data, out, oi, auctions)
    print(f"\nExcel écrit : {out}")
    if HIST_SAVE_TODAY and HIST_DIR:
        try:
            os.makedirs(HIST_DIR, exist_ok=True)
            dest = os.path.join(HIST_DIR, os.path.basename(out))
            shutil.copy2(out, dest)
            print(f"Copie dans l'historique : {dest}")
        except Exception as exc:
            print(f"[!] Copie dans l'historique impossible : {exc}")
    return out, data


def _wait_next_tick():
    """Dort jusqu'à LIVE_SETTLE_S secondes après la prochaine clôture de bougie (BAR_MINUTES)."""
    step = BAR_MINUTES * 60
    now = time.time()
    nxt = (int(now) // step + 1) * step + LIVE_SETTLE_S
    delay = max(0.0, nxt - now)
    nxt_ny = datetime.fromtimestamp(nxt, tz=timezone.utc).astimezone(NY)
    print(f"\nProchaine bougie clôturée traitée à {nxt_ny.strftime('%H:%M:%S')} NY (dans {delay:.0f} s).")
    time.sleep(delay)


def _sleep_until_session():
    """Hors séance d'alerte (bougies 9h30-14h00 NY, jours de semaine) : dort jusqu'au prochain créneau utile.
    Dors par tranches d'1 h max : le service reste réactif aux redémarrages Railway."""
    now = datetime.now(NY)
    target = now.replace(hour=9, minute=33, second=0, microsecond=0)   # réveil juste avant la 1re bougie utile (9h30)
    if now >= target:
        target += timedelta(days=1)
    while target.weekday() >= 5:                                       # samedi / dimanche -> lundi
        target += timedelta(days=1)
    delay = max(60.0, (target - now).total_seconds())
    print(f"Hors séance (calculs limités à 9h30-14h00 NY, lun-ven) : sommeil jusqu'à "
          f"{target.strftime('%A %d/%m %H:%M')} NY ({delay / 3600:.1f} h).")
    time.sleep(min(delay, 3600.0))


def live_loop():
    """Mode live : à chaque clôture de bougie -> Excel à jour + alertes sur les valeurs recalculées + push GitHub.
    Une erreur (réseau, API...) est loguée puis la boucle reprend : service H24 sans supervision."""
    global LIVE_MODE
    LIVE_MODE = True
    print(f"MODE LIVE : tick toutes les {BAR_MINUTES} min (+{LIVE_SETTLE_S} s de stabilisation). "
          f"Alertes locales : {'activées' if LIVE_ALERTS else 'désactivées (le bot les gère via le JSON publié)'}. "
          f"Fenêtre : {'séance 9h30-14h00 NY uniquement (mode allégé)' if LIVE_SESSION_ONLY else 'journée complète'}. "
          f"Historique : {HIST_DIR}")
    _client()          # vérifie les clés Alpaca au démarrage : si elles manquent, le service s'arrête tout de suite (visible dans les logs)
    _git_setup()
    _bootstrap_history()   # premier démarrage : récupère les Excel quotidiens du dépôt pour les références D / E / F
    while True:
        try:
            day = datetime.now(NY).date().isoformat()
            _wait_next_tick()
            if LIVE_SESSION_ONLY:
                now_ny = datetime.now(NY)
                mins = now_ny.hour * 60 + now_ny.minute
                # ticks utiles : 9h35 (bougie 9h30) -> 14h05/14h06 (bougie 14h00), jours de semaine uniquement
                if now_ny.weekday() >= 5 or not (9 * 60 + 35 <= mins <= 14 * 60 + 6):
                    _sleep_until_session()
                    continue
            t0 = time.time()
            out, data = run_day(day, list(FEEDS))
            last_bar = None
            feed0 = list(data)[0] if data else None
            del data                                   # libère la RAM AVANT le recalcul LibreOffice
            import gc
            gc.collect()                               # (plan gratuit = 1 Go : données + soffice ne doivent pas coexister)
            if feed0:
                res = _extract_last_bar_values(out, feed0, day)     # recalcul LibreOffice : une seule fois par tick
                if res:
                    ny_iso, vals = res
                    last_bar = (ny_iso, feed0, vals)
                    if LIVE_ALERTS:
                        _fire_alerts(ny_iso, vals)
            _git_push_tick(out, day, last_bar)
            print(f"Tick terminé en {time.time() - t0:.0f} s.")
        except SystemExit as exc:
            print(f"[!] Tick sauté ({exc}) — reprise au prochain tick.")
            time.sleep(60)
        except Exception as exc:
            import traceback
            print(f"[!] Erreur dans le tick ({exc}) — reprise dans 60 s.")
            traceback.print_exc()
            time.sleep(60)


# ----------------------------------------------------------------- main
def main():
    if len(sys.argv) > 1 and sys.argv[1].lower() == "live":
        live_loop()                                        # boucle H24 : ne rend jamais la main
        return
    day = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DAY
    feeds = list(FEEDS)
    if len(sys.argv) > 2:
        choice = sys.argv[2].upper()
        feeds = list(FEEDS) if choice in ("BOTH", "ALL") else [choice]
        if any(f not in ("IEX", "SIP") for f in feeds):
            raise SystemExit("Deuxième argument : iex, sip ou both.")
    run_day(day, feeds)                                    # mode batch : comportement d'origine, inchangé


if __name__ == "__main__":
    main()
