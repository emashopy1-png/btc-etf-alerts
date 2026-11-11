import sys, openpyxl
from openpyxl.utils import column_index_from_string as ci
ALL = ["IBIT","FBTC","BITB","ARKB","BTCO","EZBC","BRRR","HODL","BTCW","MSBT","GBTC","BTC"]
def num(x): return isinstance(x,(int,float)) and not isinstance(x,bool)
def run(f):
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    conf = list(wb["CONFLUENCE_IEX"].iter_rows(min_row=2, values_only=True))
    conf = [r for r in conf if r[0] is not None]
    F = [r[5] for r in conf]; ER=[r[ci("ER")-1] for r in conf]
    CL={}; CM={}
    for s in ALL:
        rows=list(wb[f"{s}_IEX"].iter_rows(min_row=2, values_only=True))
        rows=[r for r in rows if r[0] is not None][:len(conf)]
        CL[s]=[r[ci("CL")-1] for r in rows]; CM[s]=[r[ci("CM")-1] for r in rows]
    # règle : ETF inclus si sa réf. tape (CM) est numérique et > 0
    inc_rule=[s for s in ALL if any(num(v) and v>0 for v in CM[s])]
    print(f.split("_")[-4], "ETF avec réf. tape:", len(inc_rule), "exclus:", [s for s in ALL if s not in inc_rule])
    def calc(inc):
        bad=[];ok=0
        for i,(fv,er) in enumerate(zip(F,ER)):
            if not num(er) or er==0:
                if num(fv): bad.append((conf[i][0],fv,None))
                else: ok+=1
                continue
            v=sum((x if num(x) else 0) for x in (CL[s][i] for s in inc))/er
            if num(fv) and abs(v-fv)<1e-9: ok+=1
            elif not num(fv) and False: ok+=1
            else: bad.append((conf[i][0],fv,v))
        return ok,bad
    for name,inc in [("ancien (sans BTCO/BTCW)",[s for s in ALL if s not in("BTCO","BTCW")]),("règle CM",inc_rule)]:
        ok,bad=calc(inc); print(f"   {name}: {ok}/{len(F)} identiques, écarts: {bad[:5]}")
for f in sys.argv[1:]: run(f)
