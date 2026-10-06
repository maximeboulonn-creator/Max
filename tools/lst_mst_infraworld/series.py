"""Séries mensuelles pour le MST : composites cotés (même règle que le classeur des cours) et change."""
import json,datetime as dt,math
P='/tmp/claude-0/-home-user/fbfd82fa-9176-5404-b298-69879606bee1/scratchpad/prices/'
TICK=[('IGF','C1'),('INFR.L','C1'),('3IN.L','C1'),('HICL.L','C1'),('INPP.L','C1'),('SEQI.L','C2'),('GCP.L','C2'),('BIP',None),('088980.KS','C1'),('BKLN','C4'),('HYG','C4'),('IHYG.L','C4'),('ARCC','C3'),('BIZD','C3')]
STATS_FROM={'INFR.L':dt.date(2016,1,1)}
def load(fn):
    j=json.load(open(fn))['chart']['result'][0]; meta=j['meta']; ts=j['timestamp']; q=j['indicators']['quote'][0]
    adj=j['indicators'].get('adjclose',[{}])[0].get('adjclose',q['close']); off=meta.get('gmtoffset',0)
    return meta,[(dt.datetime.fromtimestamp(t+off,dt.UTC).date(),c,a) for t,c,a in zip(ts,q['close'],adj) if c is not None and a is not None]
def monthly_returns():
    cutoff=(2026,10)
    rets={}  # ticker -> {(y,m): simple return}
    for sym,comp in TICK:
        meta,rows=load(P+f'm_{sym}.json')
        rows=[r for r in rows if (r[0].year,r[0].month)!=cutoff]
        st=STATS_FROM.get(sym)
        if st: rows=[r for r in rows if r[0]>=st]
        d={}
        for i in range(1,len(rows)):
            a,b=rows[i][2],rows[i-1][2]
            if a>0 and b>0 and abs(math.log(a/b))<=0.6: d[(rows[i][0].year,rows[i][0].month)]=a/b-1
        rets[sym]=d
    months=sorted(set().union(*[set(d) for d in rets.values()]))
    comp={}
    for c in ('C1','C2','C3','C4'):
        syms=[s for s,k in TICK if k==c]
        ser={}
        for ym in months:
            v=[rets[s][ym] for s in syms if ym in rets[s]]
            if v: ser[ym]=sum(v)/len(v)
        comp[c]=ser
    fx={}
    for sym in ('EURUSD',):
        meta,rows=load(P+f'm_{sym}.json'); rows=[r for r in rows if (r[0].year,r[0].month)!=cutoff]
        d={}
        for i in range(1,len(rows)): d[(rows[i][0].year,rows[i][0].month)]=rows[i][1]/rows[i-1][1]-1
        fx[sym]=d
    return months,comp,fx
if __name__=='__main__':
    m,c,f=monthly_returns(); print(len(m),m[0],m[-1],{k:len(v) for k,v in c.items()},len(f['EURUSD']))
