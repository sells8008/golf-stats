#!/usr/bin/env python3
"""Build data.json from Hunter_Handicap_Tracker.xlsx. Usage: python3 build_data.py [xlsx]"""
import sys, os, json, openpyxl, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pars import par_for
# Rounds Hunter says were logged incorrectly: kept out of every chart. Key = (YYYY-MM-DD, course substring, lowercase)
EXCLUDED_ROUNDS = []  # e.g. ('2026-06-20', 'jacksonville beach'); that round is now a 13-hole partial instead
def excluded(date, course): return any(date == dt and c in (course or '').lower() for dt, c in EXCLUDED_ROUNDS)
src = sys.argv[1] if len(sys.argv) > 1 else '/workspace/hcp/Hunter_Handicap_Tracker_p1.xlsx'
wb = openpyxl.load_workbook(src, data_only=False)
def d(v): return v.strftime('%Y-%m-%d') if isinstance(v, (datetime.date, datetime.datetime)) else v
def num(v): return v if isinstance(v, (int, float)) else None
def tee(t): return t.strip().title() if isinstance(t, str) else None
def course(c):
    if c == 'Queens Harbour': return 'Queens Harbour Yacht & Country Club'
    return c
rounds = []
for sn, rt in (('18-Hole Rounds', 18), ('9-Hole Rounds', 9), ('Partial Rounds', None)):
    if sn not in wb.sheetnames: continue
    ws = wb[sn]; H = [c.value for c in ws[1]]
    ix = {h: i for i, h in enumerate(H)}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[ix['Date']]: continue
        g = lambda k: num(r[ix[k]]) if k in ix else None
        par, gross = g('Par'), g('Gross')
        rounds.append(dict(date=d(r[ix['Date']]), course=course(r[ix['Course']]), tees=tee(r[ix['Tees']]), type=rt or int(r[ix['Holes']]), teeYards=g('TeeYards'), rating=g('Rating'), differential=g('Differential'),
            par=par, gross=gross, toPar=(gross - par) if par and gross else None,
            doubles=g('Doubles (exact +2)'), triplesPlus=g('TriplesOrWorse (+3 or more)'),
            fwHit=g('FairwaysHit'), fwPossible=g('FairwaysPossible'),
            gir=g('GIR'), putts=g('Putts'), threePutts=g('ThreePuttsPlus'), onePutts=g('OnePutts'),
            udMade=g('UpDownsMade'), udAtt=g('UpDownsAttempted'), id=r[ix['RoundId']] if 'RoundId' in ix else None))
holes = []
ws = wb['Holes']; HH = [c.value for c in ws[1]]; hx = {h: i for i, h in enumerate(HH)}
for r in ws.iter_rows(min_row=2, values_only=True):
    v = lambda k: r[hx[k]] if k in hx else None
    if not v('RoundDate'): continue
    par, sc = num(v('Par')), num(v('Score'))
    holes.append(dict(id=v('RoundId'), date=d(v('RoundDate')), course=course(v('Course')), tees=tee(v('Tees')), type=v('RoundType'), hole=v('Hole'),
        par=par, score=sc, toPar=(sc - par) if par and sc else None, fir=v('FIR'), fwMiss=v('FairwayMiss'), gir=v('GIR'), grMiss=v('GreenMiss'),
        putts=num(v('Putts')), udAtt=v('UpDownAttempt'), udMade=v('UpDownMade')))
hcp = []
ws = wb['Handicap Tracker']
for r in ws.iter_rows(min_row=8, max_col=2, values_only=True):
    if r[0] and num(r[1]) is not None: hcp.append(dict(date=d(r[0]), index=r[1]))
rmap = {r['id']: r for r in rounds if r.get('id')}
ghin = []
if 'GHIN Scores' in wb.sheetnames:
    ws = wb['GHIN Scores']; GH = [c.value for c in ws[1]]; gx = {h: i for i, h in enumerate(GH)}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0] or num(r[gx['Differential']]) is None: continue
        rid = r[gx['RoundId']]; lr = rmap.get(rid, {})
        hl = int(r[gx['Holes']]); sc = num(r[gx['Score']])
        # Score to par: tracker round (gross - par) when linked, else GHIN score - verified course par (pars.py); None if par unknown
        par = lr.get('par') or par_for(r[gx['Course']], hl)
        tp = lr.get('toPar') if lr.get('toPar') is not None else ((sc - par) if par and sc else None)
        ghin.append(dict(date=d(r[0]), course=course(r[gx['Course']]), tees=tee(r[gx['Tees']]), type=lr.get('type') or hl, score=sc,
            differential=r[gx['Differential']], id=rid, rating=lr.get('rating'), teeYards=lr.get('teeYards'), par=par, toPar=tp))
ghin = [g for g in ghin if not excluded(g['date'], g['course'])]
rounds = [r for r in rounds if not excluded(r['date'], r['course'])]
holes = [h for h in holes if not excluded(h['date'], h['course'])]
out = dict(ghin=ghin, generated=datetime.datetime.now().isoformat(timespec='minutes'), rounds=rounds, holes=holes, handicap=hcp)
json.dump(out, open('data.json', 'w'), indent=1)
print(len(ghin), 'ghin', sum(g['toPar'] is not None for g in ghin), 'with toPar', len(rounds), 'rounds', len(holes), 'holes', len(hcp), 'hcp')
