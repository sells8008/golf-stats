#!/usr/bin/env python3
"""Build data.json from Hunter_Handicap_Tracker.xlsx. Usage: python3 build_data.py [xlsx]"""
import sys, json, openpyxl, datetime
src = sys.argv[1] if len(sys.argv) > 1 else '/workspace/hcp/Hunter_Handicap_Tracker_p1.xlsx'
wb = openpyxl.load_workbook(src, data_only=False)
def d(v): return v.strftime('%Y-%m-%d') if isinstance(v, (datetime.date, datetime.datetime)) else v
def num(v): return v if isinstance(v, (int, float)) else None
def tee(t): return t.strip().title() if isinstance(t, str) else None
def course(c):
    if c == 'Queens Harbour': return 'Queens Harbour Yacht & Country Club'
    return c
rounds = []
for sn, rt in (('18-Hole Rounds', 18), ('9-Hole Rounds', 9)):
    ws = wb[sn]; H = [c.value for c in ws[1]]
    ix = {h: i for i, h in enumerate(H)}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]: continue
        g = lambda k: num(r[ix[k]]) if k in ix else None
        par, gross = g('Par'), g('Gross')
        rounds.append(dict(date=d(r[0]), course=course(r[1]), tees=tee(r[2]), type=rt, teeYards=g('TeeYards'),
            par=par, gross=gross, toPar=(gross - par) if par and gross else None,
            doubles=g('Doubles (exact +2)'), triplesPlus=g('TriplesOrWorse (+3 or more)'),
            fwHit=g('FairwaysHit'), fwPossible=g('FairwaysPossible'),
            gir=g('GIR'), putts=g('Putts'), threePutts=g('ThreePuttsPlus'), onePutts=g('OnePutts'),
            udMade=g('UpDownsMade'), udAtt=g('UpDownsAttempted')))
holes = []
ws = wb['Holes']
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r[0]: continue
    par, sc = num(r[5]), num(r[6])
    holes.append(dict(date=d(r[0]), course=course(r[1]), tees=tee(r[2]), type=r[3], hole=r[4], par=par, score=sc,
        toPar=(sc - par) if par and sc else None, fir=r[8], fwMiss=r[9], gir=r[10], grMiss=r[11],
        putts=num(r[12]), udAtt=r[13], udMade=r[14]))
hcp = []
ws = wb['Handicap Tracker']
for r in ws.iter_rows(min_row=8, max_col=2, values_only=True):
    if r[0] and num(r[1]) is not None: hcp.append(dict(date=d(r[0]), index=r[1]))
out = dict(generated=datetime.datetime.now().isoformat(timespec='minutes'), rounds=rounds, holes=holes, handicap=hcp)
json.dump(out, open('data.json', 'w'), indent=1)
print(len(rounds), 'rounds', len(holes), 'holes', len(hcp), 'hcp')
