#!/usr/bin/env python3
"""Build data.json from Hunter_Handicap_Tracker.xlsx.
Usage: python3 build_data.py [tracker.xlsx]   (default: /workspace/hcp/final/Hunter_Handicap_Tracker.xlsx)
- Rounds come from '18-Hole Rounds', '9-Hole Rounds', 'Partial Rounds' (Holes column). The round type is the number of
  holes actually played: scored rows in 'Holes' for that RoundId when present, else the sheet's hole count. 18=full, 9=nine, other=partial.
- Up-and-downs and per-round counts are derived from 'Holes' when the sheet cell is blank; a WARN is printed on mismatch.
- 'GHIN Scores' is required (score history, differential, rating/slope). Score to par uses the linked round, else pars.py."""
import sys, os, json, openpyxl, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pars import par_for

DEFAULT_SRC = '/workspace/hcp/final/Hunter_Handicap_Tracker.xlsx'
# Rounds logged incorrectly: kept out of every chart. Key = (YYYY-MM-DD, lowercase course substring)
EXCLUDED_ROUNDS = []  # e.g. ('2026-06-20', 'jacksonville beach')
# Course / tee aliases so tracker and GHIN names line up (applied after strip)
COURSE_ALIAS = {'Queens Harbour': 'Queens Harbour Yacht & Country Club', 'Talbot': 'Talbot Country Club',
                'TPC Sawgrass STADIUM COURSE': 'TPC Sawgrass - STADIUM COURSE', 'Soldier Hollow': 'Soldier Hollow - Gold Course',
                'Columbia CC': 'Columbia Country Club'}
TEE_ALIAS = {'Blue / Blue': 'Blue', 'Black (Gold Crs)': 'Black'}

def excluded(date, course): return any(date == dt and c in (course or '').lower() for dt, c in EXCLUDED_ROUNDS)
def d(v): return v.strftime('%Y-%m-%d') if isinstance(v, (datetime.date, datetime.datetime)) else v
def num(v): return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None
def course(c): c = c.strip() if isinstance(c, str) else c; return COURSE_ALIAS.get(c, c)
def tee(t):
    if not isinstance(t, str): return None
    t = t.strip(); return TEE_ALIAS.get(t, t).title()
warns = []
def warn(m): warns.append(m)

src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SRC
wb = openpyxl.load_workbook(src, data_only=False)
if 'GHIN Scores' not in wb.sheetnames:
    sys.exit(f'ERROR: {src} has no "GHIN Scores" sheet - wrong/old tracker file. Pass the current tracker path.')
def table(name):
    ws = wb[name]; H = [c.value for c in ws[1]]; ix = {h: i for i, h in enumerate(H) if h}
    return ix, [r for r in ws.iter_rows(min_row=2, values_only=True)]

# ---- holes (with derived up-and-downs)
holes = []; ix, rows = table('Holes')
for r in rows:
    v = lambda k: r[ix[k]] if k in ix else None
    if not v('RoundDate'): continue
    par, sc = num(v('Par')), num(v('Score'))
    gir = v('GIR'); ua, um = num(v('UpDownAttempt')), num(v('UpDownMade'))
    if gir in ('Hit', 'Miss') and sc is not None:
        da = 1 if gir == 'Miss' else 0; dm = 1 if da and par and sc <= par else 0
        if ua is None: ua = da
        elif ua != da: warn(f'{v("RoundId")} H{v("Hole")}: UpDownAttempt {ua} vs derived {da}')
        if um is None: um = dm
        elif um != dm: warn(f'{v("RoundId")} H{v("Hole")}: UpDownMade {um} vs derived {dm}')
    holes.append(dict(id=v('RoundId'), date=d(v('RoundDate')), course=course(v('Course')), tees=tee(v('Tees')), type=v('RoundType'), hole=v('Hole'),
        par=par, score=sc, toPar=(sc - par) if par and sc is not None else None, fir=v('FIR'), fwMiss=v('FairwayMiss'), gir=gir, grMiss=v('GreenMiss'),
        putts=num(v('Putts')), udAtt=ua, udMade=um))
byid = collections.defaultdict(list)
for h in holes:
    if h['id']: byid[h['id']].append(h)

def derive(hs):
    """Per-round counts from hole rows; a count is returned only when every scored hole has the needed field."""
    s = [h for h in hs if h['score'] is not None]; out = {}
    if not s: return out
    if all(h['par'] for h in s):
        tp = [h['toPar'] for h in s]; out['doubles'] = sum(x == 2 for x in tp); out['triplesPlus'] = sum(x >= 3 for x in tp)
    if all(h['gir'] in ('Hit', 'Miss') for h in s):
        out['gir'] = sum(h['gir'] == 'Hit' for h in s); out['udAtt'] = sum(h['udAtt'] or 0 for h in s); out['udMade'] = sum(h['udMade'] or 0 for h in s)
    if all(h['putts'] is not None for h in s):
        p = [h['putts'] for h in s]; out['putts'] = sum(p); out['threePutts'] = sum(x >= 3 for x in p); out['onePutts'] = sum(x == 1 for x in p)
    f = [h for h in s if (h['par'] or 0) >= 4]
    if f and all(h['fir'] in ('Hit', 'Miss') for h in f):
        out['fwHit'] = sum(h['fir'] == 'Hit' for h in f); out['fwPossible'] = len(f)
    return out

# ---- holes actually played per RoundId, from GHIN Scores HolesPlayed (GHIN '83(17)' notation)
gx0, grows0 = table('GHIN Scores')
ghin_played = {r[gx0['RoundId']]: num(r[gx0['HolesPlayed']]) for r in grows0 if 'HolesPlayed' in gx0 and r[gx0['RoundId']] and num(r[gx0['HolesPlayed']])}
# ---- rounds
COLS = {'doubles': 'Doubles (exact +2)', 'triplesPlus': 'TriplesOrWorse (+3 or more)', 'fwHit': 'FairwaysHit', 'fwPossible': 'FairwaysPossible',
        'gir': 'GIR', 'putts': 'Putts', 'threePutts': 'ThreePuttsPlus', 'onePutts': 'OnePutts', 'udMade': 'UpDownsMade', 'udAtt': 'UpDownsAttempted'}
rounds = []
for sn, default_type in (('18-Hole Rounds', 18), ('9-Hole Rounds', 9), ('Partial Rounds', None)):
    if sn not in wb.sheetnames: continue
    ix, rows = table(sn)
    for r in rows:
        if not r[ix['Date']]: continue
        g = lambda k: num(r[ix[k]]) if k in ix else None
        rid = r[ix['RoundId']] if 'RoundId' in ix else None
        hs = byid.get(rid, [])
        scored = sum(h['score'] is not None for h in hs)
        played = ghin_played.get(rid) or default_type or int(g('Holes'))
        if default_type and played != default_type: warn(f'{rid}: on {sn} but GHIN says {played} holes played -> treated as {played}-hole (move it to Partial Rounds)')
        if hs and scored != played: warn(f'{rid}: {scored} scored hole rows vs {played} holes played (hole data incomplete)')
        par, gross = g('Par'), g('Gross')
        rd = dict(date=d(r[ix['Date']]), course=course(r[ix['Course']]), tees=tee(r[ix['Tees']]), type=played, teeYards=g('TeeYards'), rating=g('Rating'),
                  slope=g('Slope'), differential=g('Differential'), par=par, gross=gross, id=rid)
        dv = derive(hs) if hs and scored == played else {}
        if hs and played != (default_type or played) and scored == played:  # partial played on a full sheet: par/gross from scored holes
            s = [h for h in hs if h['score'] is not None]; rd['par'] = sum(h['par'] for h in s); rd['gross'] = sum(h['score'] for h in s)
        rd['toPar'] = rd['gross'] - rd['par'] if rd['par'] and rd['gross'] else None
        for k, col in COLS.items():
            sv = g(col)
            if sv is None: sv = dv.get(k)
            elif k in dv and dv[k] != sv: warn(f'{rid}: {col} sheet {sv} vs holes {dv[k]} (kept sheet)')
            rd[k] = sv
        rounds.append(rd)
# holes take their round's played count as type
rtype = {r['id']: r['type'] for r in rounds if r['id']}
for h in holes: h['type'] = rtype.get(h['id'], h['type'])

# ---- handicap
hcp = []
ws = wb['Handicap Tracker']
for r in ws.iter_rows(min_row=8, max_col=2, values_only=True):
    if r[0] and num(r[1]) is not None: hcp.append(dict(date=d(r[0]), index=r[1]))

# ---- GHIN score history
rmap = {r['id']: r for r in rounds if r.get('id')}
ghin = []; gx, rows = table('GHIN Scores')
for r in rows:
    if not r[0] or num(r[gx['Differential']]) is None: continue
    g = lambda k: r[gx[k]] if k in gx else None
    rid = g('RoundId'); lr = rmap.get(rid, {})
    if rid and not lr: warn(f'GHIN Scores RoundId {rid} has no round row')
    hl = lr.get('type') or num(g('HolesPlayed')) or int(g('Holes')); sc = num(g('Score'))
    par = lr.get('par') if lr else par_for(g('Course'), hl)
    tp = lr.get('toPar') if lr else ((sc - par) if par and sc else None)
    ghin.append(dict(date=d(r[0]), course=course(g('Course')), tees=tee(g('Tees')), type=hl, score=sc, differential=g('Differential'), id=rid,
        rating=num(g('Rating')) if num(g('Rating')) is not None else lr.get('rating'), slope=num(g('Slope')) if num(g('Slope')) is not None else lr.get('slope'),
        teeYards=lr.get('teeYards'), par=par, toPar=tp))

ghin = [x for x in ghin if not excluded(x['date'], x['course'])]
rounds = [x for x in rounds if not excluded(x['date'], x['course'])]
holes = [x for x in holes if not excluded(x['date'], x['course'])]
out = dict(ghin=ghin, generated=datetime.datetime.now().isoformat(timespec='minutes'), rounds=rounds, holes=holes, handicap=hcp)
json.dump(out, open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data.json'), 'w'), indent=1)
for w in warns: print('WARN', w)
print(len(ghin), 'ghin', sum(x['toPar'] is not None for x in ghin), 'with toPar', sum(x['rating'] is not None for x in ghin), 'with rating |',
      len(rounds), 'rounds', collections.Counter(r['type'] for r in rounds), '|', len(holes), 'holes |', len(hcp), 'hcp')
