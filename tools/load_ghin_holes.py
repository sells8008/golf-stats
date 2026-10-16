#!/usr/bin/env python3
"""Load GHIN hole-by-hole data into the tracker 'Holes' sheet and fill per-round counts.
CSV columns: RoundId,RoundDate(MM/DD/YYYY),Course,Tees,RoundType(holes played),Hole,Par,Score,FIR(hit/miss/NA),FairwayMiss,GIR(hit/miss),GreenMiss,Putts
 - RoundId must match the round row (18-Hole / 9-Hole / Partial Rounds sheet). Existing Holes rows for that RoundId are replaced.
 - Full-stats round (any putts recorded): blank GIR on a scored hole = Miss/'No Chance'; blank FIR on a par 4/5 = Miss/'No Chance'.
 - Up-and-downs derived per hole: attempt = GIR Miss (incl. No Chance); made = attempt and score <= par.
 - Per-round counts written only when every scored hole has the field.
Usage: python3 tools/load_ghin_holes.py holes.csv tracker.xlsx [out.xlsx]   (then run tools/rebuild_fairway_holes.py)"""
import sys, csv, datetime, collections, openpyxl
sys.path.insert(0, __import__('os').path.dirname(__file__))
from rebuild_fairway_holes import rebuild
CAP = {'hit': 'Hit', 'miss': 'Miss', 'na': 'NA'}
def main(csvp, xl, out):
    wb = openpyxl.load_workbook(xl); H = wb['Holes']; hc = {h.value: i + 1 for i, h in enumerate(H[1])}
    by = collections.OrderedDict()
    for r in csv.DictReader(open(csvp)): by.setdefault(r['RoundId'], []).append(r)
    for rid in by:  # replace existing rows for these rounds
        for i in range(H.max_row, 1, -1):
            if H.cell(i, hc['RoundId']).value == rid: H.delete_rows(i)
    for rid, hs in by.items():
        full = any(h['Putts'] for h in hs)
        for h in hs:
            par = int(h['Par']); sc = int(h['Score']) if h['Score'] else None
            fir = CAP.get((h['FIR'] or '').lower()); fm = h['FairwayMiss'] or None; gir = CAP.get((h['GIR'] or '').lower()); gm = h['GreenMiss'] or None
            if full and sc is not None:
                if not gir: gir, gm = 'Miss', 'No Chance'
                if par >= 4 and not fir: fir, fm = 'Miss', 'No Chance'
            if par == 3: fir = 'NA'
            att = (1 if gir == 'Miss' else 0) if gir and sc is not None else None
            made = (1 if att and sc <= par else 0) if att is not None else None
            n = H.max_row + 1
            H.append([rid, datetime.datetime.strptime(h['RoundDate'], '%m/%d/%Y'), h['Course'], h['Tees'], int(h['RoundType']), int(h['Hole']), par, sc,
                      f'=IF(COUNT(G{n}:H{n})=2,H{n}-G{n},"")', fir, fm, gir, gm, int(h['Putts']) if h['Putts'] else None, att, made])
            H.cell(n, 2).number_format = 'yyyy-mm-dd'
    # per-round counts
    rows = [r for r in H.iter_rows(min_row=2, values_only=True)]; col = {k: v - 1 for k, v in hc.items()}
    for sn in ('18-Hole Rounds', '9-Hole Rounds', 'Partial Rounds'):
        if sn not in wb.sheetnames: continue
        S = wb[sn]; sc_ = {h.value: i + 1 for i, h in enumerate(S[1])}
        for i in range(2, S.max_row + 1):
            rid = S.cell(i, sc_['RoundId']).value
            if rid not in by: continue
            hs = [r for r in rows if r[col['RoundId']] == rid and r[col['Score']] is not None]
            put = lambda k, v: S.cell(i, sc_[k], v)
            tp = [r[col['Score']] - r[col['Par']] for r in hs]
            put('Doubles (exact +2)', sum(x == 2 for x in tp)); put('TriplesOrWorse (+3 or more)', sum(x >= 3 for x in tp))
            if all(r[col['GIR']] in ('Hit', 'Miss') for r in hs):
                put('GIR', sum(r[col['GIR']] == 'Hit' for r in hs)); put('UpDownsAttempted', sum(r[col['UpDownAttempt']] or 0 for r in hs)); put('UpDownsMade', sum(r[col['UpDownMade']] or 0 for r in hs))
            if all(r[col['Putts']] is not None for r in hs):
                p = [r[col['Putts']] for r in hs]; put('Putts', sum(p)); put('ThreePuttsPlus', sum(x >= 3 for x in p)); put('OnePutts', sum(x == 1 for x in p))
            f = [r for r in hs if r[col['Par']] >= 4]
            if f and all(r[col['FIR']] in ('Hit', 'Miss') for r in f):
                put('FairwaysHit', sum(r[col['FIR']] == 'Hit' for r in f)); put('FairwaysPossible', len(f))
            print('round', rid, len(hs), 'holes')
    print(rebuild(wb), 'fairway rows'); wb.save(out)
if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else sys.argv[2])
