#!/usr/bin/env python3
"""Backfill GHIN Scores Rating / Slope / HolesPlayed from a GHIN Score History dump.
Dump lines (pipe-separated, as captured from GHIN Score History): <year|Rece>|<score text e.g. 83(17)Ai>|MM/DD/YYYY|<rating>/<slope>|<pcc>|<differential>|<course...>
'(n)' in the score text = holes actually played (GHIN posts partial rounds as 18 with net par for the rest).
Rows are matched on date + score + differential; nothing is filled when there is no unique match.
Usage: python3 tools/ghin_scores_meta.py dump.txt tracker.xlsx [out.xlsx]"""
import sys, re, datetime, openpyxl, copy
def main(dump, xl, out):
    recs = []
    for line in open(dump):
        p = line.rstrip('\n').split('|')
        if len(p) < 6: continue
        m = re.match(r'\*?\s*(\d+)(?:\((\d+)\))?', p[1].strip()); rs = re.match(r'([\d.]+)/(\d+)', p[3])
        if not m or not rs: continue
        recs.append(dict(dt=datetime.datetime.strptime(p[2], '%m/%d/%Y'), score=int(m.group(1)), played=int(m.group(2)) if m.group(2) else None,
                         diff=float(p[5]), rating=float(rs.group(1)), slope=int(rs.group(2))))
    wb = openpyxl.load_workbook(xl); ws = wb['GHIN Scores']; H = [c.value for c in ws[1]]
    for k in ('Rating', 'Slope', 'HolesPlayed'):
        if k not in H:
            H.append(k); c = ws.cell(1, len(H), k); c._style = copy.copy(ws.cell(1, 1)._style)
    ci = {h: i + 1 for i, h in enumerate(H)}; filled = blank = 0
    for i in range(2, ws.max_row + 1):
        dt = ws.cell(i, 1).value
        if not dt: continue
        sc, df = ws.cell(i, ci['Score']).value, ws.cell(i, ci['Differential']).value
        m = [r for r in recs if r['dt'] == dt and r['score'] == sc and abs(r['diff'] - float(df)) < 0.05]
        if len(m) == 1:
            ws.cell(i, ci['Rating'], m[0]['rating']); ws.cell(i, ci['Slope'], m[0]['slope'])
            if m[0]['played']: ws.cell(i, ci['HolesPlayed'], m[0]['played'])
            filled += 1
        elif ws.cell(i, ci['Rating']).value is None: blank += 1
    wb.save(out); print('rating/slope filled', filled, '; still blank', blank)
if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else sys.argv[2])
