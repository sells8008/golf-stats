#!/usr/bin/env python3
"""Rebuild the 'Fairway Holes' sheet (feeds Recent Rounds / All Time summary formulas) from 'Holes'.
Usage: python3 tools/rebuild_fairway_holes.py tracker.xlsx [out.xlsx]"""
import sys, openpyxl
def rebuild(wb):
    H = wb['Holes']; c = {h.value: i for i, h in enumerate(H[1])}
    rows = [r for r in H.iter_rows(min_row=2, values_only=True) if r[c['RoundDate']]]
    rows.sort(key=lambda r: (r[c['RoundDate']], r[c['RoundId']] or '', r[c['Hole']]))
    fh = wb['Fairway Holes']; fh.delete_rows(2, fh.max_row); i = 2
    for r in rows:
        par, sc, fir = r[c['Par']], r[c['Score']], r[c['FIR']]
        if par and par >= 4 and fir in ('Hit', 'Miss') and sc is not None:
            fh.cell(i, 1, r[c['RoundDate']]).number_format = 'yyyy-mm-dd'; fh.cell(i, 2, r[c['Course']]); fh.cell(i, 3, f"{r[c['RoundType']]}-hole")
            fh.cell(i, 4, r[c['Hole']]); fh.cell(i, 5, par); fh.cell(i, 6, sc); fh.cell(i, 7, f'=IF(COUNT(E{i}:F{i})=2,F{i}-E{i},"")'); fh.cell(i, 8, fir)
            fh.cell(i, 9, f'=IF(C{i}="9-hole",1,IF(COUNTIFS(\'18-Hole Rounds\'!$A$2:$A$1000,A{i},\'18-Hole Rounds\'!$R$2:$R$1000,1)>0,1,0))'); i += 1
    return i - 2
if __name__ == '__main__':
    wb = openpyxl.load_workbook(sys.argv[1]); n = rebuild(wb); wb.save(sys.argv[2] if len(sys.argv) > 2 else sys.argv[1]); print(n, 'fairway rows')
