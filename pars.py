"""Verified course pars (men's tees) for GHIN score history. Key = lowercase substring of GHIN course name.
(par18, front9, back9). Sources: public scorecards (golflink/bluegolf/golfpass/club sites), checked Oct 2026.
Add a course here only after verifying its scorecard. Never guess."""
PARS = [
 ('jacksonville beach golf club', 71, 35, 36),
 ('queens harbour', 72, 36, 36),
 ('talbot', 72, None, None),
 ('cimarrone', 72, None, None),
 ('club at savannah harbor', 72, None, None),
 ('tpc sawgrass', 72, None, None),
 ('canyons golf', 70, None, None),
 ('park city gc', 72, None, None),
 ('soldier hollow', 72, None, None),
 ('riomar', 70, 36, 34),
 ('columbia c', 70, 35, 35),
 ('quail valley', 72, 36, 36),
 ('hyde park golf club', 72, 36, 36),
 ('golf club at fleming island', 71, 36, 35),
 ('stillwater golf', 72, 36, 36),
 ('rookery golf club', 71, None, None),
 ('atlantic beach country club', 71, None, None),
 ('blue sky golf club', 72, 36, 36),
 ('hidden hills golf club', 72, 36, 36),
 ('bear trap dunes', 72, None, None),
 ('ponte vedra inn & club - lagoon', 70, None, None),
 ('marsh landing', 72, 36, 36),
 ('st. johns golf', 72, None, None),
 ('amelia river club', 72, 36, 36),
 ('lpga international - jones', 72, 36, 36),
 ('falls road', 70, 35, 35),
 ('hog neck', 72, None, None),
 ('blue mash', 71, 36, 35),
]
def par_for(course, holes):
    c = (course or '').lower()
    for k, p18, f, b in PARS:
        if k in c:
            if holes == 18: return p18
            if f is not None and f == b: return f   # 9 holes: only when both nines share a par
            return None
    return None
