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
 ('jekyll island golf club - great dunes', 72, None, None),  # nines differ between published cards
 ('windy harbor', 72, 36, 36),
 ('rockaway hunting club', 70, 35, 35),  # Blue 2023 72.6/133 (BlueGolf)
 ('the windsor club', 72, 36, 36),
 ("john's island club - west", 72, 36, 36),
 ('bobby jones golf club', 71, None, None),  # Sarasota Championship, Black 70.7/131 (FSGA/GolfPass)
 ('forest park golf course - redbud/dogwood', 70, 35, 35),
 ('whiskey creek golf club', 72, 36, 36),
 ('the park', 71, None, None),  # The Park West Palm, Black 71.9/134 (BlueGolf/FSGA)
 ('lpga international - hills', 72, 36, 36),
 ('gran coyote', 71, 35, 36),
 ('el camaleon mayakoba', 72, 36, 36),
 ('innisbrook resort & golf club - copperhead', 71, 36, 35),
 ('the river club golf course', 72, 36, 36),  # Suwanee GA
 ('pasadena yacht', 72, 36, 36),
 ('sand canyon country club - mountain/valley', 71, 35, 36),
 ('oak quarry', 72, 35, 37),
 ('los verdes', 71, 36, 35),
 ('washington golf & country club', 70, 35, 35),
 ('golf club at north hampton', 72, 36, 36),
 ('cobblestone park golf club - gold / garnet', 70, 35, 35),  # Gold 35 + Garnet 35
 ('links at perry cabin', 72, 35, 37),
 ('plantation lakes', 72, 36, 36),
 ('georgia southern university', 72, 36, 36),
 ('rivertowne', 72, 36, 36),
 ('kiawah island golf resort - cougar point', 72, 36, 36),
 ('stono ferry', 72, 36, 36),
 ('windsor parke', 72, 36, 36),
 ('championsgate golf resort - international', 72, None, None),
 ('king & bear', 72, 36, 36),
 ('shingle creek', 72, 36, 36),
 ('oak marsh', 72, 36, 36),
 ('ponte vedra inn & club - ocean', 71, 36, 35),
 ('angel park golf club - palm', 70, 34, 36),
 ('east potomac golf links - blue', 72, None, None),
 ("myrtle beach national - king's north", 72, 36, 36),
 ('golf club at south hampton', 72, 36, 36),
 ('arrowhead country club - cypress / lakes', 72, 36, 36),
 ('clearwater country club', 72, 36, 36),
 ('chevy chase club', 71, 36, 35),  # source: Hunter (course recently redone)
]
def par_for(course, holes):
    c = (course or '').lower()
    for k, p18, f, b in PARS:
        if k in c:
            if holes == 18: return p18
            if f is not None and f == b: return f   # 9 holes: only when both nines share a par
            return None
    return None
