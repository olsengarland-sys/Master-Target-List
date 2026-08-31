"""Wave-2 cross-reference dedupe + triage. Usage:
   python3 scripts/xref/triage_xref.py <raw_rows.json> <out.json>
Input: JSON list of rows {name,domain,hq,rev,emp,own,yr,url,desc,source}.
Output: {survivors:[...with priority/note], dropped_known:int, dropped_nodomain:int}
Same normalization and gates as wave 1 (scripts/retriage.py), plus Inven
ownership gates. Deterministic; agents must not re-implement this logic."""
import json, re, sys

K = json.load(open('/home/user/Master-Target-List/scripts/xref/known_keys.json'))
DK, NK = set(K['dkeys']), set(K['nkeys'])
LEGAL = r'\b(llc|l\.l\.c|inc|incorporated|corp|corporation|co|company|ltd|limited|lp|llp|pllc|pc|plc)\b'

def nd(u):
    if not u or not str(u).strip(): return None
    x = re.sub(r'^https?://', '', str(u).strip().lower()); x = re.sub(r'^www\.', '', x)
    return x.split('/')[0].split('?')[0].split('#')[0].strip() or None

def nn(n):
    if not n or not str(n).strip(): return None
    x = str(n).strip().lower().replace('&', ' and '); x = re.sub(r'[.,]', ' ', x)
    x = re.sub(LEGAL, ' ', x); x = re.sub(r'[^a-z0-9 ]', ' ', x)
    return re.sub(r'\s+', ' ', x).strip() or None

GATES = [
 (r'residential', 'residential'),
 (r'\bsolar\b', 'solar'),
 (r'home ?owner|home services', 'home services'),
 (r'fire alarm', 'fire alarm'),
 (r'structured cabling', 'structured cabling'),
 (r'security system', 'security systems'),
 (r'portfolio company of', 'portfolio company of'),
 (r'franchis', 'franchise'),
 (r'\bdistributor\b|distribution and resale|\breseller\b|resale of|manufacturer\'?s? representative', 'distribution/resale'),
 (r'fit-?out|tenant improvement', 'fit-out'),
 (r'low-?voltage only|solely low voltage', 'low-voltage-only'),
 (r'(designs?,? (and )?(develops?,? (and )?)?manufactur\w*)'
  r'|(manufactur\w* (and )?(sells?|distributes?|markets?|supplies))'
  r'|(is a .{0,40}manufacturer)|(manufacturer of)|(manufactures? \w)'
  r'|(product manufacturing)|(manufacturing (company|firm|business))'
  r'|((test|measurement) (equipment|instrument)s? (sales|rental|supplier))',
  'manufacturer (self)'),
 (r'dealership|retail .{0,20}dealer', 'vehicle dealership'),
 (r'\bplumbing\b|\bhvac\b|tree service|arboricultural|roofing', 'off-trade (plumbing/HVAC/roofing/tree)'),
]
MODEL = r'(maintenance|service|testing|repair|commissioning|troubleshoot|preventive|preventative|predictive|emergency|24/7|24 hour|24-hour|rewind|retrofit|inspection|calibration|restoration)'
OWN_DQ = {'private_equity': 'PE-backed', 'public': 'public company', 'subsidiary': 'subsidiary', 'public_subsidiary': 'public subsidiary'}
OWN_REVIEW = {'venture_capital', 'investor backed', 'angel'}

def quote(desc, m):
    s = max(0, m.start()-60); e = min(len(desc), m.end()+60)
    return ('...' if s > 0 else '') + desc[s:e].strip() + ('...' if e < len(desc) else '')

def triage(r):
    desc = str(r.get('desc') or ''); low = desc.lower()
    own = str(r.get('own') or '').strip().lower()
    hits = []
    for pat, label in GATES:
        m = re.search(pat, low)
        if m: hits.append((label, quote(desc, m)))
    if own in OWN_DQ:
        hits.append(('ownership: ' + OWN_DQ[own], 'source ownership field = "%s"' % own))
    if hits:
        return 'DQ-candidate', 'gate: ' + '; '.join('%s -> "%s"' % (l, q) for l, q in hits)
    flag = ' | REVIEW: ownership "%s" - verify no sponsor' % own if own in OWN_REVIEW else ''
    yr = r.get('yr')
    if yr and str(yr).isdigit() and int(yr) > 2016:
        flag += ' | AGE FLAG: founded %s (under 10 years)' % yr
    rev = r.get('rev'); emp = r.get('emp')
    has_model = bool(re.search(MODEL, low))
    if emp is not None and isinstance(emp, (int, float)) and emp < 10:
        return 'Nurture-candidate', ('under ~10 employees with model language' if has_model else 'under ~10 employees') + flag
    if has_model and rev and 5_000_000 <= rev <= 50_000_000:
        return 'P2-candidate', 'service/maintenance/testing language; revenue est. in $5-50M box' + flag
    if has_model:
        if rev and rev < 5_000_000: note = 'model language present; revenue est. below $5M box'
        elif rev and rev > 50_000_000: note = 'model language present; revenue est. above $50M box'
        else: note = 'model language present; revenue [unknown]'
        return 'P3-candidate', note + flag
    return 'P3-candidate', 'trade fits; model signals absent from description' + flag

def main(inp, outp):
    rows = json.load(open(inp))
    seen_d, seen_n = set(), set()
    survivors, known, nodom = [], 0, 0
    for r in rows:
        d, n = nd(r.get('domain')), nn(r.get('name'))
        if not d and not n: nodom += 1; continue
        if (d and d in DK) or (n and n in NK): known += 1; continue
        if (d and d in seen_d) or (n and n in seen_n): known += 1; continue
        if d: seen_d.add(d)
        if n: seen_n.add(n)
        pr, note = triage(r)
        r['priority'], r['note'] = pr, note
        survivors.append(r)
    json.dump({'survivors': survivors, 'dropped_known': known, 'dropped_nodomain': nodom}, open(outp, 'w'), indent=1)
    print('in=%d survivors=%d dropped_known=%d nodomain=%d' % (len(rows), len(survivors), known, nodom))

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
