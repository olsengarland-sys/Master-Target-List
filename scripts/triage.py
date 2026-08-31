"""Dedupe-check and triage helpers for Grata expansion candidates."""
import sys, json, re, os
sys.path.insert(0, '/home/user/Master-Target-List/scripts')
from dedupe import norm_domain, norm_name

IDX = json.load(open('/home/user/Master-Target-List/scripts/dedupe_index.json'))
DK, NK = set(IDX['dkeys']), set(IDX['nkeys'])
STORE = '/home/user/Master-Target-List/scripts/candidates.json'

# Exclusion language per system doc / prompt §2 gate check
GATES = [
    (r'residential', 'residential'),
    (r'\bsolar\b', 'solar'),
    (r'\bhome ?owner|home services\b', 'home services'),
    (r'fire alarm', 'fire alarm'),
    (r'structured cabling', 'structured cabling'),
    (r'security system', 'security systems'),
    (r'portfolio company of', 'portfolio company of'),
    (r'franchis', 'franchise'),
    (r'\bdistributor\b|distribution and resale|\breseller\b|resale of', 'distribution/resale'),
    (r'fit-?out|tenant improvement', 'fit-out'),
    (r'low-?voltage only|solely low voltage', 'low-voltage-only'),
    (r'manufactur', 'manufacturing'),
    (r'dealership|retail .{0,20}dealer', 'vehicle dealership'),
    (r'\bplumbing\b|\bhvac\b|tree service|arboricultural|roofing', 'off-trade (plumbing/HVAC/roofing/tree)'),
]
# Service/maintenance model language -> P2-candidate signal
MODEL = r'(maintenance|service|testing|repair|commissioning|troubleshoot|preventive|preventative|predictive|emergency|24/7|24 hour|24-hour|rewind|retrofit|inspection|calibration)'

def check_new(rows):
    """rows: list of dicts with name+domain. Returns (new, dropped_count, dropped_names)."""
    new, dropped = [], []
    for r in rows:
        d, n = norm_domain(r.get('domain')), norm_name(r.get('name'))
        if (d and d in DK) or (n and n in NK):
            dropped.append(r.get('name'))
        else:
            new.append(r)
    return new, len(dropped), dropped

def triage(c):
    """Assign gate flags and provisional priority from Grata fields only."""
    desc = (c.get('desc') or '')
    low = desc.lower()
    hits = [label for pat, label in GATES if re.search(pat, low)]
    emp = c.get('emp')
    if hits:
        return 'DQ-candidate', 'gate: ' + '; '.join(f'"{h}"' for h in hits)
    if isinstance(emp, (int, float)) and emp < 10:
        if re.search(MODEL, low):
            return 'Nurture-candidate', f'under ~10 employees (Grata est. {emp}) with model language'
        return 'Nurture-candidate', f'under ~10 employees (Grata est. {emp})'
    own = (c.get('own') or '')
    own_flag = ' | REVIEW: ownership "%s" (not bootstrapped) — verify no sponsor' % own if own and own.lower() != 'bootstrapped' else ''
    rev = c.get('rev')
    in_box = isinstance(rev, (int, float)) and 5_000_000 <= rev <= 50_000_000
    if re.search(MODEL, low) and in_box:
        return 'P2-candidate', 'service/maintenance/testing language; Grata revenue est. in $5-50M box' + own_flag
    if re.search(MODEL, low):
        note = 'service/maintenance language; revenue est. outside/unknown vs box'
        if isinstance(rev, (int, float)) and rev < 5_000_000:
            note += f' (Grata est. ${rev/1e6:.1f}M — below floor)'
        elif isinstance(rev, (int, float)) and rev > 50_000_000:
            note += f' (Grata est. ${rev/1e6:.1f}M — above box)'
        return 'P3-candidate', note + own_flag
    return 'P3-candidate', 'trade fits; model signals absent from Grata description' + own_flag

DESIGN_BUILD = r'(design[- ]build|design/build|design-assist|ground-?up|new construction|tenant (improvement|build-?out)|fit-?out|pre-?construction)'

def is_pure_design_build(c):
    """T5 rule: drop pure design-build/new-construction results lacking service/maintenance language."""
    low = (c.get('desc') or '').lower()
    return bool(re.search(DESIGN_BUILD, low)) and not re.search(MODEL, low)

def load_store():
    return json.load(open(STORE)) if os.path.exists(STORE) else {}

def save_bucket(bucket, source, rows, dropped_add=0):
    """Dedupe rows, triage survivors, merge into the store under `bucket`."""
    st = load_store()
    b = st.setdefault(bucket, {'candidates': [], 'dropped_known': 0, 'sources': []})
    new, ndrop, _ = check_new(rows)
    b['dropped_known'] += ndrop + dropped_add
    if bucket == 'T5':
        before = len(new)
        new = [r for r in new if not is_pure_design_build(r)]
        b['dropped_offthesis'] = b.get('dropped_offthesis', 0) + (before - len(new))
    seen = {norm_domain(x['domain']) for x in b['candidates']}
    added = 0
    for r in new:
        d = norm_domain(r.get('domain'))
        if d in seen:
            continue
        seen.add(d)
        pri, note = triage(r)
        r = dict(r); r['source'] = source; r['priority'] = pri; r['note'] = note
        b['candidates'].append(r)
        added += 1
    if source not in b['sources']:
        b['sources'].append(source)
    json.dump(st, open(STORE, 'w'), indent=1)
    print(f'{bucket} [{source}]: {len(rows)} returned -> {ndrop} known-dropped, {added} new added '
          f'(bucket totals: {len(b["candidates"])} new, {b["dropped_known"]} dropped'
          + (f', {b["dropped_offthesis"]} off-thesis design-build' if b.get('dropped_offthesis') else '') + ')')

if __name__ == '__main__':
    # stdin: JSON list of {name,domain} -> report which are new
    rows = json.load(sys.stdin)
    new, ndrop, dropped = check_new(rows)
    print(f'{len(rows)} in, {ndrop} known (dropped), {len(new)} NEW')
    print('NEW:', json.dumps([r['domain'] for r in new]))
