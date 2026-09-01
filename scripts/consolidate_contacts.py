"""Merge every enrichment shard into scripts/contacts/owner_contacts_wave123.json.

Shards are append-only JSONL written by parallel workers; a run can be killed
mid-line, so partial trailing lines are tolerated and skipped. Re-runnable:
later shards win on a duplicate domain.
"""
import json, glob, os

S = '/tmp/claude-0/-home-user-Master-Target-List/cdc2f6cb-7804-5be0-bfc8-f894363d3735/scratchpad'
QUEUE = json.load(open('scripts/contacts_enrich_list.json'))
META = {r['domain']: r for r in QUEUE}

def read_jsonl(path):
    out = []
    for line in open(path):
        line = line.strip()
        if not line:
            continue
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            continue  # truncated trailing line from a killed worker
    return out

raw = read_jsonl(f'{S}/contacts_raw.jsonl')
for p in sorted(glob.glob(f'{S}/contacts_slice_*.jsonl')):
    raw += read_jsonl(p)
# The named-lookup recovery run: Inven results for companies Grata had already
# covered. These carry the phones, so they must land in `raw` (which wins on
# phones) rather than in the Grata merge below.
for p in sorted(glob.glob(f'{S}/contacts_mob_*.jsonl')) + sorted(glob.glob(f'{S}/contacts_mob_blind.jsonl')):
    raw += read_jsonl(p)
# Grata runs last and covers companies Inven already touched, so its rows are MERGED
# into the Inven record rather than replacing it -- only Inven carries phone numbers,
# and a straight overwrite would silently drop every mobile we paid for.
grata = read_jsonl(f'{S}/contacts_grata.jsonl')
for _p in sorted(glob.glob(f'{S}/contacts_grata_final*.jsonl')) + sorted(glob.glob(f'{S}/contacts_grata_recovery.jsonl')):
    grata += read_jsonl(_p)

def key(c):
    return (c.get('name') or '').strip().lower()

companies = {}
for r in raw:
    d = r.get('domain')
    if not d:
        continue
    m = META.get(d, {})
    companies[d] = {
        'name': m.get('name'), 'bucket': m.get('bucket'), 'campaign': m.get('campaign'),
        'priority': m.get('priority'), 'wave': m.get('wave'), 'replied': m.get('replied', False),
        'source_platform': r.get('source', 'inven'), 'credits_used': r.get('credits_used', 0),
        'error': r.get('error'), 'contacts': r.get('contacts') or [],
    }

for r in grata:
    d = r.get('domain')
    if not d:
        continue
    gc = r.get('contacts') or []
    if d not in companies:
        m = META.get(d, {})
        companies[d] = {
            'name': m.get('name'), 'bucket': m.get('bucket'), 'campaign': m.get('campaign'),
            'priority': m.get('priority'), 'wave': m.get('wave'), 'replied': m.get('replied', False),
            'source_platform': 'grata', 'credits_used': 0,
            'error': r.get('error'), 'contacts': gc,
        }
        continue
    cur = companies[d]
    seen = {key(c) for c in cur['contacts']}
    added = []
    for c in gc:
        if key(c) in seen:
            # same person from both providers: keep Inven's phones, take Grata's emails
            for e in cur['contacts']:
                if key(e) == key(c) and not e.get('emails') and c.get('emails'):
                    e['emails'] = c['emails']
        else:
            added.append(c)
            seen.add(key(c))
    cur['contacts'] += added
    if added or gc:
        cur['source_platform'] = 'inven+grata' if cur['contacts'] else 'grata'
    if r.get('error') and not cur.get('error'):
        cur['error'] = r['error']

def phones_of(t):
    return sum(1 for v in companies.values() for c in v['contacts']
               for p in (c.get('phones') or []) if p.get('type') == t)

with_contacts = [v for v in companies.values() if v['contacts']]
office_only = [v for v in with_contacts
               if any(c.get('phones') for c in v['contacts'])
               and all(p.get('type') == 'office'
                       for c in v['contacts'] for p in (c.get('phones') or []))]

doc = {
    'generated': '2026-09-01',
    'platform': 'inven get_company_contacts (owner-title filtered)',
    'policy': 'inven contact-credit floor 5500 (lowered from 6000 by the client to fund the named-lookup recovery run); grata covers what inven cannot reach',
    'attempted': len(companies),
    'with_contacts': len(with_contacts),
    'credits_used': sum(v['credits_used'] for v in companies.values()),
    # count COMPANIES, not contacts - two contacts with mobiles at one company is
    # still one company we can reach.
    'companies_with_mobile': sum(1 for v in companies.values()
                                 if any(p.get('type') == 'mobile'
                                        for c in v['contacts'] for p in (c.get('phones') or []))),
    'contacts_with_mobile': sum(1 for v in companies.values() for c in v['contacts']
                                if any(p.get('type') == 'mobile' for p in (c.get('phones') or []))),
    'mobile_phones': phones_of('mobile'),
    'direct_dial_phones': phones_of('direct dial'),
    'office_only_companies': len(office_only),
    'verified_emails': sum(1 for v in companies.values() for c in v['contacts']
                           for e in (c.get('emails') or []) if e.get('is_verified')),
    'companies': companies,
}
os.makedirs('scripts/contacts', exist_ok=True)
json.dump(doc, open('scripts/contacts/owner_contacts_wave123.json', 'w'), indent=1)
print({k: v for k, v in doc.items() if k != 'companies'})
