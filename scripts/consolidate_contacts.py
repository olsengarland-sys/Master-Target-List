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
    'policy': 'inven contact-credit floor 6000; grata carries the remainder; 1200-company cap',
    'attempted': len(companies),
    'with_contacts': len(with_contacts),
    'credits_used': sum(v['credits_used'] for v in companies.values()),
    'companies_with_mobile': sum(1 for v in companies.values() for c in v['contacts']
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
