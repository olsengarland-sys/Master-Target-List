"""Merge every address shard into scripts/contacts/addresses.json.

Sources: the original mainline_map pull, the three slice files, and any
out-of-slice rows a worker paid for and parked (those are already bought;
discarding them would waste the credits).

A row with a city but no street line is NOT mailable, so street-level and
partial addresses are counted separately rather than lumped together.
"""
import json, glob, re, os

S = '/tmp/claude-0/-home-user-Master-Target-List/cdc2f6cb-7804-5be0-bfc8-f894363d3735/scratchpad'

def read_jsonl(path):
    out = []
    for line in open(path):
        line = line.strip()
        if not line:
            continue
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return out

addr = {}
# seed from the original pull
if os.path.exists('scripts/contacts/mainline_map.json'):
    for d, v in json.load(open('scripts/contacts/mainline_map.json'))['companies'].items():
        if v.get('street_address'):
            addr[d] = {'street_address': v['street_address'], 'note': None}

for p in sorted(glob.glob(f'{S}/addr_slice_*.jsonl')) + sorted(glob.glob(f'{S}/*out_of_slice.jsonl')):
    for r in read_jsonl(p):
        d, a = r.get('domain'), r.get('street_address')
        if d and a:
            addr[d] = {'street_address': a, 'note': r.get('note')}

# A mailable address needs a street line: a leading number, or a PO box.
STREET = re.compile(r'\d+\s+\w|p\.?\s*o\.?\s*box|\bsuite\b|\bste\b|\bhwy\b|\broad\b|\bstreet\b|\bave\b',
                    re.I)
for d, v in addr.items():
    v['mailable'] = bool(STREET.search(v['street_address']))

q = json.load(open('scripts/contacts_enrich_list.json'))
targets = {r['domain'] for r in q}
doc = {
    'generated': '2026-09-01',
    'source': 'inven get_company_info (ai_enrichment pool)',
    'targets_total': len(targets),
    'with_address': sum(1 for d in addr if d in targets),
    'mailable_street_level': sum(1 for d, v in addr.items() if d in targets and v['mailable']),
    'partial_city_state_only': sum(1 for d, v in addr.items() if d in targets and not v['mailable']),
    'note': ('A row with only city/state/ZIP is not mailable as-is - filter on `mailable` before a '
             'mail merge. Flagged rows carry a `note`: wrong-entity matches, registered-agent or '
             'virtual-office suites, and non-US addresses on a US-only mandate.'),
    'companies': {d: v for d, v in addr.items()},
}
json.dump(doc, open('scripts/contacts/addresses.json', 'w'), indent=1)
print({k: v for k, v in doc.items() if k != 'companies'})
