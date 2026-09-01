"""Export every dataset built in this engagement to flat CSV/JSON in deliverables/.

The workbook is the readable artifact; this is the machine-readable mirror so the
data can go into a CRM or be re-cut without opening Excel.
"""
import json, csv, os, glob, shutil

OUT = 'deliverables'
os.makedirs(OUT, exist_ok=True)

def write_csv(name, rows, fields=None):
    if not rows:
        return 0
    fields = fields or list(rows[0].keys())
    with open(f'{OUT}/{name}', 'w', newline='') as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction='ignore')
        w.writeheader()
        w.writerows(rows)
    return len(rows)

manifest = []

# 1. Enriched contacts, one row per contact
C = json.load(open('scripts/contacts/owner_contacts_wave123.json'))
rows = []
for dom, v in C['companies'].items():
    if not v['contacts']:
        rows.append({'company': v.get('name'), 'domain': dom, 'contact': '', 'title': '',
                     'mobile': '', 'direct_dial': '', 'office': '', 'email': '',
                     'email_verified': '', 'bucket': v.get('bucket'), 'campaign': v.get('campaign'),
                     'priority': v.get('priority'), 'source': v.get('source_platform'),
                     'status': v.get('error') or 'no contact found'})
        continue
    for c in v['contacts']:
        ph = lambda t: ', '.join(p['number'] for p in (c.get('phones') or []) if p.get('type') == t)
        es = sorted((c.get('emails') or []),
                    key=lambda e: (not e.get('is_verified'), e.get('type') != 'professional'))
        rows.append({'company': v.get('name'), 'domain': dom, 'contact': c.get('name'),
                     'title': c.get('title'), 'mobile': ph('mobile'), 'direct_dial': ph('direct dial'),
                     'office': ph('office'), 'email': es[0]['email'] if es else '',
                     'email_verified': 'yes' if (es and es[0].get('is_verified')) else '',
                     'bucket': v.get('bucket'), 'campaign': v.get('campaign'),
                     'priority': v.get('priority'), 'source': v.get('source_platform'),
                     'status': v.get('error') or 'ok'})
manifest.append(('01_enriched_contacts.csv', write_csv('01_enriched_contacts.csv', rows),
                 'Every enriched contact: owner name, title, mobile / direct dial / office, email'))

# 2. Direct lines (already built)
if os.path.exists('scripts/contacts/direct_lines.csv'):
    shutil.copy('scripts/contacts/direct_lines.csv', f'{OUT}/02_direct_lines.csv')
    n = sum(1 for _ in open(f'{OUT}/02_direct_lines.csv')) - 1
    manifest.append(('02_direct_lines.csv', n, 'Desk numbers that bypass the switchboard'))

# 3. Postal addresses
A = json.load(open('scripts/contacts/addresses.json'))['companies']
_tg = {r['domain'] for r in json.load(open('scripts/contacts_enrich_list.json'))}
rows = [{'domain': d, 'street_address': v['street_address'],
         'mailable': 'yes' if v.get('mailable') else 'NO - city/state only',
         'in_target_list': 'yes' if d in _tg else '',
         'flag': v.get('note') or ''}
        for d, v in sorted(A.items())]
manifest.append(('03_postal_addresses.csv', write_csv('03_postal_addresses.csv', rows),
                 'HQ addresses. Filter mailable=yes before a mail merge; city/state-only rows '
                 'cannot be posted.'))

# 4. QC flags - everything that must be checked before outreach
Q = json.load(open('scripts/contacts/contact_qc_flags.json'))
rows = []
for k, label in [('verify_before_outreach', 'WRONG-ENTITY / NON-US - do not call unverified'),
                 ('ownership_signals', 'possible undisclosed parent'),
                 ('non_owner_titles', 'contact is not a decision maker'),
                 ('us_only_violations', 'non-US domain, mandate is US-only'),
                 ('bad_domains', 'not a real company domain'),
                 ('possible_duplicate_entities', 'same phone as another target - possible duplicate')]:
    for x in Q.get(k, []):
        rows.append({'flag': label, 'domain': x.get('domain') or ', '.join(x.get('domains', [])),
                     'company': x.get('company') or x.get('name') or '',
                     'contact': x.get('contact', ''), 'title': x.get('title', ''),
                     'detail': x.get('reason') or x.get('note') or x.get('alt_email_domain') or '',
                     'priority': x.get('priority', ''), 'campaign': x.get('campaign', '')})
manifest.append(('04_qc_flags.csv', write_csv('04_qc_flags.csv', rows),
                 'Everything needing a human check before outreach'))

# 5. Conference rosters + 6. event summary
if os.path.exists('scripts/xref/conference_rosters_flat.csv'):
    shutil.copy('scripts/xref/conference_rosters_flat.csv', f'{OUT}/05_conference_rosters.csv')
    n = sum(1 for _ in open(f'{OUT}/05_conference_rosters.csv')) - 1
    manifest.append(('05_conference_rosters.csv', n,
                     'Every banked attendee record, flagged where it is one of our targets'))
ev = json.load(open('scripts/xref/conference_event_summary.json'))
manifest.append(('06_conference_events.csv', write_csv('06_conference_events.csv', ev),
                 'Events ranked by how many of our targets attend'))

# 7. Forward calendar
if os.path.exists('scripts/xref/conferences_forward.json'):
    fw = json.load(open('scripts/xref/conferences_forward.json'))
    fw = fw if isinstance(fw, list) else fw.get('events', [])
    manifest.append(('07_conference_calendar_forward.csv', write_csv('07_conference_calendar_forward.csv', fw),
                     'Forward 18-month conference calendar'))

# 8-11. M&A intelligence
M = json.load(open('scripts/xref/ma_intel.json'))
T = M['transactions']
rows = [dict(r, kind='disclosed multiple') for r in T['disclosed_multiples']] + \
       [dict(r, kind='notable recent') for r in T['notable_recent']]
manifest.append(('08_transactions.csv',
                 write_csv('08_transactions.csv', rows,
                           ['kind', 'target', 'buyer', 'date', 'type', 'ev', 'ebitda',
                            'ev_ebitda', 'note']),
                 'Closed deals. Universe: %d US electrical-contractor deals in 5yr, %d industrial '
                 'repair in 3yr. %s' % (T['universe']['electrical_contractors_5yr_US'],
                                        T['universe']['industrial_repair_3yr_US'],
                                        T['multiples_note'])))

manifest.append(('09_live_mandates.csv', write_csv('09_live_mandates.csv', M['live_deals']),
                 'Businesses currently marketed for sale. ' + M['ask_multiple_read']))

B = M['bankers']
rows = [dict(r, kind='sector specialist') for r in B['sector_specialists']] + \
       [dict(r, kind='deal-evidenced') for r in B['deal_evidenced']]
manifest.append(('10_bankers.csv',
                 write_csv('10_bankers.csv', rows, ['kind', 'name', 'domain', 'hq', 'why', 'evidence']),
                 'Advisors who sell these businesses - the intermediaries to be known by'))

Y = M['buyers']
rows = [dict(r, kind='PE sponsor') for r in Y['pe_sponsors_lmm']] + \
       [dict(r, kind='strategic acquirer') for r in Y['strategics']]
manifest.append(('11_buyers_pe_and_strategic.csv',
                 write_csv('11_buyers_pe_and_strategic.csv', rows,
                           ['kind', 'name', 'matching_portfolio', 'note']),
                 'Who you are bidding against. %d PE sponsors and %d strategics active in the trade. %s'
                 % (Y['pe_universe_count'], Y['strategic_universe_count'], Y['read'])))

# 12. The full target list with its enrichment status
q = json.load(open('scripts/contacts_enrich_list.json'))
rows = []
for r in q:
    v = C['companies'].get(r['domain'])
    has = bool(v and v['contacts'])
    mob = bool(v and any(p.get('type') == 'mobile' for c in v['contacts'] for p in (c.get('phones') or [])))
    rows.append({'company': r.get('name'), 'domain': r['domain'], 'bucket': r.get('bucket'),
                 'campaign': r.get('campaign'), 'priority': r.get('priority'), 'wave': r.get('wave'),
                 'had_email_before': 'yes' if r.get('has_email') else '',
                 'had_phone_before': 'yes' if r.get('has_phone') else '',
                 'enriched': 'yes' if has else '', 'has_mobile': 'yes' if mob else '',
                 'source': (v or {}).get('source_platform', ''),
                 'status': (v or {}).get('error') or ('ok' if has else 'not enriched')})
manifest.append(('12_all_targets_status.csv', write_csv('12_all_targets_status.csv', rows),
                 'Every target with its enrichment status'))

# 13. Held-phone verdicts
if os.path.exists('scripts/contacts/mainline_resolved.json'):
    MR = json.load(open('scripts/contacts/mainline_resolved.json'))
    rows = [{'domain': r['domain'], 'company': r['company'], 'phone_on_file': r['phone_on_file'],
             'company_switchboard': ', '.join(r['grata_office']), 'verdict': r['verdict']}
            for r in MR['results']]
    manifest.append(('13_held_phone_verdicts.csv', write_csv('13_held_phone_verdicts.csv', rows),
                     'Held phones tested against the company switchboard. ' + MR['note']))

# 14. The call list: every owner reachable by mobile, with email and address
import subprocess  # built by scripts/build_call_list.py
if os.path.exists(f'{OUT}/14_owner_mobiles.csv'):
    n = sum(1 for _ in open(f'{OUT}/14_owner_mobiles.csv')) - 1
    manifest.append(('14_owner_mobiles.csv', n,
                     'THE CALL LIST - every owner with a mobile, plus their email and mailing '
                     'address. Sorted decision-makers first; check the verify_first column before '
                     'dialling.'))

with open(f'{OUT}/00_MANIFEST.csv', 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['file', 'rows', 'what it is'])
    for n, c, d in manifest:
        w.writerow([n, c, d])
for n, c, d in manifest:
    print(f'{c:7} {n:40} {d}')
