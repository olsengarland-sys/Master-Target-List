"""Build the single master tracker workbook."""
import json, re, collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

R = '/home/user/Master-Target-List/'
DB = json.load(open('db.json'))
C = json.load(open(R + 'scripts/contacts/owner_contacts_wave123.json'))['companies']
A = json.load(open(R + 'scripts/contacts/addresses.json'))['companies']
QC = json.load(open(R + 'scripts/contacts/contact_qc_flags.json'))
FLAGS = {}
for key, lab in [('verify_before_outreach','wrong-entity / non-US - verify before contact'),
                 ('us_only_violations','non-US domain, mandate is US-only'),
                 ('bad_domains','not a real company domain'),
                 ('ownership_signals','possible undisclosed parent'),
                 ('possible_duplicate_entities','possible duplicate of another target')]:
    for x in QC.get(key, []):
        for d in ([x['domain']] if x.get('domain') else x.get('domains', [])):
            FLAGS.setdefault(d, []).append(lab)
try:
    CONF = {}
    import csv as _c
    for row in _c.DictReader(open(R + 'deliverables/05_conference_rosters.csv')):
        if row.get('in_target_list') == 'YES':
            CONF.setdefault(row['domain'], []).append(f"{row['event']} ({row['start']})")
except Exception:
    CONF = {}

OWNER = re.compile(r'owner|president|chief executive|\bceo\b|founder|principal|proprietor|'
                   r'managing (director|partner|member)|general manager|\bpartner\b', re.I)
CAMPAIGNS = ['Commercial & Industrial Ctrs','Motor & Generator Companies','Utility Line & Substation',
 'Mission critical power','Traffic Signals & Lighting','Testing & Commissioning',
 'Electrical Apparatus Servicing','EV Charging','Lightning Protection','Electrical Equipment Mfg',
 'Power Systems Engineering']
BUCKET2CAMP = {'T1':'Testing & Commissioning','T2':'Electrical Apparatus Servicing',
 'T3':'Utility Line & Substation','T4':'Commercial & Industrial Ctrs','T5':'Commercial & Industrial Ctrs',
 'T6':'Mission critical power','S7':'Motor & Generator Companies','S8':'Traffic Signals & Lighting',
 'S9':'EV Charging','S10':'Lightning Protection','M13':'Electrical Equipment Mfg',
 'E12':'Power Systems Engineering'}
BEN2CAMP = {'Commercial & Industrial Contractors':'Commercial & Industrial Ctrs',
 'Motor & Generator Companies':'Motor & Generator Companies',
 'Traffic Signals & Lighting Companies':'Traffic Signals & Lighting',
 'Substation & T&D':'Utility Line & Substation'}

def has(rec, frag):
    return any(frag in s for s in rec['sources'])

def resolve(rec):
    """Status, campaign and which file decided it - newest verdict wins."""
    d = rec['domain']
    if has(rec, 'New Targets 1 Sep / Remove - DQ'):
        return 'DQ', None, 'New Targets 1 Sep (validation)'
    nt = [s for s in rec['sources'] if s.startswith('New Targets 1 Sep /') and 'Remove' not in s]
    if nt:
        return 'ACTIVE TARGET', rec.get('v_campaign') or nt[0].split(' / ')[1], 'New Targets 1 Sep'
    ms = rec.get('master_status')
    if ms in ('Batch 1','Batch 2','Batch 3','New add B1/2','New add B2/3'):
        b = str(rec.get('bucket') or '')[:3].strip()
        return 'ACTIVE TARGET', BUCKET2CAMP.get(b[:2]) or BUCKET2CAMP.get(b), f'Master v3 / {ms}'
    if ms == 'Nurture':
        b = str(rec.get('bucket') or '')[:3].strip()
        return 'NURTURE', BUCKET2CAMP.get(b[:2]) or BUCKET2CAMP.get(b), 'Master v3 / Nurture'
    if ms == 'DQ':
        return 'DQ', None, 'Master v3 / DQ log'
    if has(rec, 'Validated 1 Sep / Remove - DQ'):
        return 'DQ', None, 'Validated 1 Sep'
    ca = [s for s in rec['sources'] if s.startswith('Campaign Assignment')]
    if ca:
        sh = ca[0].split(' / ')[1]
        if sh == 'DQ':      return 'DQ', None, 'Campaign Assignment 31 Aug'
        if sh == 'Nurture': return 'NURTURE', None, 'Campaign Assignment 31 Aug'
        return 'ACTIVE TARGET', sh, 'Campaign Assignment 31 Aug'
    if has(rec, 'Ben/Anchor'):
        return 'ACTIVE TARGET', BEN2CAMP.get(rec.get('ben_campaign') or ''), 'Ben/Anchor list'
    if has(rec, 'Trade association'):
        return 'RAW LEAD - not screened', None, 'Trade association directory'
    return 'UNCLASSIFIED', None, 'no verdict found'

rows = []
for d, rec in DB.items():
    status, campaign, decided = resolve(rec)
    v = C.get(d); ct = {}
    if v and v.get('contacts'):
        cs = v['contacts']
        best = next((x for x in cs if OWNER.search(x.get('title') or '')), cs[0])
        ph = lambda t: ', '.join(p['number'] for p in (best.get('phones') or []) if p.get('type') == t)
        es = sorted((best.get('emails') or []), key=lambda e: (not e.get('is_verified'), e.get('type') != 'professional'))
        ct = {'n': best.get('name'), 't': best.get('title'), 'm': ph('mobile'),
              'dd': ph('direct dial'), 'off': ph('office'),
              'e': es[0]['email'] if es else None,
              'ev': 'yes' if (es and es[0].get('is_verified')) else '',
              'alt': ', '.join(x.get('name') or '' for x in cs[1:3]).strip(', ')}
    a = A.get(d) or {}
    rows.append({
        'status': status, 'campaign': campaign or '', 'decided_by': decided,
        'company': rec.get('company'), 'domain': d,
        'hq': rec.get('hq'),
        'address': a.get('street_address') if a.get('mailable') else (a.get('street_address') or ''),
        'address_mailable': 'yes' if a.get('mailable') else ('partial' if a.get('street_address') else ''),
        'bucket': rec.get('bucket') or rec.get('bucket_reassessed') or rec.get('bucket_v'),
        'band': rec.get('band'),
        'priority': rec.get('priority') or rec.get('priority_v') or rec.get('priority_reassessed'),
        'rank': rec.get('rank'),
        'owner_contact': ct.get('n') or rec.get('contact') or rec.get('owner'),
        'owner_title': ct.get('t') or rec.get('contact_title'),
        'email': ct.get('e') or rec.get('contact_email'),
        'email_verified': ct.get('ev'),
        'mobile': ct.get('m'),
        'direct_dial': ct.get('dd'),
        'office_phone': ct.get('off') or rec.get('contact_phone'),
        'other_contacts': ct.get('alt'),
        'revenue': rec.get('revenue') or rec.get('size'),
        'employees': rec.get('employees'),
        'founded': rec.get('founded') or rec.get('age'),
        'ownership': rec.get('ownership'),
        'size_read': rec.get('size_read'), 'size_check': rec.get('size_check'),
        'scope': rec.get('scope'), 'model': rec.get('model'),
        'verification': rec.get('verification'), 'web_check': rec.get('web_check'),
        'confidence': rec.get('confidence') or rec.get('v_confidence'),
        'why': rec.get('why') or rec.get('v_reason'),
        'dq_reason': rec.get('dq_reason') or (rec.get('ca_reason') if status == 'DQ' else None),
        'reass_action': rec.get('reass_action'), 'reass_reason': rec.get('reass_reason'),
        'neta': rec.get('neta'), 'acq_check': rec.get('acq'), 'unknowns': rec.get('unknowns'),
        'angle': rec.get('angle'),
        'ben_tier': rec.get('ben_tier'), 'ben_reply': rec.get('ben_reply'),
        'ben_succession': rec.get('ben_succession'),
        'conferences': '; '.join(CONF.get(d, [])[:4]),
        'qc_flag': '; '.join(FLAGS.get(d, [])),
        'wave': rec.get('wave'), 'assoc_list': rec.get('assoc_list'),
        'profile': rec.get('profile'),
        'sources': ' | '.join(rec['sources']),
        'description': rec.get('description'),
    })

# description overflow -> Description 2/3 (Excel caps a cell at 32,767 chars)
CHUNK = 8000
maxparts = 1
for r in rows:
    t = str(r['description'] or '')
    parts = [t[i:i+CHUNK] for i in range(0, len(t), CHUNK)] or ['']
    maxparts = max(maxparts, len(parts))
    for i, p in enumerate(parts, 1):
        r[f'description_{i}'] = p
print('longest description:', max(len(str(r['description'] or '')) for r in rows),
      '-> description columns needed:', maxparts)
json.dump({'rows': rows, 'maxparts': maxparts}, open('rows.json', 'w'), default=str)
print('total rows:', len(rows))
print(collections.Counter(r['status'] for r in rows).most_common())
