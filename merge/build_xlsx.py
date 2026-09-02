import json, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

m = json.load(open('merged.json'))
C = json.load(open('/home/user/Master-Target-List/scripts/contacts/owner_contacts_wave123.json'))['companies']
A = json.load(open('/home/user/Master-Target-List/scripts/contacts/addresses.json'))['companies']
QC = json.load(open('/home/user/Master-Target-List/scripts/contacts/contact_qc_flags.json'))
FLAG = ({x['domain'] for x in QC['verify_before_outreach']} | {x['domain'] for x in QC['us_only_violations']}
        | {x['domain'] for x in QC['bad_domains']})
OWNER = re.compile(r'owner|president|chief executive|\bceo\b|founder|principal|proprietor|'
                   r'managing (director|partner|member)|general manager|\bpartner\b', re.I)

def norm(d):
    d = str(d or '').strip().lower()
    d = re.sub(r'^https?://', '', d); d = re.sub(r'^www\.', '', d)
    return d.split('/')[0] or None

def contact_of(dom):
    """Best owner-level contact we hold, else the best of any contact."""
    v = C.get(dom)
    if not v or not v.get('contacts'):
        return {}
    cs = v['contacts']
    best = next((c for c in cs if OWNER.search(c.get('title') or '')), cs[0])
    ph = lambda t: ', '.join(p['number'] for p in (best.get('phones') or []) if p.get('type') == t)
    es = sorted((best.get('emails') or []),
                key=lambda e: (not e.get('is_verified'), e.get('type') != 'professional'))
    return {'name': best.get('name'), 'title': best.get('title'), 'mobile': ph('mobile'),
            'direct': ph('direct dial'), 'email': es[0]['email'] if es else None,
            'verified': 'yes' if (es and es[0].get('is_verified')) else ''}

rows = []
for dom, r in m['net_new'].items():
    c = contact_of(dom)
    a = A.get(dom) or {}
    rows.append({
        'group': 'NET-NEW', 'company': r.get('company'), 'website': dom, 'hq': r.get('hq'),
        'bucket': r.get('bucket'), 'campaign': r.get('campaign'), 'priority': r.get('priority'),
        'employees': r.get('employees'), 'revenue': r.get('revenue'), 'founded': r.get('founded'),
        'ownership': r.get('ownership'), 'contact': c.get('name'), 'title': c.get('title'),
        'email': c.get('email') or r.get('contact_email'), 'email_verified': c.get('verified'),
        'mobile': c.get('mobile'), 'direct': c.get('direct'),
        'address': a.get('street_address') if a.get('mailable') else None,
        'verify': 'YES - QC flag' if dom in FLAG else '',
        'note': r.get('description') or r.get('size_read'),
        'src': ' + '.join(r.get('sources') or []),
    })
for r in m['nurture']:
    dom = norm(r.get('Website'))
    c = contact_of(dom) if dom else {}
    a = (A.get(dom) or {}) if dom else {}
    rows.append({
        'group': 'NURTURE (from master v3)', 'company': r.get('Company'), 'website': dom,
        'hq': r.get('HQ'), 'bucket': r.get('Type bucket'), 'campaign': None,
        'priority': r.get('Priority'), 'employees': None, 'revenue': r.get('Size estimate and basis'),
        'founded': r.get('Age'), 'ownership': r.get('Ownership'),
        'contact': c.get('name') or r.get('Contact'), 'title': c.get('title') or r.get('Contact title'),
        'email': c.get('email') or r.get('Contact email'), 'email_verified': c.get('verified'),
        'mobile': c.get('mobile'), 'direct': c.get('direct'),
        'address': a.get('street_address') if a.get('mailable') else None,
        'verify': 'YES - QC flag' if dom in FLAG else '',
        'note': r.get('Why / verdict reason') or r.get('List-stage note'),
        'src': 'Master v3 Nurture sheet',
    })
rows.sort(key=lambda x: (x['group'] != 'NET-NEW', not x['mobile'], not x['email'],
                         str(x['company'] or '')))

HDR = PatternFill('solid', fgColor='1F3864'); HF = Font(color='FFFFFF', bold=True, size=10)
GREEN = PatternFill('solid', fgColor='E2EFDA'); AMBER = PatternFill('solid', fgColor='FFF2CC')
RED = PatternFill('solid', fgColor='FCE4E4'); WRAP = Alignment(wrap_text=True, vertical='top')
wb = Workbook(); ws = wb.active; ws.title = 'READ ME'
notes = [
 ('Which file is the latest?', None),
 ('', None),
 ('Master Target List v3 (25 Aug) is your BASE - 1,922 companies. Everything below is measured against it.', None),
 ('Grata Expansion (31 Aug) is the raw expansion output: net-new companies found by searching Grata and Inven.', None),
 ('Validated Campaigns (1 Sep) took that expansion, validated it and filed it into 11 campaigns.', None),
 ('New Targets for Campaigns (1 Sep) is Validated Campaigns PLUS 138 companies released from the master Nurture list.', None),
 ('', None),
 ('So: "New Targets for Campaigns" SUPERSEDES "Validated Campaigns" - it contains everything that file has, plus 138 more.', True),
 ('You do not need to work from both. The Grata Expansion still holds 627 companies that never made it into a campaign sheet.', None),
 ('', None),
 ('What is in this workbook', None),
 ('', None),
 ('One sheet, "ADD TO TARGET LIST", with everything to append to master v3:', None),
 ('  A. NET-NEW (1,353) - present in the later files, absent from master v3 entirely.', None),
 ('  B. NURTURE (138) - already in master v3 but parked below the EBITDA floor; you asked for these back.', None),
 ('', None),
 ('Companies removed as DQ in the later files were excluded (402 domains). They were dropped deliberately.', None),
 ('', None),
 ('Contact data from the enrichment run is joined on where we have it:', None),
 ('  425 of these carry an owner MOBILE, 751 an email, 834 a mailable postal address.', None),
 ('', None),
 ('CAUTION on the net-new rows', None),
 ('726 of the 1,353 were validated and campaign-filed. The other 627 come only from the Grata Expansion', None),
 ('and were never validated against the thesis - treat those as leads to screen, not qualified targets.', None),
 ('Rows marked in the "Verify first" column carry a QC flag (wrong-entity match, non-US contact, or a bad domain).', None),
]
ws.cell(row=1, column=1, value='Hunter Power - companies to add to Master Target List v3').font = Font(bold=True, size=14)
r = 3
for txt, bold in notes:
    c = ws.cell(row=r, column=1, value=txt)
    if bold: c.font = Font(bold=True)
    r += 1
ws.column_dimensions['A'].width = 125

ws = wb.create_sheet('ADD TO TARGET LIST')
cols = [('Add group',22),('Company',32),('Website',26),('HQ',24),('Type bucket',12),('Campaign',26),
        ('Priority',14),('Employees',11),('Revenue est.',26),('Year founded',12),('Ownership',16),
        ('Owner / contact',20),('Title',26),('Email',30),('Email verified',13),('MOBILE',20),
        ('Direct dial',18),('Postal address',40),('Verify first?',14),('Note / reason',60),('Source file(s)',44)]
for i,(h,w) in enumerate(cols,1):
    c = ws.cell(row=1, column=i, value=h); c.fill = HDR; c.font = HF; c.alignment = WRAP
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = 'C2'
keys = ['group','company','website','hq','bucket','campaign','priority','employees','revenue','founded',
        'ownership','contact','title','email','email_verified','mobile','direct','address','verify','note','src']
for n,row in enumerate(rows,2):
    for i,k in enumerate(keys,1):
        ws.cell(row=n, column=i, value=row.get(k)).alignment = WRAP
    if row['verify']:
        for i in range(1,22): ws.cell(row=n, column=i).fill = RED
    elif row['mobile']:
        ws.cell(row=n, column=16).fill = GREEN
    if row['group'].startswith('NURTURE'):
        ws.cell(row=n, column=1).fill = AMBER
    ws.row_dimensions[n].height = 26

OUT = '/home/user/Master-Target-List/Hunter_Power_Companies_to_Add_20260902.xlsx'
wb.save(OUT)
print('rows:', len(rows),
      '| net-new:', sum(1 for r in rows if r['group']=='NET-NEW'),
      '| nurture:', sum(1 for r in rows if r['group'].startswith('NURTURE')),
      '| with mobile:', sum(1 for r in rows if r['mobile']),
      '| with email:', sum(1 for r in rows if r['email']),
      '| with address:', sum(1 for r in rows if r['address']),
      '| flagged:', sum(1 for r in rows if r['verify']))
print('wrote', OUT)
