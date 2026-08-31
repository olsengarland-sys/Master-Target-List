import json, re, statistics, collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CAND = json.load(open('scripts/candidates.json'))
COMPS = json.load(open('scripts/comps.json'))
CUTS  = json.load(open('scripts/cuts.json'))

OUT = 'Hunter_Power_Grata_Expansion_20260831.xlsx'
ORDER = ['T1','T2','T3','T4','T5','T6','S7','S8','S9','S10']
NAMES = {
 'T1':'T1 Testing & Commissioning','T2':'T2 Apparatus Service','T3':'T3 Utility Line & Substation',
 'T4':'T4 Industrial Plant Electrical','T5':'T5 Commercial & Institutional','T6':'T6 Critical Power & Data Ctr',
 'S7':'S7 Motor & Generator','S8':'S8 Traffic Signals & Lighting','S9':'S9 EV Charging',
 'S10':'S10 Lightning Protection'}

RED   = PatternFill('solid', fgColor='FCE4E4')
AMBER = PatternFill('solid', fgColor='FFF2CC')
HDR   = PatternFill('solid', fgColor='1F3864')
HDRF  = Font(color='FFFFFF', bold=True, size=10)
BOLD  = Font(bold=True)
TITLE = Font(bold=True, size=13)
WRAP  = Alignment(wrap_text=True, vertical='top')
TOP   = Alignment(vertical='top')
THIN  = Border(*[Side(style='thin', color='D0D0D0')]*4)

wb = Workbook()

def header(ws, row, cols):
    for i,c in enumerate(cols,1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.fill = HDR; cell.font = HDRF; cell.alignment = WRAP
    ws.freeze_panes = ws.cell(row=row+1, column=1)

def widths(ws, w):
    for i,x in enumerate(w,1): ws.column_dimensions[get_column_letter(i)].width = x

def money(v):
    if v is None: return '[unknown]'
    return '$%.1fM' % (v/1_000_000)

def num(v):
    return '[unknown]' if v is None else v

# ---------------------------------------------------------------- Summary
ws = wb.active; ws.title = 'Summary'
ws['A1'] = 'Hunter Power Services — Grata Bucket Expansion'; ws['A1'].font = Font(bold=True, size=15)
ws['A2'] = 'Run date 2026-08-31 · Sourcing analyst pass · Parts A (expansion), B (public comps), C (market cuts)'
ws['A3'] = 'All company facts below are Grata returned fields or the master list. Revenue and employee figures are GRATA ESTIMATES. Unknowns are written [unknown], never guessed.'
ws['A3'].font = Font(italic=True, size=9)

r = 5
ws.cell(row=r, column=1, value='NEW CANDIDATES BY BUCKET AND PROVISIONAL PRIORITY').font = TITLE; r += 1
cols = ['Bucket','Bucket name','P2-candidate','P3-candidate','Nurture-candidate','DQ-candidate','Total new','Dropped as known (dedupe)','Dropped off-thesis at triage']
header(ws, r, cols); r += 1
tot = collections.Counter()
for b in ORDER:
    v = CAND[b]; pc = collections.Counter(c['priority'] for c in v['candidates'])
    n = len(v['candidates'])
    row = [b, NAMES[b], pc['P2-candidate'], pc['P3-candidate'], pc['Nurture-candidate'], pc['DQ-candidate'],
           n, v.get('dropped_known', 0), v.get('dropped_offthesis', 0)]
    for i,x in enumerate(row,1): ws.cell(row=r, column=i, value=x).alignment = TOP
    for k in ('P2-candidate','P3-candidate','Nurture-candidate','DQ-candidate'): tot[k] += pc[k]
    tot['n'] += n; tot['dk'] += v.get('dropped_known',0); tot['do'] += v.get('dropped_offthesis',0)
    r += 1
trow = ['TOTAL','', tot['P2-candidate'], tot['P3-candidate'], tot['Nurture-candidate'], tot['DQ-candidate'], tot['n'], tot['dk'], tot['do']]
for i,x in enumerate(trow,1):
    c = ws.cell(row=r, column=i, value=x); c.font = BOLD
r += 3

ws.cell(row=r, column=1, value='GRATA TOKEN BALANCE').font = TITLE; r += 1
for lab,val in [('Balance at start of run','5,743 of 7,000 remaining'),
                ('Balance at end of run','4,758 of 7,000 remaining (32.03% of annual pool used)'),
                ('Consumed by this run','985 tokens'),
                ('Billing period','2026-08-24 to 2027-08-23'),
                ('Tools used','search_companies (3), find_similar_companies (3), get_public_comps (1), get_token_usage (0). No credit-consuming or side-effectful tool was called.')]:
    ws.cell(row=r, column=1, value=lab).font = BOLD
    ws.cell(row=r, column=2, value=val).alignment = WRAP
    r += 1
r += 2

ws.cell(row=r, column=1, value='CONFIDENCE AND GAPS').font = TITLE; r += 1
gaps = [
 ('Seller Intent unavailable',
  'seller_intent_levels: ["High Intent"] returns HTTP 403 — "Your organization does not have access to Seller Intent." Skipped per prompt instruction. No seller-intent cut exists in Part C for any bucket.'),
 ('hold_period_signal unavailable',
  'exit_readiness_signals: ["hold_period_signal"] is silently remapped by the API to transaction_signal (visible in the returned filters_used) and returns 0. Only founder_age_signal is usable. Treat the exit-readiness cut as founder-age only.'),
 ('Nurture-candidate is structurally empty',
  'The common filter employees_min: 10 removes sub-10-FTE firms before they can be returned, so the Nurture-candidate class cannot populate from these searches. Finding sub-floor nurture names requires a separate run with employees_min lowered — recommended as a follow-on.'),
 ('P1 was never assigned, by design',
  'Per the system doc §3.6, P1 requires the enrichment pass. Nothing here is above P2-candidate. Every priority in this workbook is provisional and read off a Grata description only.'),
 ('Seeds resolved by domain, not lookup_companies',
  'find_similar_companies and get_public_comps accept a domain directly as company_uid, so seeds were passed as domains. This saved ~300 tokens and 10 round trips but means no lookup_companies confirmation step ran on seed identity. Seed domains came from the master and are listed on the Public comps tab.'),
 ('Public comps map to listed PARENTS, not trade peers',
  'Grata resolves a private comp to whatever listed entity owns it. That produced non-comparable "comps" — Goldman Sachs (T6), Vistra (T6), KKR and HEICO (S7), ACCO Brands (S10), Ares and OFS Capital (S9). These are flagged amber on the Public comps tab and should be ignored for valuation.'),
 ('Scale discount is not optional',
  'Every multiple on the Public comps tab is an EMCOR/Quanta/Comfort Systems-scale listed multiple (13x-32x EV/EBITDA). These DO NOT transfer to a $2-10M EBITDA private target. Lower-middle-market electrical services trade in a materially lower band; the listed set is directional context for Olsen\'s pricing instinct, not a pricing model. The S7 seed 2 set (median 9.4x, genuine motor-service peers) is the closest thing in this run to a trade-honest read.'),
 ('locations filter deviates from the prompt literal',
  'locations: ["US"] returns HTTP 400 "Ambiguous locations term(s)". Country was passed as "United States" and states as "State, US" (e.g. "Texas, US"), which matches the prompt instruction to spell out state names.'),
 ('S9 EV Charging search is heavily off-thesis',
  'The EV-charging keyword profile returns vehicle dealerships, plumbing/HVAC and tree-service firms. 14 of 19 S9 survivors gate as DQ-candidate — consistent with the master\'s own S9 record (11 DQ of 17). S9 is not a productive hunting ground on Grata keywords alone.'),
 ('T2 bleeds into S7',
  '"electrical apparatus service" pulls EASA motor-shop results, overlapping S7. The two buckets should be read together, not as independent pipelines.'),
 ('Residential gate is literal, as the prompt specifies',
  'Any residential language in a Grata description fires the gate, including firms that list residential alongside commercial and industrial. 32 of 64 DQ-candidates fired on this gate. Several are likely C&I-dominant and worth a human re-read before they enter the DQ log permanently — the quoted phrase is on the DQ tab so the mix can be judged.'),
 ('Unknown-revenue pool is effectively zero',
  'In every bucket the five revenue bands sum to the unfiltered bucket total (variance 0-2 companies), i.e. Grata carries a revenue estimate for essentially every company in these filter sets. The pool was derived arithmetically rather than by 50 extra revenue_include_unknown calls.'),
 ('Search depth was not uniform across buckets',
  'Every bucket got a keyword search_companies pass AND a find_similar_companies pass seeded with up to 5 master P1/P2 companies '
  '(Grata warns per-seed signal dilutes past ~5 seeds). Page depth varied by how fast results went off-thesis: T2 and T3 ran 2 pages '
  'of similar results, the other buckets 1 page (25 rows). The prompt allows up to 4 pages; depth was cut where overlap with the '
  'keyword pass was already running above 40 percent, which it was in every bucket.'),
 ('S9 had no P1/P2 seed in the master',
  'S9 EV Charging holds no P1 or P2 company in the master, so the similar-to pass used the top P3 (Resound Energy, resoundenergy.com) '
  'as a fallback seed and the comps seed is the same P3. Treat S9 similarity results as anchored to a weaker seed than every other bucket.'),
 ('13 S9 results dropped off-thesis rather than logged as DQ',
  'The S9 similar-to pass returned pure LED-lighting-retrofit and energy-services firms with no EV or electrical-infrastructure service '
  'line. These were dropped at triage rather than carried into the DQ tab as noise; the names are recorded in scripts/candidates.json '
  'under S9.offthesis_note. No other bucket needed an off-thesis drop.'),
 ('Description-only triage',
  'No enrichment call was made on any candidate. Model signals (recurring service revenue, contract base, technician count, NETA accreditation) are inferred from marketing prose. Expect meaningful false-positive and false-negative rates on the P2/P3 split.'),
]
header(ws, r, ['Gap / caveat','Detail']); r += 1
for g,d in gaps:
    ws.cell(row=r, column=1, value=g).alignment = WRAP
    c = ws.cell(row=r, column=2, value=d); c.alignment = WRAP; c.fill = AMBER
    ws.row_dimensions[r].height = 42
    r += 1
widths(ws, [34, 110, 16, 18, 20, 16, 12, 24, 26])

# ---------------------------------------------------------------- bucket tabs
BCOLS = ['Company','Domain','HQ (city, state)','Grata revenue est.','Grata employee est.','Ownership',
         'Year founded','Grata description (first 300 chars)','Source (search keywords or similar-to seed)',
         'Provisional priority','Gate / flag notes','Grata profile URL']
PRI_ORDER = {'P2-candidate':0,'P3-candidate':1,'Nurture-candidate':2,'DQ-candidate':3}

for b in ORDER:
    ws = wb.create_sheet(b)
    ws['A1'] = NAMES[b] + ' — new candidates not in master, DQ log, Nurture or screened log'
    ws['A1'].font = TITLE
    v = CAND[b]
    ws['A2'] = ('%d new candidates · %d dropped as already-known at dedupe · %d dropped off-thesis at triage · '
                'revenue and employee figures are Grata estimates'
                % (len(v['candidates']), v.get('dropped_known',0), v.get('dropped_offthesis',0)))
    ws['A2'].font = Font(italic=True, size=9)
    header(ws, 4, BCOLS)
    r = 5
    for c in sorted(v['candidates'], key=lambda x: (PRI_ORDER.get(x['priority'],9), -(x.get('rev') or 0))):
        desc = (c.get('desc') or '')[:300]
        row = [c['name'], c.get('domain') or '[unknown]', c.get('hq') or '[unknown]',
               money(c.get('rev')), num(c.get('emp')), c.get('own') or '[unknown]',
               num(c.get('yr')), desc, c.get('source',''), c['priority'], c.get('note',''), c.get('url','')]
        for i,x in enumerate(row,1):
            cell = ws.cell(row=r, column=i, value=x); cell.alignment = WRAP; cell.border = THIN
        if c['priority'] == 'DQ-candidate':
            for i in range(1, len(BCOLS)+1): ws.cell(row=r, column=i).fill = RED
        elif 'REVIEW' in c.get('note','') or c['priority'] == 'Nurture-candidate':
            for i in range(1, len(BCOLS)+1): ws.cell(row=r, column=i).fill = AMBER
        ws.row_dimensions[r].height = 60
        r += 1
    widths(ws, [30,28,20,15,14,16,10,70,42,17,50,32])

# ---------------------------------------------------------------- DQ tab
ws = wb.create_sheet('DQ candidates')
ws['A1'] = 'DQ-candidates — gated in Part A triage, with the quoted phrase that fired the gate'; ws['A1'].font = TITLE
ws['A2'] = ('These names go into the DQ log so they are never re-screened cold. The quoted phrase is verbatim from the Grata description. '
            'Note: the residential gate is literal per the prompt — a firm listing residential alongside commercial/industrial fires it, '
            'so read the quote before treating a residential DQ as final.')
ws['A2'].font = Font(italic=True, size=9); ws['A2'].alignment = WRAP
header(ws, 4, ['Bucket','Company','Domain','HQ','Grata revenue est.','Grata employee est.','Gate(s) fired with quoted phrase','Source','Grata profile URL'])
r = 5
dq_n = 0
for b in ORDER:
    for c in CAND[b]['candidates']:
        if c['priority'] != 'DQ-candidate': continue
        dq_n += 1
        row = [b, c['name'], c.get('domain') or '[unknown]', c.get('hq') or '[unknown]',
               money(c.get('rev')), num(c.get('emp')), c.get('note',''), c.get('source',''), c.get('url','')]
        for i,x in enumerate(row,1):
            cell = ws.cell(row=r, column=i, value=x); cell.alignment = WRAP; cell.fill = RED; cell.border = THIN
        ws.row_dimensions[r].height = 55
        r += 1
widths(ws, [8,30,28,20,15,14,95,40,32])

# ---------------------------------------------------------------- Public comps
ws = wb.create_sheet('Public comps')
ws['A1'] = 'Part B — Grata public comps per bucket'; ws['A1'].font = TITLE
ws['A2'] = ('SCALE-DISCOUNT CAVEAT: every multiple below comes from a listed company at EMCOR / Quanta / Comfort Systems scale. '
            'These multiples DO NOT transfer to a $2-10M EBITDA private target without a significant size discount. '
            'A lower-middle-market electrical services business does not trade at 25x EBITDA. Treat this tab as directional context '
            'for pricing instinct, not as a pricing model. Amber rows are listed PARENTS of small subsidiaries that Grata surfaced as '
            '"comps" — they are not trade comparables and should be excluded from any read.')
ws['A2'].alignment = WRAP; ws['A2'].fill = AMBER
ws.row_dimensions[2].height = 58

OFF = re.compile(r'goldman sachs|vistra|kkr|heico|acco|ares|ofs capital|amadeus|kingsway|ferguson|sgs |sgs$|littelfuse', re.I)
r = 4
for b in ORDER:
    ws.cell(row=r, column=1, value='%s — %s' % (b, NAMES[b])).font = TITLE
    r += 1
    seeds = COMPS[b]['seeds']
    allc = []
    for s in seeds:
        ws.cell(row=r, column=1, value='Seed: ' + s['seed']).font = BOLD
        r += 1
        if s.get('note'):
            c = ws.cell(row=r, column=1, value='Note: ' + s['note']); c.alignment = WRAP; c.fill = AMBER
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            ws.row_dimensions[r].height = 30
            r += 1
        header(ws, r, ['Comp company (Grata)','Ticker','Listed entity','EV/EBITDA (LTM)','EV/Revenue (LTM)','Flag'])
        ws.freeze_panes = None
        r += 1
        for cp in s['comps']:
            off = bool(OFF.search(cp.get('o') or '')) or bool(OFF.search(cp.get('t') or ''))
            row = [cp['n'], cp.get('t') or '[unknown]', cp.get('o') or '[unknown]',
                   round(cp['ebitda'],2) if cp.get('ebitda') is not None else '[unknown]',
                   round(cp['rev'],2) if cp.get('rev') is not None else '[unknown]',
                   'listed parent, not a trade comparable' if off else '']
            for i,x in enumerate(row,1):
                cell = ws.cell(row=r, column=i, value=x); cell.alignment = WRAP; cell.border = THIN
                if off: cell.fill = AMBER
            if cp.get('ebitda') is not None and not off:
                allc.append((cp['n'], cp['ebitda'], cp.get('rev')))
            r += 1
        gs = s.get('grata_stats') or {}
        if gs:
            txt = 'Grata aggregate for this seed: median EV/EBITDA %s · median EV/Rev %s · min EV/EBITDA %s · max EV/EBITDA %s' % (
                round(gs.get('ev_ebitda_median'),2) if gs.get('ev_ebitda_median') is not None else '[unknown]',
                round(gs.get('ev_rev_median'),2) if gs.get('ev_rev_median') is not None else '[unknown]',
                round(gs.get('ev_ebitda_min'),2) if gs.get('ev_ebitda_min') is not None else '[unknown]',
                round(gs.get('ev_ebitda_max'),2) if gs.get('ev_ebitda_max') is not None else '[unknown]')
            ws.cell(row=r, column=1, value=txt).font = Font(italic=True, size=9)
            r += 1
        r += 1
    # union median across seeds, trade-relevant comps only
    if allc:
        seen = {}
        for n,e,v_ in allc: seen[n] = (e, v_)
        eb = [e for e,_ in seen.values() if e is not None]
        rv = [v_ for _,v_ in seen.values() if v_ is not None]
        ws.cell(row=r, column=1, value='UNION across seeds (trade-relevant comps only, de-duplicated): %d comps · median EV/EBITDA %.2fx · median EV/Rev %.2fx'
                % (len(seen), statistics.median(eb), statistics.median(rv))).font = BOLD
        r += 2
widths(ws, [46,16,34,17,17,34])

# ---------------------------------------------------------------- Market cuts
ws = wb.create_sheet('Market cuts')
ws['A1'] = 'Part C — market cuts per bucket (counts only, page 1 total from each filtered search)'; ws['A1'].font = TITLE
ws['A2'] = ('Counts are the size of the Grata universe matching each bucket\'s primary keyword search plus the common filters, '
            'varying ONE dimension at a time. These are market-structure counts, not candidate lists — they include companies already in the master. '
            'EBITDA proxy: $2-10M EBITDA maps to the $5-50M revenue box per the system doc.')
ws['A2'].alignment = WRAP; ws['A2'].font = Font(italic=True, size=9)
ws['A3'] = 'Seller Intent: ' + CUTS['_meta']['seller_intent']
ws['A4'] = 'Hold period: ' + CUTS['_meta']['hold_period_signal']
ws['A5'] = 'Unknown revenue: ' + CUTS['_meta']['unknown_revenue_note']
for rr in (3,4,5):
    ws.cell(row=rr, column=1).fill = AMBER; ws.cell(row=rr, column=1).alignment = WRAP

r = 7
header(ws, r, ['Bucket','Bucket name','Total (all filters, no cut)',
               'Rev <$5M (nurture zone)','Rev $5-15M','Rev $15-30M','Rev $30-50M','Rev $50M+ (above box)',
               'IN BOX $5-50M','Unknown-revenue pool',
               'Northeast','Southeast','Midwest','Texas + South Central','West',
               'Bootstrapped','PE-backed (consolidator gauge)','Public / public subsidiary',
               'founder_age_signal','hold_period_signal','Seller Intent High'])
r += 1
for b in ORDER:
    d = CUTS[b]; rv = d['revenue']; rg = d['region']; ow = d['ownership']; ex = d['exit']
    row = [b, NAMES[b], d['total'], rv['<5M'], rv['5-15M'], rv['15-30M'], rv['30-50M'], rv['50M+'],
           d['in_box_5_50M'], d['unknown_revenue_pool'],
           rg['Northeast'], rg['Southeast'], rg['Midwest'], rg['Texas+South Central'], rg['West'],
           ow['bootstrapped'], ow['pe_backed'], ow['public+public_sub'],
           ex['founder_age_signal'],
           'unavailable' if not ex.get('hold_period_signal') else ex['hold_period_signal'],
           'unavailable']
    for i,x in enumerate(row,1):
        cell = ws.cell(row=r, column=i, value=x); cell.alignment = TOP; cell.border = THIN
        if i == 9: cell.font = BOLD
        if i in (20,21): cell.fill = AMBER
    r += 1
tots = ['TOTAL','']
for key in range(2, 21):
    pass
for i,fn in enumerate([lambda d: d['total'], lambda d: d['revenue']['<5M'], lambda d: d['revenue']['5-15M'],
                       lambda d: d['revenue']['15-30M'], lambda d: d['revenue']['30-50M'], lambda d: d['revenue']['50M+'],
                       lambda d: d['in_box_5_50M'], lambda d: d['unknown_revenue_pool'],
                       lambda d: d['region']['Northeast'], lambda d: d['region']['Southeast'],
                       lambda d: d['region']['Midwest'], lambda d: d['region']['Texas+South Central'],
                       lambda d: d['region']['West'], lambda d: d['ownership']['bootstrapped'],
                       lambda d: d['ownership']['pe_backed'], lambda d: d['ownership']['public+public_sub'],
                       lambda d: d['exit']['founder_age_signal']]):
    tots.append(sum(fn(CUTS[b]) for b in ORDER))
tots += ['unavailable','unavailable']
for i,x in enumerate(tots,1):
    c = ws.cell(row=r, column=i, value=x); c.font = BOLD
r += 2
ws.cell(row=r, column=1, value='Note: bucket universes overlap (a firm can satisfy more than one keyword profile), so the TOTAL row is a sum of overlapping sets, not a distinct company count.').font = Font(italic=True, size=9)
widths(ws, [8,30]+[16]*19)

wb.save(OUT)
print('wrote', OUT)
print('DQ rows:', dq_n, '| new candidates:', tot['n'])
