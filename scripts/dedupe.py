"""Dedupe index + seed selection for the Hunter Power Grata expansion run."""
import re, json, pandas as pd

XLSX = '/home/user/Master-Target-List/Hunter_Power_Master_Target_List_2026-08-25_v3.xlsx'
CSV  = '/root/.claude/uploads/cdc2f6cb-7804-5be0-bfc8-f894363d3735/d9dfb564-claude_screenedcompanieslog.csv'
OUT  = '/home/user/Master-Target-List/scripts/dedupe_index.json'

LEGAL = r'\b(llc|l\.l\.c|inc|incorporated|corp|corporation|co|company|ltd|limited|lp|llp|pllc|pc|plc)\b'

def norm_domain(url):
    if not isinstance(url, str) or not url.strip():
        return None
    d = re.sub(r'^https?://', '', url.strip().lower())
    d = re.sub(r'^www\.', '', d)
    d = d.split('/')[0].split('?')[0].split('#')[0].strip()
    return d or None

def norm_name(name):
    if not isinstance(name, str) or not name.strip():
        return None
    n = name.strip().lower().replace('&', ' and ')
    n = re.sub(r'[.,]', ' ', n)
    n = re.sub(LEGAL, ' ', n)
    n = re.sub(r'[^a-z0-9 ]', ' ', n)
    return re.sub(r'\s+', ' ', n).strip() or None

# --- every master tab that holds screened names, plus the screened log ---
XL_TABS = ['Batch 1 (250)', 'Batch 2 (250)', 'Batch 3 (147)',
           'NEW - Add to batch 1 or 2 (11)', 'NEW - Add to batch 2 or 3 (51)',
           'Nurture (below floor)', 'DQ log', 'All 1922 master']

def load_all():
    frames = []
    for tab in XL_TABS:
        d = pd.read_excel(XLSX, sheet_name=tab, header=1)
        d = d.rename(columns={'Company': 'company', 'Website': 'website',
                              'Bucket code': 'bucket', 'Priority': 'priority',
                              'Band': 'band', 'Rank': 'rank', 'HQ': 'hq'})
        d['_src'] = tab
        frames.append(d[[c for c in ['company','website','bucket','priority','band','rank','hq','_src'] if c in d.columns]])
    c = pd.read_csv(CSV)
    c = c.rename(columns={'city': 'hq'})
    c['_src'] = 'screened_log_csv'
    frames.append(c[['company','website','bucket','priority','band','rank','hq','_src']])
    df = pd.concat(frames, ignore_index=True)
    df['dkey'] = df['website'].map(norm_domain)
    df['nkey'] = df['company'].map(norm_name)
    return df

def build_index(df):
    return set(df['dkey'].dropna()), set(df['nkey'].dropna())

def is_known(name, domain, dkeys, nkeys):
    """Return matching key type ('domain'/'name'), or None if the candidate is new."""
    d = norm_domain(domain)
    if d and d in dkeys:
        return 'domain'
    n = norm_name(name)
    if n and n in nkeys:
        return 'name'
    return None

if __name__ == '__main__':
    df = load_all()
    dkeys, nkeys = build_index(df)
    print(f'total rows across sources = {len(df)}')
    print(df['_src'].value_counts().to_string())
    print(f'\nunique domain keys = {len(dkeys)}   unique name keys = {len(nkeys)}')

    BUCKETS = ['T1','T2','T3','T4','T5','T6','S7','S8','S9','S10']
    order = {'P1': 0, 'P2': 1}
    seeds = {}
    print()
    for b in BUCKETS:
        sub = df[(df['bucket'] == b) & (df['priority'].isin(['P1','P2'])) & df['dkey'].notna()].copy()
        sub['po'] = sub['priority'].map(order)
        sub['bo'] = (sub['band'] != '1 Core').astype(int)
        sub = sub.sort_values(['po','bo','rank'], na_position='last').drop_duplicates('dkey')
        seeds[b] = sub[['company','dkey','priority','hq']].head(10).to_dict('records')
        print(f'{b:4s}: {len(sub):3d} unique P1/P2 -> {len(seeds[b])} seeds')

    json.dump({'seeds': seeds, 'dkeys': sorted(dkeys), 'nkeys': sorted(nkeys)}, open(OUT, 'w'))
    print(f'\nwrote {OUT}')
