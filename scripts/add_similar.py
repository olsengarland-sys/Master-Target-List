import json, sys, re
sys.path.insert(0,'scripts')
IDX=json.load(open('scripts/dedupe_index.json'))
DK,NK=set(IDX['dkeys']),set(IDX['nkeys'])
STORE='scripts/candidates.json'
LEGAL=r'\b(llc|l\.l\.c|inc|incorporated|corp|corporation|co|company|ltd|limited|lp|llp|pllc|pc|plc)\b'
def nd(u):
    if not u: return None
    d=re.sub(r'^https?://','',u.strip().lower()); d=re.sub(r'^www\.','',d)
    return d.split('/')[0].split('?')[0].split('#')[0].strip() or None
def nn(n):
    if not n: return None
    n=n.strip().lower().replace('&',' and '); n=re.sub(r'[.,]',' ',n)
    n=re.sub(LEGAL,' ',n); n=re.sub(r'[^a-z0-9 ]',' ',n)
    return re.sub(r'\s+',' ',n).strip() or None

def add(bucket, source, rows):
    d=json.load(open(STORE))
    seen_d={nd(c['domain']) for b in d for c in d[b]['candidates']}
    seen_n={nn(c['name']) for b in d for c in d[b]['candidates']}
    dropped=0; kept=[]
    for r in rows:
        dk,nk=nd(r['domain']),nn(r['name'])
        if dk in DK or nk in NK: dropped+=1; continue
        if dk in seen_d or nk in seen_n: dropped+=1; continue
        seen_d.add(dk); seen_n.add(nk)
        r['source']=source; kept.append(r)
    b=d.setdefault(bucket,{'candidates':[],'dropped_known':0,'sources':[]})
    b['candidates'].extend(kept)
    b['dropped_known']=b.get('dropped_known',0)+dropped
    b.setdefault('sources',[]).append(source)
    json.dump(d,open(STORE,'w'),indent=1)
    print('%s: kept %d, dropped %d (now %d)'%(bucket,len(kept),dropped,len(b['candidates'])))
