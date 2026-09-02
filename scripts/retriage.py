import json, re

STORE='scripts/candidates.json'
d=json.load(open(STORE))

# Context-aware gates: (regex, label). Manufacturing now requires SELF-manufacture,
# not "manufacturing" appearing as a served customer vertical.
GATES = [
 (r'residential', 'residential'),
 (r'\bsolar\b', 'solar'),
 (r'home ?owner|home services', 'home services'),
 (r'fire alarm', 'fire alarm'),
 (r'structured cabling', 'structured cabling'),
 (r'security system', 'security systems'),
 (r'portfolio company of', 'portfolio company of'),
 (r'franchis', 'franchise'),
 (r'\bdistributor\b|distribution and resale|\breseller\b|resale of', 'distribution/resale'),
 (r'fit-?out|tenant improvement', 'fit-out'),
 (r'low-?voltage only|solely low voltage', 'low-voltage-only'),
 (r'(designs?,? (and )?(develops?,? (and )?)?manufactur\w*)'
  r'|(manufactur\w* (and )?(sells?|distributes?|markets?))'
  r'|(is a .{0,40}manufacturer)|(manufacturer of)|(manufactures? \w)'
  r'|(product manufacturing)|(manufacturing (company|firm|business|operations of its own))',
  'manufacturer (self)'),
 (r'dealership|retail .{0,20}dealer', 'vehicle dealership'),
 (r'\bplumbing\b|\bhvac\b|tree service|arboricultural|roofing', 'off-trade (plumbing/HVAC/roofing/tree)'),
]

MODEL = r'(maintenance|service|testing|repair|commissioning|troubleshoot|preventive|preventative|predictive|emergency|24/7|24 hour|24-hour|rewind|retrofit|inspection|calibration)'

def quote(desc, m):
    s=max(0, m.start()-60); e=min(len(desc), m.end()+60)
    return ('...' if s>0 else '')+desc[s:e].strip()+('...' if e<len(desc) else '')

def gate_hits(desc):
    low=desc.lower(); out=[]
    for pat,label in GATES:
        m=re.search(pat, low)
        if m: out.append((label, quote(desc,m)))
    return out

changed=0
for b,v in d.items():
    for r in v['candidates']:
        desc=r.get('desc') or ''
        hits=gate_hits(desc)
        old=r.get('priority')
        # preserve any ownership review flag already appended
        own_flag=''
        mo=re.search(r'\| REVIEW: ownership.*$', r.get('note',''))
        if mo: own_flag=' '+mo.group(0)
        if hits:
            r['priority']='DQ-candidate'
            r['note']='gate: '+'; '.join('%s -> "%s"'%(l,q) for l,q in hits)+own_flag
        else:
            rev=r.get('rev'); emp=r.get('emp')
            has_model=bool(re.search(MODEL, desc.lower()))
            if emp is not None and emp<10 and not (rev and 5_000_000<=rev<=50_000_000):
                r['priority']='Nurture-candidate'; note='under ~10 employees (Grata est.); revenue not determinably in box'
            elif has_model and rev and 5_000_000<=rev<=50_000_000:
                r['priority']='P2-candidate'; note='service/maintenance/testing language; Grata revenue est. in $5-50M box'
            elif has_model:
                r['priority']='P3-candidate'
                if rev and rev<5_000_000: note='model language present; Grata revenue est. below $5M box'
                elif rev and rev>50_000_000: note='model language present; Grata revenue est. above $50M box'
                else: note='model language present; revenue [unknown] in Grata'
            else:
                r['priority']='P3-candidate'; note='trade fits; model signals absent from description'
            r['note']=note+own_flag
        if r['priority']!=old: changed+=1

json.dump(d, open(STORE,'w'), indent=1)

import collections
pc=collections.Counter(); gc=collections.Counter()
for b,v in d.items():
    for r in v['candidates']:
        pc[r['priority']]+=1
        if r['priority']=='DQ-candidate':
            for l in re.findall(r'(?:gate: |; )([^-]+?) -> ', r['note']): gc[l.strip()]+=1
print('reclassified:',changed)
print('priorities:',dict(pc))
print('gates:',dict(gc))
