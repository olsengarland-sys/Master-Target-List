"""QC over the enriched contacts. Writes scripts/contacts/contact_qc_flags.json.

Three failure modes seen in the data, each of which would cost real outreach:
  * wrong-entity matches - the provider returned a person in another country
    for a US target, so the contact belongs to a different company entirely;
  * undisclosed ownership - an owner whose verified email sits on another
    corporate domain (CLAUDE.md rule 6: platform ownership labels are unverified);
  * non-owner titles - Inven's title filter matches loosely and lets managers through.
"""
import json, re

C = json.load(open('scripts/contacts/owner_contacts_wave123.json'))['companies']
QUEUE = json.load(open('scripts/contacts_enrich_list.json'))

FREE = {'gmail.com','yahoo.com','hotmail.com','outlook.com','aol.com','icloud.com','msn.com',
        'comcast.net','sbcglobal.net','att.net','verizon.net','bellsouth.net','me.com','live.com',
        'protonmail.com','ymail.com','cox.net','charter.net','earthlink.net','mac.com','mail.com',
        'windstream.net','frontier.com','rocketmail.com'}
# .co is Colombia's TLD but reads as a US vanity domain here, so it is not treated as foreign
FOREIGN_TLD = re.compile(r'\.(co\.uk|uk|ca|au|de|fr|nl|es|it|in|cn|mx|br|ie|nz)$', re.I)
GEO = re.compile(r'\b(UK|United Kingdom|Canada|Canadian|India|Lebanon|Myanmar|Venezuela|Mexico|'
                 r'Alberta|Liberia|Australia|Germany|Ireland|Philippines|Pakistan|Nigeria|Spain|'
                 r'Italy|Netherlands|Brazil|Colombia)\b', re.I)
CAUTION = re.compile(r'wrong company|wrong entity|wrong[- ]person|likely wrong|bad match|mismatch|'
                     r'retired|ownership|acquisi|teamshares|different person|WARNING', re.I)
OWNER = re.compile(r'owner|president|chief executive|\bceo\b|founder|principal|proprietor|'
                   r'managing (director|partner|member)|general manager|\bgm\b|\bpartner\b', re.I)

def base(d):
    return re.sub(r'^www\.', '', (d or '').lower())

verify_first, ownership, non_owner, us_only = [], [], [], []

for dom, v in C.items():
    err = v.get('error') or ''
    if GEO.search(err) or CAUTION.search(err):
        verify_first.append({'domain': dom, 'company': v.get('name'), 'priority': v.get('priority'),
                             'campaign': v.get('campaign'), 'reason': err})
    for c in v.get('contacts') or []:
        title = c.get('title') or ''
        if title and not OWNER.search(title):
            non_owner.append({'domain': dom, 'contact': c.get('name'), 'title': title,
                              'priority': v.get('priority')})
        for e in c.get('emails') or []:
            ed = e['email'].split('@')[-1].lower()
            if ed in FREE or ed == base(dom):
                continue
            stem, edstem = base(dom).split('.')[0], ed.split('.')[0]
            if stem[:6] in ed or edstem[:6] in base(dom):
                continue  # sister/rebrand domain, not a parent
            ownership.append({'domain': dom, 'company': v.get('name'), 'contact': c.get('name'),
                              'title': title, 'alt_email_domain': ed, 'email': e['email'],
                              'priority': v.get('priority'), 'campaign': v.get('campaign')})

for r in QUEUE:
    if FOREIGN_TLD.search(r['domain']):
        us_only.append({'domain': r['domain'], 'name': r.get('name'), 'bucket': r.get('bucket'),
                        'priority': r.get('priority'), 'campaign': r.get('campaign'),
                        'already_enriched': r['domain'] in C})

# A website-builder or directory host is not a company domain: enrichment can never
# resolve it, and domain-keyed dedupe against it is unreliable.
JUNK_HOST = re.compile(r'(godaddysites|wixsite|squarespace|weebly|business\.site|blogspot|'
                       r'wordpress\.com|facebook\.com|linkedin\.com|yelp\.|angi\.|houzz\.|'
                       r'bbb\.org|indeed\.|sites\.google)', re.I)
bad_domains = [{'domain': r['domain'], 'name': r.get('name'), 'priority': r.get('priority'),
                'campaign': r.get('campaign'),
                'note': 'not a company domain - website-builder/directory host'}
               for r in QUEUE if JUNK_HOST.search(r['domain'])]

parents = {}
for o in ownership:
    parents.setdefault(o['alt_email_domain'], []).append(o['domain'])
multi = {k: v for k, v in parents.items() if len(set(v)) > 1}

doc = {
    'generated': '2026-09-01',
    'verify_before_outreach': sorted(verify_first, key=lambda x: str(x['priority'])),
    'verify_note': ('The provider matched a person in another country to a US target - these '
                    'contacts most likely belong to a different company that shares the name. Do '
                    'not mail them without a human check; the contact, not the company, is suspect.'),
    'ownership_signals': ownership,
    'possible_parent_domains': multi,
    'ownership_note': ('An owner whose verified email sits on another corporate domain is a proxy '
                       'for an undisclosed parent. Sister and rebrand domains are filtered out; what '
                       'remains needs the acquisition check CLAUDE.md rule 6 requires before outreach.'),
    'non_owner_titles': non_owner,
    'non_owner_note': ('Inven expands title keywords server-side and matches loosely, so managers '
                       'and VPs come through an owner-title filter. Re-work or downgrade these '
                       'rather than mailing them as owner contacts.'),
    'us_only_violations': us_only,
    'us_only_note': ('The mandate is a US company. .co was not treated as foreign - it is Colombia\'s '
                     'TLD but reads as a US vanity domain here. Domain TLD alone misses the bigger '
                     'problem, which is foreign contacts on .com targets; those are in '
                     'verify_before_outreach.'),
    'bad_domains': bad_domains,
    'bad_domain_note': ('A website-builder or directory host is not a company domain. Enrichment '
                        'can never resolve one, and dedupe keyed on it is unreliable, so the real '
                        'domain has to be found before the row is usable.'),
    'counts': {'verify_before_outreach': len(verify_first), 'ownership_signals': len(ownership),
               'non_owner_titles': len(non_owner), 'us_only_violations': len(us_only),
               'bad_domains': len(bad_domains)},
}
json.dump(doc, open('scripts/contacts/contact_qc_flags.json', 'w'), indent=1)
print(doc['counts'])
print('possible parent domains:', multi)
