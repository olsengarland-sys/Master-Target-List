# Master-Target-List — Hunter Power Services sourcing

Olsen Garland (NextGen Growth Partners) is buying ONE US electrical infrastructure
services company as an owner-operator. This repo holds the master target list, the
Grata/Inven expansion pipeline (`scripts/`), and the output workbook. The thesis
lives in the Target Screening System doc; buckets are T1-T6, S7-S10 (X9/M13/E12
are DQ logs, never hunting grounds).

## Sourcing rules — validated against outreach results 2026-09-01

An external validation of 833 campaign-assigned companies
(`scripts/validation/campaign_validation_20260901.json` holds every per-domain
verdict) found 32% correctly filed, 29% right target but wrong campaign, 31%
should never have been assigned. These rules exist so the next pass doesn't
repeat those errors. Do not weaken them without the client's say-so.

1. **Assign bucket by what the company IS, not by which search surfaced it.**
   Filing a company under the bucket whose keywords found it produced 0-7%
   validity in the specialist buckets (Lightning Protection 0%, Apparatus 5%,
   Testing & Commissioning 7%) while generalist contractors belonged in T5.
   Classify every candidate against the bucket trade definitions from its own
   description; a company surfaced by an S10 search that reads as a commercial
   contractor is T5. Any verifier/review agent MUST be given the full T1-T6 /
   S7-S10 bucket definitions so it can re-home rows (the wave-2 verifiers were
   not, and re-homing silently failed). Biggest validated bleed flows:
   T2 -> S7 (motor shops), T1/T4/T6/S8/S9/S10 -> T5 (generalists).

2. **A platform revenue estimate alone cannot carry a size call.** Median implied
   revenue/employee across the validated file was $559K against the thesis band
   of $200-350K; estimates were wrong in both directions. P2 requires the
   revenue estimate in $5-50M AND the headcount-implied band (emp x $200-350K)
   to overlap the box; contradictions are labelled SIZE CONFLICT and capped at
   P3 (implemented in `scripts/xref/triage_xref.py`). Above $700K/employee,
   flag the estimate as inflated. Headcount itself is NOT an investment
   criterion (client rule) — it is the cross-check on revenue, nothing more.

3. **Never cut a Nurture/holdback list on a single size field.** The source
   Nurture sheet was cut on revenue alone; 115 of its 212 rows have headcounts
   that imply >=$5M at $350K/emp. Re-cut with the same double-lock before
   treating Nurture as dead.

4. **Store FULL descriptions in the data files.** The 300-char truncation in
   earlier candidate stores directly caused 47 UNRESOLVED validation verdicts
   and degraded a verifier pass. Truncate for workbook display only.

5. **Validated DQ patterns now gated in triage:** cathodic protection /
   corrosion control (not lightning protection — 3 firms web-verified out),
   electrical-safety products/training firms, warranty-recovery/repair-management
   administrators. Keep the residential gate literal (client's standing choice).

6. **Ownership labels are unverified.** Platform "Bootstrapped" does not clear
   a recent acquisition; check for one before any name reaches outreach (§5).
   "Investor Backed" rows carry REVIEW flags, never clean P2s.

7. **Dedupe authority:** `scripts/xref/known_keys.json` (domains + normalized
   names) covers master, DQ log, Nurture, screened log, outreach contacts,
   waves 1-3, and must be extended with every new wave AND with the validation
   file's domains. The validator also dedupes against a Dealbuff
   active-outreach list this repo does not hold — ask the client for it before
   the next expansion, or note its absence.

8. **Platform notes:** Grata `locations` needs `"United States"` / `"State, US"`
   (bare `"US"` errors). Seller Intent and `hold_period_signal` are not in the
   Grata package. Inven searches cost ~1 visible screening credit per call plus
   1 export-volume row per returned row (export pool is NOT visible via API —
   spend it deliberately). Inven `exclude_list_ids` [138892, 138756, 138899]
   are the org's saved lists. Grata founded-year filters can translate to a
   company-age range with a silent floor (T4 lost pre-1965 firms); check
   `filters_used` on every search.

## Conventions

- Workbook fills: FCE4E4 red = DQ-candidate, FFF2CC amber = review flag.
- Unknowns are written `[unknown]`, never guessed; revenue/employee figures are
  always labelled as estimates with their source platform.
- P1 is never assigned from a search description — it requires the enrichment
  pass (system doc §3.6).
- No Grata/Inven contact, enrichment, CRM, list-write or export tool without
  asking first. The workbook is the only artifact.

## Contact enrichment — platform capabilities (verified 2026-09-01)

Established by running both platforms over ~880 companies. Do not re-litigate
these by assumption; they were measured.

9. **Phone numbers come from Inven only.** Grata's contacts tool
   (`contacts_companies_contacts_create`) returns `"phones": null` on every
   contact — it carries executives, titles, verified work emails, LinkedIn
   URLs and locations, but no phone data at all. Inven
   `get_company_contacts` returns phones already TYPED as
   `mobile` / `direct dial` / `office`, which is what makes an owner-cell
   campaign possible. So "fall back to Grata for the cell number" cannot
   work; Grata is the email/org-chart source, Inven the phone source.
   Inven's *company* records (`get_company_info`) also carry no phone — 0 of
   817 returned one — so there is no free HQ-mainline reference either.
10. **Always pass the owner-title filter to Inven `get_company_contacts`.**
   Title-filter misses and no-coverage both cost 0 credits, so the filter is
   free protection; an unfiltered pull spends credits on office staff.
   Measured economics with the filter: ~0.56 contact credits per company
   attempted, ~55% yield, ~47% of attempts producing a mobile number.
   Provider budget is ~8 units/call, and a titled domain costs
   1 + max_contacts + 1, so 2 domains per call is the ceiling.
11. **Inven `named_people` mode is free when the contact is already resolved**
   and is the most accurate path. Run named lookups FIRST in any future wave,
   not last — every named call in waves so far cost 0 credits.
12. **Both platforms return wrong-entity matches for US targets** — a person in
   the UK, India, Venezuela, Mexico, Canada or Germany attached to a US
   company that shares the name. Check the contact's `location` before any
   name reaches outreach; ~25 were caught this way. Foreign TLDs are the easy
   case; foreign contacts on `.com` targets are the dangerous one.
13. **Web egress is blocked in the remote session** (EGRESS_BLOCKED / HTTP 403
   from the proxy, confirmed against control hosts). Any plan that depends on
   fetching a company's own website — mainline-phone verification, site
   scraping — cannot run there. Do not write a placeholder result file for a
   check that never ran; an all-null file reads downstream as "checked, no
   match" and silently clears the rows it was meant to question.
