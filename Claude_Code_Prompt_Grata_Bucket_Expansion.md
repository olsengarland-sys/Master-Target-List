# Claude Code Prompt — Grata Bucket Expansion, Public Comps, and Market Cuts

Paste everything below the line into Claude Code. Files to have in the working directory are listed in the prompt itself (section 0).

---

You are the sourcing analyst for Hunter Power Services (Olsen Garland, CEO-in-Residence at NextGen Growth Partners, buying one US electrical infrastructure services company as an owner-operator). Your job in this run: use the Grata MCP to (A) expand each target bucket with new companies not already in the master list, (B) pull Grata public comps per bucket, and (C) produce market cuts per bucket by EBITDA-proxy size band, geography, ownership, and exit-readiness signals. Follow the screening system doc for the thesis; this prompt governs the mechanics.

## 0. Inputs in the working directory

1. `Hunter_Power_Master_Target_List_20260825_v3.xlsx` — the master. Relevant tabs:
   - `Buckets (13)` — bucket definitions and bands (header on row 2)
   - `All 1922 master` — every screened company; header on row 2 (index 1); key columns: Company, HQ, Website, Band, Type bucket, Priority, Bucket code, Inven ID
   - `DQ log` and `Nurture (below floor)` — same header convention
2. `Hunter_Power_Screening_System_v1.md` — the thesis, exclusions, size box, signals cheat sheet, and priority verdict rules. Read it fully before any Grata call.
3. `claude_screened-companies-log.csv` — previously screened names beyond the master; include in dedupe.

Load all tabs with pandas using `header=1` and `read_only=True` for sheet discovery with openpyxl. Build the dedupe index before any search: normalized website domain (strip protocol, `www.`, trailing slash, lowercase) as primary key; normalized company name as secondary (strip legal suffixes LLC/Inc/Corp/Co/Ltd/LP/PLLC, replace `&` with `and`, collapse whitespace, lowercase). Include master (all tabs), DQ log, Nurture, and the screened log. A candidate matching either key is not new.

## 1. Grata guardrails (non-negotiable)

- Use `search_companies` (or `ai_search_companies` where available), `find_similar_companies`, `lookup_companies`, `get_public_comps`, and `get_company` only. These are free or low-cost.
- Do NOT call `enrich_companies`, `export_search_results`, `add_search_results_to_list`, `contacts_*`, or any tool marked side-effectful or credit-consuming without stopping and asking first. If a preview shows a credit cost, report it and wait.
- Check `get_token_usage` once at the start and report the balance.
- Every company fact in the output must come from Grata's returned fields or the master; label Grata revenue/employee figures as Grata estimates. Unknowns are written `[unknown]`, never guessed.
- If `get_public_comps` returns a feature-not-enabled error, note it in the gaps section and continue; do not retry.

## 2. Part A — Bucket expansion

Expand these buckets only (core band + secondary band). Skip X9, M13 (manufacturers are a DQ), and E12 (engineering-only is a DQ) — they exist in the master as logs, not hunting grounds.

For each bucket, run `search_companies` with the keyword profile below plus the common filters, then `find_similar_companies` seeded with up to 10 P1/P2 companies from that bucket in the master (resolve seeds via `lookup_companies` on company name + domain; skip seeds that don't resolve cleanly and note them). Pull up to 4 pages (100 results) per search method per bucket, or fewer if results go off-thesis.

**Common filters, every search:**
- `locations: ["US"]`
- `employees_min: 10, employees_max: 250` (screening heuristic per the system doc §3.4: ~$200K–$350K revenue/employee, so the $5M–$50M revenue box ≈ 20–175 FTE; 10–250 gives edge room)
- `ownership_types: ["bootstrapped", "private_subsidiary" excluded]` — use the granular non-PE tokens: exclude `ib_pe_backed`, `pe_add_on`, `public`, `public_subsidiary`. If the API's `private` rollup includes PE-backed tokens, list the safe tokens explicitly instead.
- `year_founded_max: 2016` (10+ years preferred; do not hard-exclude 2016–2021 founders — run a second pass without the founded filter and tag those results "age flag")
- `keywords_exclude: ["residential electrical", "solar installation", "fire alarm", "structured cabling", "security systems", "home services"]`

**Keyword profiles per bucket** (use as `keywords_include` for search_companies; each list is OR within the tool's semantics — run 2–3 keyword variants per bucket if a single call over-narrows):

| Bucket | Keywords |
|---|---|
| T1 Testing & Commissioning | "electrical acceptance testing", "NETA testing", "power system commissioning", "relay testing and calibration", "electrical testing services" |
| T2 Apparatus Service | "circuit breaker repair", "switchgear service", "transformer repair", "electrical apparatus service", "breaker rebuild" |
| T3 Utility Line & Substation | "utility line contractor", "substation construction and maintenance", "distribution line construction", "transmission line services", "powerline maintenance" |
| T4 Industrial Plant Electrical | "industrial electrical services", "plant electrical maintenance", "industrial electrical contractor", "electrical maintenance and troubleshooting" |
| T5 Commercial & Institutional | "commercial electrical contractor", "electrical service and maintenance" (this bucket over-returns; require service/maintenance language and drop pure design-build results at triage) |
| T6 Critical Power & Data Ctr | "critical power services", "data center electrical services", "UPS maintenance", "battery systems service", "mission critical power" |
| S7 Motor & Generator | "electric motor repair", "motor rewind", "generator maintenance and repair", "rotating equipment repair" |
| S8 Traffic Signals & Lighting | "traffic signal construction", "traffic signal maintenance", "street lighting maintenance", "roadway lighting" |
| S9 EV Charging | "EV charging installation", "EVSE service and maintenance", "EV charging infrastructure services" |
| S10 Lightning Protection | "lightning protection systems", "grounding and bonding", "surge protection services" |

**Per candidate returned:** dedupe FIRST, before anything else. The moment a page of results comes back, check every row against the dedupe index (domain key, then normalized name key) and hard-drop any match — a company already in the master, DQ log, Nurture, or screened log is removed immediately and never triaged, never carried into any output tab, never mentioned again. High overlap is fine and expected; do not treat it as a problem or throttle searches because of it. Keep only a single integer per bucket (dropped-as-known count) for the Summary tab. For survivors, apply a light triage using only Grata fields and the system doc:
- Gate check from the Grata description: any exclusion language (residential, solar-only, low-voltage-only, distribution/resale, fit-out, "portfolio company of", franchise) → mark DQ-candidate with the phrase quoted.
- Provisional priority: P2-candidate if description shows service/maintenance/testing language and size fields are in box; P3-candidate if trade fits but model signals are absent from the description; Nurture-candidate if under ~10 employees with the right model language. Never assign P1 from a search description alone — P1 requires the enrichment pass in the system doc §3.6, which is a separate session.

## 3. Part B — Public comps per bucket

For each expanded bucket, pick 2 seed companies: the top-ranked P1 (or P2 if no P1) from the master's `All 1922 master` tab for that bucket code, resolved via `lookup_companies`. Call `get_public_comps` on each seed. From the returned comp sets:

- Record the public comparable companies (name, ticker), EV/Revenue (LTM), EV/EBITDA (LTM), and the aggregate stats (min/median/mean/max) Grata returns.
- Per bucket, report the union of comps across the two seeds and the median EV/EBITDA and EV/Rev, noting which seed produced which set.
- Do NOT build an Excel of GRATA.PUBF formulas — inline values only, this is a chat/workbook deliverable, not an add-in file.
- Note: public comps are valuation benchmarks from listed companies (EMCOR-scale and up); flag clearly that multiples at that scale do not transfer to $2–10M EBITDA private targets without a significant size discount. This is context for Olsen's pricing instincts, not a pricing model.

## 4. Part C — Market cuts per bucket

For each bucket, re-run the bucket's primary keyword search with the common filters, varying one dimension at a time, and record only the result **count** (`total`/`count` field from page 1 — do not paginate these; counts are the deliverable):

1. **EBITDA-proxy size bands** (Grata filters on revenue, not EBITDA; map via the system doc's $2–10M EBITDA ≈ $5–50M revenue box): revenue <$5M (nurture zone), $5–15M, $15–30M, $30–50M (in box, three slices), $50M+ (above box, DQ zone). Also run each band with `revenue_include_unknown: true` once to size the unknown-revenue pool.
2. **Geography**: five US regions — Northeast, Southeast, Midwest, Texas+South Central, West (pass state lists; spell out state names, never two-letter codes).
3. **Ownership**: bootstrapped/independent private vs. PE-backed vs. public-subsidiary counts (the PE-backed count per bucket = consolidator activity gauge).
4. **Exit-readiness**: same search with `exit_readiness_signals: ["founder_age_signal"]` and separately `["hold_period_signal"]` if the package supports it; if the org has Seller Intent, add a `seller_intent_levels: ["High Intent"]` count. If either filter errors as not-in-package, note and skip.

## 5. Output

One workbook, `Hunter_Power_Grata_Expansion_YYYYMMDD.xlsx`, via openpyxl, with the project's fill conventions (FCE4E4 red for DQ-candidates, FFF2CC amber for review flags):

1. **Summary** — new candidates per bucket by provisional priority, dropped-as-known count per bucket (informational only), token balance before/after, and the Confidence and Gaps note (which searches under-returned, which Grata features were unavailable, which seeds failed to resolve).
2. **One tab per bucket** — new candidates only: Company | Domain | HQ (city, state) | Grata revenue est. | Grata employee est. | Ownership | Year founded | Grata description (first 300 chars) | Source (search keywords or similar-to seed) | Provisional priority | Gate/flag notes | Grata profile URL.
3. **DQ candidates** — everything gated in Part A triage, with the quoted phrase that fired the gate, so these names enter the DQ log and never get re-screened cold.
4. **Public comps** — per bucket: seed used, comp companies with tickers, EV/EBITDA and EV/Rev per comp, bucket medians, and the scale-discount caveat.
5. **Market cuts** — one block per bucket: counts by revenue band, region, ownership, exit-readiness/seller-intent, plus the unknown-revenue pool size.

Do not write to any Grata list, do not sync anything to a CRM, do not export from Grata. The workbook is the only artifact. End with the counts, the top 10 new candidates across all buckets by your provisional read, and the three biggest gaps in the run.
