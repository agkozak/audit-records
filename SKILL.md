---
name: audit-records
description: Scans the Record_Collection and Sold tables for likely data-entry errors — typos, out-of-range values, missing hyperlinks, inconsistent values across copies, and lookup-table violations — and writes findings to one or more new timestamped audit sheets, sorted by row with confidence tiers.
---

## When to use

Run when the user invokes `/audit-records` or asks to audit, error-check, or find typos/anomalies in their record collection. The skill scans the `Record_Collection` table (on the `Record Collection` sheet) and the `Sold` table (on the `No Longer in Collection` sheet) and produces one or more findings sheets.

## Output

Create NEW worksheets each run. Never overwrite a previous audit sheet.

**Default**: a single sheet named `Audit YYYY-MM-DD HHMM`.

**Split into multiple sheets** when it makes the findings easier to act on — e.g., a separate sheet for hyperlink issues, one for cross-copy consistency, one for date/price anomalies. Use judgment: if one category has 50+ findings or has a very different shape from the others, give it its own sheet.

The findings table on each sheet has these columns:

| Column | Purpose |
|---|---|
| Row | Row number in the source sheet (header is row 1, so first data row is row 2) |
| Table | `Record_Collection` or `Sold` |
| Confidence | `Definite error` / `Probable error` / `Worth a look` |
| Issue type | Short tag for the rule that fired |
| Record | Value from the `Record` column for context |
| Column | The column header where the suspect value lives |
| Value | The actual cell value |
| Reason | One-sentence explanation of why this looks wrong |
| Suggested fix | If obvious, otherwise blank |
| Cell | `=HYPERLINK("#'Sheet'!A1","Sheet!A1")` so the user can click through |

**Sort by Row** within each table (Record_Collection first, then Sold). Confidence is a column, not the sort key.

Add a 3–4 row summary header above the table: title, run timestamp, total findings, breakdown by confidence tier. If any required non-core columns were missing (see "Header resolution" below), add a banner row listing the skipped checks.

## Implementation environment

**Do everything in `execute_office_js`.** The `code_execution` Python sandbox does NOT have Excel tool functions in this workbook — calling `get_range_as_csv` from Python returns `NameError`. All data reads, all checks (including Levenshtein), and all writes happen in JavaScript via `execute_office_js`.

Hyperlinks must be read via `getCellProperties` per column — CSV exports do not carry hyperlink URLs. The property name is **singular `hyperlink`** in both the request and the response (NOT `hyperlinks`):

- Request: `range.getCellProperties({address: true, format: {fill: {color: true}}, hyperlink: true})` — only `address` and `hyperlink` are actually needed for this skill, but pass them as `{address: true, hyperlink: true}`.
- Response: each cell's entry has a `hyperlink` field that is either absent/empty `{}` (no link) or an object with an `address` property holding the URL. Read `cell.hyperlink && cell.hyperlink.address` — do not look for `hyperlinks` (plural), it does not exist.

## Header resolution — read this before anything else

**Never use hard-coded column letters or indices.** Both source sheets are wide and the user adds columns over time; absolute references silently break when that happens. Resolve every column by its header name at runtime.

### How it works

1. Read row 1 of each source sheet to get the header strings actually present.
2. For each header in the lookup table below, find its column index by **trimmed, case-insensitive** match against the live headers (so a stray space or a `sold on` → `Sold On` cap change doesn't break the run, but a genuine rename like `Date Sold` does).
3. Build a `cols` map: `cols.Record`, `cols.SideADate`, `cols.PriceUSD`, etc. — variable-name keys mapping to 0-based column indices. Every check below refers to columns through this map, never by letter.
4. Compute column letters from indices when needed (for `HYPERLINK` formulas and `getCellProperties` ranges) — don't hard-code `AN`/`AO`/`AP`/`AQ`.

### What to do when a header is missing

Columns are tiered:

**Core columns** — load-bearing for grouping, identification, and cross-table logic. If ANY of these is missing in either source sheet, **abort the run** and reply with a clear message naming the missing header(s) and the sheet(s) they're missing from. Do not produce a partial audit.

Core:
`Record`, `Copy`, `Label`, `Prefix`, `Number`, `Suffix`, `Side A Title`, `Side B Title`, `Side A Date`, `Side B Date`, `Date Acquired`, `Sold On`, `Sold To`, `Original Price`, `Price (USD)`, `Sold For (USD)`, `Currency`.

**Optional columns** — power individual checks. If one is missing, **skip the checks that need it**, record which checks were skipped, and proceed. Include a banner in the summary header of the audit sheet listing the skipped checks so the user knows.

Optional (with the checks they gate, for reference):
- `Location` → Location lookup check
- `Rating` → Rating range check
- `Country` → Country-code check
- `Side A Artist(s)`, `Side A Pseudonym(s)`, `Other Side A Artist(s)`, `Side B Artist(s)`, `Side B Pseudonym(s)`, `Other Side B Artist(s)` → role-abbreviation check, artist near-duplicate check (each artist column is independently optional; checks run over whichever subset is present)
- `Upgrade?`, `Sell?`, `In Rust?`, `Has Prefix`, `Has Suffix` → boolean validation
- `Prefix` paired with `Has Prefix`, `Suffix` paired with `Has Suffix` → boolean-vs-value consistency
- `Vendor`, `Provenance` → vendor/provenance near-dup check (needs both)
- `Side A Matrix Number`, `Side B Matrix Number` → matrix-number outlier check
- `DAHR A`, `DAHR B`, `Discogs`, `45cat` → hyperlink checks (each independently optional)

### Header registry (canonical names, exactly as they appear in row 1)

Use these strings verbatim for the lookup. The keys in the bullet lists are the `cols.*` variable names; the values are the live header strings to match against.

Core:
- `Record` → `Record`
- `Copy` → `Copy`
- `Label` → `Label`
- `Prefix` → `Prefix`
- `Number` → `Number`
- `Suffix` → `Suffix`
- `SideATitle` → `Side A Title`
- `SideBTitle` → `Side B Title`
- `SideADate` → `Side A Date`
- `SideBDate` → `Side B Date`
- `DateAcquired` → `Date Acquired`
- `SoldOn` → `Sold On`
- `SoldTo` → `Sold To`
- `OriginalPrice` → `Original Price`
- `PriceUSD` → `Price (USD)`
- `SoldForUSD` → `Sold For (USD)`
- `Currency` → `Currency`

Optional:
- `Location` → `Location`
- `Rating` → `Rating`
- `Country` → `Country`
- `SideAArtist` → `Side A Artist(s)`
- `SideAPseudonym` → `Side A Pseudonym(s)`
- `OtherSideAArtist` → `Other Side A Artist(s)`
- `SideBArtist` → `Side B Artist(s)`
- `SideBPseudonym` → `Side B Pseudonym(s)`
- `OtherSideBArtist` → `Other Side B Artist(s)`
- `Upgrade` → `Upgrade?`
- `Sell` → `Sell?`
- `InRust` → `In Rust?`
- `HasPrefix` → `Has Prefix`
- `HasSuffix` → `Has Suffix`
- `Vendor` → `Vendor`
- `Provenance` → `Provenance`
- `SideAMatrix` → `Side A Matrix Number`
- `SideBMatrix` → `Side B Matrix Number`
- `DAHR_A` → `DAHR A`
- `DAHR_B` → `DAHR B`
- `Discogs` → `Discogs`
- `Cat45` → `45cat`

Headers that exist in the workbook but the skill does not need (do NOT add lookups for these, do NOT depend on their position): `S`, `Details`, `Size`, `On Shelves`, `Side A Recording Location`, `Side B Recording Location`, `Vendor Grade`, `Grade`, `Playback Notes`, `Notes`, `Sale Notes`, `Label Sort`.

## Workflow

### Step 1 — Resolve headers, then read both tables and lookups in one `execute_office_js` call

For each source sheet (`Record Collection`, last row 2520; `No Longer in Collection`, last row 190):
- Read row 1, build the `cols` map (per "Header resolution" above), apply the core/optional logic. If a core column is missing on either sheet, abort and report.
- Load `A1:{lastColLetter}{lastRow}` values, where `{lastColLetter}` is computed from the actual width of row 1 (not hard-coded).
- For each hyperlink column that resolved (`DAHR A`, `DAHR B`, `Discogs`, `45cat`), call `getCellProperties({address: true, hyperlink: true})` on that column's range using its resolved letter. Read the URL as `cell.hyperlink && cell.hyperlink.address` (singular — see "Implementation environment" above).

Also load:
- `Locations!A1:A9`
- `Roles!A1:B40` (column A = role abbreviations)

Do NOT load `Label Abbreviations` (not used for validation) or `Grading` (Grade check is removed).

### Step 2 — Run all checks (see "Checks" section)

Collect findings as objects with: `table`, `row`, `col`, `colLetter`, `confidence`, `issue`, `record`, `column`, `value`, `reason`, `fix`.

**Deduplicate** before writing. The dedup key is `(table, row, col, issue)` — i.e. one finding per (cell, issue type). If a single check naturally produces multiple sub-findings for the same cell (e.g. multiple bad role segments in one artist cell), the check itself must consolidate them into one finding before adding to the list (see the role-abbreviation check for the pattern). Different issue types on the same cell are not duplicates and should each be reported.

Track which optional checks were skipped due to missing headers; pass the list to Step 3.

### Step 3 — Write findings sheet

Sort by Table (Record_Collection first), then by Row.

Summary header rows 1–3 (title, timestamp, totals/breakdown). If any checks were skipped, add a row 4 banner like `Skipped checks (missing columns): Rating range, Country code` and shift the table down accordingly. Table headers on the next row, data starting the row after.

Color-code Confidence column via **conditional formatting** on the whole column (NOT per-cell fills — slow on large outputs):
- `Definite error` → `#F4CCCC`
- `Probable error` → `#FFF2CC`
- `Worth a look` → `#D9EAD3`

Freeze the top header rows. Autofit columns, then cap widths: Reason = 320, Record = 180, Value = 200.

### Step 4 — Reply concisely

Short paragraph: total, breakdown by confidence, citation link to new sheet, top-issue counts, and a note if any checks were skipped. Don't paste findings into chat.

## Date parsing — CRITICAL

The date columns (`Side A Date`, `Side B Date`, `Date Acquired`, `Sold On`) mix two formats. The user enters only a year (e.g. `1929`) when the exact date is unknown. Other cells contain full Excel serial dates.

**Parser heuristic**: if cell value is an integer 1880–2030, treat as a year-only marker. Otherwise treat as Excel serial.

**Use UTC throughout** — Excel serial dates convert to UTC midnight. Local-time methods (`getMonth()`, `getDate()`) can be off by one day depending on the user's timezone. Use `getUTCFullYear()`, `getUTCMonth()`, `getUTCDate()`.

**Jan 1 of year Y is ALSO an "imprecise year" marker** — when the user enters `2022` Excel may store it as serial 44562 (= 2022-01-01). Treat any precise Jan 1 the same as a year-only value for the purpose of comparison against more precise dates.

**Never** flag two dates as inconsistent when one is imprecise (year-only OR Jan 1) and the years are equal — even if one side is a precise mid-year date. "Bought Aug 23 2021, sold sometime in 2021" is logically consistent.

## Checks

Calibrate confidence honestly:
- **Definite error**: rule violation with no plausible legitimate cause.
- **Probable error**: looks wrong but could be intentional.
- **Worth a look**: unusual pattern worth a glance, frequent false positives expected.

Every check below names the columns it needs by `cols.*` key. If any required column for a check is missing (per the optional-column rules), skip that check and record it in the skipped list.

### Lookup-table validation
- **Location** (`cols.Location`): must match `Locations!A:A`. Blank allowed. → `Definite error` if non-blank and not in the list.
- **Rating** (`cols.Rating`): blank or integer 1–5. → `Definite error` otherwise.
- **Country** (`cols.Country`): should be a 2-letter ISO country code. Exception: there is one known record from Scandinavia using a non-ISO label. Accept any value the user has used in ≥ 2 rows. Only flag a value that appears exactly once and isn't 2 uppercase letters. → `Probable error`.
- **Currency** (`cols.Currency`): 3-letter ISO. → `Probable error`.

**Do NOT validate**:
- Label (against `Label Abbreviations` or otherwise — that table is a reference for short forms, not exhaustive).
- Grade (the user has historically not had grade-entry errors; check is disabled).
- Vendor Grade (freeform third-party prose, not the user's lingo).

### Date sanity

- **Recording dates** (`cols.SideADate`, `cols.SideBDate`): year must be 1910–today.
  - Before 1910 or in the future → `Probable error`.
  - After 1960 → `Probable error` (user is surprised if there are any).
- **Date Acquired** (`cols.DateAcquired`): year must be 2014–today. Outside → `Definite error`.
- **Sold On** (`cols.SoldOn`):
  - In the future → `Definite error`.
  - Compared to `Date Acquired`: see "Date parsing" above. Only flag when **logically impossible**. If either side is imprecise (year-only OR Jan 1), compare years and only flag if `Sold On` year < `Date Acquired` year. If both are precise dates, flag if `Sold On.date < Date Acquired.date`.
- **Side A vs Side B recording-year gap** (needs both `cols.SideADate` and `cols.SideBDate`):
  - ≤ 2 years → no flag.
  - 2–5 years → `Worth a look`.
  - 5–10 years → `Probable error`.
  - > 10 years → `Probable error` (highly likely a typo but not certain).

### Prices
- Negative values in `cols.OriginalPrice`, `cols.PriceUSD`, or `cols.SoldForUSD` → `Definite error`.
- `cols.OriginalPrice` set but `cols.PriceUSD` blank → `Probable error`.
- Exchange-rate sanity when both `cols.OriginalPrice` and `cols.PriceUSD` are present (and `cols.Currency` is present):
  - GBP: implied rate must be 1.06–1.72. Outside → `Probable error`.
  - EUR: implied rate must be 1.00–1.26. Outside → `Probable error`.
  - Other currencies: skip until ranges are added.

### Booleans
- `cols.Upgrade`, `cols.Sell`, `cols.InRust`, `cols.HasPrefix`, `cols.HasSuffix`: must be TRUE/FALSE/blank. Other values → `Definite error`.
- `cols.HasPrefix` value vs whether `cols.Prefix` is non-blank (same for Suffix). Mismatch → `Probable error`. Skip this sub-check for either pair if its boolean column is missing.

### Cross-table sanity
- Rows in `Record_Collection` with non-blank `cols.SoldOn` or `cols.SoldTo` → `Definite error` (belong in `Sold`).
- Rows in `Sold` with blank `cols.SoldOn` or `cols.SoldTo` → `Probable error`.

### Role abbreviations in artist columns

Artist columns use the format `Name-roleabbr` or `Name-r1-r2`. Slashes separate artists.

For each cell in the six artist columns that resolved (`cols.SideAArtist`, `cols.SideAPseudonym`, `cols.OtherSideAArtist`, `cols.SideBArtist`, `cols.SideBPseudonym`, `cols.OtherSideBArtist` — process whichever subset is present):
1. Split on `/` to get individual artist tokens.
2. For each token: strip `+N`, `(uncredited)`, `(as ...)`, trailing `-pic`, leading `?`.
3. Split remainder on `-`. Walk segments after the first: if a segment matches the regex `/^[a-z]{1,4}\.?$/` **(case-sensitive — no `/i` flag)** AND is not in `validRoles` AND is not in `UNIDENTIFIED_PERFORMERS` → record the bad segment for this cell.
4. Emit **one finding per cell**, not one per bad segment. If a cell has multiple bad segments, list them together in the Reason (e.g., `Unknown role abbreviations: zzz, qqq`). Use a single issue tag: `Unknown role abbreviation`. Confidence: `Probable error`.

**Why case-sensitive matters**: role abbreviations are always lowercase by convention (`v`, `p`, `tb`, `dir.`). Hyphenated surnames like `Scott-Wood`, `Lloyd-Jones`, or `Conn-Smythe` contain capitalized segments that look role-shaped but aren't. Lowercase-only matching handles those naturally with no allowlist needed. The allowlist exists only for genuine lowercase ambiguities (e.g. `Million-airs`).

`validRoles`:
- All entries from `Roles!A:A` (column A).
- Plus `"scat"` (vocal style — treated as a role).

`UNIDENTIFIED_PERFORMERS` (not roles, but legitimate tokens that DO NOT trigger the warning):
- `girl`, `child`, `another` (descriptors for otherwise-unidentified performers — `John Thorne-girl-v` means John Thorne plus an unidentified girl on vocals).

Allowlist of names containing hyphen-something-that-looks-like-a-role:
- `Million-airs` (band name; the `-airs` is part of the name, not a role).

### Artist near-duplicate detection (Option A + B)

Goal: detect typos like "Jack Buchanon" (1×) vs "Jack Buchanan" (9×) without flagging legitimate distinct names like "Ray Starita" (5×) vs "Rudy Starita" (2×, brothers).

Runs over whichever of the six artist columns resolved.

For each artist NAME (parsed via the role tokenizer, then stripping role segments to the bare name; collapse to lowercase letters+digits for comparison):

1. Build a map: collapsed-name → { count, original spellings, list of occurrences with **source table tag** (`"RC"` or `"Sold"`) + row + opposite-side artist + titles on both sides }. See "Implementation notes" for why the source-table tag is mandatory.
2. For each name with count ≤ 2 ("rare"):
   - **(A) Asymmetry**: find a "canonical" name appearing ≥ 8× the rare's count, with Levenshtein distance 1–2 on collapsed forms.
   - **(B) Co-occurrence** (must ALSO hold): at least one of the rare name's occurrences must share context with the canonical:
     - the opposite-side artist on that record equals the canonical, OR
     - one of the record's titles also appears on a record under the canonical name.
   - Both (A) and (B) must hold. → `Probable error` (issue: `Artist near-duplicate`).
3. Names shorter than 4 collapsed characters: skip.

### Title cross-copy mismatch

Goal: detect title typos using high-signal cross-copy comparison (a typo of "Puttin' on the Ritz" on copy 1 vs the correct spelling on copy 2 of the SAME record).

1. Group `Record_Collection` rows by `(cols.Label, cols.Prefix, cols.Number, cols.Suffix)` — different copies of the same record.
2. Within each group of 2+ copies, for `cols.SideATitle` and `cols.SideBTitle` separately:
   - Collapse each title to lowercase letters+digits (so `"Puttin' on the Ritz"` and `"Puttin' On the Ritz"` collapse the same).
   - If two copies have different collapsed forms → `Probable error` (issue: `Title cross-copy mismatch`).
   - Flag the minority spelling(s); suggest the majority as the fix.
3. The old singleton-vs-common Levenshtein title check is removed — it produced too many false positives across legitimate different songs.

### Vendor / Provenance namespace

Needs both `cols.Vendor` and `cols.Provenance`. Treat them as ONE combined namespace.

1. Pool unique values from both columns across both tables; compute combined frequency.
2. Suppress pairs differing only by a trailing `?` (uncertainty marker — `Henry Parsons` vs `Henry Parsons?` is intentional).
3. Pairs differing only by case/punctuation (collapsed forms match): → `Definite error` (issue: `Vendor/Provenance case/punct dup`).
4. Pairs with Levenshtein ≤ 2 on collapsed forms where one appears ≥ 3× and the other appears ≤ 2×: → `Probable error` (issue: `Vendor/Provenance near-dup`). Report the less-frequent value with the canonical as suggested fix.

### Matrix-number outliers (conservative)

Needs `cols.SideAMatrix` and/or `cols.SideBMatrix` (runs over whichever resolved). Goal: surface the occasional typo, not a category-wide pattern report.

1. Per-label format signature: replace letter runs with "L", digit runs with "N" (e.g., `W-204` → `L-N`).
2. For each label, count format frequencies (skip `x` — a sentinel meaning "arrived, needs cataloguing").
3. Only flag when ALL of:
   - The label has ≥ 100 matrix entries.
   - The candidate format appears in < 0.5% of that label's matrices.
   - The candidate format appears only 1 or 2 times in absolute terms for that label.
4. → `Worth a look`.

If this still produces many findings, tighten further (raise min-entries threshold or lower share threshold).

### Hyperlink checks (`cols.DAHR_A`, `cols.DAHR_B`, `cols.Discogs`, `cols.Cat45`)

Compute the column letter for each resolved hyperlink column from its index; pass that letter into `getCellProperties`. Each column is independently optional — process whichever subset resolved.

- **Missing hyperlink**: cell has text but no native hyperlink → `Definite error`.
- **Cross-copy hyperlink mismatch**: group by `(cols.Label, cols.Prefix, cols.Number, cols.Suffix)`. Within groups of 2+ copies, per column: if some have URLs and some don't, OR distinct URLs across copies → `Probable error`. Report once per group + column.

## Extending the skill

When the user describes a new check:
1. Add a bullet under the appropriate section (or create a new section).
2. Assign confidence honestly.
3. Give it a short, distinct `Issue type` tag so findings group cleanly.
4. Note known exceptions/allowlist entries directly in the check.
5. If the check needs a new column, add it to the header registry above and mark it core or optional.
6. After running, ask the user whether the new findings were useful before declaring the check stable.

## Implementation notes

- **Header resolution comes first.** Every column reference in generated code goes through `cols.*` — never a literal letter or index. To go from index to letter (for `HYPERLINK` formulas and `getCellProperties` ranges), use a small helper that handles two-letter columns (`AA`–`ZZ`).
- **Cross-table data structures must carry source-table tags.** Any map or list that pools rows from both `Record_Collection` and `Sold` (artist near-duplicate, vendor/provenance, and any future cross-table check) must store each entry with its source table identifier (`"RC"` or `"Sold"`) alongside the row number. When looking up a row later, dispatch to the correct data array based on that tag — never assume an occurrence came from Record_Collection. Indexing `rcData[occ.row - 2]` for a Sold-table occurrence silently pulls the wrong record.
- **Levenshtein**: small JS function with early exit when length difference > 3.
- **Collapse function**: lowercase, then replace anything that isn't a letter or digit — strips all non-alphanumeric for similarity comparison.
- **Header matching**: trim whitespace, lowercase, then compare. Build the live-headers map once per sheet from row 1.
- **Sheet names with spaces** in HYPERLINK formulas need single quotes inside the # reference.

## Done when

- One or more sheets named `Audit ...` (with timestamps) exist.
- Each sheet has the summary header (with a skipped-checks banner if applicable), then the findings table sorted by Table then Row.
- All ten columns populated; `Cell` column has working HYPERLINK formulas.
- Confidence column color-coded via conditional formatting.
- Chat reply is a short paragraph with totals, confidence breakdown, top-issue counts, a citation link to the new sheet, and a note if any checks were skipped.
