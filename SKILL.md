---
name: audit-records
description: Scans the Record_Collection and Sold tables for likely data-entry errors — typos, out-of-range values, missing hyperlinks, inconsistent values across copies, and lookup-table violations — and writes findings to one or more new timestamped audit sheets, sorted by row with confidence tiers.
---

## When to use

Run when the user invokes `/audit-records` or asks to audit, error-check, or find typos/anomalies in their record collection. The skill scans the `Record_Collection` table (on the `Record Collection` sheet) and the `Sold` table (on the `No Longer in Collection` sheet) and produces one or more findings sheets.

The user's tolerance for false positives is LOW. They have repeatedly asked for fewer, higher-signal findings. When in doubt, suppress rather than flag.

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

Add a 3–4 row summary header above the table: title, run timestamp, total findings, breakdown by confidence tier.

## Implementation environment

**Do everything in `execute_office_js`.** The `code_execution` Python sandbox does NOT have Excel tool functions in this workbook — calling `get_range_as_csv` from Python returns `NameError`. All data reads, all checks (including Levenshtein), and all writes happen in JavaScript via `execute_office_js`.

Hyperlinks must be read via `getCellProperties({address:true, hyperlink:true})` per column — CSV exports do not carry hyperlink URLs.

## Workflow

### Step 1 — Read both tables and lookups in one `execute_office_js` call

For each of `Record Collection` (last row 2520) and `No Longer in Collection` (last row 190):
- Load `A1:AY{lastRow}` values.
- Loop columns AN, AO, AP, AQ separately and call `getCellProperties({address:true, hyperlink:true})` on each.

Also load:
- `Locations!A1:A9`
- `Roles!A1:B40` (column A = role abbreviations)

Do NOT load `Label Abbreviations` (not used for validation) or `Grading` (Grade check is removed).

### Step 2 — Run all checks (see "Checks" section)

Collect findings as objects with: `table`, `row`, `col`, `colLetter`, `confidence`, `issue`, `record`, `column`, `value`, `reason`, `fix`.

Deduplicate: don't report the same (cell, issue) twice.

### Step 3 — Write findings sheet

Sort by Table (Record_Collection first), then by Row.

Summary header rows 1–3, table headers row 5, data from row 6.

Color-code Confidence column via **conditional formatting** on the whole column (NOT per-cell fills — slow on large outputs):

[code block showing JS for conditional formatting on col C: Definite=#F4CCCC, Probable=#FFF2CC, Worth a look=#D9EAD3]

Freeze top 5 rows. Autofit columns, then cap widths: H (Reason) = 320, E (Record) = 180, G (Value) = 200.

### Step 4 — Reply concisely

Short paragraph: total, breakdown by confidence, citation link to new sheet, top-issue counts. Don't paste findings into chat.

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

### Lookup-table validation
- **Location**: must match `Locations!A:A`. Blank allowed. → `Definite error` if non-blank and not in the list.
- **Rating**: blank or integer 1–5. → `Definite error` otherwise.
- **Country**: should be a 2-letter ISO country code. Exception: there is one known record from Scandinavia using a non-ISO label. Accept any value the user has used in ≥ 2 rows. Only flag a value that appears exactly once and isn't 2 uppercase letters. → `Probable error`.
- **Currency**: 3-letter ISO. → `Probable error`.

**Do NOT validate**:
- Label (against `Label Abbreviations` or otherwise — that table is a reference for short forms, not exhaustive).
- Grade (the user has historically not had grade-entry errors; check is disabled).
- Vendor Grade (freeform third-party prose, not the user's lingo).

### Date sanity

- **Recording dates** (Side A/B Date): year must be 1910–today.
  - Before 1910 or in the future → `Definite error`.
  - After 1960 → `Worth a look` (user is surprised if there are any).
- **Date Acquired**: year must be 2014–today. Outside → `Definite error`.
- **Sold On**:
  - In the future → `Definite error`.
  - Compared to `Date Acquired`: see "Date parsing" above. Only flag when **logically impossible**. If either side is imprecise (year-only OR Jan 1), compare years and only flag if `Sold On` year < `Date Acquired` year. If both are precise dates, flag if `Sold On.date < Date Acquired.date`.
- **Side A vs Side B recording-year gap**:
  - ≤ 2 years → no flag.
  - 2–5 years → `Worth a look`.
  - 5–10 years → `Probable error`.
  - > 10 years → `Probable error` (highly likely a typo but not certain).

### Prices
- Negative values in `Original Price`, `Price (USD)`, or `Sold For (USD)` → `Definite error`.
- `Original Price` set but `Price (USD)` blank → `Probable error`.
- Exchange-rate sanity when both `Original Price` and `Price (USD)` are present:
  - GBP: implied rate must be 1.06–1.72. Outside → `Probable error`.
  - EUR: implied rate must be 1.00–1.26. Outside → `Probable error`.
  - Other currencies: skip until ranges are added.

### Booleans
- `Upgrade?`, `Sell?`, `In Rust?`, `Has Prefix`, `Has Suffix`: must be TRUE/FALSE/blank. Other values → `Definite error`.
- `Has Prefix` value vs whether `Prefix` is non-blank (same for Suffix). Mismatch → `Probable error`.

### Cross-table sanity
- Rows in `Record_Collection` with non-blank `Sold On` or `Sold To` → `Definite error` (belong in `Sold`).
- Rows in `Sold` with blank `Sold On` or `Sold To` → `Probable error`.

### Role abbreviations in artist columns

Artist columns use the format `Name-roleabbr` or `Name-r1-r2`. Slashes separate artists.

For each cell in `Side A Artist(s)`, `Side A Pseudonym(s)`, `Other Side A Artist(s)`, and the three Side B equivalents:
1. Split on `/` to get individual artist tokens.
2. For each token: strip `+N`, `(uncredited)`, `(as ...)`, trailing `-pic`, leading `?`.
3. Split remainder on `-`. Walk segments after the first: if a segment matches `/^[a-z]{1,4}\.?$/` AND is not in `validRoles` AND is not in `UNIDENTIFIED_PERFORMERS` → `Probable error` (issue: `Unknown role abbreviation`).

`validRoles`:
- All entries from `Roles!A:A` (column A).
- Plus `"scat"` (vocal style — treated as a role).

`UNIDENTIFIED_PERFORMERS` (not roles, but legitimate tokens that DO NOT trigger the warning):
- `girl`, `boy`, `man`, `woman`, `lady`, `kid`, `child` (descriptors for otherwise-unidentified performers — `John Thorne-girl-v` means John Thorne plus an unidentified girl on vocals).

Allowlist of names containing hyphen-something-that-looks-like-a-role:
- `Million-airs` (band name; the `-airs` is part of the name, not a role).

### Artist near-duplicate detection (Option A + B)

Goal: detect typos like "Jack Buchanon" (1×) vs "Jack Buchanan" (9×) without flagging legitimate distinct names like "Ray Starita" (5×) vs "Rudy Starita" (2×, brothers).

For each artist NAME (parsed via the role tokenizer, then stripping role segments to the bare name; collapse to lowercase letters+digits for comparison):

1. Build a map: collapsed-name → { count, original spellings, list of occurrences with row + opposite-side artist + titles on both sides }.
2. For each name with count ≤ 2 ("rare"):
   - **(A) Asymmetry**: find a "canonical" name appearing ≥ 8× the rare's count, with Levenshtein distance 1–2 on collapsed forms.
   - **(B) Co-occurrence** (must ALSO hold): at least one of the rare name's occurrences must share context with the canonical:
     - the opposite-side artist on that record equals the canonical, OR
     - one of the record's titles also appears on a record under the canonical name.
   - Both (A) and (B) must hold. → `Probable error` (issue: `Artist near-duplicate`).
3. Names shorter than 4 collapsed characters: skip.

### Title cross-copy mismatch

Goal: detect title typos using high-signal cross-copy comparison (a typo of "Puttin' on the Ritz" on copy 1 vs the correct spelling on copy 2 of the SAME record).

1. Group `Record_Collection` rows by `(Label, Prefix, Number, Suffix)` — different copies of the same record.
2. Within each group of 2+ copies, for `Side A Title` and `Side B Title` separately:
   - Collapse each title to lowercase letters+digits (so `"Puttin' on the Ritz"` and `"Puttin' On the Ritz"` collapse the same).
   - If two copies have different collapsed forms → `Probable error` (issue: `Title cross-copy mismatch`).
   - Flag the minority spelling(s); suggest the majority as the fix.
3. The old singleton-vs-common Levenshtein title check is removed — it produced too many false positives across legitimate different songs.

### Vendor / Provenance namespace

Treat `Vendor` and `Provenance` as ONE combined namespace.

1. Pool unique values from both columns across both tables; compute combined frequency.
2. Suppress pairs differing only by a trailing `?` (uncertainty marker — `Henry Parsons` vs `Henry Parsons?` is intentional).
3. Pairs differing only by case/punctuation (collapsed forms match): → `Definite error` (issue: `Vendor/Provenance case/punct dup`).
4. Pairs with Levenshtein ≤ 2 on collapsed forms where one appears ≥ 3× and the other appears ≤ 2×: → `Probable error` (issue: `Vendor/Provenance near-dup`). Report the less-frequent value with the canonical as suggested fix.

### Matrix-number outliers (conservative)

Goal: surface the occasional typo, not a category-wide pattern report.

1. Per-label format signature: replace letter runs with "L", digit runs with "N" (e.g., `W-204` → `L-N`).
2. For each label, count format frequencies (skip `x` — a sentinel meaning "arrived, needs cataloguing").
3. Only flag when ALL of:
   - The label has ≥ 100 matrix entries.
   - The candidate format appears in < 0.5% of that label's matrices.
   - The candidate format appears only 1 or 2 times in absolute terms for that label.
4. → `Worth a look`.

If this still produces many findings, tighten further (raise min-entries threshold or lower share threshold).

### Hyperlink checks (DAHR A, DAHR B, Discogs, 45cat — columns AN/AO/AP/AQ)

- **Missing hyperlink**: cell has text but no native hyperlink → `Definite error`.
- **Cross-copy hyperlink mismatch**: group by `(Label, Prefix, Number, Suffix)`. Within groups of 2+ copies, per column: if some have URLs and some don't, OR distinct URLs across copies → `Probable error`. Report once per group + column.

## Extending the skill

When the user describes a new check:
1. Add a bullet under the appropriate section (or create a new section).
2. Assign confidence honestly. The user's threshold is HIGH — default to `Worth a look` unless you're sure.
3. Give it a short, distinct `Issue type` tag so findings group cleanly.
4. Note known exceptions/allowlist entries directly in the check.
5. After running, ask the user whether the new findings were useful before declaring the check stable.

## Implementation notes

- **Levenshtein**: small JS function with early exit when length difference > 3.
- **Collapse function**: lowercase, then replace anything that isn't a letter or digit — strips all non-alphanumeric for similarity comparison.
- **Column indices (0-based)** for the 51-column tables: 0=Record, 3=Copy, 4=Location, 6=Label, 7=Prefix, 8=Number, 9=Suffix, 11=SideAMatrix, 12=SideAArtist, 13=SideAPseudonym, 14=OtherSideAArtist, 15=SideATitle, 17=SideADate, 18=SideBMatrix, 19=SideBArtist, 20=SideBPseudonym, 21=OtherSideBArtist, 22=SideBTitle, 24=SideBDate, 25=Country, 26=VendorGrade, 27=Grade, 28=Upgrade?, 30=Rating, 31=InRust?, 32=Vendor, 33=DateAcquired, 34=Currency, 35=OriginalPrice, 36=PriceUSD, 37=Provenance, 39=DAHRA (AN), 40=DAHRB (AO), 41=Discogs (AP), 42=45cat (AQ), 43=Sell?, 45=SoldOn, 46=SoldTo, 47=SoldForUSD, 49=HasPrefix, 50=HasSuffix.
- **Sheet names with spaces** in HYPERLINK formulas need single quotes inside the # reference.

## Done when

- One or more sheets named `Audit ...` (with timestamps) exist.
- Each sheet has the 3–4 row summary header, then the findings table sorted by Table then Row.
- All ten columns populated; `Cell` column has working HYPERLINK formulas.
- Confidence column color-coded via conditional formatting.
- Chat reply is a short paragraph with totals, confidence breakdown, top-issue counts, and a citation link to the new sheet.
