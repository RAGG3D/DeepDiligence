# Catalyst and Historical Events workflow

This is a required, fail-closed part of every new-ticker build. A workbook is not
deliverable when the official press-release audit, full drug×indication Catalyst
framework, source hyperlinks, or zero-residue checks are incomplete.
The pinned canonical tree is `00_WORKFLOW_STRUCTURE.md`; every workflow change
must update it in the same change.

## Always-active research database

Every research workflow scans the shared fact ledger before external research,
passes existing facts and completeness gaps into the research brief, and
immediately upserts every new/corrected fact into both the durable JSON seed and
DuckDB. The scan remains active throughout research. For every in-scope
competitor, maximize ORR, CR, PR, mPFS/PFS, mOS/OS, safety overview, named
adverse events and rates, dose/regimen, actual price and forecast price. Also
capture reported DoR, EFS, DFS and RFS. Store missing fields as
`unavailable` with source and reason; never invent them. Preserve conflicts.
Test research receives only clinical/dose/trial fields even though the database
may retain commercial facts for other workflows.

## “event EVENT”

This command builds a price-blind ticker universe for a completed or future
clinical meeting. It does not modify a workbook.

1. Resolve the exact official meeting name, edition, dates, embargo and
   presentation schedule.
2. For a completed event, crawl the official program/abstract archive and every
   candidate issuer to verify an actual quantitative human clinical disclosure.
   For a future event, require an official announcement or accepted abstract
   explicitly promising quantitative human clinical data.
3. Require at least one important patient-level metric: ORR/CR/PR, DCR/CBR,
   DoR, PFS/mPFS, OS/mOS, EFS/DFS/RFS, MRD, validated patient outcome or
   quantitative safety result. Attendance-only, trial-design, enrollment-only,
   preclinical and nonquantitative items are excluded.
4. Verify event-cutoff listing on Nasdaq, NYSE or NYSE American and biotech or
   biopharma classification. Use the historical ticker for past events.
5. Run the `<USD 1B` capital gate in isolation. It returns only
   `capital_under_1000m`; numeric cap and raw security data never enter the
   clinical branch, artifact or output. Past screens use the last completed
   session before earliest disclosure; future screens use the latest completed
   session as of screening.
6. Keep the database scan active and upsert new clinical/commercial facts to
   their proper modules. Only clinical/trial/dose fields may enter this blind
   screen.
7. Save `artifacts/event_screens/{event_slug}.json`, run the
   `event-clinical-screen` validator, and fail closed on an unverified gate.
8. Return only alphabetized linked tickers, or `NONE`; provide no market figures,
   commentary, company summaries or excluded candidates.

## New ticker: clean model gate

1. Copy the master template; never reuse another ticker's delivered workbook.
2. Run the company-facts gate and store `ir_base_url` and `ir_news_url` from an
   official issuer page. Establish issuer identity/corporate-action boundaries
   before using vendor price history.
3. Clear Catalyst event metadata and analyst-input cells. Historical Events is
   rebuilt from blank event/category cells for all four calendar-year blocks.
4. Build Scenarios first. Derive the Catalyst universe from every Scenarios
   `Absolute` market-share row: one target for every drug × indication. Never
   aggregate a tail into “Other Pipeline” or copy a prior ticker's targets.
5. Build the full-universe neutral Catalyst framework first. During an active
   catalyst run, rebuild it in v7 active mode using researched
   `relevant_targets`. Delete Catalyst Target / Outcome / Scenario Summary and
   use the fixed order `Scenario → Base Case → Final Market Price → Upside → RJConv.
   → one outcome column per relevant target → target blocks`. Retain every non-relevant
   target block and its Table-3 inputs, but place those groups after the active
   blocks and keep all four columns visible with a grey background and grey
   data font.
6. Table 3 sits below the main scenario rows. Each target owns an independent
   four-column group: Market Share Change / LOA Change / Conv. / spacer. Apply
   the visual mask to the whole non-catalyst group without hiding any column.
   Preserve all analyst inputs by target name on every rebuild.
7. Conviction is not a valuation weight. Exclude each
   relevant target's outcomes below 10%, then generate the Cartesian product of
   every surviving target/outcome list. `RJConv.` means Raw Joint Conviction.
   Base RJConv. is 100%; each scenario RJConv.
   is the raw product of its selected target-level convictions under the
   independence assumption. Do not
   renormalize; sort the complete scenario rows by this product descending and
   use a stable outcome key for exact ties. Fail closed if any relevant target
   has no surviving outcome. Use the same contiguous scenario IDs in Catalyst and
   `Scenarios > Catalyst Scenarios`.
8. The Catalyst price What-If table lives directly in `Catalyst!B8:C[last]`.
   B8 is the adjusted `VALUATION!C48` output corner displayed as `ID`, C8 is the
   visible 3%, B9 is scenario 1 displayed as `Base`, and C9:C[last] are native
   Excel Data Table results. Local hidden bridges B6/C6 feed VALUATION C3/C5 so
   Excel's same-sheet input-cell rule is respected. `VALUATION!O4:P200` must be
   empty; other Valuation sensitivity tables remain unchanged.
9. Positive current-year RCFS Ending Cash divided by basic shares outstanding
   is included once in the post-catalyst equity value; negative cash contributes
   zero. Do not retain a stale valuation-date cash component.
10. The framework must scale beyond the four legacy coloured SOTP panels. Every
   drug × indication remains visible in the main table and Table 3; v7 greys
   the background and displayed data of non-relevant groups during an active
   run. Main-table row 6 stores each target's Scenario-1 breakdown contribution
   ratio; every catalyst scenario's breakdown value equals that scenario's Base
   Case price times the locked row-6 ratio. Scenario LOA equals row-9 base LOA plus the
   matching Table-3 LOA change. Scenarios catalyst MS equals Absolute MS plus
   the matching Table-3 Market Share change. The scenario rows,
   `Scenarios > Catalyst Scenarios`, the embedded Catalyst data table, lifecycle
   manifest and final post-catalyst price must contain every target. Never stop
   at four, aggregate a tail, or deliver a partial model.
11. Each Catalyst Scenarios block contains exactly one `# + catalyst title`
   header followed immediately by its asset/market-share rows. Numeric-only
   duplicate header rows are forbidden. Asset-title Y cells such as Y373/Y376
   must be blank.
12. Every relevant target's scenario Y market-share cell starts from its
   `Absolute` Y value and looks up that scenario's selected outcome in Catalyst
   Table 3. Suspension returns zero; other outcomes use
   `MAX(0, Absolute MS + looked-up MS change)`. Non-relevant targets carry the
   Absolute value unchanged. All row and column anchors must be explicit.

Production commands are wired in `build_dcf.py` as `catalyst`,
`catalystframework`, and `catalystclean`.

## “TICKER catalyst run”

1. Run `research/gpt_research.py --brief catalyst`. Search the latest official
   company press releases/presentations and official recent biomedical meeting
   schedules/notices. Confirm the nearest future event and exactly what will be
   disclosed: drug, indication, trial/phase, analysis type, cohort/population,
   endpoints and announced metrics. Do not invent a date or conference.
2. The research JSON must contain official source URLs and exact
   `relevant_targets` names from that ticker's Catalyst manifest.
3. Run `tools/catalyst_workflow.py run`. When research contains
   `relevant_targets`, the command first rebuilds the v7 active layout and its
   conviction-filtered Cartesian scenarios while preserving current Table-3
   inputs. It then records an open run in the durable
   `datastore/catalyst_history.duckdb`, writes event metadata, preserves the
   original colour scheme for relevant target columns, and changes every
   unrelated target column to light grey.
4. Analyst-entered probabilities, market-share deltas and LOA deltas remain in
   the workbook until a later post-catalyst command. After any Conv./MS/LOA
   edit, run `tools/catalyst_workflow.py refresh --ticker TICKER`; it preserves
   the open lifecycle run, reapplies the active mask and rebuilds the Cartesian
   scenario set from the current inputs.
5. Every Catalyst build/refresh must finish with Excel-native scoped calculation
   of the rebuilt Scenarios/Catalyst dependency chain and package-level
   calc-state normalization. Do not use a whole-workbook
   `CalculateFullRebuild`: legacy external dependencies can keep it busy after
   the Catalyst Data Table is already complete. Delivery is blocked when `calcMode`
   is Manual, `calcCompleted/calcOnSave` is zero, stale `calcFeatures` survives,
   or any stored font/conditional-format strike flag exists. This prevents
   Excel's visual stale-value strikethrough from returning after a late
   Catalyst edit.

## “TICKER test EVENT”

This command is a historical clinical-data interpretation test. For example,
`CMPX test ASCO2026` creates `Test-ASCO2026` in the existing model.

1. The event must already have occurred. Resolve the earliest public disclosure
   date separately from later conference presentation dates, plus data cutoff,
   drug × indication, trial/phase, population, endpoints, enrolled/evaluable
   denominators, efficacy, durability, safety and follow-up.
2. Clinical interpretation may retrieve only clinical evidence: full company
   clinical releases, conference abstracts/posters, trial registries,
   peer-reviewed papers, regulator clinical documents and competitor clinical
   readouts. Use the clinical-domain allowlist. Filter all price/commercial facts
   from the interpretation context.
3. Scan the database in parallel, then read only its clinical/dose/trial rows
   into the Test interpretation. Compare line-, biomarker-, population-, dose-,
   follow-up- and maturity-matched competitors; state cross-trial limitations.
4. For every directly affected target, the model—not the analyst—must assign
   Increase / Remain / Decrease / Suspension MS change, LOA change, conviction
   and clinical rationale. MS/LOA changes are absolute percentage-point changes;
   convictions sum to 100%. Lock these judgments before reading prices.
5. Exclude each outcome below 10% Conviction and generate the Cartesian product
   of all surviving outcomes. Conviction filters scenarios; it never weights a
   result.
6. Only after step 4 is locked, retrieve raw closes for completed sessions in
   `[earliest disclosure date − 7 calendar days, earliest disclosure date)`.
   During the blind stage, never retrieve or use disclosure-day/post-event
   price, reaction, return, volume, market cap, beta or analyst-target data.
   Average the eligible closes.
7. Estimate Base breakdown LOAs from the existing un-risked target values and
   prior LOA structure. Require the sum of every target Base Market Price to
   equal the pre-disclosure average close within a strict `<0.5%` relative
   difference. Fail closed outside the feasible LOA range.
8. Add a standalone `Test-EVENT` tab. Put `RJConv.` immediately after `Upside`.
   Base RJConv. is 100%; each scenario RJConv. is the raw product of its selected
   target-level Table-3 convictions. Sort scenario rows by this raw product
   descending without renormalizing after the >=10% filter. Event targets appear
   first in color. Every
   other drug × indication remains visible with grey background and grey data;
   no target group may be hidden. The page contains the observed readout,
   scenario grid, exact auto-generated MS/LOA/Conv. inputs, matched competitors,
   limitations, clickable clinical sources and the pre-disclosure calibration
   evidence.
9. Append or replace `Scenarios > Test Scenarios - EVENT`. Its scenario IDs are
   allocated after the maximum existing actual/Test ID and match the Test tab.
   Each active target's Y formula starts from Absolute Y and looks up the Test
   tab's selected outcome and MS change; Suspension is zero. Non-event targets
   remain exactly at Absolute Y, and asset-title Y cells stay blank. Rebuilding
   actual Catalyst scenarios must preserve these modules and renumber them only
   when needed to maintain workbook-global uniqueness.
10. `Test-EVENT` reproduces Catalyst's model-implied valuation table and may use
   internal formulas to the existing DCF/VALUATION/operating-model sheets.
   External-workbook formulas remain forbidden. The only observed security
   input allowed inside the blind phase is the bounded pre-disclosure raw-close
   calibration from steps 6–7; it cannot alter the already locked clinical
   judgments.
11. The Test builder must calculate the native Test Data Table and its local
   dependency chain, save the calculated caches, then normalize package calc
   metadata to Automatic/non-stale before its audit. It must not invoke a
   whole-workbook full rebuild. A Test tab may never leave the parent model in
   Manual/stale state.
12. Audit the filtered Cartesian count, joint-Conviction formulas and descending
   order, visible-grey mask, scenario IDs,
   per-target calibrated LOAs, Base reconciliation `<0.5%`, formula caches,
   forbidden same/post-event data, strikethrough residue and ZIP integrity.
13. Execute with:

   ```bash
   python tools/test_catalyst_event.py --ticker TICKER --event EVENT
   ```

   Company name is read from the company-facts artifact. `--research-file` is
   permitted only for a previously validated price-blind clinical payload; the
   command always refreshes the bounded pre-disclosure calibration.

14. Freeze and hash the clinical artifact, selected MS/LOA/Conviction inputs,
   pre-event calibration and original Test sheet before retrieving any
   post-release market data. Start a different agent context; never expose its
   prices or review to the blind interpretation agent.
15. Resolve the earliest public disclosure timestamp. Include the release-day
   close only when a reliable source proves a pre-close release; otherwise start
   with the next completed trading session. Retrieve exactly three eligible
   unadjusted closes and no later market observations.
16. The independent agent gives each active drug × indication one integer score
   from 1–10 (`1=miss`, `10=perfect`) with one short reason. Write the three
   closes in `D2:F2` and each score in row 2 of that target's final-LOA column.
17. For the same three sessions, retrieve regular-session 60-minute unadjusted
   High bars. Retain every exact tie for highest RJConv. For each tied
   scenario, compare its frozen Final Market Price with the three-day real peak:
   - miss: report `blind−peak` USD/share and `(blind−peak)/peak`;
   - reached: report `peak−blind` USD/share and `(peak−blind)/peak`.
18. Append the closes, target scores, daily intraday highs and peak test below
   the blind evidence. Upsert prices to `price` and evaluations to `backtest`.
   Recheck the blind sheet digest, native Data Table, formulas, strikes and ZIP.
19. Execute the isolated overlay with:

   ```bash
   python tools/score_test_catalyst_event.py \
     --ticker TICKER --event EVENT --scores-file INDEPENDENT_SCORE.json
   ```

## “TICKER post-catalyst”

1. Run `research/gpt_research.py --brief postcatalyst --context-file ...` after
   reading the full company result release and, for a meeting disclosure, the
   conference abstract/poster/report. The JSON must include the result versus
   expectation, exact data, limitations, price-reaction interpretation and URLs.
2. Run `tools/catalyst_workflow.py post`. It calculates the last pre-event close
   and first three post-event trading-session returns, then stores:
   - the complete XLSX binary;
   - exact Catalyst worksheet XML;
   - formula/input/cached-value/style JSON for the whole sheet;
   - price reaction and web interpretation.
3. Only after the DuckDB transaction commits does the tool clear analyst-input
   cells, clear event metadata and restore all original target colours. If the
   snapshot or interpretation is missing, the reset is refused.

## Historical Events

1. Enumerate every company-site press release in all four displayed calendar
   years, traversing pagination/RSS/archive pages. Preserve multiple releases on
   one date. Reconcile the fetched count and keep canonical PR URLs.
2. Open every PR. A clinical-data event begins with its disclosure venue in
   square brackets (`[ASCO GU Poster]`, `[AACR Abstract]`, `[CALL]`, etc.) and is
   summarized in this compact order using exact reported figures only:
   Phase / data type / N / ORR / CR / survival / safety.
3. When data are presented at an academic or other major meeting, read both the
   company PR and meeting abstract/poster/report. Link the Excel EVT cell to the
   meeting source first; otherwise link it to the company data release. If the
   meeting source is unavailable, record that explicitly and do not fabricate.
4. The JSON artifact retains every release separately. Because the worksheet has
   one EVT cell per date, same-day items are joined with ` | ` only at display
   time; they are never deduplicated away in storage.
5. Secondary news and large-price-move research may supplement the official
   archive but may not replace it. A zero-result official crawl blocks delivery.
