# DeepDiligence workflow structure

> **Pinned canonical map.** Every workflow/code change must update this file in
> the same change. If this map and another document conflict, stop and reconcile.

## Parallel workflow tree

```mermaid
flowchart TD
    A[Command received] --> B[Resolve ticker + intent]
    B --> C{Workflow}
    B --> DB0[Start database scan]

    DB0 --> DB1[Load ticker facts]
    DB1 --> DB2[Load peer facts]
    DB2 --> DB3[Find completeness gaps]
    DB3 --> DB4[Search primary sources]
    DB4 --> DB5[Extract new facts]
    DB5 --> DB6[Validate source + date]
    DB6 --> DB7[Upsert JSON seed]
    DB7 --> DB8[Upsert DuckDB]
    DB8 --> DB9[Rescan gaps]
    DB9 --> DB4

    C --> N[New ticker]
    C --> R[Research]
    C --> CA[TICKER catalyst]
    C --> TE[TICKER test EVENT]
    C --> EV[event EVENT]
    C --> PC[TICKER post-catalyst]
    C --> HE[Historical Events]

    N --> N1[Copy clean template]
    N1 --> N2[Lock company identity]
    N2 --> N3[Fill financials]
    N3 --> N4[Build pipeline]
    N4 --> N5[Build Scenarios]
    N5 --> N6[Build Catalyst universe]
    N6 --> QA[Shared delivery QA]

    R --> R1[Define scope]
    R1 --> R2[Read internal facts]
    R2 --> R3[Research primary sources]
    R3 --> R4[Match populations]
    R4 --> R5[Complete competitors]
    R5 --> R6[Write report + DB updates]
    R6 --> QA

    CA --> CA1[Confirm nearest event]
    CA1 --> CA2[Map affected targets]
    CA2 --> CA3[Preserve Table 3 inputs]
    CA3 --> CA4[Filter Conv. >=10%]
    CA4 --> CA5[Build Cartesian outcomes]
    CA5 --> CA6[Reorder + grey mask]
    CA6 --> CA7[Rebuild Catalyst Scenarios]
    CA7 --> CA8[Calculate local Data Table]
    CA8 --> QA

    TE --> TE1[Confirm completed event]
    TE1 --> TE2[Find earliest disclosure]
    TE2 --> TE3[Load clinical-only DB]
    TE3 --> TE4[Research actual readout]
    TE4 --> TE5[Match clinical competitors]
    TE5 --> TE6[Lock MS + LOA + Conv.]
    TE6 --> TE7[Filter Conv. >=10%]
    TE7 --> TE8[Build Cartesian outcomes]
    TE6 --> TP1[Start price calibration]
    TP1 --> TP2[Fetch prior 7-day raw closes]
    TP2 --> TP3[Exclude disclosure day + after]
    TP3 --> TP4[Average completed sessions]
    TP4 --> TP5[Estimate Base LOAs]
    TP5 --> TP6{Breakdown error <0.5%?}
    TP6 -- No --> TP5
    TP6 -- Yes --> TE9[Build Test-EVENT]
    TE8 --> TE9
    TE9 --> TE10[Build Test Scenarios]
    TE10 --> TE11[Calculate local Data Table]
    TE11 --> TF1[Freeze blind prediction]
    TF1 --> TF2[Hash clinical + sheet]
    TF2 --> PS1[Start isolated scorer]
    PS1 --> PS2[Resolve release timestamp]
    PS2 --> PS3[Fetch 3 eligible closes]
    PS3 --> PS3A[Fetch 60m raw highs]
    PS3A --> PS4[Compare locked prediction]
    PS4 --> PS4A[Find highest RJConv.]
    PS4A --> PS5[Score each target 1-10]
    PS5 --> PS6[Write D2:F2 closes]
    PS6 --> PS7[Write row-2 LOA scores]
    PS7 --> PS8[Append brief review]
    PS8 --> PS8A[Test peak vs blind price]
    PS8A --> PS9[Verify blind digest]
    PS9 --> QA

    EV --> EV1[Resolve edition + dates]
    EV1 --> EV2{Completed?}
    EV2 -- Yes --> EV3[Crawl actual clinical disclosures]
    EV2 -- No --> EV4[Crawl announced clinical readouts]
    EV3 --> EV5[Verify US listing]
    EV4 --> EV5
    EV5 --> EV6[Run isolated capital gate]
    EV6 --> EV7[Keep capital < USD 1B]
    EV7 --> EV8[Validate price-blind artifact]
    EV8 --> EV9[Return ticker links only]

    PC --> PC1[Read actual disclosure]
    PC1 --> PC2[Measure price reaction]
    PC2 --> PC3[Interpret result]
    PC3 --> PC4[Archive workbook + XML]
    PC4 --> PC5[Commit lifecycle record]
    PC5 --> PC6[Reset Catalyst inputs]
    PC6 --> QA

    HE --> HE1[Crawl official archive]
    HE1 --> HE2[Open every release]
    HE2 --> HE3[Read meeting source]
    HE3 --> HE4[Preserve same-day items]
    HE4 --> HE5[Write linked events]
    HE5 --> QA

    DB8 -. latest facts .-> R3
    DB8 -. latest facts .-> CA1
    DB8 -. clinical fields only .-> TE3
    DB8 -. clinical fields only .-> EV3
    DB8 -. clinical fields only .-> EV4
    DB8 -. price/backtest fields .-> PS3
    DB8 -. latest facts .-> PC1
    DB8 -. latest facts .-> HE2

    QA --> Q1[Formula audit]
    Q1 --> Q2[Style + strike audit]
    Q2 --> Q3[Scenario-ID audit]
    Q3 --> Q4[Calc-cache audit]
    Q4 --> Q5[ZIP integrity]
    Q5 --> Q6[Write manifest]
    Q6 --> Q7[Sync this file]
    Q7 --> Z[Deliver]
```

## Always-active database branch

1. Scan `research_fact`, peers and internal modules.
2. Identify new, stale, missing and conflicting facts.
3. Search official/primary sources first.
4. Store every new or corrected fact immediately.
5. Preserve value, unit, population, dose, date and URL.
6. Separate reported, estimated, unavailable and conflict.
7. Never invent a missing value.
8. Rebuilds retain facts through `research_facts.json`.
9. Rescan after every upsert.
10. Keep scanning while research remains active.

The branch is mandatory in GPT, Gemini, Opus, TAM-indication, ClinicalTrials.gov
and model-judgement entrypoints; no research entrypoint may bypass it.

For every in-scope competitor, maximize these fields:

1. ORR.
2. CR.
3. PR.
4. mPFS/PFS.
5. mOS/OS.
6. DoR/EFS/DFS/RFS when reported.
7. Safety overview.
8. Named side effects.
9. Side-effect rates.
10. Dose/regimen.
11. Actual price.
12. Forecast price.

Unreported fields are stored as `unavailable` with a source and explanation.
Actual and forecast prices remain separate.

## Event screen gates

1. Trigger on `event EVENT` or `EVENT event`.
2. Resolve the exact official edition and dates.
3. Completed event: require disclosed quantitative human clinical data.
4. Future event: require an official announced quantitative clinical readout.
5. Reject attendance-only, trial-design, enrollment-only and preclinical items.
6. Require Nasdaq, NYSE or NYSE American listing at the cutoff.
7. Run the sub-USD-1B capital test in an isolated branch.
8. Pass only `capital_under_1000m` back to the clinical branch.
9. Never expose price, OHLC, return, reaction, volume, beta or targets.
10. Update the research database while the scan remains active.
11. Validate the price-blind evidence artifact.
12. Return only alphabetized ticker links; return `NONE` if empty.

## Test EVENT gates

1. Use completed clinical events only.
2. Use the earliest public disclosure date.
3. Keep the clinical interpretation price-blind.
4. Filter all commercial/price fields from clinical context.
5. Store commercial facts in the database when found elsewhere.
6. Lock clinical MS/LOA/Conv. before price calibration.
7. Fetch raw closes only in `[disclosure−7 days, disclosure)`.
8. Use completed sessions only.
9. Never use same-day/post-event data in the blind stage.
10. Scale target Base LOAs to the pre-event average close.
11. Keep every target breakdown visible.
12. Require `abs(sum(breakdown)/average−1) < 0.5%`.
13. Build Test formulas from the matching Test Scenarios module.
14. Fail on stale caches, strikes, broken references or duplicate IDs.

## Post-release scoring gates

1. Finish the blind workbook first.
2. Hash the clinical artifact and prediction fields.
3. Hash the original Test sheet below row 2.
4. Start a separate agent context.
5. Never feed post-release prices back to the blind agent.
6. Resolve the earliest public release timestamp.
7. Include release-day close only after a pre-close release.
8. Otherwise start with the next completed session.
9. Use exactly three eligible raw closes.
10. Fetch regular-session 60m raw High for those sessions.
11. Use no price or interpretation after the third close.
12. Score every active target from 1 to 10.
13. Treat 1 as miss and 10 as perfect.
14. Write closes in `D2:F2`.
15. Write each score in row 2 of its final-LOA column.
16. Find every exact highest-RJConv. scenario tie.
17. Test each blind price against the three-day peak.
18. Report USD/share and percent-of-peak difference.
19. Keep each review to one short sentence.
20. Append a separate audit trail below the blind evidence.
21. Store closes/highs under `price` and scores/tests under `backtest`.
22. Recheck the blind digest after saving.
23. Never revise locked LOA, MS, Conviction or breakdown.

## Catalyst/Test table invariants

1. Order: Scenario, Base Case, Final Market Price, Upside, RJConv.
2. Define RJConv. as Raw Joint Conviction.
3. Calculate RJConv. as the raw product of selected outcome convictions.
4. Keep Base RJConv. at 100%.
5. Sort scenario rows by RJConv. descending.
6. Put active outcome columns next.
7. Put active target groups next.
8. Put inactive target groups last.
9. Mask inactive groups grey; never hide them.
10. Keep inactive LOA and market share unchanged.
11. Use target Conviction as a >=10% filter and RJConv. input.
12. Never renormalize surviving RJConv. values.
13. Use every surviving outcome combination.
14. Sum all target Market Price columns.
15. Keep row/column anchors explicit.
16. Keep title-row index cells blank.
17. Calculate only rebuilt dependency ranges.

## Required synchronization

Any workflow modification must update, as applicable:

1. `00_WORKFLOW_STRUCTURE.md`.
2. `information/CATALYST_HISTORICAL_EVENTS_WORKFLOW.md`.
3. `information/PIPELINE_RESEARCH_REQUIREMENTS.md`.
4. The applicable skill `SKILL.md`.
5. Builders, validators and manifests.
6. Regression audits.
