-- ============================================================================
--  DD DATA CENTER — LAYER 2 : DERIVED VIEWS  (computed, reference Layer 1)
-- ----------------------------------------------------------------------------
--  Every object here is a VIEW: it stores nothing and is always consistent
--  with Layer 1. This is the "second layer" that references the raw drug data
--  to compute TAM and the model parameters.
--
--  These views REPLACE the live, interdependent cell formulas that used to live
--  in the TAM Solid sheet (e.g. =SUMIF($D$9:$D$399,$D406,P$9:P$399)). Because
--  they aggregate by KEY, they are immune to row insertions — adding a drug
--  changes the inputs, never these definitions.
-- ============================================================================

-- 1) Effective split share per drug × indication.
--    Resolves the two split methods into a single fraction in [0,1].
CREATE VIEW v_drug_split AS
SELECT
    drug_id,
    indication_code,
    method,
    CASE
        WHEN method = 'share' THEN share
        ELSE incidence_weight
             / NULLIF(SUM(incidence_weight) OVER (PARTITION BY drug_id), 0)
    END AS eff_share
FROM drug_indication_split;

-- 2) Revenue attributed to each drug × indication × year.
CREATE VIEW v_drug_indication_revenue AS
SELECT
    r.drug_id,
    s.indication_code,
    r.year,
    r.revenue_usd_m * s.eff_share AS revenue_usd_m
FROM drug_revenue r
JOIN v_drug_split s ON s.drug_id = r.drug_id;

-- 3) THE TAM. Total addressable market by indication × year
--    = sum of every drug's revenue attributed to that indication.
--    (This is the GROUP BY that replaces the sheet's SUMIF column.)
CREATE VIEW v_tam_by_indication_year AS
SELECT
    indication_code,
    year,
    SUM(revenue_usd_m) AS tam_usd_m
FROM v_drug_indication_revenue
GROUP BY indication_code, year;

-- 3b) TAM rolled up to the solid / blood world (drug.tam_group).
CREATE VIEW v_tam_by_group_year AS
SELECT
    d.tam_group,
    v.indication_code,
    v.year,
    SUM(v.revenue_usd_m) AS tam_usd_m
FROM v_drug_indication_revenue v
JOIN drug d ON d.drug_id = v.drug_id
GROUP BY d.tam_group, v.indication_code, v.year;

-- 4) Derived parameter: incidence (passed through from inputs, exposed for Excel).
CREATE VIEW v_param_incidence AS
SELECT indication_code, incidence_rate, incidence_global_annual
FROM indication;

-- 5) Derived parameter: growth tiers from peer-drug sales ramps.
--    AVG = mean year-over-year growth across all reference drugs.
--    BIC = Tagrisso ramp (best-in-class launch trajectory).
--    T1  = Alunbrig ramp (tier-one trajectory).
CREATE VIEW v_param_growth AS
WITH yoy AS (
    SELECT
        item,
        year,
        value / NULLIF(LAG(value) OVER (PARTITION BY item ORDER BY year), 0) - 1
            AS growth
    FROM reference_drug_sale
    WHERE purpose = 'growth'
)
SELECT 'AVG' AS tier, AVG(growth) AS growth_factor FROM yoy WHERE growth IS NOT NULL
UNION ALL
SELECT 'BIC', AVG(growth) FROM yoy WHERE item = 'Tagrisso'  AND growth IS NOT NULL
UNION ALL
SELECT 'T1',  AVG(growth) FROM yoy WHERE item = 'Alunbrig'  AND growth IS NOT NULL;

-- 6) Derived parameter: COGS / Price, reproduced from Xpovio's cost lines
--    exactly as the sheet does (R556-R562):
--      selling     = selling_and_ga - ga
--      total_cogs  = selling + manufacturing
--      cogs_price  = total_cogs / net_sale
CREATE VIEW v_param_cogs_price AS
WITH x AS (
    SELECT
        MAX(value) FILTER (WHERE item = 'xpovio_net_sale')       AS net_sale,
        MAX(value) FILTER (WHERE item = 'xpovio_manufacturing')  AS manufacturing,
        MAX(value) FILTER (WHERE item = 'xpovio_selling_and_ga') AS selling_and_ga,
        MAX(value) FILTER (WHERE item = 'xpovio_ga')             AS ga
    FROM reference_drug_sale
    WHERE purpose = 'cogs'
)
SELECT
    (((selling_and_ga - ga) + manufacturing) / NULLIF(net_sale, 0)) AS cogs_price
FROM x;

-- 7) Clean Peer Views rating lookup: drug -> BIC/T1/AVG by section, color decoded
--    into text. Excel XLOOKUPs this instead of reading cell fill colors.
CREATE VIEW v_peer_rating AS
SELECT section_id, section, drug, ticker, rating
FROM peer_drug
WHERE rating IS NOT NULL;
