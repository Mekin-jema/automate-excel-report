-- Query: Regional Second Recharge Cohorts with Standardized Region Ordering
SELECT
    CASE
        WHEN L.DS_EXEC_REGION = 'West Addis'  THEN 'A.West Addis'
        WHEN L.DS_EXEC_REGION = 'East Addis'  THEN 'B.East Addis'
        WHEN L.DS_EXEC_REGION = 'Central'     THEN 'C.Central'
        WHEN L.DS_EXEC_REGION = 'South'       THEN 'D.South'
        WHEN L.DS_EXEC_REGION = 'North West'  THEN 'E.North West'
        WHEN L.DS_EXEC_REGION = 'North East'  THEN 'F.North East'
        WHEN L.DS_EXEC_REGION = 'East 1'      THEN 'G.East 1'
        WHEN L.DS_EXEC_REGION = 'East 2'      THEN 'H.East 2'
        WHEN L.DS_EXEC_REGION = 'West'        THEN 'I.West'
        WHEN L.DS_EXEC_REGION = 'North'       THEN 'J.North'
        WHEN L.DS_EXEC_REGION = 'Afar'        THEN 'K.Afar'
        ELSE 'L.Unknown'
    END AS REGION,
    L.*
FROM BETELHEM_GIRMA.TBL_DAILY_SECOND_RECHARGE_COHORT_RPT L
WHERE L.DS_EXEC_REGION IN (
    'East Addis',
    'West Addis',
    'Central',
    'South',
    'West',
    'East 1',
    'East 2',
    'North West',
    'North East',
    'North'
)
AND L.REPORT_TYPE = '2nd Purchase'
ORDER BY REGION, ACTIVATION_DATE
