-- ================================================================
-- VOICE_QUERY:PURCHASE
-- ================================================================
WITH
    BASE_DATA AS (
        SELECT
            /*+ PARALLEL(A, 32) */
            A.ID_DATE,
            CASE A.BUNDLE_NAME
                WHEN '20B_700MIN_WEEKLY_ONNET_VOICE' THEN 'A.20B_700MIN_WEEKLY_ONNET_VOICE'
                WHEN '20B_700MIN_ONNET_WEEKLY_VOICE_FREE_USER' THEN 'A.20B_700MIN_WEEKLY_ONNET_VOICE'
                WHEN '5Br_Unlimited_Safari_Minutes_FUP_DEFAULT' THEN 'C.5Br_Unlimited_Safari_Minutes_FUP_DEFAULT'
                WHEN 'CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN' THEN 'E.CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN'
                WHEN 'CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN' THEN 'D.CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN'
                WHEN 'CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN' THEN 'B.CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN'
                WHEN 'FriendsAndFamily' THEN 'F.FriendsAndFamily'
            END AS BUNDLE_NAME,
            COUNT(DISTINCT A.MSISDN) AS TOTAL_SUBS,
            SUM(A.PURCHASE_QTY) AS TOTAL_QUANTITY,
            SUM(A.PURCHASE_AMT) AS TOTAL_REVENUE
        FROM
            BUNDLE_PURCHASE_2026_09_28 A
        WHERE
            A.ID_DATE BETWEEN {{START_DATE}} AND {{END_DATE}}
            AND (
                A.BUNDLE_NAME IN (
                    /* ====================================================
                    VOICE
                    ==================================================== */
                    'CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN',
                    'CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN',
                    'CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN',
                    '5Br_Unlimited_Safari_Minutes_FUP_DEFAULT',
                    '20B_700MIN_ONNET_WEEKLY_VOICE_FREE_USER',
                    '20B_700MIN_WEEKLY_ONNET_VOICE',
                    'FriendsAndFamily'
                )
            )
        GROUP BY
            A.ID_DATE,
            CASE A.BUNDLE_NAME
                WHEN '20B_700MIN_WEEKLY_ONNET_VOICE' THEN 'A.20B_700MIN_WEEKLY_ONNET_VOICE'
                WHEN '20B_700MIN_ONNET_WEEKLY_VOICE_FREE_USER' THEN 'A.20B_700MIN_WEEKLY_ONNET_VOICE'
                WHEN '5Br_Unlimited_Safari_Minutes_FUP_DEFAULT' THEN 'C.5Br_Unlimited_Safari_Minutes_FUP_DEFAULT'
                WHEN 'CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN' THEN 'E.CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN'
                WHEN 'CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN' THEN 'D.CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN'
                WHEN 'CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN' THEN 'B.CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN'
                WHEN 'FriendsAndFamily' THEN 'F.FriendsAndFamily'
            END
    ),
    /* ================================================================
    STEP 2: UNPIVOT METRICS
    ================================================================ */
    UNPIVOT_DATA AS (
        SELECT
            ID_DATE,
            BUNDLE_NAME,
            METRIC,
            VALUE
        FROM
            BASE_DATA UNPIVOT (
                VALUE FOR METRIC IN (
                    TOTAL_SUBS AS 'A.SUBS',
                    TOTAL_QUANTITY AS 'B.QUANTITY',
                    TOTAL_REVENUE AS 'C.REVENUE'
                )
            )
    )
    /* ================================================================
    FINAL OUTPUT
    ================================================================ */
SELECT
    ID_DATE,
    BUNDLE_NAME,
    METRIC,
    VALUE
FROM
    UNPIVOT_DATA
ORDER BY
    ID_DATE,
    BUNDLE_NAME,
    CASE
        WHEN METRIC = 'A.SUBS' THEN 1
        WHEN METRIC = 'B.QUANTITY' THEN 2
        WHEN METRIC = 'C.REVENUE' THEN 3
    END
-- VOICE_QUERY:USAGE
WITH
    /* ============================================================
    1. NORMALIZE PRODUCT NAME
    ============================================================ */
    PRODUCT_DATA AS (
        SELECT
            /*+ PARALLEL(FT, 64) */
            FT.ID_DATE,
            REGEXP_REPLACE (
                TRIM(RP.DS_NCC_ID),
                '(_+(SAF_(OF|RC)|RC(_V[0-9]+)?|OF(_V[0-9]+)?|V[0-9]+))+$',
                '',
                1,
                0,
                'i'
            ) AS PRODUCT_NAME,
            FT.ID_TRAFFIC_TYPE,
            FT.VL_RATE_USAGE
        FROM
            BI.FCT_TRAFFIC FT
            JOIN BI.REF_PRODUCT RP ON FT.ID_PRODUCT = RP.ID_PRODUCT
        WHERE
            FT.ID_DATE BETWEEN {{START_DATE}} AND {{END_DATE}}
            AND REGEXP_REPLACE (
                TRIM(RP.DS_NCC_ID),
                '(_+(SAF_(OF|RC)|RC(_V[0-9]+)?|OF(_V[0-9]+)?|V[0-9]+))+$',
                '',
                1,
                0,
                'i'
            ) IN (
                '20B_700MIN_WEEKLY_ONNET_VOICE',
                '20B_700MIN_ONNET_WEEKLY_VOICE_FREE_USER',
                'CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN',
                '5Br_Unlimited_Safari_Minutes_FUP_DEFAULT',
                'CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN',
                'CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN',
                'FriendsAndFamily'
            )
    ),
    /* ============================================================
    2. MAP PRODUCTS TO REPORTING BUNDLE NAMES
    ============================================================ */
    BASE_DATA AS (
        SELECT
            ID_DATE,
            CASE
                WHEN PRODUCT_NAME IN (
                    '20B_700MIN_WEEKLY_ONNET_VOICE',
                    '20B_700MIN_ONNET_WEEKLY_VOICE_FREE_USER'
                ) THEN 'A.20B_700MIN_WEEKLY_ONNET_VOICE'
                WHEN PRODUCT_NAME = 'CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN' THEN 'B.CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN'
                WHEN PRODUCT_NAME = '5Br_Unlimited_Safari_Minutes_FUP_DEFAULT' THEN 'C.5Br_Unlimited_Safari_Minutes_FUP_DEFAULT'
                WHEN PRODUCT_NAME = 'CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN' THEN 'D.CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN'
                WHEN PRODUCT_NAME = 'CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN' THEN 'E.CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN'
                WHEN PRODUCT_NAME = 'FriendsAndFamily' THEN 'F.FriendsAndFamily'
            END AS BUNDLE_NAME,
            /* ====================================================
            VOICE USAGE
            Traffic type = 1
            Convert seconds → minutes
            ==================================================== */
            SUM(
                CASE
                    WHEN ID_TRAFFIC_TYPE = 1 THEN VL_RATE_USAGE
                    ELSE 0
                END
            ) / 60 AS VOICE_MINUTES
        FROM
            PRODUCT_DATA
        GROUP BY
            ID_DATE,
            PRODUCT_NAME
    ),
    METRICS_DATA AS (
        SELECT
            ID_DATE,
            BUNDLE_NAME,
            'USAGE' AS METRIC,
            CASE
            /* -----------------------------------------------
            VOICE BUNDLES → MINUTES
            ----------------------------------------------- */
                WHEN BUNDLE_NAME IN (
                    'A.20B_700MIN_WEEKLY_ONNET_VOICE',
                    'B.CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN',
                    'C.5Br_Unlimited_Safari_Minutes_FUP_DEFAULT',
                    'D.CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN',
                    'E.CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN',
                    'F.FriendsAndFamily'
                ) THEN VOICE_MINUTES
            END AS VALUE
        FROM
            BASE_DATA
    )
SELECT
    ID_DATE,
    BUNDLE_NAME,
    METRIC,
    VALUE
FROM
    METRICS_DATA
ORDER BY
    ID_DATE,
    CASE
        WHEN BUNDLE_NAME LIKE 'A.%' THEN 1
        WHEN BUNDLE_NAME LIKE 'B.%' THEN 2
        WHEN BUNDLE_NAME LIKE 'C.%' THEN 3
        WHEN BUNDLE_NAME LIKE 'D.%' THEN 4
        WHEN BUNDLE_NAME LIKE 'E.%' THEN 5
        WHEN BUNDLE_NAME LIKE 'F.%' THEN 6
        WHEN BUNDLE_NAME LIKE 'G.%' THEN 7
        WHEN BUNDLE_NAME LIKE 'H.%' THEN 8
    END
-- VOICE_QUERY:FIVE_BIRR
WITH
    BASE_DATA AS (
        SELECT
            ID_DATE,
            BUNDLE_NAME,
            CASE
                WHEN BUNDLE_NAME IN ('5Br_Unlimited_Safari_Minutes_FUP_DEFAULT') THEN 'A. VOICE'
            END AS CATEGORY,
            COUNT(DISTINCT MSISDN) AS SUBS,
            SUM(MINUTES) AS USAGE
        FROM
            MELKAMU_WALE.BIGD_VOICE_DAILY_REGION_MSISDN_USAGE
        WHERE
            BUNDLE_NAME IN ('5Br_Unlimited_Safari_Minutes_FUP_DEFAULT')
            AND ID_DATE BETWEEN {{START_DATE}} AND {{END_DATE}}
        GROUP BY
            ID_DATE,
            BUNDLE_NAME,
            CASE
                WHEN BUNDLE_NAME IN ('5Br_Unlimited_Safari_Minutes_FUP_DEFAULT') THEN 'A. VOICE'
            END
    ),
    METRICS_DATA AS (
        SELECT
            ID_DATE,
            CATEGORY,
            BUNDLE_NAME,
            'USAGE' AS METRIC,
            USAGE AS VALUE
        FROM
            BASE_DATA
        WHERE
            CATEGORY = 'A. VOICE'
    )
SELECT
    ID_DATE,
    CATEGORY,
    BUNDLE_NAME,
    METRIC,
    VALUE
FROM
    METRICS_DATA
ORDER BY
    ID_DATE,
    BUNDLE_NAME,
    CASE METRIC
        WHEN 'USAGE' THEN 2
        ELSE 99
    END