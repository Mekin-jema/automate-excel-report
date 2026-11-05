-- DATA_QUERY:PURCHASE
--   bundel purchase
WITH
    BASE_DATA_UNPIVOT AS (
        SELECT
            /*+ PARALLEL(A, 32) */
            A.ID_DATE,
            /* ============================================================
            BUNDLE NAME
            01-18 PREFIX IS INTENTIONAL FOR EXCEL PIVOT SORTING
            ============================================================ */
            CASE
            /* =========================
            GIZE+
            ========================= */
                WHEN A.BUNDLE_NAME = 'CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR' THEN '01. CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR'
                WHEN A.BUNDLE_NAME = 'CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR' THEN '02. CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR'
                WHEN A.BUNDLE_NAME = 'CBU_TIMEBASED_UNLTD_DATA_DAY_1HR' THEN '03. CBU_TIMEBASED_UNLTD_DATA_DAY_1HR'
                WHEN A.BUNDLE_NAME = 'CBU_TIMEBASED_UNLTD_DATA_DAY_2HR' THEN '04. CBU_TIMEBASED_UNLTD_DATA_DAY_2HR'
                WHEN A.BUNDLE_NAME = 'CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR' THEN '05. CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR'
                WHEN A.BUNDLE_NAME = 'CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR' THEN '06. CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR'
                /* =========================
                DAILY UNLIMITED
                ========================= */
                WHEN A.BUNDLE_NAME = 'DAILY_UNLTD_D_18GB_FUP' THEN '07. DAILY_UNLTD_D_18GB_FUP'
                WHEN A.BUNDLE_NAME = 'CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION' THEN '08. CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION'
                WHEN A.BUNDLE_NAME = 'CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION' THEN '09. CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION'
                /* =========================
                WEEKLY UNLIMITED
                ========================= */
                WHEN A.BUNDLE_NAME = 'WEEKLY_UNLTD_D_70GB_FUP' THEN '10. WEEKLY_UNLTD_D_70GB_FUP'
                WHEN A.BUNDLE_NAME = 'CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION' THEN '11. CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION'
                WHEN A.BUNDLE_NAME = 'CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION' THEN '12. CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION'
                /* =========================
                MONTHLY UNLIMITED
                ========================= */
                WHEN A.BUNDLE_NAME = 'MONTHLY_UNLTD_D_300GB_FUP' THEN '13. MONTHLY_UNLTD_D_300GB_FUP'
                WHEN A.BUNDLE_NAME = 'CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION' THEN '14. CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION'
                WHEN A.BUNDLE_NAME = 'CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION' THEN '15. CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION'
                /* =========================
                WULLO / MATA
                ========================= */
                WHEN A.BUNDLE_NAME = 'EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP' THEN '16. EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP'
                WHEN A.BUNDLE_NAME = 'NIGHT_UNLTD_D_5GB_FUP' THEN '17. NIGHT_UNLTD_D_5GB_FUP'
                WHEN A.BUNDLE_NAME = 'EBU_DATA_WULLO_75GB_MONTHLY' THEN '18. EBU_DATA_WULLO_75GB_MONTHLY'
            END AS BUNDLE_NAME,
            /* ============================================================
            BUNDLE CATEGORIES
            ============================================================ */
            CASE
            /* ---------------- GIZE+ ---------------- */
                WHEN A.BUNDLE_NAME IN (
                    'CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR',
                    'CBU_TIMEBASED_UNLTD_DATA_DAY_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_DAY_2HR',
                    'CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR'
                ) THEN 'Gize+'
                /* ---------------- DAILY ---------------- */
                WHEN A.BUNDLE_NAME IN (
                    'DAILY_UNLTD_D_18GB_FUP',
                    'CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION'
                ) THEN 'Daily Unlimited'
                /* ---------------- WEEKLY ---------------- */
                WHEN A.BUNDLE_NAME IN (
                    'WEEKLY_UNLTD_D_70GB_FUP',
                    'CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION'
                ) THEN 'Weekly Unlimited'
                /* ---------------- MONTHLY ---------------- */
                WHEN A.BUNDLE_NAME IN (
                    'MONTHLY_UNLTD_D_300GB_FUP',
                    'CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION'
                ) THEN 'Monthly Unlimited'
                /* ---------------- WULLO ---------------- */
                WHEN A.BUNDLE_NAME = 'EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP' THEN 'Wullo'
                /* ---------------- MATA ---------------- */
                WHEN A.BUNDLE_NAME = 'NIGHT_UNLTD_D_5GB_FUP' THEN 'Mata'
                /* ---------------- MONTHLY WULLO ---------------- */
                WHEN A.BUNDLE_NAME = 'EBU_DATA_WULLO_75GB_MONTHLY' THEN 'Monthly Wullo'
            END AS BUNDLE_CATEGORY,
            /* ============================================================
            METRICS
            ============================================================ */
            COUNT(DISTINCT A.MSISDN) AS TOTAL_SUBS,
            SUM(A.PURCHASE_QTY) AS TOTAL_QUANTITY,
            SUM(A.PURCHASE_AMT) AS TOTAL_REVENUE
        FROM
            BUNDLE_PURCHASE_2026_09_28 A
            /* ============================================================
            DATE RANGE
            ============================================================ */
        WHERE
            A.ID_DATE BETWEEN {{START_DATE}} AND {{END_DATE}}
            /* ============================================================
            ONLY 7 CATEGORIES / 18 BUNDLES
            ============================================================ */
            AND A.BUNDLE_NAME IN (
                /* ============================================================
                GIZE+
                ============================================================ */
                'CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR',
                'CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR',
                'CBU_TIMEBASED_UNLTD_DATA_DAY_1HR',
                'CBU_TIMEBASED_UNLTD_DATA_DAY_2HR',
                'CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR',
                'CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR',
                /* ============================================================
                DAILY UNLIMITED
                ============================================================ */
                'DAILY_UNLTD_D_18GB_FUP',
                'CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION',
                'CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION',
                /* ============================================================
                WEEKLY UNLIMITED
                ============================================================ */
                'WEEKLY_UNLTD_D_70GB_FUP',
                'CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION',
                'CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION',
                /* ============================================================
                MONTHLY UNLIMITED
                ============================================================ */
                'MONTHLY_UNLTD_D_300GB_FUP',
                'CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION',
                'CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION',
                /* ============================================================
                WULLO / MATA
                ============================================================ */
                'EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP',
                'NIGHT_UNLTD_D_5GB_FUP',
                'EBU_DATA_WULLO_75GB_MONTHLY'
            )
            /* ============================================================
            GROUP BY
            ============================================================ */
        GROUP BY
            A.ID_DATE,
            A.BUNDLE_NAME,
            CASE
                WHEN A.BUNDLE_NAME IN (
                    'CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR',
                    'CBU_TIMEBASED_UNLTD_DATA_DAY_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_DAY_2HR',
                    'CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR'
                ) THEN 'Gize+'
                WHEN A.BUNDLE_NAME IN (
                    'DAILY_UNLTD_D_18GB_FUP',
                    'CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION'
                ) THEN 'Daily Unlimited'
                WHEN A.BUNDLE_NAME IN (
                    'WEEKLY_UNLTD_D_70GB_FUP',
                    'CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION'
                ) THEN 'Weekly Unlimited'
                WHEN A.BUNDLE_NAME IN (
                    'MONTHLY_UNLTD_D_300GB_FUP',
                    'CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION'
                ) THEN 'Monthly Unlimited'
                WHEN A.BUNDLE_NAME = 'EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP' THEN 'Wullo'
                WHEN A.BUNDLE_NAME = 'NIGHT_UNLTD_D_5GB_FUP' THEN 'Mata'
                WHEN A.BUNDLE_NAME = 'EBU_DATA_WULLO_75GB_MONTHLY' THEN 'Monthly Wullo'
            END
    ),
    /* ================================================================
    UNPIVOT METRICS
    ================================================================ */
    UNPIVOT_DATA AS (
        SELECT
            ID_DATE,
            BUNDLE_CATEGORY,
            BUNDLE_NAME,
            METRIC,
            VALUE
        FROM
            BASE_DATA_UNPIVOT UNPIVOT (
                VALUE FOR METRIC IN (
                    TOTAL_SUBS AS 'A.SUBS',
                    TOTAL_QUANTITY AS 'B.QUANTITY',
                    TOTAL_REVENUE AS 'C.REVENUE'
                )
            )
    )
    /* ================================================================
    FINAL RESULT
    ================================================================ */
SELECT
    ID_DATE,
    BUNDLE_CATEGORY,
    BUNDLE_NAME,
    METRIC,
    VALUE
FROM
    UNPIVOT_DATA
ORDER BY
    ID_DATE,
    BUNDLE_NAME,
    CASE METRIC
        WHEN 'A.SUBS' THEN 1
        WHEN 'B.QUANTITY' THEN 2
        WHEN 'C.REVENUE' THEN 3
    END
    --

-- DATA_QUERY:USAGE
WITH
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
    ),
    MAPPED_DATA AS (
        SELECT
            A.ID_DATE,
            A.PRODUCT_NAME,
            /* NUMBERED BUNDLE NAME */
            CASE
            /* GIZE+ */
                WHEN A.PRODUCT_NAME = 'CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR' THEN '01. CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR'
                WHEN A.PRODUCT_NAME = 'CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR' THEN '02. CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR'
                WHEN A.PRODUCT_NAME = 'CBU_TIMEBASED_UNLTD_DATA_DAY_1HR' THEN '03. CBU_TIMEBASED_UNLTD_DATA_DAY_1HR'
                WHEN A.PRODUCT_NAME = 'CBU_TIMEBASED_UNLTD_DATA_DAY_2HR' THEN '04. CBU_TIMEBASED_UNLTD_DATA_DAY_2HR'
                WHEN A.PRODUCT_NAME = 'CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR' THEN '05. CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR'
                WHEN A.PRODUCT_NAME = 'CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR' THEN '06. CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR'
                /* DAILY UNLIMITED */
                WHEN A.PRODUCT_NAME = 'DAILY_UNLTD_D_18GB_FUP' THEN '07. DAILY_UNLTD_D_18GB_FUP'
                WHEN A.PRODUCT_NAME = 'CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION' THEN '08. CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION'
                WHEN A.PRODUCT_NAME = 'CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION' THEN '09. CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION'
                /* WEEKLY UNLIMITED */
                WHEN A.PRODUCT_NAME = 'WEEKLY_UNLTD_D_70GB_FUP' THEN '10. WEEKLY_UNLTD_D_70GB_FUP'
                WHEN A.PRODUCT_NAME = 'CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION' THEN '11. CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION'
                WHEN A.PRODUCT_NAME = 'CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION' THEN '12. CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION'
                /* MONTHLY UNLIMITED */
                WHEN A.PRODUCT_NAME = 'MONTHLY_UNLTD_D_300GB_FUP' THEN '13. MONTHLY_UNLTD_D_300GB_FUP'
                WHEN A.PRODUCT_NAME = 'CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION' THEN '14. CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION'
                WHEN A.PRODUCT_NAME = 'CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION' THEN '15. CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION'
                /* WULLO / MATA */
                WHEN A.PRODUCT_NAME = 'EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP' THEN '16. EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP'
                WHEN A.PRODUCT_NAME = 'NIGHT_UNLTD_D_5GB_FUP' THEN '17. NIGHT_UNLTD_D_5GB_FUP'
                WHEN A.PRODUCT_NAME = 'EBU_DATA_WULLO_75GB_MONTHLY' THEN '18. EBU_DATA_WULLO_75GB_MONTHLY'
            END AS BUNDLE_NAME,
            /* BUNDLE CATEGORY */
            CASE
                WHEN A.PRODUCT_NAME IN (
                    'CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR',
                    'CBU_TIMEBASED_UNLTD_DATA_DAY_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_DAY_2HR',
                    'CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR',
                    'CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR'
                ) THEN 'Gize+'
                WHEN A.PRODUCT_NAME IN (
                    'DAILY_UNLTD_D_18GB_FUP',
                    'CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION'
                ) THEN 'Daily Unlimited'
                WHEN A.PRODUCT_NAME IN (
                    'WEEKLY_UNLTD_D_70GB_FUP',
                    'CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION'
                ) THEN 'Weekly Unlimited'
                WHEN A.PRODUCT_NAME IN (
                    'MONTHLY_UNLTD_D_300GB_FUP',
                    'CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION',
                    'CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION'
                ) THEN 'Monthly Unlimited'
                WHEN A.PRODUCT_NAME = 'EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP' THEN 'Wullo'
                WHEN A.PRODUCT_NAME = 'NIGHT_UNLTD_D_5GB_FUP' THEN 'Mata'
                WHEN A.PRODUCT_NAME = 'EBU_DATA_WULLO_75GB_MONTHLY' THEN 'Monthly Wullo'
            END AS BUNDLE_CATEGORY,
            A.ID_TRAFFIC_TYPE,
            A.VL_RATE_USAGE
        FROM
            PRODUCT_DATA A
        WHERE
            A.PRODUCT_NAME IN (
                /* GIZE+ */
                'CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR',
                'CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR',
                'CBU_TIMEBASED_UNLTD_DATA_DAY_1HR',
                'CBU_TIMEBASED_UNLTD_DATA_DAY_2HR',
                'CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR',
                'CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR',
                /* DAILY UNLIMITED */
                'DAILY_UNLTD_D_18GB_FUP',
                'CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION',
                'CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION',
                /* WEEKLY UNLIMITED */
                'WEEKLY_UNLTD_D_70GB_FUP',
                'CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION',
                'CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION',
                /* MONTHLY UNLIMITED */
                'MONTHLY_UNLTD_D_300GB_FUP',
                'CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION',
                'CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION',
                /* WULLO / MATA */
                'EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP',
                'NIGHT_UNLTD_D_5GB_FUP',
                'EBU_DATA_WULLO_75GB_MONTHLY'
            )
    ),
    /* ============================================================
    3. AGGREGATE USAGE BY DATE, PRODUCT AND CATEGORY
    ============================================================ */
    BASE_DATA AS (
        SELECT
            ID_DATE,
            BUNDLE_NAME,
            SUM(
                CASE
                    WHEN ID_TRAFFIC_TYPE = 3 THEN VL_RATE_USAGE
                    ELSE 0
                END
            ) / (1024 * 1024) AS MB
        FROM
            MAPPED_DATA
        GROUP BY
            ID_DATE,
            BUNDLE_NAME
    ),
    /* ============================================================
    4. CONVERT TO REPORTING METRIC
    ============================================================ */
    METRICS_DATA AS (
        SELECT
            ID_DATE,
            BUNDLE_NAME,
            'USAGE' AS METRIC,
            MB AS VALUE
        FROM
            BASE_DATA
    )
    /* ============================================================
    5. FINAL OUTPUT
    ============================================================ */
SELECT
    ID_DATE,
    BUNDLE_NAME,
    METRIC,
    VALUE
FROM
    METRICS_DATA
ORDER BY
    ID_DATE,
    BUNDLE_NAME

-- DATA_QUERY:PURCHASE_TABLE
INSERT INTO
    BUNDLE_PURCHASE_2026_09_28 (
        ID_DATE,
        DS_SITE_ID,
        MSISDN,
        CHANNEL,
        CHANNEL_DETAILS,
        BUNDLE_NAME,
        PURCHASE_QTY,
        PURCHASE_AMT,
        "Bundle_Owner",
        "Validity",
        "Bundle_Type",
        ATL_BTL,
        COMBO_STATUS,
        BUNDLE_GROUP_TYPE,
        VALIDTY_GROUP
    )
WITH
    TEMP_K AS (
        SELECT
            /*+ DRIVING_SITE(m) PARALLEL(32) */
            id_date,
            DS_SITE_ID,
            msisdn,
            channel,
            channel_details,
            purchase_type,
            REGEXP_REPLACE (
                TRIM(t.Bundle_name),
                '(_+(SAF_(OF|RC)|RC(_V[0-9]+)?|OF(_V[0-9]+)?|V[0-9]+))+$',
                '',
                1,
                0,
                'i'
            ) AS bundle_name,
            COUNT(DISTINCT X_SERIAL_NO) AS purchase_qty,
            SUM(vl_accnt_blnc) / 100 AS purchase_amt
        FROM
            (
                SELECT
                    m.id_date,
                    m.X_SERIAL_NO,
                    rl.DS_SITE_ID,
                    m.msisdn,
                    CASE
                        WHEN m.X_TRANSACTION_ID LIKE '%MPESA%'
                        OR ip.DS_NCC_ID LIKE '%MP%' THEN 'MPESA'
                        ELSE 'AIRTIME'
                    END AS channel,
                    CASE
                        WHEN m.X_TRANSACTION_ID LIKE 'MPESA_USSD_CVM%' THEN 'MPESA_USSD_BESH'
                        WHEN m.X_TRANSACTION_ID LIKE 'MPESA_SuperAPP_CVM_BESH%' THEN 'MPESA_SUPERAPP_BESH'
                        WHEN m.X_TRANSACTION_ID LIKE 'MPESA_DXL_CVM%' THEN 'MPESA_DXL_BESH'
                        WHEN m.X_TRANSACTION_ID LIKE 'MPESA_USSD%' THEN 'MPESA_USSD'
                        WHEN m.X_TRANSACTION_ID LIKE 'MPESA_SuperAPP%' THEN 'MPESA_SUPERAPP'
                        WHEN m.X_TRANSACTION_ID LIKE 'MPESA_DXL%' THEN 'MPESA_DXL'
                        WHEN m.X_TRANSACTION_ID LIKE 'USSD%' THEN 'USSD'
                        WHEN m.X_TRANSACTION_ID LIKE 'UnifiedBot%' THEN 'Bot'
                        WHEN m.X_TRANSACTION_ID LIKE 'DXL%' THEN 'DXL'
                        WHEN m.X_TRANSACTION_ID LIKE 'CRM%' THEN 'CRM'
                        WHEN m.X_TRANSACTION_ID LIKE 'PRETUPS%' THEN 'PRETUPS'
                        WHEN m.X_TRANSACTION_ID LIKE 'Web_Selfcare%' THEN 'Web_Selfcare'
                        WHEN m.X_TRANSACTION_ID LIKE 'AAS_PC_USSD%' THEN 'AAS_PC_USSD'
                        WHEN m.X_TRANSACTION_ID LIKE '4GMiniApp%' THEN '4GMiniApp'
                        WHEN ip.DS_NCC_ID LIKE 'EBU_POSTPAY%' THEN 'EBU_POSTPAY'
                        ELSE 'Unknown'
                    END AS channel_details,
                    CASE
                        WHEN m.ID_TRAFFIC_TYPE = 4
                        AND m.ID_FIN_BLNC_MVMN_TYPE = 2 THEN 'Bundle_PURCHASE'
                        WHEN m.ID_TRAFFIC_TYPE = 4
                        AND m.ID_FIN_BLNC_MVMN_TYPE = 3 THEN 'Renewals'
                        WHEN m.ID_TRAFFIC_TYPE = 6
                        AND m.ID_FIN_BLNC_MVMN_TYPE = 4 THEN 'Bundle_LOAN'
                    END AS purchase_type,
                    CASE
                        WHEN ip.id_product = -1
                        AND m.X_ADDITIONAL_INFO LIKE '%J4U%' THEN REGEXP_SUBSTR (m.X_ADDITIONAL_INFO, '/([^~]+)~', 1, 1, NULL, 1)
                        ELSE ip.DS_NCC_ID
                    END AS bundle_name,
                    CASE
                        WHEN m.ID_FIN_BLNC_MVMN_TYPE = 3 THEN NVL (m.VL_ACCNT_BLNC, m.VL_FEE)
                        ELSE m.VL_ACCNT_BLNC
                    END AS vl_accnt_blnc
                FROM
                    BI.FCT_FIN_BLNC_MVMN m
                    LEFT JOIN BI.REF_PRODUCT ip ON ip.ID_PRODUCT = m.ID_PRODUCT
                    LEFT JOIN BI.REF_LOCATION rl ON rl.ID_LOCATION = m.ID_LOCATION
                WHERE
                    m.ID_DATE BETWEEN 20260929 AND 20260929
                    AND m.ID_RESULTS = 1
                    AND (
                        (
                            m.ID_TRAFFIC_TYPE = 4
                            AND m.ID_FIN_BLNC_MVMN_TYPE IN (2, 3)
                        )
                        OR (
                            m.ID_TRAFFIC_TYPE = 6
                            AND m.ID_FIN_BLNC_MVMN_TYPE = 4
                            AND (
                                m.X_ADDITIONAL_INFO LIKE '%DAILY%'
                                OR m.X_ADDITIONAL_INFO LIKE '%WEEKLY%'
                                OR m.X_ADDITIONAL_INFO LIKE '%MONTHLY%'
                            )
                        )
                    )
            ) t
        GROUP BY
            id_date,
            DS_SITE_ID,
            msisdn,
            channel,
            channel_details,
            purchase_type,
            REGEXP_REPLACE (
                TRIM(t.Bundle_name),
                '(_+(SAF_(OF|RC)|RC(_V[0-9]+)?|OF(_V[0-9]+)?|V[0-9]+))+$',
                '',
                1,
                0,
                'i'
            )
    )
SELECT
    /*+
    APPEND
    PARALLEL(a 32)
    PARALLEL(b 32)
    PARALLEL(c 32)
    USE_HASH(b)
    USE_HASH(c)
    */
    a.ID_DATE,
    a.DS_SITE_ID,
    a.MSISDN,
    a.CHANNEL,
    a.CHANNEL_DETAILS,
    a.BUNDLE_NAME,
    a.PURCHASE_QTY,
    a.PURCHASE_AMT,
    b."Bundle_Owner",
    b."Validity",
    b."Bundle_Type",
    b.ATL_BTL,
    b.COMBO_STATUS,
    c.BUNDLE_GROUP_TYPE,
    c.VALIDTY_GROUP
FROM
    TEMP_K a
    LEFT JOIN (
        SELECT DISTINCT
            BUNDLE_NAME,
            "Bundle_Owner",
            "Validity",
            "Bundle_Type",
            ATL_BTL,
            COMBO_STATUS
        FROM
            MELKAMU_WALE.BUNDLE_GROUP_V2
    ) b ON a.BUNDLE_NAME = b.BUNDLE_NAME
    LEFT JOIN (
        SELECT DISTINCT
            BUNDLE_NAME,
            BUNDLE_GROUP_TYPE,
            VALIDTY_GROUP
        FROM
            MELKAMU_WALE.MW_BUNDLE_MAPPING
    ) c ON a.BUNDLE_NAME = c.BUNDLE_NAME
WHERE
    a.BUNDLE_NAME IS NOT NULL;
