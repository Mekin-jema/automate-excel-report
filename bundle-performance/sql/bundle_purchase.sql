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