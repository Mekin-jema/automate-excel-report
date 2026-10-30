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
