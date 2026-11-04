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
