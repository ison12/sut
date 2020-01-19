SELECT
    TRIM(S.SCHEMA_NAME)     AS SCHEMA_NAME
   ,TRIM(SC.COMMENT_VALUE)  AS SCHEMA_COMMENT
FROM
    RDBII_SYSTEM.RDBII_SCHEMA  S
        LEFT OUTER JOIN RDBII_SYSTEM.RDBII_COMMENT SC ON S.DB_CODE = SC.DB_CODE
                                                     AND S.SCHEMA_CODE = SC.SCHEMA_CODE
                                                     AND SC.COMMENT_TYPE = 'SC'
                                                     AND S.DB_NAME = ${db_name}
ORDER BY
  S.SCHEMA_NAME