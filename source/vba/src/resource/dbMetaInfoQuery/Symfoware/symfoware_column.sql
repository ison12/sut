SELECT
     TRIM(S.SCHEMA_NAME)        AS TABLE_SCHEMA
    ,TRIM(T.TABLE_NAME)         AS TABLE_NAME
    ,TRIM(TCOM.COMMENT_VALUE)   AS TABLE_COMMENT
    ,TRIM(C.COLUMN_NAME)        AS COLUMN_NAME
    ,CASE TRIM(C.DATA_TYPE) 
     WHEN 'CH' THEN TRIM('CHAR')
     WHEN 'CV' THEN TRIM('VARCHAR')
     WHEN 'BL' THEN TRIM('BLOB')
     WHEN 'CN' THEN TRIM('NCHAR')
     WHEN 'NV' THEN TRIM('NCHAR VARYING')
     WHEN 'IN' THEN TRIM('INT')
     WHEN 'SI' THEN TRIM('SMALLINT')
     WHEN 'NU' THEN TRIM('NUMERIC')
     WHEN 'DE' THEN TRIM('DECIMAL')
     WHEN 'FL' THEN TRIM('FLOAT')
     WHEN 'DP' THEN TRIM('DOUBLE PRECISION')
     WHEN 'RE' THEN TRIM('REAL')
     WHEN 'TM' THEN TRIM('TIMESTAMP')
     WHEN 'DT' THEN TRIM('DATE')
     WHEN 'TI' THEN TRIM('TIME')
     WHEN 'IT' THEN CASE C.NUMERIC_RADIX WHEN 1 THEN TRIM('INTERVAL YEAR')
                                         WHEN 2 THEN TRIM('INTERVAL MONTH')
                                         WHEN 3 THEN TRIM('INTERVAL DAY')
                                         WHEN 4 THEN TRIM('INTERVAL HOUR')
                                         WHEN 5 THEN TRIM('INTERVAL MINUTE')
                                         WHEN 6 THEN TRIM('INTERVAL SECOND')
                                         WHEN 7 THEN TRIM('INTERVAL YEAR TO MONTH')
                                         WHEN 8 THEN TRIM('INTERVAL DAY TO HOUR')
                                         WHEN 9 THEN TRIM('INTERVAL DAY TO MINUTE')
                                         WHEN 10 THEN TRIM('INTERVAL DAY TO SECOND')
                                         WHEN 11 THEN TRIM('INTERVAL HOUR TO MINUTE')
                                         WHEN 12 THEN TRIM('INTERVAL HOUR TO SECOND')
                                         WHEN 13 THEN TRIM('INTERVAL MINUTE TO SECOND')
                                         ELSE TRIM('INTERVAL') END
               ELSE TRIM(C.DATA_TYPE) END AS COLUMN_TYPE
    ,TRIM(C.DATA_TYPE)                                  AS COLUMN_TYPE_OMISSION
    ,CASE TRIM(C.NULLABLE) WHEN 'NO' THEN TRIM('N')
                                     ELSE TRIM('Y') END AS IS_NULL
    ,TRIM(C.DEFAULT_VALUE)                              AS DEFAULT_VALUE
    ,C.CHAR_OCTET_LENGTH                                AS "CHAR_LENGTH"
    ,C.NUMERIC_PRECISION                                AS DATA_PRECISION
    ,C.NUMERIC_SCALE                                    AS DATA_SCALE
    ,C.NUMERIC_PRECISION                                AS DATETIME_PRECISION
    ,C.NUMERIC_PRECISION                                AS INTERVAL_PRECISION
    ,TRIM(CCOM.COMMENT_VALUE)                           AS COLUMN_COMMENT
FROM
    RDBII_SYSTEM.RDBII_SCHEMA S
   ,RDBII_SYSTEM.RDBII_TABLE  T LEFT OUTER JOIN
       (SELECT
            DB_CODE
           ,SCHEMA_CODE
           ,TABLE_CODE
           ,COMMENT_VALUE
        FROM
            RDBII_SYSTEM.RDBII_COMMENT
        WHERE
            COMMENT_TYPE = 'TV') TCOM ON T.DB_CODE = TCOM.DB_CODE
                                     AND T.SCHEMA_CODE = TCOM.SCHEMA_CODE
                                     AND T.TABLE_CODE = TCOM.TABLE_CODE
   ,RDBII_SYSTEM.RDBII_COLUMN C LEFT OUTER JOIN
       (SELECT
            DB_CODE
           ,SCHEMA_CODE
           ,TABLE_CODE
           ,COLUMN_CODE
           ,COMMENT_VALUE
        FROM
            RDBII_SYSTEM.RDBII_COMMENT
        WHERE
            COMMENT_TYPE = 'CL') CCOM ON C.DB_CODE   = CCOM.DB_CODE
                                     AND C.SCHEMA_CODE = CCOM.SCHEMA_CODE
                                     AND C.TABLE_CODE = CCOM.TABLE_CODE
                                     AND C.COLUMN_CODE = CCOM.COLUMN_CODE
WHERE
    S.DB_CODE   = T.DB_CODE
AND S.SCHEMA_CODE = T.SCHEMA_CODE
AND T.DB_CODE   = C.DB_CODE
AND T.SCHEMA_CODE = C.SCHEMA_CODE
AND T.TABLE_CODE = C.TABLE_CODE
AND S.DB_NAME = ${db_name}
AND ${condition}
ORDER BY
    S.SCHEMA_NAME
   ,T.TABLE_NAME
   ,C.ORDINAL_POSITION