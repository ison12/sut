SELECT
    C.TABLE_SCHEMA   AS TABLE_SCHEMA
   ,C.TABLE_NAME     AS TABLE_NAME
   ,T.TABLE_COMMENT  AS TABLE_COMMENT
   ,C.COLUMN_NAME    AS COLUMN_NAME
   ,C.COLUMN_TYPE    AS COLUMN_TYPE
   ,C.COLUMN_TYPE    AS COLUMN_TYPE_FORMAL
   ,CASE C.IS_NULLABLE
        WHEN 'YES' THEN 'Y'
        ELSE 'N'
    END AS IS_NULL
   ,C.COLUMN_DEFAULT            AS DEFAULT_VALUE
   ,C.CHARACTER_MAXIMUM_LENGTH  AS CHAR_LENGTH
   ,C.NUMERIC_PRECISION         AS DATA_PRECISION
   ,C.NUMERIC_SCALE             AS DATA_SCALE
   ,C.NUMERIC_PRECISION         AS DATETIME_PRECISION
   ,C.NUMERIC_PRECISION         AS INTERVAL_PRECISION
   ,C.COLUMN_COMMENT            AS COLUMN_COMMENT
FROM
    information_schema.TABLES T
        INNER JOIN information_schema.COLUMNS C
                ON T.TABLE_SCHEMA = C.TABLE_SCHEMA
               AND T.TABLE_NAME   = C.TABLE_NAME
WHERE
${condition}
ORDER BY
    C.TABLE_SCHEMA
   ,C.TABLE_NAME
   ,C.ORDINAL_POSITION