SELECT
    c.table_schema             as TABLE_SCHEMA
   ,c.table_name               as TABLE_NAME
   ,table_comm.table_comment   as TABLE_COMMENT
   ,c.column_name              as COLUMN_NAME
   ,c.data_type                as COLUMN_TYPE
   ,CASE
        WHEN position('CHAR'      in UPPER(c.data_type)) >= 1 AND c.character_maximum_length IS NOT NULL THEN c.data_type || '(' || CAST(c.character_maximum_length AS VARCHAR(32)) || ')'
        WHEN position('TEXT'      in UPPER(c.data_type)) >= 1 AND c.character_maximum_length IS NOT NULL THEN c.data_type || '(' || CAST(c.character_maximum_length AS VARCHAR(32)) || ')'
        WHEN position('BIT'       in UPPER(c.data_type)) >= 1 AND c.character_maximum_length IS NOT NULL THEN c.data_type || '(' || CAST(c.character_maximum_length AS VARCHAR(32)) || ')'
        WHEN position('NUMERIC'   in UPPER(c.data_type)) >= 1 AND (c.numeric_precision IS NOT NULL AND c.numeric_precision > 0)
                                                              AND (c.numeric_scale     IS NOT NULL AND c.numeric_scale     > 0)
                                                              THEN c.data_type || '(' || CAST(c.numeric_precision AS VARCHAR(32)) || ',' || CAST(c.numeric_scale AS VARCHAR(32)) || ')'
        WHEN position('NUMERIC'   in UPPER(c.data_type)) >= 1 AND (c.numeric_precision IS NOT NULL AND c.numeric_precision > 0)
                                                              AND (c.numeric_scale     IS NULL      OR c.numeric_scale     = 0)
                                                              THEN c.data_type || '(' || CAST(c.numeric_precision AS VARCHAR(32)) || ')'
        WHEN position('TIMESTAMP' in UPPER(c.data_type)) >= 1 AND (c.datetime_precision IS NOT NULL AND c.datetime_precision > 0)
                                                              THEN c.data_type || '(' || CAST(c.datetime_precision AS VARCHAR(32)) || ')'
        WHEN position('TIME'      in UPPER(c.data_type)) >= 1 AND (c.datetime_precision IS NOT NULL AND c.datetime_precision > 0)
                                                              THEN c.data_type || '(' || CAST(c.datetime_precision AS VARCHAR(32)) || ')'
        WHEN position('INTERVAL'  in UPPER(c.data_type)) >= 1 AND (c.datetime_precision IS NOT NULL AND c.datetime_precision > 0)
                                                              THEN c.data_type || '(' || CAST(c.datetime_precision AS VARCHAR(32)) || ')'
        ELSE c.data_type
    END                        as COLUMN_TYPE_FORMAL
   ,CASE
       WHEN c.is_nullable = 'YES' THEN 'Y'
       ELSE                            'N'
    END                        as IS_NULL
   ,c.column_default           as DEFAULT_VALUE
   ,c.character_maximum_length as CHAR_LENGTH
   ,c.numeric_precision        as DATA_PRECISION
   ,c.numeric_scale            as DATA_SCALE
   ,c.datetime_precision       as DATETIME_PRECISION
   ,c.interval_precision       as INTERVAL_PRECISION
   ,column_comm.column_comment as COLUMN_COMMENT
FROM
    information_schema.columns c
      LEFT JOIN
        (SELECT
            ns.oid           as schema_id
           ,a.oid            as table_id
           ,ns.nspname       as table_schema
           ,a.relname        as table_name
           ,des.description  as table_comment
        FROM
          pg_catalog.pg_class a INNER JOIN pg_catalog.pg_namespace   ns   ON a.relnamespace = ns.oid
                                INNER JOIN pg_catalog.pg_description des  ON a.oid          = des.objoid
                                                                         AND des.objsubid   = 0
        WHERE
            a.relkind IN ('r', 'v')) table_comm on c.table_schema = table_comm.table_schema
                                               and c.table_name   = table_comm.table_name
      LEFT JOIN
        (SELECT
            ns.oid           as schema_id
           ,a.oid            as table_id
           ,attr.attnum      as column_id
           ,ns.nspname       as table_schema
           ,a.relname        as table_name
           ,attr.attname     as column_name
           ,des2.description as column_comment
        FROM
          pg_catalog.pg_class a INNER JOIN pg_catalog.pg_namespace   ns   ON a.relnamespace = ns.oid
                                INNER JOIN pg_catalog.pg_attribute   attr ON a.oid          = attr.attrelid
                                                                         AND attr.attnum > 0
                                INNER JOIN pg_catalog.pg_description des2 ON a.oid          = des2.objoid
                                                                         AND attr.attnum    = des2.objsubid
        WHERE
            a.relkind   IN ('r', 'v')) column_comm ON c.table_schema = column_comm.table_schema
                                                  AND c.table_name   = column_comm.table_name
                                                  AND c.column_name  = column_comm.column_name
WHERE
${condition}
ORDER BY
    c.table_schema
   ,c.table_name
   ,c.ordinal_position