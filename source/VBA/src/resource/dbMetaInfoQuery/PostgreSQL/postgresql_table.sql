SELECT
    c.table_schema             as TABLE_SCHEMA
   ,c.table_name               as TABLE_NAME
   ,table_comm.table_comment   as TABLE_COMMENT
FROM
    information_schema.tables c
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
WHERE
    ${condition}
ORDER BY
    c.table_schema
   ,c.table_name