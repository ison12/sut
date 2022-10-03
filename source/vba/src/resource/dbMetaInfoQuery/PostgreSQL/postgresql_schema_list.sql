SELECT
    schema.oid        as SCHEMA_ID
   ,schema.nspname    as SCHEMA_NAME
   ,des.description   as SCHEMA_COMMENT
FROM
    pg_catalog.pg_namespace schema LEFT JOIN pg_catalog.pg_description des ON schema.oid = des.objoid
ORDER BY
    schema.nspname