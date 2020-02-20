select
    SCHEMA_NAME  as SCHEMA_NAME
   ,''           as SCHEMA_COMMENT
from
    information_schema.schemata
order by
    SCHEMA_NAME