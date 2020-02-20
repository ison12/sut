select
    distinct(owner) as SCHEMA_NAME
   ,''  as SCHEMA_COMMENT
from
    all_tables
order by
    owner