select
    SCHEMA_NAME     as SCHEMA_NAME
   ,EP_SCHEMA.VALUE as SCHEMA_COMMENT
from
    information_schema.schemata S
        INNER JOIN sys.schemas SS ON SS.NAME = S.SCHEMA_NAME
        LEFT  JOIN sys.extended_properties EP_SCHEMA ON SS.SCHEMA_ID  = EP_SCHEMA.MAJOR_ID
                                                    AND EP_SCHEMA.MINOR_ID= 0
                                                    AND EP_SCHEMA.CLASS_DESC  = 'SCHEMA'
                                                    AND UPPER(EP_SCHEMA.NAME) = 'COMMENT_DESCRIPTION'
where
  exists (select 1 from information_schema.tables where table_schema = S.schema_name)
order by
  S.SCHEMA_NAME