select
    c.owner          as TABLE_SCHEMA
   ,c.table_name     as TABLE_NAME
   ,tc.comments      as TABLE_COMMENT
   ,c.column_name    as COLUMN_NAME
   ,c.data_type      as COLUMN_TYPE
   ,c.nullable       as IS_NULL
   ,c.data_default   as DEFAULT_VALUE
   ,c.char_length    as CHAR_LENGTH
   ,c.data_precision as DATA_PRECISION
   ,c.data_scale     as DATA_SCALE
   ,c.data_precision as DATETIME_PRECISION
   ,c.data_precision as INTERVAL_PRECISION
   ,cc.comments      as COLUMN_COMMENT
from
    all_tab_columns c
        left join all_tab_comments tc on c.owner = tc.owner
         and c.table_name  = tc.table_name
        left join all_col_comments cc on c.owner   = cc.owner
         and c.table_name  = cc.table_name
         and c.column_name = cc.column_name
where
${condition}
order by
    c.owner
   ,c.table_name
   ,c.column_id