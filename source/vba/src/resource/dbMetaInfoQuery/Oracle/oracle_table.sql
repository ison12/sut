select
    c.owner  as TABLE_SCHEMA
   ,c.table_name as TABLE_NAME
   ,tc.comments  as TABLE_COMMENT
from
    all_tables c
        left join all_tab_comments tc on c.owner = tc.owner
                                     and c.table_name  = tc.table_name
where
${condition}
order by
    c.table_name