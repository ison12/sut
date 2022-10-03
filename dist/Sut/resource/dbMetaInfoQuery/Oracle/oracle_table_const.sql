select
    const.owner              as CONSTRAINT_SCHEMA
   ,const.constraint_name    as CONSTRAINT_NAME
   ,const.owner              as TABLE_SCHEMA
   ,const.table_name         as TABLE_NAME
   ,key_column.column_name   as COLUMN_NAME
   ,ref_key.owner            as REF_TABLE_SCHEMA
   ,ref_key.table_name       as REF_TABLE_NAME
   ,ref_key.column_name      as REF_COLUMN_NAME
   ,DECODE(const.constraint_type
       , 'P', 'P'
       , 'U', 'U'
       , 'R', 'F'
       , const.constraint_type) as CONSTRAINT_TYPE
from
    all_constraints const
        left join all_cons_columns key_column on const.owner   = key_column.owner
                                             and const.constraint_name = key_column.constraint_name
                                             and const.table_name  = key_column.table_name
        left join all_constraints  ref_const on const.r_owner   = ref_const.owner
                                            and const.r_constraint_name = ref_const.constraint_name
        left join all_cons_columns ref_key on const.r_owner   = ref_key.owner
                                          and const.r_constraint_name = ref_key.constraint_name
                                          and ref_const.table_name= ref_key.table_name
                                          and key_column.position = ref_key.position
where
${condition}
order by
    const.owner
   ,const.constraint_name
   ,const.table_name
   ,const.constraint_type
   ,key_column.position