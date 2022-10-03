SELECT
    const.constraint_name      as CONSTRAINT_NAME
   ,const.table_schema         as TABLE_SCHEMA
   ,const.table_name           as TABLE_NAME
   ,key.column_name            as COLUMN_NAME
   ,ref_key.table_schema       as REF_TABLE_SCHEMA
   ,ref_key.table_name         as REF_TABLE_NAME
   ,ref_key.column_name        as REF_COLUMN_NAME
   ,case 
      when const.constraint_type = 'PRIMARY KEY' then 'P'
      when const.constraint_type = 'UNIQUE'      then 'U'
      when const.constraint_type = 'FOREIGN KEY' then 'F'
      else const.constraint_type
    end as CONSTRAINT_TYPE
FROM
    information_schema.table_constraints const LEFT JOIN information_schema.key_column_usage key on
                                                            key.constraint_schema = const.constraint_schema
                                                        and key.constraint_name   = const.constraint_name
                                             LEFT JOIN information_schema.referential_constraints ref_const on
                                                            key.constraint_schema = ref_const.constraint_schema
                                                        and key.constraint_name   = ref_const.constraint_name
                                             LEFT JOIN information_schema.key_column_usage ref_key on
                                                            ref_const.unique_constraint_schema = ref_key.constraint_schema
                                                        and ref_const.unique_constraint_name   = ref_key.constraint_name
                                                        and key.ordinal_position               = ref_key.ordinal_position
WHERE
${condition}
ORDER BY
    const.constraint_schema
   ,const.constraint_name
   ,const.table_schema
   ,const.table_name
   ,const.constraint_type
   ,key.ordinal_position