-- 実行条件
-- テストテーブル.sql が事前に実行されていること

-- system/password で接続する
GRANT CREATE ANY VIEW TO TEST;

-- test/password で接続する
CREATE VIEW view_data_type_oracle as select
    col1
   ,col2
   ,col3
   ,col4
   ,col5_1
   ,col5_2
   ,col5_3
   ,col6
   ,col7
   ,col8
   ,col9_1
   ,col9_2
   ,col10_1
   ,col10_2
   ,col11_1
   ,col11_2
   ,col12
   ,col13_1
   ,col13_2
   ,col14_1
   ,col14_2
   ,col15
   ,col17
   ,col18
   ,col19
   ,col20
   ,col21
   ,col22
   ,col23
from data_type_oracle;

-- test/password で接続する
DROP VIEW view_fk_uk_;
CREATE VIEW view_fk_uk_ as select
    fk_.col1 as fk_col1
   ,uk_.col1 as uk_col1
   ,fk_.col2 as fk_col2
from fk_ inner join uk_ on trim(fk_.col1) = uk_.col1
;
