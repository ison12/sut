-- UKを持つテーブル
DROP TABLE uk_;

CREATE TABLE uk_ (
    col1    TEXT(30)
   ,col2    NUMERIC(10)
   ,PRIMARY KEY (col1)
   ,CONSTRAINT uk_uk_ UNIQUE(col2)
);

-- UKを複数個持つテーブル
DROP TABLE uk_multiple;

CREATE TABLE uk_multiple (
    col1    TEXT(30)
   ,col2    TEXT(30)
   ,col3    TEXT(30)
   ,col4    TEXT(30)
   ,col5    TEXT(30)
   ,col6    TEXT(30)
   ,col7    TEXT(30)
   ,col8    TEXT(30)
   ,col9    TEXT(30)
   ,col10   TEXT(30)
   ,col11   TEXT(30)
   ,col12   NUMERIC(10)
   ,PRIMARY KEY (col1)
   ,CONSTRAINT uk_uk_multiple_1 UNIQUE(col2, col3, col4, col5, col6)
   ,CONSTRAINT uk_uk_multiple_2 UNIQUE(col7, col8, col9, col10, col11)
);

-- UKを複数個持つテーブル
DROP TABLE uk_multiple2;

CREATE TABLE uk_multiple2 (
    col1    TEXT(30)
   ,col2    TEXT(30)
   ,col3    TEXT(30)
   ,col4    TEXT(30)
   ,col5    TEXT(30)
   ,col6    TEXT(30)
   ,col7    TEXT(30)
   ,col8    TEXT(30)
   ,col9    TEXT(30)
   ,col10   TEXT(30)
   ,col11   TEXT(30)
   ,col12   NUMERIC(10)
   ,PRIMARY KEY (col1)
   ,CONSTRAINT uk_uk_multiple2_1 UNIQUE(col2, col3,  col6)
   ,CONSTRAINT uk_uk_multiple2_2 UNIQUE(col4, col8)
);