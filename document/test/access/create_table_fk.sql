-- FKを持つテーブル
DROP TABLE fk_;

CREATE TABLE fk_ (
    col1 TEXT(30)
   ,col2 TEXT(30)
   ,PRIMARY KEY (col1)
   ,CONSTRAINT fk_fk_ FOREIGN KEY (col2) REFERENCES notnull_(col1)
);


-- FKを複数個持つテーブル
DROP TABLE fk_multiple1;

CREATE TABLE fk_multiple1 (
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
   ,CONSTRAINT fk_multiple1 FOREIGN KEY (col2, col3, col4, col5, col6, col7, col8, col9, col10, col11) REFERENCES pk_multiple1(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
   ,PRIMARY KEY (col1)
);

-- FKを複数個・複数テーブルを参照するテーブル
DROP TABLE fk_multiple2;

CREATE TABLE fk_multiple2 (
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
   ,col12   TEXT(30)
   ,col13   TEXT(30)
   ,col14   TEXT(30)
   ,col15   TEXT(30)
   ,col16   TEXT(30)
   ,CONSTRAINT fk_multiple2_1 FOREIGN KEY (col2, col3, col4, col5, col6, col7, col8, col9, col10, col11) REFERENCES pk_multiple1(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
   ,CONSTRAINT fk_multiple2_2 FOREIGN KEY (col12, col13, col14, col15, col16) REFERENCES uk_multiple(col2, col3, col4, col5, col6)
   ,PRIMARY KEY (col1)
);
