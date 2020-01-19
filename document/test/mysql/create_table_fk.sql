-- FKを持つテーブル
DROP TABLE fk_;

CREATE TABLE fk_ (
    col1 VARCHAR(30)
   ,col2 VARCHAR(30)
   ,PRIMARY KEY (col1)
   ,CONSTRAINT fk_fk_ FOREIGN KEY (col2) REFERENCES notnull_(col1)
) ENGINE=INNODB;


-- FKを複数個持つテーブル
DROP TABLE fk_multiple1;

CREATE TABLE fk_multiple1 (
    col1    VARCHAR(30)
   ,col2    VARCHAR(30)
   ,col3    VARCHAR(30)
   ,col4    VARCHAR(30)
   ,col5    VARCHAR(30)
   ,col6    VARCHAR(30)
   ,col7    VARCHAR(30)
   ,col8    VARCHAR(30)
   ,col9    VARCHAR(30)
   ,col10   VARCHAR(30)
   ,col11   VARCHAR(30)
   ,CONSTRAINT fk_multiple1 FOREIGN KEY (col2, col3, col4, col5, col6, col7, col8, col9, col10, col11) REFERENCES pk_multiple1(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
   ,PRIMARY KEY (col1)
) ENGINE=INNODB;

-- FKを複数個・複数テーブルを参照するテーブル
DROP TABLE fk_multiple2;

CREATE TABLE fk_multiple2 (
    col1    VARCHAR(30)
   ,col2    VARCHAR(30)
   ,col3    VARCHAR(30)
   ,col4    VARCHAR(30)
   ,col5    VARCHAR(30)
   ,col6    VARCHAR(30)
   ,col7    VARCHAR(30)
   ,col8    VARCHAR(30)
   ,col9    VARCHAR(30)
   ,col10   VARCHAR(30)
   ,col11   VARCHAR(30)
   ,col12   VARCHAR(30)
   ,col13   VARCHAR(30)
   ,col14   VARCHAR(30)
   ,col15   VARCHAR(30)
   ,col16   VARCHAR(30)
   ,CONSTRAINT fk_multiple2_1 FOREIGN KEY (col2, col3, col4, col5, col6, col7, col8, col9, col10, col11) REFERENCES pk_multiple1(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
   ,CONSTRAINT fk_multiple2_2 FOREIGN KEY (col12, col13, col14, col15, col16) REFERENCES uk_multiple(col2, col3, col4, col5, col6)
   ,PRIMARY KEY (col1)
) ENGINE=INNODB;
