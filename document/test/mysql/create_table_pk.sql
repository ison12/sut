-- PK無しテーブル
DROP TABLE pk_none;

CREATE TABLE pk_none (
    col1    VARCHAR(30)
   ,col2    VARCHAR(30)
) ENGINE=INNODB;

-- PKを複数個持つテーブル
DROP TABLE pk_multiple1;

CREATE TABLE pk_multiple1 (
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
   ,col11   NUMERIC(10,3)
   ,PRIMARY KEY (col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
) ENGINE=INNODB;

-- PKを複数個持つテーブル
DROP TABLE pk_multiple2;

CREATE TABLE pk_multiple2 (
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
   ,col12   NUMERIC(10,3)
   ,PRIMARY KEY (col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11)
) ENGINE=INNODB;

