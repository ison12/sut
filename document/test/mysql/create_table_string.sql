use test;

drop table data_type_string;

create table data_type_string (
    primarycol    BIGINT  primary key
   ,col1_1    CHAR
   ,col1_2    CHAR(255)
   ,col1_3    CHAR(255) CHARACTER SET sjis
   ,col2_1    NATIONAL CHAR
   ,col2_2    NATIONAL CHAR(255)
   ,col3_1    VARCHAR(255)
   ,col3_2    VARCHAR(255) CHARACTER SET sjis
   ,col4_1    NATIONAL VARCHAR(255)
   ,col5_1    CHARACTER VARYING(255)
   ,col5_2    CHARACTER VARYING(255) CHARACTER SET sjis
   ,col6_1    NATIONAL CHARACTER VARYING(255)
   ,col7_1    CHAR BYTE
   ,col7_2    BINARY
   ,col7_3    BINARY(100)
   ,col8_1    VARBINARY(100)
   ,col9_1    TINYBLOB
   ,col10_1    TINYTEXT
   ,col10_2    TINYTEXT CHARACTER SET sjis
   ,col11_1    BLOB
   ,col11_2    BLOB(1000)
   ,col12_1    TEXT
   ,col12_2    TEXT(1000)
   ,col12_3    TEXT(1000) CHARACTER SET sjis
   ,col13_1    MEDIUMBLOB
   ,col14_1    MEDIUMTEXT
   ,col14_2    MEDIUMTEXT CHARACTER SET sjis
   ,col15_1    LONGBLOB
   ,col16_1    LONGTEXT
   ,col16_2    LONGTEXT CHARACTER SET sjis
   ,col17_1    ENUM('あ', 'い')
   ,col17_2    ENUM('あ', 'い') CHARACTER SET sjis
   ,col18_1    SET('あ', 'い', 'う')
   ,col18_2    SET('あ', 'い', 'う') CHARACTER SET sjis
) ENGINE=INNODB;


-- test2データベースで、test1と同名のテーブルを異なるカラムで登録する
use test2;

drop table data_type_string;

create table data_type_string (
    primarycol    BIGINT  primary key
   ,test2 varchar(30)
   ,test3 varchar(30)
   ,col7_2 char(255)
) ENGINE=INNODB;
