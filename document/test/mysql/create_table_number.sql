DROP TABLE data_type_number;

create table data_type_number (
    primarycol    BIGINT primary key
   ,col1_1    BIT
   ,col1_2    BIT(3)
   ,col2_1    TINYINT
   ,col2_2    TINYINT(3)
   ,col2_3    TINYINT UNSIGNED
   ,col2_4    TINYINT UNSIGNED ZEROFILL
   ,col3_1    BOOL
   ,col3_2    BOOLEAN
   ,col4_1    SMALLINT
   ,col4_2    SMALLINT(3)
   ,col4_3    SMALLINT UNSIGNED
   ,col4_4    SMALLINT UNSIGNED ZEROFILL
   ,col5_1    MEDIUMINT
   ,col5_2    MEDIUMINT(3)
   ,col5_3    MEDIUMINT UNSIGNED
   ,col5_4    MEDIUMINT UNSIGNED ZEROFILL
   ,col6_1    INT
   ,col6_2    INT(3)
   ,col6_3    INT UNSIGNED
   ,col6_4    INT UNSIGNED ZEROFILL
   ,col7_1    INTEGER
   ,col7_2    INTEGER(3)
   ,col7_3    INTEGER UNSIGNED
   ,col7_4    INTEGER UNSIGNED ZEROFILL
   ,col8_1    BIGINT
   ,col8_2    BIGINT(3)
   ,col8_3    BIGINT UNSIGNED
   ,col8_4    BIGINT UNSIGNED ZEROFILL
   ,col9_1    FLOAT
   ,col9_2    FLOAT(10)
   ,col9_3    FLOAT(10,3)
   ,col9_4    FLOAT UNSIGNED
   ,col9_5    FLOAT UNSIGNED ZEROFILL
   ,col10_1    DOUBLE
   ,col10_2    DOUBLE(10,3)
   ,col10_3    DOUBLE UNSIGNED
   ,col10_4    DOUBLE UNSIGNED ZEROFILL
   ,col11_1    DOUBLE PRECISION
   ,col11_2    DOUBLE PRECISION(10,3)
   ,col11_3    DOUBLE PRECISION UNSIGNED
   ,col11_4    DOUBLE PRECISION UNSIGNED ZEROFILL
   ,col12_1    REAL
   ,col12_2    REAL(10,3)
   ,col12_3    REAL UNSIGNED
   ,col12_4    REAL UNSIGNED ZEROFILL
   ,col13_1    DECIMAL
   ,col13_2    DECIMAL(10)
   ,col13_3    DECIMAL(10,3)
   ,col13_4    DECIMAL UNSIGNED
   ,col13_5    DECIMAL UNSIGNED ZEROFILL
   ,col14_1    DEC
   ,col14_2    DEC(10)
   ,col14_3    DEC(10,3)
   ,col15_4    DEC UNSIGNED
   ,col15_5    DEC UNSIGNED ZEROFILL
   ,col16_1    NUMERIC
   ,col16_2    NUMERIC(10)
   ,col16_3    NUMERIC(10,3)
   ,col16_4    NUMERIC UNSIGNED
   ,col16_5    NUMERIC UNSIGNED ZEROFILL
   ,col17_1    FIXED
   ,col17_2    FIXED(10)
   ,col17_3    FIXED(10,3)
   ,col17_4    FIXED UNSIGNED
   ,col17_5    FIXED UNSIGNED ZEROFILL
) ENGINE=INNODB;
