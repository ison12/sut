drop table data_type_date;

create table data_type_date (
    primarycol   BIGINT primary key
   ,col1_1   DATE
   ,col2_2   DATETIME
   ,col3_3   TIMESTAMP
   ,col4_4   TIME
   ,col5_1   YEAR
   ,col5_2   YEAR(2)
   ,col5_3   YEAR(4)
) ENGINE=INNODB;
