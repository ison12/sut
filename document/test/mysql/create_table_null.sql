-- NOT NULLêßñÒ
DROP TABLE notnull_;

CREATE TABLE notnull_ (
    col1    VARCHAR(30)
   ,col2    VARCHAR(30) NOT NULL
   ,col3    VARCHAR(30) NOT NULL
   ,col4    VARCHAR(30)
   ,col5    VARCHAR(30) NOT NULL
   ,PRIMARY KEY (col1)
) ENGINE=INNODB;
