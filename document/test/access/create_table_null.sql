-- NOT NULLêßñÒ
DROP TABLE notnull_;

CREATE TABLE notnull_ (
    col1    TEXT(30)
   ,col2    TEXT(30) NOT NULL
   ,col3    TEXT(30) NOT NULL
   ,col4    TEXT(30)
   ,col5    TEXT(30) NOT NULL
   ,CONSTRAINT notnull_pk PRIMARY KEY (col1)
);
