-- UK‚ðŽ‚Âƒe[ƒuƒ‹
DROP TABLE default_;

CREATE TABLE default_ (
    col1    VARCHAR(30)
   ,col2    NUMERIC(10) DEFAULT 10
   ,col3    VARCHAR(30) DEFAULT 'TEST'
   ,PRIMARY KEY (col1)
) ENGINE=INNODB;

