-- ���[�U���쐬����
CREATE USER test
  IDENTIFIED BY password
  DEFAULT   TABLESPACE users
  TEMPORARY TABLESPACE temp;
  
-- ���[�U�Ɍ�����t�^����
GRANT CONNECT TO test;
GRANT RESOURCE TO test;

-- Oracle�̑g�ݍ��݌^�̑S�f�[�^�^�C�v
DROP TABLE data_type_oracle;

CREATE TABLE data_type_oracle (
    col1 CHAR(30)
   ,col2 VARCHAR2(30) DEFAULT 'default'
   ,col3 NCHAR(30) DEFAULT 'default'
   ,col4 NVARCHAR2(30) DEFAULT 'default'
   ,col5_1 NUMBER(10, 3) DEFAULT 1
   ,col5_2 NUMBER(10) DEFAULT 1
   ,col5_3 NUMBER DEFAULT 1
   ,col6 BINARY_FLOAT DEFAULT 1
   ,col7 BINARY_DOUBLE DEFAULT 1
   ,col8 DATE DEFAULT sysdate
   ,col9_1 TIMESTAMP(3) DEFAULT sysdate
   ,col9_2 TIMESTAMP DEFAULT sysdate
   ,col10_1 TIMESTAMP(3) WITH TIME ZONE DEFAULT sysdate
   ,col10_2 TIMESTAMP WITH TIME ZONE DEFAULT sysdate
   ,col11_1 TIMESTAMP(3) WITH LOCAL TIME ZONE DEFAULT sysdate
   ,col11_2 TIMESTAMP WITH LOCAL TIME ZONE DEFAULT sysdate
   ,col12 TIMESTAMP WITH TIME ZONE DEFAULT sysdate
   ,col13_1 INTERVAL YEAR(4) TO MONTH
   ,col13_2 INTERVAL YEAR(2) TO MONTH
   ,col14_1 INTERVAL DAY(4) TO SECOND(3)
   ,col14_2 INTERVAL DAY(4) TO SECOND(2)
   ,col15 LONG
   -- ,col16 LONG RAW LONG�^�̓e�[�u���Ɉ�����܂߂��Ȃ�
   ,col17 RAW(1000)
   ,col18 ROWID
   ,col19 UROWID(1000)
   ,col20 BLOB
   ,col21 CLOB
   ,col22 NCLOB
   ,col23 BFILE
   --,col24 XML Type
   ,PRIMARY KEY (col1)
);

-- �R�����g�����e�[�u��
COMMENT ON TABLE  data_type_oracle IS 'ORACLE�̑g�ݍ��݃f�[�^�^';
COMMENT ON COLUMN data_type_oracle.col1 IS '�J����1';
COMMENT ON COLUMN data_type_oracle.col2 IS '�J����2';
COMMENT ON COLUMN data_type_oracle.col3 IS '�J����3';
COMMENT ON COLUMN data_type_oracle.col4 IS '�J����4';
COMMENT ON COLUMN data_type_oracle.col6 IS '�J����6';
COMMENT ON COLUMN data_type_oracle.col7 IS '�J����7';
COMMENT ON COLUMN data_type_oracle.col8 IS '�J����8';


-- ���̑��̃f�[�^�^�C�v
DROP TABLE data_type_other;

CREATE TABLE data_type_other (
    col1    CHARACTER(30)
   ,col2    CHAR(30)
   ,col3    CHARACTER VARYING(30)
   ,col4    CHAR VARYING(30)
   ,col5    NATIONAL CHARACTER(30)
   ,col6    NATIONAL CHAR(30)
   ,col7    NATIONAL CHARACTER VARYING(30)
   ,col8    NATIONAL CHAR VARYING(30)
   ,col9    NCHAR VARYING(30)
   ,col10    NUMERIC(10,3)
   ,col11    DECIMAL(10,3)
   ,col12    INTEGER
   ,col13    INT
   ,col14    SMALLINT
   ,col15    FLOAT(10)
   ,col18    VARCHAR(30)
   ,PRIMARY KEY (col1)
);

   --,col16    DOUBLE PRECISION(128)
   --,col17    REAL(63)
   --,col19    LONG VARCHAR(30)

-- PK�����e�[�u��
DROP TABLE pk_none;

CREATE TABLE pk_none (
    col1    VARCHAR2(30)
   ,col2    VARCHAR2(30)
);

-- PK�𕡐����e�[�u��
DROP TABLE pk_multiple1;

CREATE TABLE pk_multiple1 (
    col1    VARCHAR2(30)
   ,col2    VARCHAR2(30)
   ,col3    VARCHAR2(30)
   ,col4    VARCHAR2(30)
   ,col5    VARCHAR2(30)
   ,col6    VARCHAR2(30)
   ,col7    VARCHAR2(30)
   ,col8    VARCHAR2(30)
   ,col9    VARCHAR2(30)
   ,col10   VARCHAR2(30)
   ,col11   NUMERIC(10,3)
   ,PRIMARY KEY (col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
);

-- PK�𕡐����e�[�u��
DROP TABLE pk_multiple2;

CREATE TABLE pk_multiple2 (
    col1    VARCHAR2(30)
   ,col2    VARCHAR2(30)
   ,col3    VARCHAR2(30)
   ,col4    VARCHAR2(30)
   ,col5    VARCHAR2(30)
   ,col6    VARCHAR2(30)
   ,col7    VARCHAR2(30)
   ,col8    VARCHAR2(30)
   ,col9    VARCHAR2(30)
   ,col10   VARCHAR2(30)
   ,col11   VARCHAR2(30)
   ,col12   NUMERIC(10,3)
   ,PRIMARY KEY (col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11)
);

-- UK�����e�[�u��
DROP TABLE uk_;

CREATE TABLE uk_ (
    col1    VARCHAR2(30)
   ,col2    NUMERIC(10)
   ,PRIMARY KEY (col1)
   ,CONSTRAINT uk_uk_ UNIQUE(col2)
);

-- UK�𕡐����e�[�u��
DROP TABLE uk_multiple;

CREATE TABLE uk_multiple (
    col1    VARCHAR2(30)
   ,col2    VARCHAR2(30)
   ,col3    VARCHAR2(30)
   ,col4    VARCHAR2(30)
   ,col5    VARCHAR2(30)
   ,col6    VARCHAR2(30)
   ,col7    VARCHAR2(30)
   ,col8    VARCHAR2(30)
   ,col9    VARCHAR2(30)
   ,col10   VARCHAR2(30)
   ,col11   VARCHAR2(30)
   ,col12   NUMERIC(10)
   ,PRIMARY KEY (col1)
   ,CONSTRAINT uk_uk_multiple_1 UNIQUE(col2, col3, col4, col5, col6)
   ,CONSTRAINT uk_uk_multiple_2 UNIQUE(col7, col8, col9, col10, col11)
);

-- FK�����e�[�u��
DROP TABLE fk_;

CREATE TABLE fk_ (
    col1 CHAR(30)
   ,col2 CHAR(30)
   ,PRIMARY KEY (col1)
   ,CONSTRAINT fk_fk_ FOREIGN KEY (col2) REFERENCES data_type_oracle(col1)
);

-- FK�𕡐����e�[�u��
DROP TABLE fk_multiple1;

CREATE TABLE fk_multiple1 (
    col1    VARCHAR2(30)
   ,col2    VARCHAR2(30)
   ,col3    VARCHAR2(30)
   ,col4    VARCHAR2(30)
   ,col5    VARCHAR2(30)
   ,col6    VARCHAR2(30)
   ,col7    VARCHAR2(30)
   ,col8    VARCHAR2(30)
   ,col9    VARCHAR2(30)
   ,col10   VARCHAR2(30)
   ,col11   VARCHAR2(30)
   ,col12   NUMERIC(10,3)
   ,CONSTRAINT fk_multiple1 FOREIGN KEY (col2, col3, col4, col5, col6, col7, col8, col9, col10, col11) REFERENCES pk_multiple1(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
   ,PRIMARY KEY (col1)
);

-- NOT NULL����
DROP TABLE notnull_;

CREATE TABLE notnull_ (
    col1    VARCHAR2(30)
   ,col2    VARCHAR2(30) NOT NULL
   ,col3    VARCHAR2(30) NOT NULL
   ,col4    VARCHAR2(30)
   ,col5    VARCHAR2(30) NOT NULL
   ,PRIMARY KEY (col1)
);
