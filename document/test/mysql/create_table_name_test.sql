-- �e�[�u������31����
DROP TABLE table_name_test1234567891234567;

CREATE TABLE table_name_test1234567891234567 (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- �e�[�u������32����
DROP TABLE table_name_test12345678912345678;

CREATE TABLE table_name_test12345678912345678 (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- �e�[�u������32������
DROP TABLE table_name_test1234567891234567890123456789456123456798;

CREATE TABLE table_name_test1234567891234567890123456789456123456798 (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- �e�[�u������Excel�ł͈����Ȃ��s���ȕ�����i����ɑ}���j
DROP TABLE `table_name_test\[:]*/?`;

CREATE TABLE `table_name_test\[:]*/?` (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- �e�[�u������Excel�ł͈����Ȃ��s���ȕ�����i�O���ɑ}���j
DROP TABLE `\[:]*/?table_name_test`;

CREATE TABLE `\[:]*/?table_name_test` (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- �e�[�u������Excel�ł͈����Ȃ��s���ȕ�����i���Ԃɑ}���j
DROP TABLE `table_name_\[:]*/?test`;

CREATE TABLE `table_name_\[:]*/?test` (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);
