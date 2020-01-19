-- テーブル名が31文字
DROP TABLE table_name_test1234567891234567;

CREATE TABLE table_name_test1234567891234567 (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- テーブル名が32文字
DROP TABLE table_name_test12345678912345678;

CREATE TABLE table_name_test12345678912345678 (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- テーブル名が32文字超
DROP TABLE table_name_test1234567891234567890123456789456123456798;

CREATE TABLE table_name_test1234567891234567890123456789456123456798 (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- テーブル名にExcelでは扱えない不正な文字列（後方に挿入）
DROP TABLE `table_name_test\[:]*/?`;

CREATE TABLE `table_name_test\[:]*/?` (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- テーブル名にExcelでは扱えない不正な文字列（前方に挿入）
DROP TABLE `\[:]*/?table_name_test`;

CREATE TABLE `\[:]*/?table_name_test` (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);

-- テーブル名にExcelでは扱えない不正な文字列（中間に挿入）
DROP TABLE `table_name_\[:]*/?test`;

CREATE TABLE `table_name_\[:]*/?test` (
    col1 CHAR(30)
   ,PRIMARY KEY (col1)
);
