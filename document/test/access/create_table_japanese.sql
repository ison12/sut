-- 日本語カラムを持つテーブル
DROP TABLE [テーブル];

CREATE TABLE [テーブル] (
    [カラム１]    VARCHAR(30)
   ,[カラム２]    NUMERIC(10)
   ,[カラム３]  NUMERIC(10)
   ,PRIMARY KEY ([カラム１])
);


-- PKを複数個持つテーブル
DROP TABLE [日本語テーブル];

CREATE TABLE [日本語テーブル] (
    [カラム１]    VARCHAR(30)
   ,[カラム２]    VARCHAR(30)
   ,[カラム３]    VARCHAR(30)
   ,[カラム４]    VARCHAR(30)
   ,[カラム数値５]    NUMERIC(10,3)
   ,CONSTRAINT [プライマリキー] PRIMARY KEY ([カラム１])
   ,CONSTRAINT [ユニークキー] UNIQUE([カラム２])
);
