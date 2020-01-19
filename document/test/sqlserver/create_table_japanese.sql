

CREATE SCHEMA [テストスキーマ];

CREATE TABLE [テストスキーマ].[テーブル] (
  [テストカラム１] nchar(10)
 ,[テストカラム２] nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
 ,c5 nchar(10)
 ,constraint [テストPK] primary key ([テストカラム１])
 ,constraint [テストUK] unique ([テストカラム２])
);


CREATE TABLE [テストスキーマ].[テーブル２] (
  []][[] nchar(10)
);


CREATE SCHEMA [テストスキーマ]][];

CREATE TABLE [テストスキーマ]][].[テーブル]][] (
  [テストカラム１] nchar(10)
 ,[テストカラム２]][] nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
 ,c5 nchar(10)
 ,constraint [テストPK]][] primary key ([テストカラム１])
 ,constraint [テストUK]][] unique ([テストカラム２]][])
);

