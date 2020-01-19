・sqlリファレンス
http://msdn.microsoft.com/ja-jp/library/bb510741.aspx

・information_schemaリファレンス
http://www.microsoft.com/japan/technet/prodtechnol/sql/2000/books/progref12.mspx

・データ型
http://msdn.microsoft.com/ja-jp/library/ms187752.aspx

CREATE DATABASE test;


sqlcmd -S (local)\SQLEXPRESS -d test

CREATE TABLE pk (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
);

CREATE TABLE pk_multiple (
  c1 nchar(10)
 ,c2 nchar(10)
 ,c3 nchar(20)
  constraint pk_multiple_pk primary key (c1, c2)
);

CREATE TABLE default_ (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) DEFAULT 'DEFAULT'
);

CREATE TABLE uk (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) UNIQUE
);

CREATE TABLE uk_multiple (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
 ,c3 nchar(10)
  constraint uk_multiple_uk unique (c2, c3)
);

CREATE TABLE uk_multiple2 (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
 ,c5 nchar(10)
  constraint uk_multiple2_uk1 unique (c2, c3)
 ,constraint uk_multiple2_uk2 unique (c2, c3, c4, c5)
);

CREATE TABLE not_null (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) NOT NULL
 ,c3 nchar(10)
 ,c4 nchar(10) NOT NULL
);

CREATE TABLE not_null (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) NOT NULL
 ,c3 nchar(10)
 ,c4 nchar(10) NOT NULL
);

CREATE TABLE fk (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
  constraint fk_fk foreign key references uk(c2) 
);

CREATE TABLE fk2 (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) UNIQUE
 ,c3 nchar(10)
 ,c4 nchar(10)
  constraint fk2_fk foreign key(c2) references pk(c1) 
);

CREATE TABLE fk_multiple (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) UNIQUE
 ,c3 nchar(10) UNIQUE
 ,c4 nchar(10)
  constraint fk_multiple_uk unique (c2, c3)
 ,constraint fk_multiple_fk foreign key(c2, c3) references uk_multiple(c2, c3) 
);

CREATE TABLE comment (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
);

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'スキーマコメント' , 
@level0type=N'SCHEMA', @level0name=N'dbo';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'コメント' , 
@level0type=N'SCHEMA', @level0name=N'dbo', 
@level1type=N'TABLE',  @level1name=N'comment';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'カラム1 PKです' , 
@level0type=N'SCHEMA', @level0name=N'dbo', 
@level1type=N'TABLE',  @level1name=N'comment',
@level2type=N'COLUMN', @level2name=N'c1';
