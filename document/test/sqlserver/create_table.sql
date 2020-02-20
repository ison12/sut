DROP   DATABASE test;
CREATE DATABASE test;

CREATE TABLE data_all_type (
   -- 数値
     cInt	bigint	primary key
   , cInt_2	bit	
   , cInt_3	decimal(18, 0)	
   , cInt_3_2	decimal(18, 3)	
   , cInt_4	int	
   , cInt_5	money	
   , cInt_6	numeric(18, 2)	
   , cInt_7	smallint	
   , cInt_8	smallmoney	
   , cInt_9	tinyint	
   , cDbl	float(23)	
   , cDbl_2	float(30)	
   , cDbl_3	real	
   -- 文字
   , cStr	char(10)	
   , cStr_2	varchar(50)	
   , cStr_3	varchar(MAX)	
   , cStr_4	text	
   , cStr_5	nchar(10)	
   , cStr_6	nvarchar(50)	
   , cStr_7	nvarchar(MAX)	
   , cStr_8	ntext	
   -- 日付
   , cDate	date	
   , cDate_2	datetime	
   , cDate_3_1	datetime2	
   , cDate_3_2	datetime2(3)	
   , cDate_4	smalldatetime	
   , cDate_5_1	datetimeoffset	
   , cDate_5_2	datetimeoffset(3)	
   , cDate_6_1	time	
   , cDate_6_2	time(3)	
);

CREATE TABLE data_all_type_other (
     cPri	bigint	primary key
   -- バイナリ
   , cBin	binary(50)	
   , cBin_2	varbinary(50)	
   , cBin_3	varbinary(MAX)	
   , cBin_4	image	
   , cOth	timestamp	
   , cOth_2	sql_variant	
   , cOth_3	hierarchyid	
   , cOth_4	uniqueidentifier	
   , cOth_5	xml	
   , cOth_6_1	geometry
   , cOth_6_2	geography
);

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

-- ------------------------------------------------------------------
-- 別スキーマ
CREATE SCHEMA test_schema;


CREATE TABLE test_schema.sample (
  one   nchar(10) PRIMARY KEY
 ,two   nchar(10)
 ,three datetime
);

CREATE TABLE test_schema.sample_any (
  any1   nchar(10) PRIMARY KEY
 ,any2   nchar(10)
 ,any3   datetime
);

-- ------------------------------------------------------------------
-- 日本語
CREATE SCHEMA [テストスキーマ];
CREATE SCHEMA [テストスキーマ]][];

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


CREATE TABLE [テストスキーマ]][].[テーブル]][] (
  [テストカラム１] nchar(10)
 ,[テストカラム２]][] nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
 ,c5 nchar(10)
 ,constraint [テストPK]][] primary key ([テストカラム１])
 ,constraint [テストUK]][] unique ([テストカラム２]][])
);
-- ------------------------------------------------------------------

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'スキーマコメント______________________________________________________-' , 
@level0type=N'SCHEMA', @level0name=N'dbo';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'コメント______________________________________________________-' , 
@level0type=N'SCHEMA', @level0name=N'dbo', 
@level1type=N'TABLE',  @level1name=N'comment';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'カラム1 PKです______________________________________________________-' , 
@level0type=N'SCHEMA', @level0name=N'dbo', 
@level1type=N'TABLE',  @level1name=N'comment',
@level2type=N'COLUMN', @level2name=N'c1';
