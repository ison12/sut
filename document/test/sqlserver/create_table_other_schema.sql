-- コメントしているのは SQL Server2008 以降に対応しているデータ型？

CREATE SCHEMA test_schema;

CREATE TABLE test_schema.data_all_type (
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

   , cStr	char(10)	
   , cStr_2	varchar(50)	
   , cStr_3	varchar(MAX)	
   , cStr_4	text	
   , cStr_5	nchar(10)	
   , cStr_6	nvarchar(50)	
   , cStr_7	nvarchar(MAX)	
   , cStr_8	ntext	

--   , cDate	date	
   , cDate_2	datetime	
--   , cDate_3	datetime2	
   , cDate_4	smalldatetime	
--   , cDate_5	datetimeoffset	
--   , cDate_6	time	

);

CREATE TABLE test_schema.data_all_type_other (
     cPri	bigint	primary key
   , cBin	binary(50)	
   , cBin_2	varbinary(50)	
   , cBin_3	varbinary(MAX)	
   , cBin_4	image	
   , cOth	timestamp	
   , cOth_2	sql_variant	
--   , cOth_3     hierarchyid 
   , cOth_4	uniqueidentifier	
   , cOth_5	xml	
--   , cOth_6	table	
);

CREATE TABLE test_schema.pk (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
);

CREATE TABLE test_schema.pk_multiple (
  c1 nchar(10)
 ,c2 nchar(10)
 ,c3 nchar(20)
  constraint pk_multiple_pk primary key (c1, c2)
);

CREATE TABLE test_schema.default_ (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) DEFAULT 'DEFAULT'
);

CREATE TABLE test_schema.uk (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) UNIQUE
);

CREATE TABLE test_schema.uk_multiple (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
 ,c3 nchar(10)
  constraint uk_multiple_uk unique (c2, c3)
);

CREATE TABLE test_schema.uk_multiple2 (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
 ,c5 nchar(10)
  constraint uk_multiple2_uk1 unique (c2, c3)
 ,constraint uk_multiple2_uk2 unique (c2, c3, c4, c5)
);

CREATE TABLE test_schema.not_null (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) NOT NULL
 ,c3 nchar(10)
 ,c4 nchar(10) NOT NULL
);

CREATE TABLE test_schema.fk (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
  constraint fk_fk foreign key references test_schema.uk(c2) 
);

CREATE TABLE test_schema.fk2 (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) UNIQUE
 ,c3 nchar(10)
 ,c4 nchar(10)
  constraint fk2_fk foreign key(c2) references test_schema.pk(c1) 
);

CREATE TABLE test_schema.fk_multiple (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10) UNIQUE
 ,c3 nchar(10) UNIQUE
 ,c4 nchar(10)
  constraint fk_multiple_uk unique (c2, c3)
 ,constraint fk_multiple_fk foreign key(c2, c3) references test_schema.uk_multiple(c2, c3) 
);

CREATE TABLE test_schema.comment (
  c1 nchar(10) PRIMARY KEY
 ,c2 nchar(10)
 ,c3 nchar(10)
 ,c4 nchar(10)
);

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'スキーマコメント' , 
@level0type=N'SCHEMA', @level0name=N'test_schema';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'コメント' , 
@level0type=N'SCHEMA', @level0name=N'test_schema', 
@level1type=N'TABLE',  @level1name=N'comment';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'カラム1 PKです' , 
@level0type=N'SCHEMA', @level0name=N'test_schema', 
@level1type=N'TABLE',  @level1name=N'comment',
@level2type=N'COLUMN', @level2name=N'c1';
