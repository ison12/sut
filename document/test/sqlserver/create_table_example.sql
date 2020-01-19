
sqlcmd -S (local)\SQLEXPRESS -d sampleDb

CREATE DATABASE sampleDb;

CREATE SCHEMA sampleSchema;

CREATE TABLE sampleSchema.sample (
  one   nchar(10) PRIMARY KEY
 ,two   nchar(10)
 ,three datetime
);

CREATE TABLE sampleSchema.sample_any (
  any1   nchar(10) PRIMARY KEY
 ,any2   nchar(10)
 ,any3   datetime
);

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'サンプルスキーマ' , 
@level0type=N'SCHEMA', @level0name=N'sampleSchema';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'サンプルテーブル' , 
@level0type=N'SCHEMA', @level0name=N'sampleSchema', 
@level1type=N'TABLE',  @level1name=N'sample';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'カラム1' , 
@level0type=N'SCHEMA', @level0name=N'sampleSchema', 
@level1type=N'TABLE',  @level1name=N'sample',
@level2type=N'COLUMN', @level2name=N'one';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'カラム2' , 
@level0type=N'SCHEMA', @level0name=N'sampleSchema', 
@level1type=N'TABLE',  @level1name=N'sample',
@level2type=N'COLUMN', @level2name=N'two';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'カラム3' , 
@level0type=N'SCHEMA', @level0name=N'sampleSchema', 
@level1type=N'TABLE',  @level1name=N'sample',
@level2type=N'COLUMN', @level2name=N'three';
