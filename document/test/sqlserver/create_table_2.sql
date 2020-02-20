DROP   DATABASE test2;
CREATE DATABASE test2;

CREATE SCHEMA test2Schema;

CREATE TABLE test2Schema.test2_1 (
  one   nchar(10) PRIMARY KEY
 ,two   nchar(10)
 ,three datetime
);

CREATE TABLE test2Schema.test2_2 (
  any1   nchar(10) PRIMARY KEY
 ,any2   nchar(10)
 ,any3   datetime
);

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'�e�X�g2�̃X�L�[�}' , 
@level0type=N'SCHEMA', @level0name=N'test2Schema';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'�e�X�g2�̃e�[�u��2_1' , 
@level0type=N'SCHEMA', @level0name=N'test2Schema', 
@level1type=N'TABLE',  @level1name=N'test2_1';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'�J����1' , 
@level0type=N'SCHEMA', @level0name=N'test2Schema', 
@level1type=N'TABLE',  @level1name=N'test2_1',
@level2type=N'COLUMN', @level2name=N'one';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'�J����2' , 
@level0type=N'SCHEMA', @level0name=N'test2Schema', 
@level1type=N'TABLE',  @level1name=N'test2_1',
@level2type=N'COLUMN', @level2name=N'two';

EXEC sys.sp_addextendedproperty 
@name=N'comment_Description', 
@value=N'�J����3' , 
@level0type=N'SCHEMA', @level0name=N'test2Schema', 
@level1type=N'TABLE',  @level1name=N'test2_1',
@level2type=N'COLUMN', @level2name=N'three';
