-- ���{��J���������e�[�u��
DROP TABLE [�e�[�u��];

CREATE TABLE [�e�[�u��] (
    [�J�����P]    VARCHAR(30)
   ,[�J�����Q]    NUMERIC(10)
   ,[�J�����R]  NUMERIC(10)
   ,PRIMARY KEY ([�J�����P])
);


-- PK�𕡐����e�[�u��
DROP TABLE [���{��e�[�u��];

CREATE TABLE [���{��e�[�u��] (
    [�J�����P]    VARCHAR(30)
   ,[�J�����Q]    VARCHAR(30)
   ,[�J�����R]    VARCHAR(30)
   ,[�J�����S]    VARCHAR(30)
   ,[�J�������l�T]    NUMERIC(10,3)
   ,CONSTRAINT [�v���C�}���L�[] PRIMARY KEY ([�J�����P])
   ,CONSTRAINT [���j�[�N�L�[] UNIQUE([�J�����Q])
);
