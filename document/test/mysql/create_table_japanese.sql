CREATE DATABASE IF NOT EXISTS `�X�L�[�}```
    DEFAULT CHARACTER SET 'utf8';

-- ���{��J���������e�[�u��
DROP TABLE `�X�L�[�}```.`�e�[�u��```;

CREATE TABLE `�X�L�[�}```.`�e�[�u��``` (
    `�J�����P`    VARCHAR(30)
   ,`�J�����Q`    NUMERIC(10)
   ,`�J�����R```  NUMERIC(10)
   ,PRIMARY KEY (`�J�����P`)
) ENGINE=INNODB;


-- PK�𕡐����e�[�u��
DROP TABLE `�X�L�[�}```.`���{��e�[�u��`;

CREATE TABLE `�X�L�[�}```.`���{��e�[�u��` (
    `�J�����P`    VARCHAR(30)
   ,`�J�����Q`    VARCHAR(30)
   ,`�J�����R`    VARCHAR(30)
   ,`�J�����S`    VARCHAR(30)
   ,`�J�������l�T`    NUMERIC(10,3)
   ,PRIMARY KEY (`�J�����P`)
   ,CONSTRAINT `���j�[�N�L�[` UNIQUE(`�J�����Q`)
) ENGINE=INNODB;
