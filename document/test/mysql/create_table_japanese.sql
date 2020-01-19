CREATE DATABASE IF NOT EXISTS `スキーマ```
    DEFAULT CHARACTER SET 'utf8';

-- 日本語カラムを持つテーブル
DROP TABLE `スキーマ```.`テーブル```;

CREATE TABLE `スキーマ```.`テーブル``` (
    `カラム１`    VARCHAR(30)
   ,`カラム２`    NUMERIC(10)
   ,`カラム３```  NUMERIC(10)
   ,PRIMARY KEY (`カラム１`)
) ENGINE=INNODB;


-- PKを複数個持つテーブル
DROP TABLE `スキーマ```.`日本語テーブル`;

CREATE TABLE `スキーマ```.`日本語テーブル` (
    `カラム１`    VARCHAR(30)
   ,`カラム２`    VARCHAR(30)
   ,`カラム３`    VARCHAR(30)
   ,`カラム４`    VARCHAR(30)
   ,`カラム数値５`    NUMERIC(10,3)
   ,PRIMARY KEY (`カラム１`)
   ,CONSTRAINT `ユニークキー` UNIQUE(`カラム２`)
) ENGINE=INNODB;
