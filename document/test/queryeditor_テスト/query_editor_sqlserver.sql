-- �e�[�u���쐬

create table "�e�[�u��" (
	"�J�����P"  varchar(30),
	"�J�����Q"  varchar(30),
	"�J�����R"  varchar(30),
	primary key ("�J�����P")
)
;

select * from "�e�[�u��";

-- INSERT
INSERT INTO "�e�[�u��" ("�J�����P", "�J�����Q", "�J�����R") VALUES ('A1', 'B', 'C');
INSERT INTO "�e�[�u��" ("�J�����P", "�J�����Q", "�J�����R") VALUES ('BB1', 'BA', 'CA');
INSERT INTO "�e�[�u��" ("�J�����P", "�J�����Q", "�J�����R") VALUES ('BB2', 'BB', 'CB');


-- UPDATE

UPDATE "�e�[�u��"
SET
    "�J�����Q" = '�J�����Q����'
   ,"�J�����R" = '�J�����R����'
WHERE
    "�J�����P" = 'BB1';


UPDATE "�e�[�u��"
SET
    "�J�����Q" = '�J�����Q�����H'
   ,"�J�����R" = '�J�����R�ł񂪂ȁH'
WHERE
    "�J�����P" = 'A1';


UPDATE "�e�[�u��"
SET
    "�J�����Q" = '�J�����Q���ȁH'
   ,"�J�����R" = '�J�����R�܂񂪂ȁH'
WHERE
    "�J�����P" = 'BB2';

-- DELETE

DELETE FROM "�e�[�u��"
WHERE
    "�J�����P" = 'BB1';


DELETE FROM "�e�[�u��"
WHERE
    "�J�����P" = 'A1';


DELETE FROM "�e�[�u��"
WHERE
    "�J�����P" = 'BB2';



-- SELECT
SELECT * FROM "�e�[�u��";

SELECT
    "�J�����P"
   ,"�J�����Q"
   ,"�J�����R"
FROM
    (SELECT
        "�J�����P"
       ,"�J�����Q"
       ,"�J�����R"
        ,ROW_NUMBER() OVER (ORDER BY
         "�J�����P" ASC) rn
    FROM
        "�e�[�u���Qaa"
    ORDER BY
         "�J�����P" ASC) "�e�[�u���Qaa"
WHERE
    1 <= rn AND rn  <= 100
ORDER BY
     "�J�����P" ASC