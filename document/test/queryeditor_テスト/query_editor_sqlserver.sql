-- テーブル作成

create table "テーブル" (
	"カラム１"  varchar(30),
	"カラム２"  varchar(30),
	"カラム３"  varchar(30),
	primary key ("カラム１")
)
;

select * from "テーブル";

-- INSERT
INSERT INTO "テーブル" ("カラム１", "カラム２", "カラム３") VALUES ('A1', 'B', 'C');
INSERT INTO "テーブル" ("カラム１", "カラム２", "カラム３") VALUES ('BB1', 'BA', 'CA');
INSERT INTO "テーブル" ("カラム１", "カラム２", "カラム３") VALUES ('BB2', 'BB', 'CB');


-- UPDATE

UPDATE "テーブル"
SET
    "カラム２" = 'カラム２だよ'
   ,"カラム３" = 'カラム３だよ'
WHERE
    "カラム１" = 'BB1';


UPDATE "テーブル"
SET
    "カラム２" = 'カラム２じゃん？'
   ,"カラム３" = 'カラム３でんがな？'
WHERE
    "カラム１" = 'A1';


UPDATE "テーブル"
SET
    "カラム２" = 'カラム２だな？'
   ,"カラム３" = 'カラム３まんがな？'
WHERE
    "カラム１" = 'BB2';

-- DELETE

DELETE FROM "テーブル"
WHERE
    "カラム１" = 'BB1';


DELETE FROM "テーブル"
WHERE
    "カラム１" = 'A1';


DELETE FROM "テーブル"
WHERE
    "カラム１" = 'BB2';



-- SELECT
SELECT * FROM "テーブル";

SELECT
    "カラム１"
   ,"カラム２"
   ,"カラム３"
FROM
    (SELECT
        "カラム１"
       ,"カラム２"
       ,"カラム３"
        ,ROW_NUMBER() OVER (ORDER BY
         "カラム１" ASC) rn
    FROM
        "テーブル２aa"
    ORDER BY
         "カラム１" ASC) "テーブル２aa"
WHERE
    1 <= rn AND rn  <= 100
ORDER BY
     "カラム１" ASC