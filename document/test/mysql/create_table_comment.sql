drop table comments1;
drop table comments2;

create table comments1 (
  col1  CHAR      primary key comment 'col1',
  col2  CHAR                  comment '“ú–{Œê‚ÌƒRƒƒ“ƒg',
  col3  CHAR                  comment '’·‚¢ƒRƒƒ“ƒg‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O',
  col4  CHAR                  comment '‡@‡A‡B‡C‡D‡E‡F‡G‡H‡I',
  col5  CHAR                  comment '‰ü\ns'
) comment = 'ƒRƒƒ“ƒg';

create table comments2 (
  col1  CHAR      primary key comment 'col1'
) comment = '’·‚¢ƒRƒƒ“ƒg‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O‚O' ;

