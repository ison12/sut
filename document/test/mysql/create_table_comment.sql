drop table comments1;
drop table comments2;

create table comments1 (
  col1  CHAR      primary key comment 'col1',
  col2  CHAR                  comment '���{��̃R�����g',
  col3  CHAR                  comment '�����R�����g�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O',
  col4  CHAR                  comment '�@�A�B�C�D�E�F�G�H�I',
  col5  CHAR                  comment '��\n�s'
) comment = '�R�����g';

create table comments2 (
  col1  CHAR      primary key comment 'col1'
) comment = '�����R�����g�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O�O' ;

