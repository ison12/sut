VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbMetaInfoQueryCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�̃��^�����擾����N�G���𐶐�����C���^�[�t�F�[�X�B
' �Ή�����DB���ɖ{�C���^�[�t�F�[�X��Implements����B
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/15�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' =========================================================
' ���X�L�[�}���X�g�擾�p�N�G���̐���
'
' �T�v�@�@�@�F�X�L�[�}���X�g�擾�p�N�G���̐����𐶐����ĕԋp����B
' �����@�@�@�F
'
' �߂�l�@�@�F�X�L�[�}���X�g�擾�p�N�G���̐���������
'
' ���L�����@�F
'
'
' =========================================================
Public Function createQuerySelectSchemaList() As String

End Function


' =========================================================
' ���C�ӂ̃X�L�[�}�z���̃e�[�u�����X�g�擾�p�N�G���̐���
'
' �T�v�@�@�@�F�C�ӂ̃X�L�[�}�z���̃e�[�u�����X�g�擾�p�N�G���̐����𐶐����ĕԋp����B
' �����@�@�@�FschemaName �X�L�[�}��
'
' �߂�l�@�@�F�e�[�u�����X�g�擾�p�N�G���̐���������
'
' ���L�����@�F
'
'
' =========================================================
Public Function createQuerySelectTableList(ByVal schemaList As ValCollection) As String

End Function


' =========================================================
' ���C�ӂ̃X�L�[�}�z���̃e�[�u�����񃊃X�g�擾�p�N�G���̐���
'
' �T�v�@�@�@�F�C�ӂ̃X�L�[�}�z���̃e�[�u�����񃊃X�g�擾�p�N�G���̐����𐶐����ĕԋp����B
' �����@�@�@�FschemaName �X�L�[�}��
'
' �߂�l�@�@�F�e�[�u�����񃊃X�g�擾�p�N�G���̐���������
'
' ���L�����@�F
'
'
' =========================================================
Public Function createQuerySelectTableConstList(ByVal schemaList As ValCollection) As String

End Function


' =========================================================
' ���C�ӂ̃e�[�u�����擾�p�N�G���̐���
'
' �T�v�@�@�@�F�C�ӂ̃e�[�u�����擾�p�N�G���̐����𐶐����ĕԋp����B
' �����@�@�@�FschemaName �X�L�[�}��
' �@�@�@�@�@�F tableName �e�[�u����
'
' �߂�l�@�@�F�e�[�u�����擾�p�N�G���̐���������
'
' ���L�����@�F
'
'
' =========================================================
Public Function createQuerySelectTable(ByVal schemaName As String, ByVal tableName As String) As String

End Function


' =========================================================
' ���C�ӂ̃e�[�u���̃J�������擾�p�N�G���̐���
'
' �T�v�@�@�@�F�C�ӂ̃e�[�u���̃J�������擾�p�N�G���̐����𐶐����ĕԋp����B
' �����@�@�@�FschemaName �X�L�[�}��
' �@�@�@�@�@�F tableName �e�[�u����
'
' �߂�l�@�@�F�J�������擾�p�N�G���̐���������
'
' ���L�����@�F
'
'
' =========================================================
Public Function createQuerySelectColumn(ByVal schemaName As String, ByVal tableName As String) As String

End Function


' =========================================================
' ���C�ӂ̃e�[�u���̃e�[�u��������擾�p�N�G���̐���
'
' �T�v�@�@�@�F�C�ӂ̃e�[�u���̃e�[�u��������擾�p�N�G���̐����𐶐����ĕԋp����B
' �����@�@�@�FschemaName �X�L�[�}��
' �@�@�@�@�@�F tableName �e�[�u����
'
' �߂�l�@�@�F�e�[�u��������擾�p�N�G���̐���������
'
' ���L�����@�F
'
'
' =========================================================
Public Function createQuerySelectTableConst(ByVal schemaName As String, ByVal tableName As String) As String

End Function

