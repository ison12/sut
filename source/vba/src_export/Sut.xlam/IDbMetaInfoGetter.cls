VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbMetaInfoGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�̃��^�����擾����C���^�[�t�F�[�X�B
' �Ή�����DB���ɖ{�C���^�[�t�F�[�X��Implements����B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/02/11�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' =========================================================
' ��DB�R�l�N�V�����ݒ�
'
' �T�v�@�@�@�FDB�R�l�N�V������ݒ肷��B
' �����@�@�@�FvNewValue DB�R�l�N�V����
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let dbConn(ByVal vNewValue As Variant)

End Property

' =========================================================
' ���X�L�[�}��񃊃X�g�擾
'
' �T�v�@�@�@�F�ڑ����Ă���DB�̑S�X�L�[�}�����擾����
' �����@�@�@�F
'
' �߂�l�@�@�F�X�L�[�}���X�g�B
'
' ���L�����@�F
'
'
' =========================================================
Public Function getSchemaList() As ValCollection

End Function

' =========================================================
' ���e�[�u�����X�g�擾
'
' �T�v�@�@�@�F�C�ӂ̃X�L�[�}�Ɋ܂܂��e�[�u������S�Ď擾����B
' �����@�@�@�Fconn       DB�R�l�N�V����
' �@�@�@�@�@�@schemaList �X�L�[�}���X�g
'
' �߂�l�@�@�F�e�[�u����`���X�g�B
'
' ���L�����@�F
'
'
' =========================================================
Public Function getTableList(ByRef schemaList As ValCollection) As ValCollection

End Function

' =========================================================
' ���e�[�u�����擾�i�{�J�������j
'
' �T�v�@�@�@�F�X�L�[�}�E�e�[�u�������L�[�ɂ��ăe�[�u�������擾����B
' �����@�@�@�FschemaName �X�L�[�}��
' �@�@�@�@�@�@tableName  �e�[�u����
'
' �߂�l�@�@�F�e�[�u����`���B
'
' ���L�����@�F
'
'
' =========================================================
Public Function getTable(ByRef schemaName As String, ByRef tableName As String) As ValDbDefineTable

End Function