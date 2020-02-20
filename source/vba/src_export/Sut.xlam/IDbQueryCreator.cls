VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbQueryCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�̃N�G���𐶐�����C���^�[�t�F�[�X�B
' �Ή�����DB���ɖ{�C���^�[�t�F�[�X��Implements����B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/03/28�@�V�K�쐬
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
' ���󔒕����ǂݎ������ݒ�
'
' �T�v�@�@�@�F�󔒕����ǂݎ�������ݒ肷��B
' �����@�@�@�FvNewValue �󔒕����ǂݎ�����
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let emptyValueReading(ByVal vNewValue As Long)

End Property

' =========================================================
' �����ړ��͕����ݒ�
'
' �T�v�@�@�@�F���ړ��͕�����ݒ肷��B
' �����@�@�@�FvNewValue ���ړ��͕���
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let directInputChar(ByVal vNewValue As String)

End Property

' =========================================================
' ���J�����������̐ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FRHS �J�����������I�u�W�F�N�g
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let columnFormatList(ByVal RHS As Variant)

End Property

' =========================================================
' ���X�L�[�}�g�p�̐ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvNewValue �X�L�[�}�g�p�L��
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let schemaUse(ByVal vNewValue As Variant)

End Property

' =========================================================
' ���e�[�u���E�J�������G�X�P�[�v�̐ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvNewValue �e�[�u���E�J�������G�X�P�[�v
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let tableColumnEscape(ByVal vNewValue As Variant)

End Property

' =========================================================
' ��INSERT������
'
' �T�v�@�@�@�FINSERT SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@record ���R�[�h
'
' �߂�l�@�@�FINSERT��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createInsert(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

End Function

' =========================================================
' ��UPDATE������
'
' �T�v�@�@�@�FUPDATE SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@record ���R�[�h
'
' �߂�l�@�@�FUPDATE��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createUpdate(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

End Function

' =========================================================
' ��DELETE������
'
' �T�v�@�@�@�FDELETE SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@record ���R�[�h
'
' �߂�l�@�@�FDELETE��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createDelete(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

End Function

' =========================================================
' ��DELETE�������i�e�[�u����̑S���R�[�h���폜�j
'
' �T�v�@�@�@�FDELETE SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@record ���R�[�h
'
' �߂�l�@�@�FDELETE��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createDeleteOfAllTable(ByRef table As ValDbDefineTable) As String

End Function

' =========================================================
' ��SELECT������
'
' �T�v�@�@�@�FINSERT SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@condition ������
'
' �߂�l�@�@�FSELECT��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createSelect(ByRef table As ValDbDefineTable, ByRef condition As ValSelectCondition) As String

End Function