VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValRecordInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���R�[�h��\������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/02/01�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public columnInfoList  As collection  ' ValColumnInfo�N���X��v�f�Ɏ���Collection�I�u�W�F�N�g

Public tableName       As String      ' �e�[�u����

Public targetSheetName As String      ' �ΏۃV�[�g��
Public targetRow       As Long        ' �Ώۍs

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set columnInfoList = New collection
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set columnInfoList = Nothing
    
End Sub

' =========================================================
' ���J�������ǉ�
'
' �T�v�@�@�@�F�v���p�e�B�ɃJ��������ǉ�����
' �����@�@�@�FvalColumnInfo �J�������
'             key           �L�[
'
' =========================================================
Public Sub addColumnInfo(ByRef valColumnInfo, Optional ByVal key As Variant = "")

    'VBUtil.setItem columnInfoList, valColumnInfo, key
End Sub

' =========================================================
' ���J�������擾
'
' �T�v�@�@�@�F�J���������擾����
' �����@�@�@�Fkey �L�[
'
' =========================================================
Public Function getColumnInfo(ByVal key As Variant)

    'VBUtil.getItem columnInfoList, key
End Function