VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbDefineTableConstraints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�e�[�u�������`�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public constraintName As String
Public constraintType As TABLE_CONSTANTS_TYPE
Public schemaName As String
Public tableName As String
Public columnList As ValCollection
Public fkColumnList As ValCollection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    Set columnList = New ValCollection
    Set fkColumnList = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    Set columnList = Nothing
    Set fkColumnList = Nothing
End Sub

Public Sub addColumn(ByVal column As ValDbDefineColumn)
    columnList.setItem column, column.columnName
End Sub

Public Sub addFkColumn(ByVal column As ValDbDefineColumn)
    fkColumnList.setItem column, column.columnName
End Sub
