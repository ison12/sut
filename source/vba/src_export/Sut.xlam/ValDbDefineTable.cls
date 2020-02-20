VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbDefineTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�e�[�u����`�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public schemaName As String
Public tableName As String
Public tableComment As String
Public columnList As ValCollection
Public constraintsList As ValCollection
Public isLoad As Boolean

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    Set columnList = New ValCollection
    Set constraintsList = New ValCollection
    isLoad = False
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    Set columnList = Nothing
    Set constraintsList = Nothing
End Sub

Public Sub addColumn(ByVal column As ValDbDefineColumn)
    columnList.setItem column, column.columnName
End Sub

Public Sub addConstraint(ByVal tableConstraints As ValDbDefineTableConstraints)
    constraintsList.setItem tableConstraints, tableConstraints.constraintName
End Sub

Public Property Get schemaTableName() As String
    schemaTableName = schemaName & "." & tableName
End Property