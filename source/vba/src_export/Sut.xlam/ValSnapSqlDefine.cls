VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSnapSqlDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�i�b�v�V���b�g�̒�`�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public name As String
Public sqlList As ValCollection
Public paramList As ValCollection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    Set sqlList = New ValCollection
    Set paramList = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    Set sqlList = Nothing
    Set paramList = Nothing
End Sub

Public Sub addSql(ByVal val As ValSnapSqlDefineSql)
    sqlList.setItem val
End Sub

Public Sub addParam(ByVal val As ValSnapSqlDefineParam)
    paramList.setItem val, val.name
End Sub