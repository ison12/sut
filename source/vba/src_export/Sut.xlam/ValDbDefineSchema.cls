VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbDefineSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�X�L�[�}��`�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public schemaName As String
Public schemaComment As String
Public tableList As ValCollection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    Set tableList = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    Set tableList = Nothing
End Sub

Public Sub addTable(ByVal table As ValDbDefineTable)
    tableList.setItem table, table.tableName
End Sub