VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValTableWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �e�[�u�����[�N�V�[�g�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public sheetName As String
Public sheet As Worksheet
Public table As ValDbDefineTable
Public omitsSchema As Boolean
Public recFormat As REC_FORMAT

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
End Sub

Public Property Get sheetNameOrSheetTableName() As String

    If sheetName = schemaTableName Or _
       sheetName = tableName Then
        sheetNameOrSheetTableName = sheetName
        
    Else
        sheetNameOrSheetTableName = sheetName & " ( " & schemaTableName & " ) "
    
    End If

End Property

Public Property Get schemaName() As String

    If Not table Is Nothing Then
        schemaName = table.schemaName
    End If

End Property

Public Property Get schemaTableName() As String

    If Not table Is Nothing Then
        schemaTableName = table.schemaTableName
    End If

End Property

Public Property Get tableName() As String

    If Not table Is Nothing Then
        tableName = table.tableName
    End If

End Property

Public Property Get tableComment() As String

    If Not table Is Nothing Then
        tableComment = table.tableComment
    End If

End Property