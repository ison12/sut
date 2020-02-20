VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbQueryBatchTableWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�N�G���o�b�`�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public dbQueryBatchType     As DB_QUERY_BATCH_TYPE
Public tableWorksheet       As ValTableWorksheet

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    dbQueryBatchType = DB_QUERY_BATCH_TYPE.none
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
End Sub

Public Property Get dbQueryBatchTypeName() As String

    dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbQueryBatchType)

End Property

Public Property Get sheetNameOrSheetTableName() As String

    If Not tableWorksheet Is Nothing Then
        
        sheetNameOrSheetTableName = tableWorksheet.sheetNameOrSheetTableName
        
    End If

End Property

Public Property Get sheetName() As String

    If Not tableWorksheet Is Nothing Then
        sheetName = tableWorksheet.sheetName
    End If

End Property

Public Property Get schemaTableName() As String

    If Not tableWorksheet Is Nothing Then
        If Not tableWorksheet.table Is Nothing Then
            schemaTableName = tableWorksheet.table.schemaTableName
        End If
    End If

End Property

Public Property Get tableName() As String

    If Not tableWorksheet Is Nothing Then
        If Not tableWorksheet.table Is Nothing Then
            tableName = tableWorksheet.table.tableName
        End If
    End If

End Property

Public Property Get tableComment() As String

    If Not tableWorksheet Is Nothing Then
        If Not tableWorksheet.table Is Nothing Then
            tableComment = tableWorksheet.table.tableComment
        End If
    End If

End Property