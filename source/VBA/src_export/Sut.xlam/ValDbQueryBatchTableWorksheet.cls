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
' DBクエリバッチクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public dbQueryBatchTypeObj  As ValDbQueryBatchType
Public tableWorksheet       As ValTableWorksheet

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    Set dbQueryBatchTypeObj = New ValDbQueryBatchType
    dbQueryBatchTypeObj.dbQueryBatchType = DB_QUERY_BATCH_TYPE.none
    dbQueryBatchTypeObj.dbQueryBatchTypeName = ""
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
End Sub

Public Property Get dbQueryBatchType() As DB_QUERY_BATCH_TYPE

    If Not dbQueryBatchTypeObj Is Nothing Then
        
        dbQueryBatchType = dbQueryBatchTypeObj.dbQueryBatchType
        
    End If

End Property

Public Property Let dbQueryBatchType(ByVal vNewValue As DB_QUERY_BATCH_TYPE)

    If Not dbQueryBatchTypeObj Is Nothing Then
        
        dbQueryBatchTypeObj.dbQueryBatchType = vNewValue
        
    End If
    
End Property


Public Property Get dbQueryBatchTypeName() As String

    If Not dbQueryBatchTypeObj Is Nothing Then
        
        dbQueryBatchTypeName = dbQueryBatchTypeObj.dbQueryBatchTypeName
        
    End If

End Property

Public Property Let dbQueryBatchTypeName(ByVal vNewValue As String)

    If Not dbQueryBatchTypeObj Is Nothing Then
        
        dbQueryBatchTypeObj.dbQueryBatchTypeName = vNewValue
        
    End If
    
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
