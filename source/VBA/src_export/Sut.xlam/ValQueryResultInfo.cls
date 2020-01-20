VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryResultInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリ結果情報クラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public processCount As Long
Public successCount As Long
Public errorCount   As Long

Public dbQueryBatchType As DB_QUERY_BATCH_TYPE

Public tableWorksheet As ValTableWorksheet
Public detailList     As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    
    Set detailList = New ValCollection
    
    dbQueryBatchType = DB_QUERY_BATCH_TYPE.none
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    Set detailList = Nothing
End Sub

Public Function addDetail(ByVal val As ValQueryResultDetailInfo)
    detailList.setItem val
End Function

Public Function clearDetails()
    detailList.removeAll
End Function

Public Function getDetail(ByVal index As Long) As ValQueryResultDetailInfo
    Set getDetail = detailList.getItemByIndex(index)
End Function

Public Function getDetailsCount() As Long
    getDetailsCount = detailList.count
End Function

Public Property Get processErrorCount() As String

    processErrorCount = processCount & " / " & errorCount
    
End Property

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

