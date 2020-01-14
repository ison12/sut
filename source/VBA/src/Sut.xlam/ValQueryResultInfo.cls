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
Public errorCount As Long
Public queryType As BATCH_QUERY_TYPE
Public tableWorksheet As ValTableWorksheet
Public detailList As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    Set detailList = New ValCollection
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
