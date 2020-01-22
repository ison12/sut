VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryResultSetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリ結果セット情報クラス
'
' 作成者　：Ison
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public queryResultInfoList As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    Set queryResultInfoList = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    Set queryResultInfoList = Nothing
End Sub

' =========================================================
' ▽クエリ結果情報の追加
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub addInfo(info As ValQueryResultInfo)
    
    queryResultInfoList.setItem info

End Sub

' =========================================================
' ▽クエリ結果情報のクリア
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub clearInfos()
    
    Set queryResultInfoList = New ValCollection

End Sub

' =========================================================
' ▽クエリ結果情報の取得
'
' 概要　　　：
' 引数　　　：index インデックス
' 戻り値　　：
'
' =========================================================
Public Function getInfo(index As Long) As ValQueryResultInfo

    Set getInfo = queryResultInfoList.getItemByIndex(index)

End Function

' =========================================================
' ▽クエリ結果情報の件数取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：クエリ結果情報の件数
'
' =========================================================
Public Function getInfoCount() As Long

    getInfoCount = queryResultInfoList.count

End Function
