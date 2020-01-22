VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryResultDetailInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBクエリ結果詳細情報クラス
'
' 作成者　：Ison
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public cell As String
Public message As String
Public query As String
Public sqlState As String

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
End Sub
