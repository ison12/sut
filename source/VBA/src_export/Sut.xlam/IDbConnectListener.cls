VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbConnectListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB接続時の情報を取得するインターフェース。
'
' 作成者　：Ison
' 履歴　　：2020/01/16　新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' ▽接続時の処理
'
' 概要　　　：接続された時の処理
' 引数　　　：connectinfo 接続情報
'
' 戻り値　　：
'
' 特記事項　：
'
'
' =========================================================
Public Sub connect(connectInfo As ValDBConnectInfo)

End Sub

' =========================================================
' ▽接続キャンセル時の処理
'
' 概要　　　：接続キャンセル時の処理
' 引数　　　：
'
' 戻り値　　：
'
' 特記事項　：
'
'
' =========================================================
Public Sub cancel()

End Sub

