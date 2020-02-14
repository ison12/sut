VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSnapRecordsSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スナップショットのレコードセット情報
'
' 作成者　：Ison
' 履歴　　：2019/01/06　新規作成
'
' 特記事項：
' *********************************************************

' 名前
Public name       As String
' 日時
Public getDate     As Date
' レコード数
Public recordCount As Long

' リスト
Public list As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set list = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    
    Set list = Nothing
End Sub
