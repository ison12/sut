VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSnapSqlDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スナップショットの定義クラス
'
' 作成者　：Ison
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public name As String
Public sqlList As ValCollection
Public paramList As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    Set sqlList = New ValCollection
    Set paramList = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    Set sqlList = Nothing
    Set paramList = Nothing
End Sub

Public Sub addSql(ByVal val As ValSnapSqlDefine)
    sqlList.setItem val, val.name
End Sub

Public Sub addParam(ByVal val As ValSnapSqlDefineParam)
    paramList.setItem val, val.name
End Sub
