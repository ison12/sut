VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSnapSqlDefineSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スナップショットのSQL一覧定義クラス
'
' 作成者　：Ison
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public sql As String
Public memo As String
Public pkList As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    Set pkList = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    Set pkList = Nothing
End Sub

Public Sub addPk(ByVal val As ValSnapSqlDefineSqlPk)
    pkList.setItem val, val.pk
End Sub
