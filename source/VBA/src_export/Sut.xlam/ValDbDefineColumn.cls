VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbDefineColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBカラム定義クラス
'
' 作成者　：Ison
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public schemaName As String
Public tableName As String
Public columnName As String
Public columnComment As String
Public dataType As String
Public dataTypeFormal As String
Public charLength As String
Public dataPrecision As String
Public dataScale As String
Public dateTimePrecision As String
Public default As String
Public isNull As Boolean

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
