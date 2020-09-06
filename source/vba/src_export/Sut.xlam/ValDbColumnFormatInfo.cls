VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbColumnFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのカラム書式情報を表現するクラス
'
' 作成者　：Ison
' 履歴　　：2008/08/01　新規作成
'
' 特記事項：
' *********************************************************

' DB名
Public dbName    As String
' カラムリスト
Public columnList  As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    ' カラムリストを初期化する
    Set columnList = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set columnList = Nothing
    
End Sub
