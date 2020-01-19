VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValRecordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' レコードリストを表現するクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/02/01　新規作成
'
' 特記事項：
' *********************************************************

' 配列
Public values           As Variant
' Range
Public valuesCellOffset As Range

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    values = Empty
    Set valuesCellOffset = Nothing
    
End Sub
