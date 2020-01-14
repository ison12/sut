VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDBConnectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB接続情報クラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public name As String
Public type_ As String
Public dsn As String
Public host As String
Public port As String
Public db As String
Public user As String
Public password As String
Public option_ As String

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
