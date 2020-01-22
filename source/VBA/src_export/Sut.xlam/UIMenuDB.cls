VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' メニュー操作時の処理 - DB
'
' 作成者　：Ison
' 履歴　　：2009/01/24　新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' メンバ変数
' ________________________________________________________
' ■DB接続フォーム（イベント受信用に定義）
Private WithEvents frmDBConnectVar As frmDBConnect
Attribute frmDBConnectVar.VB_VarHelpID = -1

' ________________________________________________________
' イベント
' ________________________________________________________
' =========================================================
' ▽DB接続フォームでOKボタンが押下された場合に発生するイベント
'
' 概要　　　：
'
' =========================================================
Private Sub frmDBConnectVar_ok(ByVal connStr As String, ByVal connSimpleStr As String, ByVal connectInfo As ValDBConnectInfo)

    On Error GoTo err

    ' DB接続されている場合
    If Not Main.dbConn Is Nothing Then
    
        ' 切断する
        Main.SutUpdateDbConn Nothing, Empty, Empty
    End If
    
    ' DB接続し変数に設定する
    Dim conn As Object
    Set conn = ADOUtil.connectDb(connStr)
    
    ' DB接続文字列を設定する
    Main.SutUpdateDbConn conn, connStr, connSimpleStr
    
    Exit Sub
    
err:

    ' エラー発生
    Main.SutUpdateDbConn Nothing, Empty, Empty

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽DB接続フォームで処理がキャンセルされた場合に発生するイベント
'
' 概要　　　：
'
' =========================================================
Private Sub frmDBConnectVar_cancel()

    On Error GoTo err

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' イベント受信用のDB接続フォームを破棄する
    Set frmDBConnectVar = Nothing
End Sub


' =========================================================
' ▽初期化
'
' 概要　　　：初期化処理を実行する。
' 引数　　　：
'
' =========================================================
Public Sub init()

    ' DB接続フォームをロードする
    Load frmDBConnect

    ' イベント受信用のDB接続フォームを初期化する
    Set frmDBConnectVar = frmDBConnect

End Sub

' =========================================================
' ▽DB接続
'
' 概要　　　：DB接続 押下時の処理
'
' =========================================================
Public Sub connectDb()

    ' DB接続フォームを表示する
    frmDBConnect.ShowExt vbModal
End Sub

' =========================================================
' ▽DB切断
'
' 概要　　　：DB切断 押下時の処理
'
' =========================================================
Public Sub disconnectDb()

    ' 切断する
    Main.SutUpdateDbConn Nothing, Empty, Empty
    
End Sub
