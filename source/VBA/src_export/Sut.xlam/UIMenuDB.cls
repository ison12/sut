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
    If VBUtil.unloadFormIfChangeActiveBook(frmDBConnect) Then Unload frmDBConnect
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
    frmDBConnectVar.ShowExt vbModal
End Sub

' =========================================================
' ▽DB切断
'
' 概要　　　：DB切断 押下時の処理
'
' =========================================================
Public Sub disconnectDB()

    ' 切断する
    Main.SutUpdateDbConn Nothing, Empty, Empty
    
End Sub

' =========================================================
' ▽DB接続情報表示
'
' 概要　　　：
'
' =========================================================
Public Sub showDBConnectInfo(ByVal dbConn As Object)

    Const MSG_TITLE As String = "DB接続情報"

    If dbConn Is Nothing Then
        ' 切断状態
        VBUtil.showMessageBoxForInformation "DBに接続されていません", MSG_TITLE
    
    ElseIf ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        ' 切断状態
        VBUtil.showMessageBoxForInformation dbConn.ConnectionString & vbNewLine & "DBとの接続が切れています", MSG_TITLE
    Else
        ' 接続状態
        
        ' ------------------------------------------------------
        ' 接続文字列を分解してタブ区切り・改行ごとに分割する
        Dim connStrArr     As Variant
        Dim connStrInfo    As Variant
        Dim connStrInfoArr As Variant
        
        Dim connStrBuf As New StringBuilder
        
        Dim i As Long
        
        connStrArr = Split(dbConn.ConnectionString, ";")
        
        For i = LBound(connStrArr) To UBound(connStrArr)
        
            connStrInfo = Trim$(connStrArr(i))
            If connStrInfo <> Empty Then
            
                connStrInfoArr = Split(connStrInfo, "=")
                
                If VBUtil.arraySize(connStrInfoArr) >= 2 Then
                    connStrBuf.append "  " & Trim$(connStrInfoArr(0)) & "=" & Trim$(connStrInfoArr(1))
                Else
                    connStrBuf.append "  " & Trim$(connStrInfoArr(0)) & "="
                End If
            
                connStrBuf.append vbLf
            End If
        
        Next
        
        If connStrBuf.length > 0 Then
            connStrBuf.remove connStrBuf.length, 1
        End If
        ' ------------------------------------------------------
        
        VBUtil.showMessageBoxForInformation _
            "接続文字列" & vbTab & ": " & vbLf & _
            connStrBuf.str & vbLf & _
            "接続タイムアウト" & vbTab & ": " & dbConn.ConnectionTimeout & vbLf & _
            "コマンドタイムアウト" & vbTab & ": " & dbConn.CommandTimeout & vbLf & _
            "分離レベル" & vbTab & ": " & dbConn.IsolationLevel & vbLf & _
            "デフォルトDB" & vbTab & ": " & dbConn.defaultdatabase & vbLf & _
            "ADOバージョン" & vbTab & ": " & dbConn.version & _
            "", MSG_TITLE
    End If

End Sub

