VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWindowProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' Callback関数 WindowProc のインターフェース
' Windowをサブクラス化し独自にメッセージ処理する場合、本インターフェースをImplementsする。
'
' 作成者　：Ison
' 履歴　　：2008/10/11　新規作成
'
' 特記事項：
' 　関連モジュールを以下に示す。
' 　①．IWindowProc.cls
' 　②．WinAPI_CallbackWindowProc.bas
' 　③．WinAPI_User.bas
'
' *********************************************************

' =========================================================
' ▽メインメソッド
'
' 概要　　　：ウィンドウメッセージを処理するメソッド。
' 引数　　　：hWnd   ウィンドウハンドル
' 　　　　　　msg    メッセージ
' 　　　　　　wParam パラメータその1
' 　　　　　　lParam パラメータその2
' 　　　　　　result 処理結果コード
' 戻り値　　：処理したかどうかを表すフラグ
'
' =========================================================
Public Function process(ByVal hWnd As Long _
                         , ByVal msg As Long _
                         , ByVal wParam As Long _
                         , ByVal lParam As Long _
                         , ByRef result As Long) As Boolean

End Function

