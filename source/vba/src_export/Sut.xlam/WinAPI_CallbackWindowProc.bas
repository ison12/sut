Attribute VB_Name = "WinAPI_CallbackWindowProc"
Option Explicit

' *********************************************************
' ウィンドウプロシージャの処理を実装するIWindowProcへと
' 処理を振り分けるコントローラの役割を果たすモジュール。
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

' IWindowProcオブジェクトを格納するリストオブジェクト
Private list               As ValCollection
' 設定前のウィンドウプロシージャを格納するリストオブジェクト
Private listPrevWindowProc As ValCollection


' =========================================================
' ▽IWindowProcオブジェクト登録
'
' 概要　　　：IWindowProcオブジェクトを登録する。
' 引数　　　：proc   IWindowProcを実装したオブジェクト変数
' 　　　　　　hWnd   ウィンドウハンドル
'
' 戻り値　　：
'
' =========================================================
Public Sub registWindowProc(ByRef proc As IWindowProc _
                          , ByVal hwnd As Long)


    ' リストが初期化されていない場合、初期化を実施する
    If list Is Nothing Then
    
        ' 初期化する
        Set list = New ValCollection
        Set listPrevWindowProc = New ValCollection
    End If
    
    Dim prevWindowProc As Long
    
    ' 設定前のウィンドウプロシージャ
    prevWindowProc = WinAPI_User.GetWindowLong(hwnd, WinAPI_User.GWL_WNDPROC)
    
    ' エラーチェック
    If prevWindowProc = 0 Then
    
        err.Raise 5000 _
                , 0 _
                , "APIエラー"
    
    End If
    
    
    ' ウィンドウハンドルをキーに、IWindowProcオブジェクトを設定する
    list.setItem proc, hwnd
    ' ウィンドウハンドルをキーに、設定前のウィンドウプロシージャを設定する
    list.setItem prevWindowProc, hwnd
    
    ' サブクラス化開始
    If WinAPI_User.SetWindowLong(hwnd _
                                , WinAPI_User.GWL_WNDPROC _
                                , AddressOf windowProcedure) = 0 Then
                            
        err.Raise 5000 _
                , 0 _
                , "APIエラー"
                                        
    End If

End Sub

' =========================================================
' ▽IWindowProcオブジェクト削除
'
' 概要　　　：IWindowProcオブジェクトを削除する
' 引数　　　：hWnd           ウィンドウハンドル
' 　　　　　　prevWindowProc 最初に設定されていたウィンドウプロシージャ
'
' 戻り値　　：処理に成功したかどうかを表すフラグ
'
' =========================================================
Public Sub unregistWindowProc(ByVal hwnd As Long)

    ' ウィンドウハンドルをキーに、IWindowProcオブジェクトを削除する
    list.remove hwnd
    
    ' 設定前のウィンドウプロシージャ
    Dim prevWindowProc As Long
    
    ' ウィンドウハンドルをキーに、設定前のウィンドウプロシージャを削除する
    prevWindowProc = listPrevWindowProc.getItem(hwnd, vbLong)
    
    ' 最初に設定されていたウィンドウプロシージャを再セットする
    WinAPI_User.SetWindowLong hwnd, WinAPI_User.GWL_WNDPROC, prevWindowProc

    listPrevWindowProc.remove hwnd

End Sub

' =========================================================
' ▽ウィンドウプロシージャ
'
' 概要　　　：メッセージをIWindowProcに振り分ける。
' 引数　　　：hWnd   ウィンドウハンドル
' 　　　　　　msg    メッセージ
' 　　　　　　wParam パラメータその1
' 　　　　　　lParam パラメータその2
'
' 戻り値　　：結果コード
'
' =========================================================
Private Function windowProcedure(ByVal hwnd As Long _
                               , ByVal msg As Long _
                               , ByVal wParam As Long _
                               , ByVal lParam As Long) As Long

    ' 結果コード
    Dim resultCode As Long
    
    ' IWindowProcオブジェクト
    Dim windowProc As IWindowProc
    
    ' 最初に設定されていたウィンドウプロシージャ
    Dim prevWindowProc As Long

    ' オブジェクトを取得する
    Set windowProc = list.getItem(hwnd)

    ' ウィンドウプロシージャオブジェクトのチェック
    If Not windowProc Is Nothing Then
    
        ' IWindowProcオブジェクトに処理を振り分ける
        If windowProc.process(hwnd, msg, wParam, lParam, resultCode) = False Then
        
            ' 最初に設定されていたウィンドウプロシージャを取得する
            prevWindowProc = listPrevWindowProc.getItem(hwnd, vbLong)
            ' デフォルトメッセージ処理
            windowProcedure = CallWindowProc(prevWindowProc, hwnd, msg, wParam, lParam)
            
        Else
        
            ' 処理結果コードを返す
            windowProcedure = resultCode
        End If

    End If
    
End Function

