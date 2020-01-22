VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7560
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スプラッシュフォーム
'
' 作成者　：Ison
' 履歴　　：2020/01/21　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event closed()

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal               モーダルまたはモードレス表示指定
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants)
    
    activate
    
    Me.Show modal

End Sub

' =========================================================
' ▽フォーム非表示
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub HideExt()

    deactivate

    Me.Hide
    
End Sub

' =========================================================
' ▽フォームアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()
    
    ' バージョン情報を設定
    lblVersionInfo.Caption = Main.getVersionInfo
    
End Sub

' =========================================================
' ▽フォームディアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub deactivate()
    
End Sub

' =========================================================
' ▽閉じるボタンのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdClose_Click()

    On Error GoTo err
    
    ' フォームを閉じる
    HideExt
    
    ' 閉じるイベントを送信する
    RaiseEvent closed
        
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽フォーム初期化時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
    ' 初期化処理を実行する
    initial
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォーム破棄時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Terminate()

    On Error GoTo err
    
    ' 破棄処理を実行する
    unInitial
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォームキー押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii.value = vbKeyEscape Then
        cmdClose_Click
    End If
    
End Sub

' =========================================================
' ▽スプラッシュ画像クリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub imgSplash_Click()
        cmdClose_Click
End Sub

' =========================================================
' ▽フォームの閉じる時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' 本処理では処理自体をキャンセルする
        cancel = True
        ' 以下のイベント経由で閉じる
        cmdClose_Click
    End If
    
End Sub

' =========================================================
' ▽初期化処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub initial()

    ExcelUtil.setNonTitleBarWindowStyle Me
    
End Sub

' =========================================================
' ▽後始末処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub unInitial()
    
End Sub

