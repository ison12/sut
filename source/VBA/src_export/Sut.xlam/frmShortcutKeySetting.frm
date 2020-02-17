VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShortcutKeySetting 
   Caption         =   "キー設定"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4215
   OleObjectBlob   =   "frmShortcutKeySetting.frx":0000
End
Attribute VB_Name = "frmShortcutKeySetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


' *********************************************************
' ショートカットキーの設定（子画面）
'
' 作成者　：Ison
' 履歴　　：2009/06/02　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽決定した際に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：applicationSetting アプリケーション設定情報
'
' =========================================================
Public Event ok(ByVal KeyCode As String, ByVal keyLabel As String)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' ショートカットキーリスト
Private shortcutKeyList As CntListBox

' キーコード（フォーム表示時点でのキーコード）
Private keyCodeBefore As String

' 対象ブック
Private targetBook As Workbook
' 対象ブックを取得する
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal モーダルまたはモードレス表示指定
' 　　　　　　var   アプリケーション設定情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal KeyCode As String)

    keyCodeBefore = KeyCode
    
    activate
    
    ' デフォルトフォーカスコントロールを設定する
    cboKey.SetFocus
    
    Main.restoreFormPosition Me.name, Me
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
    
    Main.storeFormPosition Me.name, Me
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

    ' キーコードを分解した後に格納する変数
    Dim shiftCtrl  As Boolean
    Dim shiftShift As Boolean
    Dim shiftAlt   As Boolean
    Dim keyName    As String
    
    ' キーコードを分解し、対応する変数に格納する
    VBUtil.resolveAppOnKey keyCodeBefore _
                                 , shiftCtrl _
                                 , shiftShift _
                                 , shiftAlt _
                                 , keyName

    chbCtrl.value = shiftCtrl
    chbShift.value = shiftShift
    chbAlt.value = shiftAlt
    cboKey.value = keyName
    
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
' ▽フォーム初期化時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
    ' ロード時点のアクティブブックを保持しておく
    Set targetBook = ExcelUtil.getActiveWorkbook
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
' ▽フォームアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Activate()

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
        cmdCancel_Click
    End If
    
End Sub

' =========================================================
' ▽削除ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDelete_Click()

    chbCtrl.value = False
    chbShift.value = False
    chbAlt.value = False
    cboKey.ListIndex = -1

End Sub

' =========================================================
' ▽OKボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err
    
    ' フォームを閉じる
    HideExt
    
    ' キーコードを分解し、対応する変数に格納する
    Dim KeyCode As String
    KeyCode = VBUtil.getAppOnKeyCodeBySomeParams( _
                                   chbCtrl.value _
                                 , chbShift.value _
                                 , chbAlt.value _
                                 , cboKey.value)

    Dim keyName As String
    keyName = VBUtil.getAppOnKeyNameBySomeParams( _
                                   chbCtrl.value _
                                 , chbShift.value _
                                 , chbAlt.value _
                                 , cboKey.value)

    ' OKイベントを送信する
    RaiseEvent ok(KeyCode, keyName)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽キャンセルボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdCancel_Click()

    On Error GoTo err
    
    ' フォームを閉じる
    HideExt

    ' キャンセルイベントを送信する
    RaiseEvent cancel
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
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

    Set shortcutKeyList = New CntListBox: shortcutKeyList.init cboKey
    shortcutKeyList.addAll VBUtil.getAppOnKeyCodeList

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

    Set shortcutKeyList = Nothing
    
End Sub
