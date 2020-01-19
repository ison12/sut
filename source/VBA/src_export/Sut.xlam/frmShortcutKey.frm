VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShortcutKey 
   Caption         =   "ショートカットキーの設定"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6585
   OleObjectBlob   =   "frmShortcutKey.frx":0000
End
Attribute VB_Name = "frmShortcutKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ショートカットキーの設定
'
' 作成者　：Hideki Isobe
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
Public Event ok(ByRef applicationSetting As ValApplicationSettingShortcut)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' ショートカットキー設定情報
Private WithEvents frmShortcutKeySettingVar As frmShortcutKeySetting
Attribute frmShortcutKeySettingVar.VB_VarHelpID = -1

' アプリケーション設定情報（ショートカットキー）
Private applicationSetting As ValApplicationSettingShortcut

' 機能リスト コントロール
Private appMenuList As CntListBox

' 機能リストでの選択項目インデックス
Private appMenuListSelectedIndex As Long
' 機能リストでの選択項目オブジェクト
Private appMenuListSelectedItem As ValShortcutKey

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal モーダルまたはモードレス表示指定
' 　　　　　　var   アプリケーション設定情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef var As ValApplicationSettingShortcut)

    ' メンバ変数にアプリケーション設定情報を設定する
    Set applicationSetting = var
    
    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show vbModal
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

    Load frmShortcutKeySetting

    restoreShortcut
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

    ' フォームクローズ後にイベントを受信しないようにフォーム変数をクリアしておく
    Set frmShortcutKeySettingVar = Nothing
    
    Set appMenuList = Nothing
    
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
' ▽OKボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err
    
    ' 情報を記録する
    storeShortcut
    
    ' フォームを閉じる
    HideExt
    
    ' OKイベントを送信する
    RaiseEvent ok(applicationSetting)
    
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
' ▽リセットボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdReset_Click()

    On Error GoTo err
    
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽機能リストボックスダブルクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstAppList_DblClick(ByVal cancel As MSForms.ReturnBoolean)

    editAppShortcutKey
End Sub

' =========================================================
' ▽編集ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdEdit_Click()

    editAppShortcutKey
End Sub

' =========================================================
' ▽消去ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDelete_Click()

    ' 現在選択されているインデックスを取得
    appMenuListSelectedIndex = lstAppList.ListIndex

    ' 未選択の場合
    If appMenuListSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' ショートカット情報の取得
    Set appMenuListSelectedItem = appMenuList.getItem(appMenuListSelectedIndex)

    appMenuListSelectedItem.shortcutKeyCode = ""
    appMenuListSelectedItem.shortcutKeyLabel = ""
    
    lstAppList.list(appMenuListSelectedIndex, 1) = ""
    
    lstAppList.SetFocus

End Sub

Private Sub editAppShortcutKey()

    ' 現在選択されているインデックスを取得
    appMenuListSelectedIndex = lstAppList.ListIndex

    ' 未選択の場合
    If appMenuListSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' ショートカット情報の取得
    Set appMenuListSelectedItem = appMenuList.getItem(appMenuListSelectedIndex)

    Set frmShortcutKeySettingVar = frmShortcutKeySetting
    ' ショートカットキー設定用のフォームを開く
    frmShortcutKeySettingVar.ShowExt vbModal, appMenuListSelectedItem.shortcutKeyCode
    Set frmShortcutKeySettingVar = Nothing

End Sub

' =========================================================
' ▽ショートカットキーの設定ダイアログでOKボタンが押下された場合のイベント
' =========================================================
Private Sub frmShortcutKeySettingVar_ok(ByVal keyCode As String, ByVal keyLabel As String)

    appMenuListSelectedItem.shortcutKeyCode = keyCode
    appMenuListSelectedItem.shortcutKeyLabel = keyLabel
    
    lstAppList.list(appMenuListSelectedIndex, 1) = keyLabel
    
    lstAppList.SetFocus
End Sub

' =========================================================
' ▽ショートカットキーの設定ダイアログでキャンセルボタンが押下された場合のイベント
' =========================================================
Private Sub frmShortcutKeySettingVar_cancel()

    lstAppList.SetFocus
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

' =========================================================
' ▽オプション情報を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeShortcut()

    ' 既存のショートカットを削除する
    applicationSetting.clearShortcutKey
    
    ' ここで設定されたショートカット情報をアプリケーションオブジェクトに設定し、レジストリに登録する
    Set applicationSetting.shortcutAppList = appMenuList.collection
    applicationSetting.writeForRegistryForShortcut
    
    ' 新たに設定されたショートカットを登録する
    applicationSetting.updateShortcutKey
    
End Sub

' =========================================================
' ▽オプション情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreShortcut()

    ' 機能リストをリセットする
    lstAppList.clear
    
    ' 機能リストの初期化
    Set appMenuList = New CntListBox: appMenuList.init lstAppList
    
    ' ショートカットリストを取得する
    ' ※Cloneメソッドを使用して情報をコピーする。
    ' 　ここでは、ApplicationSetting#ShortcutAppListに格納されているValShortCut要素を直接変更せずに
    ' 　クローンを生成し編集を行う。
    Dim shortCutList As ValCollection
    Set shortCutList = applicationSetting.CloneShortcutAppList
    
    ' 機能リストに反映する
    appMenuList.addAll shortCutList, "commandName", "shortcutKeyLabel"

End Sub
