VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPopupMenu 
   Caption         =   "ポップアップメニューの設定"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6390
   OleObjectBlob   =   "frmPopupMenu.frx":0000
End
Attribute VB_Name = "frmPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ポップアップメニューの設定
'
' 作成者　：Ison
' 履歴　　：2009/06/07　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽決定した際に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：appSettingShortcut アプリケーション設定ショートカット
' 　　　　　　selectedItemList 選択済み項目リスト
' 　　　　　　menuName 新しいメニュー名
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
Public Event Cancel()

' ポップアップメニューの新規作成時のデフォルト文字列
Private Const POPUP_MENU_NEW_CREATED_STR As String = "Popup Menu"

' ポップアップメニューの新規作成最大数
Private Const POPUP_MENU_NEW_CREATED_OVER_SIZE As String = "ポップアップは最大${count}まで登録可能です。"

' メニュー設定情報
Private WithEvents frmMenuSettingVar As frmMenuSetting
Attribute frmMenuSettingVar.VB_VarHelpID = -1

' ショートカットキー設定情報
Private WithEvents frmShortcutKeySettingVar As frmShortcutKeySetting
Attribute frmShortcutKeySettingVar.VB_VarHelpID = -1

' アプリケーション設定情報（ショートカットキー）
Private applicationSetting As ValApplicationSettingShortcut

' ポップアップメニューリスト コントロール
Private popupMenuList As CntListBox

' ポップアップメニューリストでの選択項目インデックス
Private popupMenuListSelectedIndex As Long
' ポップアップメニューリストでの選択項目オブジェクト
Private popupMenuListSelectedItem As ValPopupmenu

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
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef var As ValApplicationSettingShortcut)

    ' メンバ変数にアプリケーション設定情報を設定する
    Set applicationSetting = var
    
    activate
    
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

    restoreShortcut
    
    If VBUtil.unloadFormIfChangeActiveBook(frmMenuSetting) Then Unload frmMenuSetting
    Load frmMenuSetting
    If VBUtil.unloadFormIfChangeActiveBook(frmShortcutKeySetting) Then Unload frmShortcutKeySetting
    Load frmShortcutKeySetting
    
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

    Set popupMenuList = Nothing
    
    ' Nothingを設定することでイベントを受信しないようにする
    Set frmMenuSettingVar = Nothing
    Set frmShortcutKeySettingVar = Nothing
    
    Main.storeFormPosition Me.name, Me

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
' ▽ポップアップメニューリストボックスダブルクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstPopupMenu_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    editPopup
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
    RaiseEvent Cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽新規ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdAdd_Click()

    ' リストボックスのサイズ
    Dim cnt As Long
    ' リストボックスのサイズを取得する
    cnt = popupMenuList.collection.count
    
    ' ポップアップの数が最大登録数を超えているかチェックする
    If cnt >= ConstantsCommon.POPUP_MENU_NEW_CREATED_MAX_SIZE Then
    
        ' メッセージを表示する
        Dim mess As String
        mess = replace(POPUP_MENU_NEW_CREATED_OVER_SIZE, "${count}", ConstantsCommon.POPUP_MENU_NEW_CREATED_MAX_SIZE)
        
        VBUtil.showMessageBoxForInformation mess _
                                          , ConstantsCommon.APPLICATION_NAME
        Exit Sub
    End If
    
    ' ポップアップメニューオブジェクトをリストに追加する
    Dim popupMenu As ValPopupmenu
    Set popupMenu = New ValPopupmenu: popupMenu.init ConstantsCommon.COMMANDBAR_MENU_NAME
    
    '
    popupMenu.popupMenuName = POPUP_MENU_NEW_CREATED_STR & " " & (cnt + 1)
    
    popupMenuList.addItem popupMenu.popupMenuName, popupMenu
    
    lstPopupMenu.ListIndex = cnt
    lstPopupMenu.SetFocus
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

    editPopup
End Sub

Private Sub editPopup()

    ' 現在選択されているインデックスを取得
    popupMenuListSelectedIndex = lstPopupMenu.ListIndex

    ' 未選択の場合
    If popupMenuListSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' 現在選択されている項目を取得
    Set popupMenuListSelectedItem = popupMenuList.getItem(popupMenuListSelectedIndex)
    
    Set frmMenuSettingVar = frmMenuSetting
    frmMenuSettingVar.ShowExt Me _
                            , vbModal _
                            , applicationSetting _
                            , popupMenuListSelectedItem.itemList _
                            , "" _
                            , "ポップアップメニューの設定をします。" _
                            , popupMenuListSelectedItem.popupMenuName
    Set frmMenuSettingVar = Nothing

End Sub

' =========================================================
' ▽メニュー設定フォームのOKボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmMenuSettingVar_ok(appSettingShortcut As ValApplicationSettingShortcut _
                               , selectedItemList As ValCollection _
                               , ByVal menuName As String)

    popupMenuListSelectedItem.itemList = selectedItemList
    popupMenuListSelectedItem.popupMenuName = menuName
    
    lstPopupMenu.list(popupMenuListSelectedIndex, 0) = menuName
    lstPopupMenu.SetFocus

End Sub

' =========================================================
' ▽メニュー設定フォームのキャンセルボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmMenuSettingVar_cancel()

    lstPopupMenu.SetFocus
End Sub

' =========================================================
' ▽メニュー設定フォームのリセットボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmMenuSettingVar_reset(appSettingShortcut As ValApplicationSettingShortcut _
                                  , ByRef Cancel As Boolean)

End Sub

' =========================================================
' ▽削除ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDelete_Click()

    Dim selectedIndex As Long
    
    ' 現在選択されているインデックスを取得
    selectedIndex = lstPopupMenu.ListIndex

    ' 未選択の場合
    If selectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    popupMenuList.removeItem selectedIndex
    
    lstPopupMenu.SetFocus

End Sub

' =========================================================
' ▽ショートカットキー設定ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdShortcut_Click()

    ' ショートカットキーの設定
    editShortcutKey
End Sub

Private Sub editShortcutKey()

    ' 現在選択されているインデックスを取得
    popupMenuListSelectedIndex = lstPopupMenu.ListIndex

    ' 未選択の場合
    If popupMenuListSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' ショートカット情報の取得
    Set popupMenuListSelectedItem = popupMenuList.getItem(popupMenuListSelectedIndex)

    Set frmShortcutKeySettingVar = frmShortcutKeySetting
    ' ショートカットキー設定用のフォームを開く
    frmShortcutKeySettingVar.ShowExt vbModal, popupMenuListSelectedItem.shortcutKeyCode
    Set frmShortcutKeySettingVar = Nothing
    
End Sub

' =========================================================
' ▽ショートカットキーの設定ダイアログでOKボタンが押下された場合のイベント
' =========================================================
Private Sub frmShortcutKeySettingVar_ok(ByVal KeyCode As String, ByVal keyLabel As String)

    popupMenuListSelectedItem.shortcutKeyCode = KeyCode
    popupMenuListSelectedItem.shortcutKeyLabel = keyLabel
    
    lstPopupMenu.list(popupMenuListSelectedIndex, 1) = keyLabel
    
    lstPopupMenu.SetFocus
End Sub

' =========================================================
' ▽ショートカットキーの設定ダイアログでキャンセルボタンが押下された場合のイベント
' =========================================================
Private Sub frmShortcutKeySettingVar_cancel()

    lstPopupMenu.SetFocus
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

    applicationSetting.clearPopupMenu
    
    Set applicationSetting.popupMenuList = popupMenuList.collection
    applicationSetting.writeForDataPopupMenu

    applicationSetting.updatePopupMenu

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

    Set popupMenuList = New CntListBox: popupMenuList.init lstPopupMenu
    
    popupMenuList.addAll applicationSetting.ClonePopupMenuList _
                       , "popupMenuName" _
                       , "shortcutKeyLabel"
    
End Sub

