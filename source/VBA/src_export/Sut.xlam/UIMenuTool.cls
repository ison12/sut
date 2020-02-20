VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' メニュー操作時の処理 - オプション
'
' 作成者　：Ison
' 履歴　　：2009/03/15　新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' メンバ変数
' ________________________________________________________
' ■オプションフォーム（イベント受信用に定義）
Private WithEvents frmOptionVar As frmOption
Attribute frmOptionVar.VB_VarHelpID = -1
' ■右クリックメニュー設定フォーム（イベント受信用に定義）
Private WithEvents frmMenuSettingVar As frmMenuSetting
Attribute frmMenuSettingVar.VB_VarHelpID = -1
' ■ショートカットキー設定フォーム（イベント受信用に定義）
Private WithEvents frmShortcutKeyVar As frmShortcutKey
Attribute frmShortcutKeyVar.VB_VarHelpID = -1
' ■ポップアップメニュー設定フォーム（イベント受信用に定義）
Private WithEvents frmPopupMenuVar As frmPopupMenu
Attribute frmPopupMenuVar.VB_VarHelpID = -1

' ________________________________________________________
' イベント
' ________________________________________________________
' =========================================================
' ▽オプションフォームでOKボタンが押下された場合に発生するイベント
'
' 概要　　　：
' 引数　　　：applicationSetting アプリケーション設定情報
'
' =========================================================
Private Sub frmOptionVar_ok(ByRef applicationSetting As ValApplicationSetting)

End Sub

' =========================================================
' ▽オプションフォームでキャンセルボタンが押下された場合に発生するイベント
'
' 概要　　　：
'
' =========================================================
Private Sub frmOptionVar_cancel()

End Sub

' =========================================================
' ▽右クリックメニュー設定フォームでOKボタンが押下された場合に発生するイベント
'
' 概要　　　：
' 引数　　　：applicationSetting アプリケーション設定情報
'
' =========================================================
Private Sub frmMenuSettingVar_ok(ByRef appSettingShortcut As ValApplicationSettingShortcut _
                               , ByRef selectedItemList As ValCollection _
                               , ByVal menuName As String)

    ' 現在の情報で右クリックメニューを一旦削除する
    appSettingShortcut.clearRClickMenu

    ' アプリケーション設定情報に右クリックメニューに登録する項目を設定しレジストリに登録する
    Set appSettingShortcut.rclickMenuItemList = selectedItemList
    appSettingShortcut.writeForDataRClick

    ' 登録された情報で右クリックメニューを更新する
    appSettingShortcut.updateRClickMenu

End Sub

' =========================================================
' ▽右クリックメニュー設定フォームでキャンセルボタンが押下された場合に発生するイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Private Sub frmMenuSettingVar_cancel()

End Sub

' =========================================================
' ▽右クリックメニュー設定フォームでリセットボタンが押下された場合に発生するイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Private Sub frmMenuSettingVar_reset(appSettingShortcut As ValApplicationSettingShortcut _
                                  , ByRef Cancel As Boolean)

    Dim result As Long
    result = VBUtil.showMessageBoxForYesNo("右クリックメニューが初期状態にリセットされます。" & vbNewLine & _
                    "初期状態とは、Excelがインストールされた直後の事を指します。" & vbNewLine & _
                    "したがって、他のマクロによって追加された項目があれば、それらもリセットの対象になりますがよろしいですか？", ConstantsCommon.APPLICATION_NAME)
    
    If result = WinAPI_User.IDNO Then
    
        Cancel = True
        Exit Sub
    End If
    
    ' 右クリックメニューをデフォルト状態にする
    appSettingShortcut.resetRClickMenu

    ' 空のコレクションを設定しレジストリに登録する（＝情報を全て削除する）
    Set appSettingShortcut.rclickMenuItemList = New ValCollection
    appSettingShortcut.writeForDataRClick

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

End Sub

' =========================================================
' ▽オプション設定
'
' 概要　　　：オプション 押下時の処理
'
' =========================================================
Public Sub settingOption()

    On Error GoTo err

    ' フォームをロードする
    If VBUtil.unloadFormIfChangeActiveBook(frmOption) Then Unload frmOption
    Load frmOption
    ' イベント受信用のフォームを初期化する
    Set frmOptionVar = frmOption
    
    ' オプションフォームを表示する
    frmOptionVar.ShowExt vbModal, Main.getApplicationSetting, Main.getApplicationSettingColFormat

    ' イベント受信用のオプションフォームを破棄する
    Set frmOptionVar = Nothing
    
    Exit Sub

err:

    Main.ShowErrorMessage

    ' イベント受信用のオプションフォームを破棄する
    Set frmOptionVar = Nothing
    
End Sub

' =========================================================
' ▽右クリックメニューの設定
'
' 概要　　　：
'
' =========================================================
Public Sub settingRClickMenu()

    On Error GoTo err

    Dim appSettingShortcut As ValApplicationSettingShortcut
    Set appSettingShortcut = Main.getApplicationSettingShortcut

    ' フォームをロードする
    If VBUtil.unloadFormIfChangeActiveBook(frmMenuSetting) Then Unload frmMenuSetting
    Load frmMenuSetting
    ' イベント受信用のフォームを初期化する
    Set frmMenuSettingVar = frmMenuSetting

    ' メニュー設定フォームを表示する
    frmMenuSettingVar.ShowExt Nothing _
                            , vbModal _
                            , appSettingShortcut _
                            , appSettingShortcut.cloneRClickMenuList _
                            , "右クリックメニューの設定" _
                            , "シート上で右クリックした際に表示されるポップアップをカスタマイズします。" _
                            , "右クリックメニュー" _
                            , True

    ' イベント受信用のメニュー設定フォームを破棄する
    Set frmMenuSettingVar = Nothing

    Exit Sub

err:

    Main.ShowErrorMessage

    ' イベント受信用のメニュー設定フォームを破棄する
    Set frmMenuSettingVar = Nothing

End Sub

' =========================================================
' ▽ショートカットキーの設定
'
' 概要　　　：
'
' =========================================================
Public Sub settingShortCutKey()

    On Error GoTo err

    ' フォームをロードする
    If VBUtil.unloadFormIfChangeActiveBook(frmShortcutKey) Then Unload frmShortcutKey
    Load frmShortcutKey
    ' イベント受信用のフォームを初期化する
    Set frmShortcutKeyVar = frmShortcutKey
    ' ショートカットキーフォームを表示する
    frmShortcutKeyVar.ShowExt vbModal, Main.getApplicationSettingShortcut

    ' イベント受信用のショートカットキーフォームを破棄する
    Set frmShortcutKeyVar = Nothing
    
    Exit Sub

err:

    Main.ShowErrorMessage

    ' イベント受信用のショートカットキーフォームを破棄する
    Set frmShortcutKeyVar = Nothing
    
End Sub

' =========================================================
' ▽ポップアップメニューの設定
'
' 概要　　　：
'
' =========================================================
Public Sub settingPopupMenu()

    On Error GoTo err

    ' フォームをロードする
    If VBUtil.unloadFormIfChangeActiveBook(frmPopupMenu) Then Unload frmPopupMenu
    Load frmPopupMenu
    ' イベント受信用のフォームを初期化する
    Set frmPopupMenuVar = frmPopupMenu
    ' ショートカットキーフォームを表示する
    frmPopupMenuVar.ShowExt vbModal, Main.getApplicationSettingShortcut

    ' イベント受信用のポップアップメニューフォームを破棄する
    Set frmPopupMenuVar = Nothing

    Exit Sub

err:

    Main.ShowErrorMessage

    ' イベント受信用のポップアップメニューフォームを破棄する
    Set frmPopupMenuVar = Nothing
    
End Sub

