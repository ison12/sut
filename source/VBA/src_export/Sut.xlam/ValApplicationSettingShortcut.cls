VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValApplicationSettingShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' アプリケーションのオプション情報
' ショートカットキー関連
'
' 作成者　：Ison
' 履歴　　：2009/06/04　新規作成
'
' 特記事項：
' *********************************************************

' =========================================================
' ▽メンバ
' =========================================================

' 本アプリケーションのCommandBarControlには一意なIDをTagプロパティに格納している
' CommandBarにはFindControlメソッドがあり、Tagを条件にしてコントロールを検索することが可能。
' このFindControlを利用するために、Tagプロパティに一意なIDを格納している。

' ショートカットキーに割り当て可能な機能リスト
'   [ key   ] : CommandBarControl.Tag
'   [ value ] : ValShortcutKey
Public shortcutAppList   As ValCollection

' 右クリックメニューリスト
'   [ key   ] : CommandBarControl.Tag
'   [ value ] : CommandBarControl.Tag
Public rclickMenuItemList As ValCollection

' ポップアップメニューリスト
'   [ key   ] : CommandBarControl.Tag
'   [ value ] : ValPopupMenu
Public popupMenuList As ValCollection

' =========================================================
' ▽定数プロパティ
' 　Classモジュールでは、Public Constが定義できないのでプロパティで代用する
' =========================================================
' キー値 右クリックメニューショートカット
Public Property Get REG_SUB_KEY_RCLICK_SHORTCUT() As String

    REG_SUB_KEY_RCLICK_SHORTCUT = "application_setting_rclick_shortcut"
End Property

' キー値 ショートカット
Public Property Get REG_SUB_KEY_SHORTCUT() As String

    REG_SUB_KEY_SHORTCUT = "application_setting_shortcut"
End Property

' キー値 ポップアップメニュー
Public Property Get REG_SUB_KEY_POPUPMENU() As String

    REG_SUB_KEY_POPUPMENU = "application_setting_popupmenu"
End Property

' キー値 ポップアップメニュー（ショートカットキー）
Public Property Get REG_SUB_KEY_POPUPMENU_SHORTCUTKEY() As String

    REG_SUB_KEY_POPUPMENU_SHORTCUTKEY = "shortcutkey"
End Property

' キー値 ポップアップメニュー（項目）
Public Property Get REG_SUB_KEY_POPUPMENU_ITEM() As String

    REG_SUB_KEY_POPUPMENU_ITEM = "item"
End Property

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set shortcutAppList = New ValCollection
    Set rclickMenuItemList = New ValCollection
    Set popupMenuList = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' 右クリックメニューから追加された項目を消去する
    clearRClickMenu
    ' ショートカットキーから追加された項目を消去する
    clearShortcutKey
    ' 追加されたポップアップメニューを消去する
    clearPopupMenu

    Set shortcutAppList = Nothing
    Set rclickMenuItemList = Nothing
    Set popupMenuList = Nothing
    
End Sub

Public Sub init()
    
    ' ショートカット機能リストを生成する
    
    Dim shortcutInfo As ValShortcutKey
    
    ' Sutメニュー
    Dim sutMenu As CommandBar
    ' Sutメニューの要素
    Dim sutMenuItem As commandBarControl
    
    ' Sutメニューを取得する
    Set sutMenu = Application.CommandBars(ConstantsCommon.COMMANDBAR_MENU_NAME)
    
    ' 第1階層
    For Each sutMenuItem In sutMenu.Controls
    
        ' 第1階層のボタン
        If sutMenuItem.Type = msoControlButton And sutMenuItem.DescriptionText <> "" Then
        
            Set shortcutInfo = New ValShortcutKey
            shortcutInfo.commandBarControl = sutMenuItem

            shortcutAppList.setItem shortcutInfo, sutMenuItem.Tag
        
        ' 第2階層
        ElseIf sutMenuItem.Type = msoControlPopup Then
        
            Dim sutMenu2Popup As commandBarPopup
            Dim sutMenu2Item  As commandBarControl
            
            Set sutMenu2Popup = sutMenuItem
            
            For Each sutMenu2Item In sutMenu2Popup.Controls
            
                ' 第2階層のボタン
                If sutMenu2Item.Type = msoControlButton And sutMenu2Item.DescriptionText <> "" Then
                
                    Set shortcutInfo = New ValShortcutKey
                    shortcutInfo.commandBarControl = sutMenu2Item
        
                    shortcutAppList.setItem shortcutInfo, sutMenu2Item.Tag
                    
                ' 第3階層
                ElseIf sutMenu2Item.Type = msoControlPopup Then
                
                    Dim sutMenu3Popup As commandBarPopup
                    Dim sutMenu3Item  As commandBarControl
                    
                    Set sutMenu3Popup = sutMenu2Item
                    
                    For Each sutMenu3Item In sutMenu3Popup.Controls
                    
                        ' 第3階層のボタン
                        If sutMenu3Item.Type = msoControlButton And sutMenu3Item.DescriptionText <> "" Then
                        
                            Set shortcutInfo = New ValShortcutKey
                            shortcutInfo.commandBarControl = sutMenu3Item
                
                            shortcutAppList.setItem shortcutInfo, sutMenu3Item.Tag
                        
                        End If
                        
                    Next
                
                End If
            Next
        End If
        
    Next
    
    ' レジストリから情報を読み込む
    readForRegistryForRClick
    readForRegistryForShortcut
    readForRegistryForPopupMenu
    
    ' 右クリックメニューに項目を追加する
    updateRClickMenu
    ' ショートカットキーを登録する
    updateShortcutKey
    ' ポップアップメニューを更新する
    updatePopupMenu

End Sub

' =========================================================
' ▽レジストリから情報を読み込む（右クリックメニュー）
' =========================================================
Public Sub readForRegistryForRClick()

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_RCLICK_SHORTCUT) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    ' ----------------------------------------------
    ' 右クリックメニューのショートカットキーの情報を取得する（特定のキーの配下にある値を全て取得する）
    Dim valNameList As ValCollection
    Dim valList     As ValCollection
    
    registry.getValueList valNameList, valList
    
    ' 右クリックメニューにレジストリから取得したリストを設定する
    Set rclickMenuItemList = valNameList
    
    ' 不要なリストを削除する（念のため）
    Set valNameList = Nothing
    Set valList = Nothing
                
    Set registry = Nothing

    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽レジストリに情報を書き込む（右クリックメニュー）
' =========================================================
Public Sub writeForRegistryForRClick()

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_RCLICK_SHORTCUT) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' ----------------------------------------------
    ' 右クリックメニューのショートカットキーの情報をレジストリから一旦全て削除する
    Dim valNameList As ValCollection
    Dim valList     As ValCollection
    
    Dim valName As Variant
    
    registry.getValueList valNameList, valList
    
    For Each valName In valNameList.col
    
        ' レジストリから情報を削除する
        registry.deleteValue valName
        
    Next
                
    ' ----------------------------------------------
    ' 右クリックメニューのショートカットキー情報をレジストリに追加する
    For Each valName In rclickMenuItemList.col
    
        ' レジストリに情報を書き込む
        registry.setValue valName, "True"
        
    Next
    
    ' 不要なリストを削除する（念のため）
    Set valNameList = Nothing
    Set valList = Nothing
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽レジストリから情報を読み込む（ショートカットキー）
' =========================================================
Public Sub readForRegistryForShortcut()

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SHORTCUT) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    ' ----------------------------------------------
    ' 右クリックメニューのショートカットキーの情報を取得する（特定のキーの配下にある値を全て取得する）
    Dim valNameList As ValCollection
    Dim valList     As ValCollection
    
    registry.getValueList valNameList, valList
    
    Dim key As Variant
    Dim val As Variant
    
    Dim shortcutInfo As ValShortcutKey
    
    For Each key In valNameList.col
    
        val = valList.getItem(key, vbString)
        
        Set shortcutInfo = shortcutAppList.getItem(key)
        
        If Not shortcutInfo Is Nothing Then
        
            shortcutInfo.shortcutKeyCode = val
            shortcutInfo.shortcutKeyLabel = VBUtil.getAppOnKeyNameByMultipleCode(val)
        End If
    
    Next
    
    ' 不要なリストを削除する（念のため）
    Set valNameList = Nothing
    Set valList = Nothing
                
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽レジストリに情報を書き込む（ショートカットキー）
' =========================================================
Public Sub writeForRegistryForShortcut()

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SHORTCUT) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' ----------------------------------------------
    ' ショートカットキーの情報をレジストリから一旦全て削除する
    Dim valNameList As ValCollection
    Dim valList     As ValCollection
    
    Dim valName As Variant
    
    registry.getValueList valNameList, valList
    
    For Each valName In valNameList.col
    
        ' レジストリから情報を削除する
        registry.deleteValue valName
        
    Next
                
    ' ----------------------------------------------
    ' ショートカットキー情報をレジストリに追加する
    Dim shortcutInfo As ValShortcutKey
    
    For Each shortcutInfo In shortcutAppList.col
    
        ' レジストリに情報を書き込む
        registry.setValue shortcutInfo.commandBarControl.Tag _
                        , shortcutInfo.shortcutKeyCode
        
    Next
    
    ' 不要なリストを削除する（念のため）
    Set valNameList = Nothing
    Set valList = Nothing
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽レジストリから情報を読み込む（ポップアップメニュー）
' 　REG_SUB_KEY_POPUPMENU
' 　 |
' 　 + ポップアップメニューその1
' 　   |
' 　   + ショートカットキー REG_SUB_KEY_POPUPMENU_SHORTCUTKEY
' 　   + 項目               REG_SUB_KEY_POPUPMENU_ITEM
' 　     |
' 　     + 項目1
' 　     + 項目1
' 　     + ・・・・
' 　 + ポップアップメニューその2
' 　 + ポップアップメニューその3
' =========================================================
Public Sub readForRegistryForPopupMenu()

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_POPUPMENU) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    Dim key     As Variant
    Dim keyList As ValCollection
    
    Set keyList = registry.getKeyList
    
    For Each key In keyList.col
    
        Dim popupMenu As ValPopupmenu
        Set popupMenu = New ValPopupmenu: popupMenu.init ConstantsCommon.COMMANDBAR_MENU_NAME
        
        popupMenu.popupMenuName = key
        
        readForRegistryForPopupMenuSub popupMenu
        
        popupMenuList.setItem popupMenu, popupMenu.popupMenuName
    Next
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

Public Sub readForRegistryForPopupMenuSub(ByRef popupMenu As ValPopupmenu)

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_POPUPMENU & "\" & popupMenu.popupMenuName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    ' ショートカットキーを取得する
    Dim shortcutKey As String
                
    registry.getValue REG_SUB_KEY_POPUPMENU_SHORTCUTKEY, shortcutKey
                    
    popupMenu.shortcutKeyCode = shortcutKey
    popupMenu.shortcutKeyLabel = VBUtil.getAppOnKeyNameByMultipleCode(shortcutKey)
                
                
    ' ----------------------------------------------
    ' ポップアップメニューに含まれる機能IDを取得する
    
    ' レジストリ操作クラス
    Dim registrySub As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registrySub.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_POPUPMENU & "\" & popupMenu.popupMenuName & "\" & REG_SUB_KEY_POPUPMENU_ITEM) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    Dim valNameList As ValCollection
    Dim valList     As ValCollection
    
    registrySub.getValueList valNameList, valList
    
    Dim key As Variant
    Dim val As Variant
    
    Dim itemList As ValCollection
    Set itemList = popupMenu.itemList
    
    For Each key In valNameList.col
    
        itemList.setItem key, key
    Next
    
    ' 不要なリストを削除する（念のため）
    Set valNameList = Nothing
    Set valList = Nothing
                
    Set registry = Nothing
    Set registrySub = Nothing
                
    Exit Sub
    
err:
    
    Set valNameList = Nothing
    Set valList = Nothing
    
    Set registry = Nothing
    Set registrySub = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽レジストリに情報を書き込む（ポップアップメニュー）
' 　REG_SUB_KEY_POPUPMENU
' 　 |
' 　 + ポップアップメニューその1
' 　   |
' 　   + ショートカットキー REG_SUB_KEY_POPUPMENU_SHORTCUTKEY
' 　   + 項目               REG_SUB_KEY_POPUPMENU_ITEM
' 　     |
' 　     + 項目1
' 　     + 項目1
' 　     + ・・・・
' 　 + ポップアップメニューその2
' 　 + ポップアップメニューその3
' =========================================================
Public Sub writeForRegistryForPopupMenu()

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_POPUPMENU) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' -------------------------------------------------------
    ' 全てのポップアップメニューの情報をレジストリから削除する
    ' -------------------------------------------------------
    Dim key     As Variant
    Dim keyList As ValCollection
    
    Set keyList = registry.getKeyList
    
    For Each key In keyList.col
    
        registry.delete key & "\" & REG_SUB_KEY_POPUPMENU_ITEM
        registry.delete key
        
    Next
    
    ' -------------------------------------------------------
    ' 設定されているポップアップメニューの情報でレジストリを更新する
    ' -------------------------------------------------------
    Dim popupMenu As ValPopupmenu
    
    For Each popupMenu In popupMenuList.col
    
        writeForRegistryForPopupMenuSub popupMenu
    Next
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

Public Sub writeForRegistryForPopupMenuSub(ByRef popupMenu As ValPopupmenu)

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_POPUPMENU & "\" & popupMenu.popupMenuName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    ' ショートカットキーの設定
    Dim shortcutKey As String
                
    registry.setValue REG_SUB_KEY_POPUPMENU_SHORTCUTKEY, popupMenu.shortcutKeyCode

    ' ----------------------------------------------
    ' ポップアップメニューに含まれる機能IDを取得する
    
    ' レジストリ操作クラス
    Dim registrySub As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registrySub.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_POPUPMENU & "\" & popupMenu.popupMenuName & "\" & REG_SUB_KEY_POPUPMENU_ITEM) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    Dim popupMenuItemList As ValCollection
    Dim popupMenuId       As Variant
    
    Set popupMenuItemList = popupMenu.itemList
    
    For Each popupMenuId In popupMenuItemList.col
    
        registrySub.setValue popupMenuId, "True"
    
    Next
    
    Set registry = Nothing
    Set registrySub = Nothing
                
    Exit Sub
    
err:
    
    Set registry = Nothing
    Set registrySub = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽右クリックメニューのリセット
'
' 概要　　　：右クリックメニューを元の状態に戻す。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub resetRClickMenu()

    On Error Resume Next
    
    Dim targetBar1    As CommandBar
    Dim targetBar2    As CommandBar
    Dim targetBar3    As CommandBar
    
    Set targetBar1 = Application.CommandBars("Cell")
    Set targetBar2 = Application.CommandBars("Row")
    Set targetBar3 = Application.CommandBars("Column")
    
    targetBar1.reset
    targetBar2.reset
    targetBar3.reset
    
    On Error GoTo 0

End Sub

' =========================================================
' ▽右クリックメニューに追加した項目の削除
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub clearRClickMenu()

    On Error Resume Next
    
    Dim sutBar        As CommandBar
    Dim targetBar1    As CommandBar
    Dim targetBar2    As CommandBar
    Dim targetBar3    As CommandBar
    
    Set sutBar = Application.CommandBars(ConstantsCommon.COMMANDBAR_MENU_NAME)
    
    Set targetBar1 = Application.CommandBars("Cell")
    Set targetBar2 = Application.CommandBars("Row")
    Set targetBar3 = Application.CommandBars("Column")
    
    Dim shortcutInfo As ValShortcutKey
    
    Dim val   As commandBarControl
    Dim val2  As commandBarControl
    
    For Each shortcutInfo In shortcutAppList.col
    
        Set val = shortcutInfo.commandBarControl
        
        Set val2 = targetBar1.FindControl(Tag:=val.Tag, recursive:=True)
        If Not val2 Is Nothing Then
        
            val2.delete
        End If
        Set val2 = targetBar2.FindControl(Tag:=val.Tag, recursive:=True)
        If Not val2 Is Nothing Then
        
            val2.delete
        End If
        Set val2 = targetBar3.FindControl(Tag:=val.Tag, recursive:=True)
        If Not val2 Is Nothing Then
        
            val2.delete
        End If
    
    Next
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ▽右クリックメニューへの項目追加
'
' 概要　　　：現在のプロパティの状態を基に右クリックメニューに項目を追加する。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub updateRClickMenu()

    On Error Resume Next
    
    Dim sutBar        As CommandBar
    Dim targetBar1    As CommandBar
    Dim targetBar2    As CommandBar
    Dim targetBar3    As CommandBar
    
    Set sutBar = Application.CommandBars(ConstantsCommon.COMMANDBAR_MENU_NAME)
    
    Set targetBar1 = Application.CommandBars("Cell")
    Set targetBar2 = Application.CommandBars("Row")
    Set targetBar3 = Application.CommandBars("Column")
    
    Dim val   As Variant
    
    Dim srcControl  As CommandBarButton
    Dim desControl  As CommandBarButton
    
    For Each val In rclickMenuItemList.col
    
        Set srcControl = sutBar.FindControl(Tag:=val, recursive:=True)
        
        If srcControl Is Nothing Then
        
            GoTo continue
        
        End If
        
        Set desControl = targetBar1.FindControl(Tag:=val)
        If desControl Is Nothing Then
        
            Set desControl = targetBar1.Controls.Add(Type:=msoControlButton)
            ExcelUtil.copyCommandBarControl srcControl, desControl
            
        End If
        
        Set desControl = targetBar2.FindControl(Tag:=val)
        If desControl Is Nothing Then
        
            Set desControl = targetBar2.Controls.Add(Type:=msoControlButton)
            ExcelUtil.copyCommandBarControl srcControl, desControl
            
        End If
        
        Set desControl = targetBar3.FindControl(Tag:=val)
        If desControl Is Nothing Then
        
            Set desControl = targetBar3.Controls.Add(Type:=msoControlButton)
            ExcelUtil.copyCommandBarControl srcControl, desControl
            
        End If
        
continue:

    Next
    
    If err.Number <> 0 Then
    
        ' ショートカットキーの設定に失敗したエラー情報を設定し、メッセージボックスに表示する
        err.Number = ERR_NUMBER_RCLICKMENU_SETTING_FAILED
        err.Description = ERR_DESC_RCLICKMENU_SETTING_FAILED
        Main.ShowErrorMessage
    
        err.Number = 0
    End If
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ▽ショートカットキーの更新
'
' 概要　　　：現在のプロパティの状態を基にショートカットキーを更新する。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub updateShortcutKey()

    On Error Resume Next

    Dim targetBar1    As CommandBar
    Dim targetBar2    As CommandBar
    Dim targetBar3    As CommandBar
    Set targetBar1 = Application.CommandBars("Cell")
    Set targetBar2 = Application.CommandBars("Row")
    Set targetBar3 = Application.CommandBars("Column")
    
    Dim tmp As commandBarControl
    
    Dim shortcutInfo       As ValShortcutKey
    Dim shortcutCommandBar As commandBarControl
    
    For Each shortcutInfo In shortcutAppList.col
    
        Set shortcutCommandBar = shortcutInfo.commandBarControl
        
        If shortcutInfo.shortcutKeyCode <> "" Then
        
            Application.OnKey _
                shortcutInfo.shortcutKeyCode _
              , shortcutCommandBar.OnAction
              
            ' エラー発生時
            If err.Number <> 0 Then
            
                ' ショートカットキーの設定に失敗したエラー情報を設定し、メッセージボックスに表示する
                err.Number = ERR_NUMBER_SHORTCUT_SETTING_FAILED
                err.Description = ERR_DESC_SHORTCUT_SETTING_FAILED & vbNewLine & _
                                  "(" & shortcutInfo.commandName & ") << (" & shortcutInfo.shortcutKeyLabel & ")"
                Main.ShowErrorMessage
            
                ' ショートカットキー情報を削除する
                shortcutInfo.shortcutKeyCode = ""
                shortcutInfo.shortcutKeyLabel = ""
            
                ' エラーをクリアする
                err.Number = 0
                
            ' 正常時
            Else
            
                ' ---------------------------------------------------
                ' ショートカットキーテキストの設定
                ' Excel2007ではショートカットキーテキストが表示されないようなので、テキストの設定は行わない
                'setShortcutTextForCommandBar shortcutInfo, targetBar1, targetBar2, targetBar3, shortcutInfo.shortcutKeyLabel
    
            End If
              
        End If

continue:
    Next
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ▽ショートカットキーの消去
'
' 概要　　　：現在のプロパティの状態を基にショートカットキーをデフォルトの挙動に戻す
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub clearShortcutKey()

    Dim targetBar1    As CommandBar
    Dim targetBar2    As CommandBar
    Dim targetBar3    As CommandBar
    Set targetBar1 = Application.CommandBars("Cell")
    Set targetBar2 = Application.CommandBars("Row")
    Set targetBar3 = Application.CommandBars("Column")

    Dim shortcutInfo       As ValShortcutKey
    Dim shortcutCommandBar As commandBarControl

    For Each shortcutInfo In shortcutAppList.col
    
        If shortcutInfo.shortcutKeyCode <> "" Then
        
            ' Application#OnKeyのProcedureを指定しないことでデフォルトの挙動に戻す
            Application.OnKey _
                shortcutInfo.shortcutKeyCode
                
            ' ---------------------------------------------------
            ' ショートカットキーテキストの設定
            ' Excel2007ではショートカットキーテキストが表示されないようなので、テキストの設定は行わない
            'setShortcutTextForCommandBar shortcutInfo, targetBar1, targetBar2, targetBar3, ""
            
        End If
            
    Next
    
End Sub

' =========================================================
' ▽ショートカットキーの無効化
'
' 概要　　　：現在のプロパティの状態を基にショートカットキーを無効化する
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub disableShortcutKey()

    Dim targetBar1    As CommandBar
    Dim targetBar2    As CommandBar
    Dim targetBar3    As CommandBar
    Set targetBar1 = Application.CommandBars("Cell")
    Set targetBar2 = Application.CommandBars("Row")
    Set targetBar3 = Application.CommandBars("Column")

    Dim shortcutInfo       As ValShortcutKey
    Dim shortcutCommandBar As commandBarControl
    
    For Each shortcutInfo In shortcutAppList.col
    
        If shortcutInfo.shortcutKeyCode <> "" Then
        
            ' Application#OnKeyのProcedureに空文字列を指定することでキーを無効化する
            Application.OnKey _
                shortcutInfo.shortcutKeyCode, ""
                
            ' ---------------------------------------------------
            ' ショートカットキーテキストの設定
            ' Excel2007ではショートカットキーテキストが表示されないようなので、テキストの設定は行わない
            ' setShortcutTextForCommandBar shortcutInfo, targetBar1, targetBar2, targetBar3, ""
            
        End If
            
    Next
    
End Sub

' =========================================================
' ▽ポップアップメニューの更新
'
' 概要　　　：現在のプロパティの状態を基にポップアップメニューを更新する。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub updatePopupMenu()

    On Error Resume Next

    ' 生成に成功したかどうか
    Dim created As Boolean
    
    Dim i    As Long
    Dim size As Long
    
    ' リストのサイズを取得する
    size = popupMenuList.count
    
    ' リストの要素を1つずつ処理する
    For i = 1 To size
    
        ' ポップアップメニューオブジェクトの取得
        Dim popupMenu As ValPopupmenu
        Set popupMenu = popupMenuList.getItemByIndex(i)
        
        created = popupMenu.createPopup
        
        If created = True Then
        
            If popupMenu.shortcutKeyCode <> "" Then
            
                Application.OnKey _
                    popupMenu.shortcutKeyCode _
                  , ConstantsCommon.POPUP_MENU_CALL_FUNC_PREFIX & i
                  
                ' エラー発生時
                If err.Number <> 0 Then
                
                    ' ショートカットキーの設定に失敗したエラー情報を設定し、メッセージボックスに表示する
                    err.Number = ERR_NUMBER_POPUP_SETTING_FAILED
                    err.Description = ERR_DESC_POPUP_SETTING_FAILED & vbNewLine & _
                                      "(" & popupMenu.popupMenuName & ") << (" & popupMenu.shortcutKeyLabel & ")"
                    Main.ShowErrorMessage
                
                    ' ショートカットキー情報を削除する
                    popupMenu.shortcutKeyCode = ""
                    popupMenu.shortcutKeyLabel = ""
                
                    ' エラーをクリアする
                    err.Number = 0
                
                ' 正常時
                Else
                
                
                End If

            End If
            
        Else
        
            ' ショートカットキーの設定に失敗したエラー情報を設定し、メッセージボックスに表示する
            err.Number = ERR_NUMBER_POPUP_SETTING_FAILED
            err.Description = ERR_DESC_POPUP_SETTING_FAILED & " (" & popupMenu.popupMenuName & ")"
            Main.ShowErrorMessage
        
            err.Number = 0
        End If

    Next
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ▽ポップアップメニューの消去
'
' 概要　　　：現在のプロパティの状態を基にポップアップメニューをデフォルトの挙動に戻す
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub clearPopupMenu()

    Dim popupMenu       As ValPopupmenu
    
    Dim i As Long
    
    For Each popupMenu In popupMenuList.col
    
        i = i + 1
        
        Dim popupMenuControl As CommandBar
        
        Set popupMenuControl = popupMenu.commandBarPopup
        
        If Not popupMenuControl Is Nothing Then
        
            popupMenuControl.delete
            
            If popupMenu.shortcutKeyCode <> "" Then
            
                Application.OnKey _
                    popupMenu.shortcutKeyCode
            End If
    
        End If
    Next
    
End Sub

Private Sub setShortcutTextForCommandBar(ByRef shortcutInfo As ValShortcutKey _
                                       , ByRef targetBar1 As CommandBar _
                                       , ByRef targetBar2 As CommandBar _
                                       , ByRef targetBar3 As CommandBar _
                                       , ByVal text As String)

    Dim tmp As commandBarControl

    Dim shortcutCommandBar As commandBarControl
    Set shortcutCommandBar = shortcutInfo.commandBarControl

    ' ---------------------------------------------------
    ' ショートカットキーテキストの設定
    ' 自身に設定する
    If shortcutCommandBar.Type = msoControlButton Then
    
        ' ショートカットキーテキストを設定する
        shortcutCommandBar.ShortcutText = text
        
    End If
    
    ' 右クリックメニューへの設定
    Set tmp = targetBar1.FindControl(Tag:=shortcutCommandBar.Tag)
    If Not tmp Is Nothing Then
        With tmp
            If .Type = msoControlButton Then
            
                ' ショートカットキーテキストを設定する
                .ShortcutText = text
                
            End If
        End With
    End If
    
    ' 右クリックメニューへの設定
    Set tmp = targetBar2.FindControl(Tag:=shortcutCommandBar.Tag)
    If Not tmp Is Nothing Then
        With tmp
            If .Type = msoControlButton Then
            
                ' ショートカットキーテキストを設定する
                .ShortcutText = text
                
            End If
        End With
    End If
    
    ' 右クリックメニューへの設定
    Set tmp = targetBar3.FindControl(Tag:=shortcutCommandBar.Tag)
    If Not tmp Is Nothing Then
        With tmp
            If .Type = msoControlButton Then
            
                ' ショートカットキーテキストを設定する
                .ShortcutText = text
                
            End If
        End With
    End If

End Sub

Public Function CloneShortcutAppList() As ValCollection

    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim tmp  As ValShortcutKey
    Dim tmp2 As ValShortcutKey
    
    For Each tmp In shortcutAppList.col
    
        Set tmp2 = tmp.Clone
        
        ret.setItem tmp2, tmp2.commandBarControl.Tag
    Next
    
    Set CloneShortcutAppList = ret

End Function

Public Function CloneRClickMenuList() As ValCollection

    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim tmp  As Variant
    
    For Each tmp In rclickMenuItemList.col
    
        ret.setItem tmp, tmp
    Next
    
    Set CloneRClickMenuList = ret

End Function

Public Function ClonePopupMenuList() As ValCollection

    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim tmp  As ValPopupmenu
    Dim tmp2 As ValPopupmenu
    
    For Each tmp In popupMenuList.col
    
        Set tmp2 = tmp.Clone
        ret.setItem tmp2, tmp2.popupMenuName
    Next
    
    Set ClonePopupMenuList = ret

End Function
