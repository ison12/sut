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
    
    ' 情報を読み込む
    readForDataRClick
    readForDataShortcut
    readForDataPopupMenu
    
    ' 右クリックメニューに項目を追加する
    updateRClickMenu
    ' ショートカットキーを登録する
    updateShortcutKey
    ' ポップアップメニューを更新する
    updatePopupMenu

End Sub

' =========================================================
' ▽設定情報の生成（右クリックメニュー）
' =========================================================
Private Function createApplicationPropertiesForRClick() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingRClick.ini"

    Set createApplicationPropertiesForRClick = appProp
    
End Function

' =========================================================
' ▽設定情報の生成（ショートカット）
' =========================================================
Private Function createApplicationPropertiesForShortcut() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingShortcut.ini"

    Set createApplicationPropertiesForShortcut = appProp
    
End Function

' =========================================================
' ▽設定情報の生成（ポップアップメニュー）
' =========================================================
Private Function createApplicationPropertiesForPopupMenu() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingPopupmenu.ini"

    Set createApplicationPropertiesForPopupMenu = appProp
    
End Function

' =========================================================
' ▽設定情報の生成（ポップアップサブメニュー）
' =========================================================
Public Function createApplicationPropertiesForPopupMenuSub(ByVal popupMenuName As String) As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingPopupmenu_" & popupMenuName & ".ini"

    Set createApplicationPropertiesForPopupMenuSub = appProp
    
End Function

' =========================================================
' ▽情報を読み込む（右クリックメニュー）
' =========================================================
Public Sub readForDataRClick()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForRClick

    ' データを読み込む
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    Dim itemList As New ValCollection
    
    For Each val In values.col
    
        If val(2) = CStr(True) Then
            itemList.setItem val(1), val(1)
        End If
    
    Next
    
    Set rclickMenuItemList = itemList

    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を書き込む（右クリックメニュー）
' =========================================================
Public Sub writeForDataRClick()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForRClick

    ' 書き込みデータ
    Dim values As New ValCollection
    Dim val As Variant
    
    For Each val In rclickMenuItemList.col
    
        values.setItem Array(val, CStr(True))
        
    Next
    
    ' データを書き込む
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽情報を読み込む（ショートカットキー）
' =========================================================
Public Sub readForDataShortcut()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForShortcut

    ' データを読み込む
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    For Each val In values.col
        
        Dim shortcutInfo As ValShortcutKey
        Set shortcutInfo = shortcutAppList.getItem(val(1))
        
        If Not shortcutInfo Is Nothing Then
        
            shortcutInfo.shortcutKeyCode = val(2)
            shortcutInfo.shortcutKeyLabel = VBUtil.getAppOnKeyNameByMultipleCode(val(2))
        End If
    
    Next
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を書き込む（ショートカットキー）
' =========================================================
Public Sub writeForDataShortcut()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForShortcut
    
    ' 書き込みデータ
    Dim values As New ValCollection
    
    Dim shortcutInfo As ValShortcutKey
    For Each shortcutInfo In shortcutAppList.col
    
        values.setItem Array(shortcutInfo.commandBarControl.Tag, shortcutInfo.shortcutKeyCode)
    Next
    
    ' データを書き込む
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を読み込む（ポップアップメニュー）
' =========================================================
Public Sub readForDataPopupMenu()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForPopupMenu
    
    ' データを読み込む
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
        
    Dim i As Long: i = 1
    Do While True
    
        val = values.getItem(i & "_" & "popupMenuName", vbVariant)
        If Not IsArray(val) Then
            Exit Do
        End If
    
        Dim popupMenu As ValPopupmenu
        Set popupMenu = New ValPopupmenu: popupMenu.init ConstantsCommon.COMMANDBAR_MENU_NAME
        
        val = values.getItem(i & "_" & "popupMenuName", vbVariant): If IsArray(val) Then popupMenu.popupMenuName = val(2)
        val = values.getItem(i & "_" & "shortcutKeyCode", vbVariant): If IsArray(val) Then popupMenu.shortcutKeyCode = val(2)
        val = values.getItem(i & "_" & "shortcutKeyLabel", vbVariant): If IsArray(val) Then popupMenu.shortcutKeyLabel = val(2)
        
        readForDataPopupMenuSub popupMenu
        
        popupMenuList.setItem popupMenu, popupMenu.popupMenuName
    
        i = i + 1
    Loop
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を読み込む（ポップアップメニューに含まれる機能リスト）
' =========================================================
Public Sub readForDataPopupMenuSub(ByRef popupMenu As ValPopupmenu)

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForPopupMenuSub(popupMenu.popupMenuName)
    
    ' データを読み込む
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    Dim itemList As ValCollection
    Set itemList = popupMenu.itemList
    
    For Each val In values.col
    
        itemList.setItem val(1), val(1)
    Next
                
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を書き込む（ポップアップメニュー）
' =========================================================
Public Sub writeForDataPopupMenu()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForPopupMenu
    
    ' 書き込みデータ
    Dim values As New ValCollection
    Dim val As ValPopupmenu
    
    Dim i As Long
    i = 1
    For Each val In popupMenuList.col
    
        values.setItem Array(i & "_" & "popupMenuName", val.popupMenuName)
        values.setItem Array(i & "_" & "shortcutKeyCode", val.shortcutKeyCode)
        values.setItem Array(i & "_" & "shortcutKeyLabel", val.shortcutKeyLabel)
        
        writeForDataPopupMenuSub val
        
        i = i + 1
    Next
        
    ' データを書き込む
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を書き込む（ポップアップメニューに含まれる機能リスト）
' =========================================================
Public Sub writeForDataPopupMenuSub(ByRef popupMenu As ValPopupmenu)

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForPopupMenuSub(popupMenu.popupMenuName)

    ' 書き込みデータ
    Dim val As Variant
    Dim values As New ValCollection
    
    For Each val In popupMenu.itemList.col
    
        values.setItem Array(val, CStr(True))
    Next

    ' データを書き込む
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
                
    Exit Sub
    
err:

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

Public Function cloneShortcutAppList() As ValCollection

    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim tmp  As ValShortcutKey
    Dim tmp2 As ValShortcutKey
    
    For Each tmp In shortcutAppList.col
    
        Set tmp2 = tmp.clone
        
        ret.setItem tmp2, tmp2.commandBarControl.Tag
    Next
    
    Set cloneShortcutAppList = ret

End Function

Public Function cloneRClickMenuList() As ValCollection

    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim tmp  As Variant
    
    For Each tmp In rclickMenuItemList.col
    
        ret.setItem tmp, tmp
    Next
    
    Set cloneRClickMenuList = ret

End Function

Public Function clonePopupMenuList() As ValCollection

    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim tmp  As ValPopupmenu
    Dim tmp2 As ValPopupmenu
    
    For Each tmp In popupMenuList.col
    
        Set tmp2 = tmp.clone
        ret.setItem tmp2, tmp2.popupMenuName
    Next
    
    Set clonePopupMenuList = ret

End Function
