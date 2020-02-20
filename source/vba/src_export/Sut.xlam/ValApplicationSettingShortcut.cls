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
' �A�v���P�[�V�����̃I�v�V�������
' �V���[�g�J�b�g�L�[�֘A
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/06/04�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' =========================================================
' �������o
' =========================================================

' �{�A�v���P�[�V������CommandBarControl�ɂ͈�ӂ�ID��Tag�v���p�e�B�Ɋi�[���Ă���
' CommandBar�ɂ�FindControl���\�b�h������ATag�������ɂ��ăR���g���[�����������邱�Ƃ��\�B
' ����FindControl�𗘗p���邽�߂ɁATag�v���p�e�B�Ɉ�ӂ�ID���i�[���Ă���B

' �V���[�g�J�b�g�L�[�Ɋ��蓖�ĉ\�ȋ@�\���X�g
'   [ key   ] : CommandBarControl.Tag
'   [ value ] : ValShortcutKey
Public shortcutAppList   As ValCollection

' �E�N���b�N���j���[���X�g
'   [ key   ] : CommandBarControl.Tag
'   [ value ] : CommandBarControl.Tag
Public rclickMenuItemList As ValCollection

' �|�b�v�A�b�v���j���[���X�g
'   [ key   ] : CommandBarControl.Tag
'   [ value ] : ValPopupMenu
Public popupMenuList As ValCollection

' =========================================================
' ���萔�v���p�e�B
' �@Class���W���[���ł́APublic Const����`�ł��Ȃ��̂Ńv���p�e�B�ő�p����
' =========================================================

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set shortcutAppList = New ValCollection
    Set rclickMenuItemList = New ValCollection
    Set popupMenuList = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' �E�N���b�N���j���[����ǉ����ꂽ���ڂ���������
    clearRClickMenu
    ' �V���[�g�J�b�g�L�[����ǉ����ꂽ���ڂ���������
    clearShortcutKey
    ' �ǉ����ꂽ�|�b�v�A�b�v���j���[����������
    clearPopupMenu

    Set shortcutAppList = Nothing
    Set rclickMenuItemList = Nothing
    Set popupMenuList = Nothing
    
End Sub

Public Sub init()
    
    ' �V���[�g�J�b�g�@�\���X�g�𐶐�����
    
    Dim shortcutInfo As ValShortcutKey
    
    ' Sut���j���[
    Dim sutMenu As CommandBar
    ' Sut���j���[�̗v�f
    Dim sutMenuItem As commandBarControl
    
    ' Sut���j���[���擾����
    Set sutMenu = Application.CommandBars(ConstantsCommon.COMMANDBAR_MENU_NAME)
    
    ' ��1�K�w
    For Each sutMenuItem In sutMenu.Controls
    
        ' ��1�K�w�̃{�^��
        If sutMenuItem.Type = msoControlButton And sutMenuItem.DescriptionText <> "" Then
        
            Set shortcutInfo = New ValShortcutKey
            shortcutInfo.commandBarControl = sutMenuItem

            shortcutAppList.setItem shortcutInfo, sutMenuItem.Tag
        
        ' ��2�K�w
        ElseIf sutMenuItem.Type = msoControlPopup Then
        
            Dim sutMenu2Popup As commandBarPopup
            Dim sutMenu2Item  As commandBarControl
            
            Set sutMenu2Popup = sutMenuItem
            
            For Each sutMenu2Item In sutMenu2Popup.Controls
            
                ' ��2�K�w�̃{�^��
                If sutMenu2Item.Type = msoControlButton And sutMenu2Item.DescriptionText <> "" Then
                
                    Set shortcutInfo = New ValShortcutKey
                    shortcutInfo.commandBarControl = sutMenu2Item
        
                    shortcutAppList.setItem shortcutInfo, sutMenu2Item.Tag
                    
                ' ��3�K�w
                ElseIf sutMenu2Item.Type = msoControlPopup Then
                
                    Dim sutMenu3Popup As commandBarPopup
                    Dim sutMenu3Item  As commandBarControl
                    
                    Set sutMenu3Popup = sutMenu2Item
                    
                    For Each sutMenu3Item In sutMenu3Popup.Controls
                    
                        ' ��3�K�w�̃{�^��
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
    
    ' ����ǂݍ���
    readForDataRClick
    readForDataShortcut
    readForDataPopupMenu
    
    ' �E�N���b�N���j���[�ɍ��ڂ�ǉ�����
    updateRClickMenu
    ' �V���[�g�J�b�g�L�[��o�^����
    updateShortcutKey
    ' �|�b�v�A�b�v���j���[���X�V����
    updatePopupMenu

End Sub

' =========================================================
' ���ݒ���̐����i�E�N���b�N���j���[�j
' =========================================================
Private Function createApplicationPropertiesForRClick() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingRClick.ini"

    Set createApplicationPropertiesForRClick = appProp
    
End Function

' =========================================================
' ���ݒ���̐����i�V���[�g�J�b�g�j
' =========================================================
Private Function createApplicationPropertiesForShortcut() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingShortcut.ini"

    Set createApplicationPropertiesForShortcut = appProp
    
End Function

' =========================================================
' ���ݒ���̐����i�|�b�v�A�b�v���j���[�j
' =========================================================
Private Function createApplicationPropertiesForPopupMenu() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingPopupmenu.ini"

    Set createApplicationPropertiesForPopupMenu = appProp
    
End Function

' =========================================================
' ���ݒ���̐����i�|�b�v�A�b�v�T�u���j���[�j
' =========================================================
Public Function createApplicationPropertiesForPopupMenuSub(ByVal popupMenuName As String) As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingPopupmenu_" & popupMenuName & ".ini"

    Set createApplicationPropertiesForPopupMenuSub = appProp
    
End Function

' =========================================================
' ������ǂݍ��ށi�E�N���b�N���j���[�j
' =========================================================
Public Sub readForDataRClick()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForRClick

    ' �f�[�^��ǂݍ���
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
' �������������ށi�E�N���b�N���j���[�j
' =========================================================
Public Sub writeForDataRClick()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForRClick

    ' �������݃f�[�^
    Dim values As New ValCollection
    Dim val As Variant
    
    For Each val In rclickMenuItemList.col
    
        values.setItem Array(val, CStr(True))
        
    Next
    
    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ������ǂݍ��ށi�V���[�g�J�b�g�L�[�j
' =========================================================
Public Sub readForDataShortcut()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForShortcut

    ' �f�[�^��ǂݍ���
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
' �������������ށi�V���[�g�J�b�g�L�[�j
' =========================================================
Public Sub writeForDataShortcut()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForShortcut
    
    ' �������݃f�[�^
    Dim values As New ValCollection
    
    Dim shortcutInfo As ValShortcutKey
    For Each shortcutInfo In shortcutAppList.col
    
        values.setItem Array(shortcutInfo.commandBarControl.Tag, shortcutInfo.shortcutKeyCode)
    Next
    
    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ������ǂݍ��ށi�|�b�v�A�b�v���j���[�j
' =========================================================
Public Sub readForDataPopupMenu()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForPopupMenu
    
    ' �f�[�^��ǂݍ���
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
' ������ǂݍ��ށi�|�b�v�A�b�v���j���[�Ɋ܂܂��@�\���X�g�j
' =========================================================
Public Sub readForDataPopupMenuSub(ByRef popupMenu As ValPopupmenu)

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForPopupMenuSub(popupMenu.popupMenuName)
    
    ' �f�[�^��ǂݍ���
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
' �������������ށi�|�b�v�A�b�v���j���[�j
' =========================================================
Public Sub writeForDataPopupMenu()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForPopupMenu
    
    ' �������݃f�[�^
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
        
    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' �������������ށi�|�b�v�A�b�v���j���[�Ɋ܂܂��@�\���X�g�j
' =========================================================
Public Sub writeForDataPopupMenuSub(ByRef popupMenu As ValPopupmenu)

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationPropertiesForPopupMenuSub(popupMenu.popupMenuName)

    ' �������݃f�[�^
    Dim val As Variant
    Dim values As New ValCollection
    
    For Each val In popupMenu.itemList.col
    
        values.setItem Array(val, CStr(True))
    Next

    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
                
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���E�N���b�N���j���[�̃��Z�b�g
'
' �T�v�@�@�@�F�E�N���b�N���j���[�����̏�Ԃɖ߂��B
' �����@�@�@�F
' �߂�l�@�@�F
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
' ���E�N���b�N���j���[�ɒǉ��������ڂ̍폜
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
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
' ���E�N���b�N���j���[�ւ̍��ڒǉ�
'
' �T�v�@�@�@�F���݂̃v���p�e�B�̏�Ԃ���ɉE�N���b�N���j���[�ɍ��ڂ�ǉ�����B
' �����@�@�@�F
' �߂�l�@�@�F
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
    
        ' �V���[�g�J�b�g�L�[�̐ݒ�Ɏ��s�����G���[����ݒ肵�A���b�Z�[�W�{�b�N�X�ɕ\������
        err.Number = ERR_NUMBER_RCLICKMENU_SETTING_FAILED
        err.Description = ERR_DESC_RCLICKMENU_SETTING_FAILED
        Main.ShowErrorMessage
    
        err.Number = 0
    End If
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�̍X�V
'
' �T�v�@�@�@�F���݂̃v���p�e�B�̏�Ԃ���ɃV���[�g�J�b�g�L�[���X�V����B
' �����@�@�@�F
' �߂�l�@�@�F
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
              
            ' �G���[������
            If err.Number <> 0 Then
            
                ' �V���[�g�J�b�g�L�[�̐ݒ�Ɏ��s�����G���[����ݒ肵�A���b�Z�[�W�{�b�N�X�ɕ\������
                err.Number = ERR_NUMBER_SHORTCUT_SETTING_FAILED
                err.Description = ERR_DESC_SHORTCUT_SETTING_FAILED & vbNewLine & _
                                  "(" & shortcutInfo.commandName & ") << (" & shortcutInfo.shortcutKeyLabel & ")"
                Main.ShowErrorMessage
            
                ' �V���[�g�J�b�g�L�[�����폜����
                shortcutInfo.shortcutKeyCode = ""
                shortcutInfo.shortcutKeyLabel = ""
            
                ' �G���[���N���A����
                err.Number = 0
                
            ' ���펞
            Else
            
                ' ---------------------------------------------------
                ' �V���[�g�J�b�g�L�[�e�L�X�g�̐ݒ�
                ' Excel2007�ł̓V���[�g�J�b�g�L�[�e�L�X�g���\������Ȃ��悤�Ȃ̂ŁA�e�L�X�g�̐ݒ�͍s��Ȃ�
                'setShortcutTextForCommandBar shortcutInfo, targetBar1, targetBar2, targetBar3, shortcutInfo.shortcutKeyLabel
    
            End If
              
        End If

continue:
    Next
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�̏���
'
' �T�v�@�@�@�F���݂̃v���p�e�B�̏�Ԃ���ɃV���[�g�J�b�g�L�[���f�t�H���g�̋����ɖ߂�
' �����@�@�@�F
' �߂�l�@�@�F
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
        
            ' Application#OnKey��Procedure���w�肵�Ȃ����ƂŃf�t�H���g�̋����ɖ߂�
            Application.OnKey _
                shortcutInfo.shortcutKeyCode
                
            ' ---------------------------------------------------
            ' �V���[�g�J�b�g�L�[�e�L�X�g�̐ݒ�
            ' Excel2007�ł̓V���[�g�J�b�g�L�[�e�L�X�g���\������Ȃ��悤�Ȃ̂ŁA�e�L�X�g�̐ݒ�͍s��Ȃ�
            'setShortcutTextForCommandBar shortcutInfo, targetBar1, targetBar2, targetBar3, ""
            
        End If
            
    Next
    
End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�̖�����
'
' �T�v�@�@�@�F���݂̃v���p�e�B�̏�Ԃ���ɃV���[�g�J�b�g�L�[�𖳌�������
' �����@�@�@�F
' �߂�l�@�@�F
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
        
            ' Application#OnKey��Procedure�ɋ󕶎�����w�肷�邱�ƂŃL�[�𖳌�������
            Application.OnKey _
                shortcutInfo.shortcutKeyCode, ""
                
            ' ---------------------------------------------------
            ' �V���[�g�J�b�g�L�[�e�L�X�g�̐ݒ�
            ' Excel2007�ł̓V���[�g�J�b�g�L�[�e�L�X�g���\������Ȃ��悤�Ȃ̂ŁA�e�L�X�g�̐ݒ�͍s��Ȃ�
            ' setShortcutTextForCommandBar shortcutInfo, targetBar1, targetBar2, targetBar3, ""
            
        End If
            
    Next
    
End Sub

' =========================================================
' ���|�b�v�A�b�v���j���[�̍X�V
'
' �T�v�@�@�@�F���݂̃v���p�e�B�̏�Ԃ���Ƀ|�b�v�A�b�v���j���[���X�V����B
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub updatePopupMenu()

    On Error Resume Next

    ' �����ɐ����������ǂ���
    Dim created As Boolean
    
    Dim i    As Long
    Dim size As Long
    
    ' ���X�g�̃T�C�Y���擾����
    size = popupMenuList.count
    
    ' ���X�g�̗v�f��1����������
    For i = 1 To size
    
        ' �|�b�v�A�b�v���j���[�I�u�W�F�N�g�̎擾
        Dim popupMenu As ValPopupmenu
        Set popupMenu = popupMenuList.getItemByIndex(i)
        
        created = popupMenu.createPopup
        
        If created = True Then
        
            If popupMenu.shortcutKeyCode <> "" Then
            
                Application.OnKey _
                    popupMenu.shortcutKeyCode _
                  , ConstantsCommon.POPUP_MENU_CALL_FUNC_PREFIX & i
                  
                ' �G���[������
                If err.Number <> 0 Then
                
                    ' �V���[�g�J�b�g�L�[�̐ݒ�Ɏ��s�����G���[����ݒ肵�A���b�Z�[�W�{�b�N�X�ɕ\������
                    err.Number = ERR_NUMBER_POPUP_SETTING_FAILED
                    err.Description = ERR_DESC_POPUP_SETTING_FAILED & vbNewLine & _
                                      "(" & popupMenu.popupMenuName & ") << (" & popupMenu.shortcutKeyLabel & ")"
                    Main.ShowErrorMessage
                
                    ' �V���[�g�J�b�g�L�[�����폜����
                    popupMenu.shortcutKeyCode = ""
                    popupMenu.shortcutKeyLabel = ""
                
                    ' �G���[���N���A����
                    err.Number = 0
                
                ' ���펞
                Else
                
                
                End If

            End If
            
        Else
        
            ' �V���[�g�J�b�g�L�[�̐ݒ�Ɏ��s�����G���[����ݒ肵�A���b�Z�[�W�{�b�N�X�ɕ\������
            err.Number = ERR_NUMBER_POPUP_SETTING_FAILED
            err.Description = ERR_DESC_POPUP_SETTING_FAILED & " (" & popupMenu.popupMenuName & ")"
            Main.ShowErrorMessage
        
            err.Number = 0
        End If

    Next
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ���|�b�v�A�b�v���j���[�̏���
'
' �T�v�@�@�@�F���݂̃v���p�e�B�̏�Ԃ���Ƀ|�b�v�A�b�v���j���[���f�t�H���g�̋����ɖ߂�
' �����@�@�@�F
' �߂�l�@�@�F
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
    ' �V���[�g�J�b�g�L�[�e�L�X�g�̐ݒ�
    ' ���g�ɐݒ肷��
    If shortcutCommandBar.Type = msoControlButton Then
    
        ' �V���[�g�J�b�g�L�[�e�L�X�g��ݒ肷��
        shortcutCommandBar.ShortcutText = text
        
    End If
    
    ' �E�N���b�N���j���[�ւ̐ݒ�
    Set tmp = targetBar1.FindControl(Tag:=shortcutCommandBar.Tag)
    If Not tmp Is Nothing Then
        With tmp
            If .Type = msoControlButton Then
            
                ' �V���[�g�J�b�g�L�[�e�L�X�g��ݒ肷��
                .ShortcutText = text
                
            End If
        End With
    End If
    
    ' �E�N���b�N���j���[�ւ̐ݒ�
    Set tmp = targetBar2.FindControl(Tag:=shortcutCommandBar.Tag)
    If Not tmp Is Nothing Then
        With tmp
            If .Type = msoControlButton Then
            
                ' �V���[�g�J�b�g�L�[�e�L�X�g��ݒ肷��
                .ShortcutText = text
                
            End If
        End With
    End If
    
    ' �E�N���b�N���j���[�ւ̐ݒ�
    Set tmp = targetBar3.FindControl(Tag:=shortcutCommandBar.Tag)
    If Not tmp Is Nothing Then
        With tmp
            If .Type = msoControlButton Then
            
                ' �V���[�g�J�b�g�L�[�e�L�X�g��ݒ肷��
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