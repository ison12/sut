VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPopupMenu 
   Caption         =   "�|�b�v�A�b�v���j���[�̐ݒ�"
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
' �|�b�v�A�b�v���j���[�̐ݒ�
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/06/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FappSettingShortcut �A�v���P�[�V�����ݒ�V���[�g�J�b�g
' �@�@�@�@�@�@selectedItemList �I���ςݍ��ڃ��X�g
' �@�@�@�@�@�@menuName �V�������j���[��
'
' =========================================================
Public Event ok(ByRef applicationSetting As ValApplicationSettingShortcut)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' �|�b�v�A�b�v���j���[�̐V�K�쐬���̃f�t�H���g������
Private Const POPUP_MENU_NEW_CREATED_STR As String = "Popup Menu"

' �|�b�v�A�b�v���j���[�̐V�K�쐬�ő吔
Private Const POPUP_MENU_NEW_CREATED_OVER_SIZE As String = "�|�b�v�A�b�v�͍ő�${count}�܂œo�^�\�ł��B"

' ���j���[�ݒ���
Private WithEvents frmMenuSettingVar As frmMenuSetting
Attribute frmMenuSettingVar.VB_VarHelpID = -1

' �V���[�g�J�b�g�L�[�ݒ���
Private WithEvents frmShortcutKeySettingVar As frmShortcutKeySetting
Attribute frmShortcutKeySettingVar.VB_VarHelpID = -1

' �A�v���P�[�V�����ݒ���i�V���[�g�J�b�g�L�[�j
Private applicationSetting As ValApplicationSettingShortcut

' �|�b�v�A�b�v���j���[���X�g �R���g���[��
Private popupMenuList As CntListBox

' �|�b�v�A�b�v���j���[���X�g�ł̑I�����ڃC���f�b�N�X
Private popupMenuListSelectedIndex As Long
' �|�b�v�A�b�v���j���[���X�g�ł̑I�����ڃI�u�W�F�N�g
Private popupMenuListSelectedItem As ValPopupmenu

' �Ώۃu�b�N
Private targetBook As Workbook
' �Ώۃu�b�N���擾����
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@var   �A�v���P�[�V�����ݒ���
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef var As ValApplicationSettingShortcut)

    ' �����o�ϐ��ɃA�v���P�[�V�����ݒ����ݒ肷��
    Set applicationSetting = var
    
    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal
End Sub

' =========================================================
' ���t�H�[����\��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub HideExt()

    deactivate
    Me.Hide

End Sub

' =========================================================
' ���t�H�[���A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
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
' ���t�H�[���f�B�A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub deactivate()

    Set popupMenuList = Nothing
    
    ' Nothing��ݒ肷�邱�ƂŃC�x���g����M���Ȃ��悤�ɂ���
    Set frmMenuSettingVar = Nothing
    Set frmShortcutKeySettingVar = Nothing
    
    Main.storeFormPosition Me.name, Me

End Sub

' =========================================================
' ���t�H�[�����������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
    ' ���[�h���_�̃A�N�e�B�u�u�b�N��ێ����Ă���
    Set targetBook = ExcelUtil.getActiveWorkbook
    ' ���������������s����
    initial
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���t�H�[���j�����̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Terminate()

    On Error GoTo err
    
    ' �j�����������s����
    unInitial
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���t�H�[���A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Activate()

End Sub

' =========================================================
' ���|�b�v�A�b�v���j���[���X�g�{�b�N�X�_�u���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstPopupMenu_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    editPopup
End Sub

' =========================================================
' ��OK�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err
    
    ' �����L�^����
    storeShortcut
    
    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(applicationSetting)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub


' =========================================================
' ���L�����Z���{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdCancel_Click()

    On Error GoTo err
    
    ' �t�H�[�������
    HideExt
    
    ' �L�����Z���C�x���g�𑗐M����
    RaiseEvent Cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���V�K�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdAdd_Click()

    ' ���X�g�{�b�N�X�̃T�C�Y
    Dim cnt As Long
    ' ���X�g�{�b�N�X�̃T�C�Y���擾����
    cnt = popupMenuList.collection.count
    
    ' �|�b�v�A�b�v�̐����ő�o�^���𒴂��Ă��邩�`�F�b�N����
    If cnt >= ConstantsCommon.POPUP_MENU_NEW_CREATED_MAX_SIZE Then
    
        ' ���b�Z�[�W��\������
        Dim mess As String
        mess = replace(POPUP_MENU_NEW_CREATED_OVER_SIZE, "${count}", ConstantsCommon.POPUP_MENU_NEW_CREATED_MAX_SIZE)
        
        VBUtil.showMessageBoxForInformation mess _
                                          , ConstantsCommon.APPLICATION_NAME
        Exit Sub
    End If
    
    ' �|�b�v�A�b�v���j���[�I�u�W�F�N�g�����X�g�ɒǉ�����
    Dim popupMenu As ValPopupmenu
    Set popupMenu = New ValPopupmenu: popupMenu.init ConstantsCommon.COMMANDBAR_MENU_NAME
    
    '
    popupMenu.popupMenuName = POPUP_MENU_NEW_CREATED_STR & " " & (cnt + 1)
    
    popupMenuList.addItem popupMenu.popupMenuName, popupMenu
    
    lstPopupMenu.ListIndex = cnt
    lstPopupMenu.SetFocus
End Sub

' =========================================================
' ���ҏW�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdEdit_Click()

    editPopup
End Sub

Private Sub editPopup()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    popupMenuListSelectedIndex = lstPopupMenu.ListIndex

    ' ���I���̏ꍇ
    If popupMenuListSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set popupMenuListSelectedItem = popupMenuList.getItem(popupMenuListSelectedIndex)
    
    Set frmMenuSettingVar = frmMenuSetting
    frmMenuSettingVar.ShowExt Me _
                            , vbModal _
                            , applicationSetting _
                            , popupMenuListSelectedItem.itemList _
                            , "" _
                            , "�|�b�v�A�b�v���j���[�̐ݒ�����܂��B" _
                            , popupMenuListSelectedItem.popupMenuName
    Set frmMenuSettingVar = Nothing

End Sub

' =========================================================
' �����j���[�ݒ�t�H�[����OK�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
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
' �����j���[�ݒ�t�H�[���̃L�����Z���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmMenuSettingVar_cancel()

    lstPopupMenu.SetFocus
End Sub

' =========================================================
' �����j���[�ݒ�t�H�[���̃��Z�b�g�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmMenuSettingVar_reset(appSettingShortcut As ValApplicationSettingShortcut _
                                  , ByRef Cancel As Boolean)

End Sub

' =========================================================
' ���폜�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDelete_Click()

    Dim selectedIndex As Long
    
    ' ���ݑI������Ă���C���f�b�N�X���擾
    selectedIndex = lstPopupMenu.ListIndex

    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    popupMenuList.removeItem selectedIndex
    
    lstPopupMenu.SetFocus

End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�ݒ�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdShortcut_Click()

    ' �V���[�g�J�b�g�L�[�̐ݒ�
    editShortcutKey
End Sub

Private Sub editShortcutKey()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    popupMenuListSelectedIndex = lstPopupMenu.ListIndex

    ' ���I���̏ꍇ
    If popupMenuListSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' �V���[�g�J�b�g���̎擾
    Set popupMenuListSelectedItem = popupMenuList.getItem(popupMenuListSelectedIndex)

    Set frmShortcutKeySettingVar = frmShortcutKeySetting
    ' �V���[�g�J�b�g�L�[�ݒ�p�̃t�H�[�����J��
    frmShortcutKeySettingVar.ShowExt vbModal, popupMenuListSelectedItem.shortcutKeyCode
    Set frmShortcutKeySettingVar = Nothing
    
End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�̐ݒ�_�C�A���O��OK�{�^�����������ꂽ�ꍇ�̃C�x���g
' =========================================================
Private Sub frmShortcutKeySettingVar_ok(ByVal KeyCode As String, ByVal keyLabel As String)

    popupMenuListSelectedItem.shortcutKeyCode = KeyCode
    popupMenuListSelectedItem.shortcutKeyLabel = keyLabel
    
    lstPopupMenu.list(popupMenuListSelectedIndex, 1) = keyLabel
    
    lstPopupMenu.SetFocus
End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�̐ݒ�_�C�A���O�ŃL�����Z���{�^�����������ꂽ�ꍇ�̃C�x���g
' =========================================================
Private Sub frmShortcutKeySettingVar_cancel()

    lstPopupMenu.SetFocus
End Sub

' =========================================================
' ������������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub initial()

End Sub

' =========================================================
' ����n������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub unInitial()

End Sub

' =========================================================
' ���I�v�V��������ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeShortcut()

    applicationSetting.clearPopupMenu
    
    Set applicationSetting.popupMenuList = popupMenuList.collection
    applicationSetting.writeForDataPopupMenu

    applicationSetting.updatePopupMenu

End Sub

' =========================================================
' ���I�v�V��������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreShortcut()

    Set popupMenuList = New CntListBox: popupMenuList.init lstPopupMenu
    
    popupMenuList.addAll applicationSetting.ClonePopupMenuList _
                       , "popupMenuName" _
                       , "shortcutKeyLabel"
    
End Sub

