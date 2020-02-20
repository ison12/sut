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
' ���j���[���쎞�̏��� - �I�v�V����
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/03/15�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ________________________________________________________
' �����o�ϐ�
' ________________________________________________________
' ���I�v�V�����t�H�[���i�C�x���g��M�p�ɒ�`�j
Private WithEvents frmOptionVar As frmOption
Attribute frmOptionVar.VB_VarHelpID = -1
' ���E�N���b�N���j���[�ݒ�t�H�[���i�C�x���g��M�p�ɒ�`�j
Private WithEvents frmMenuSettingVar As frmMenuSetting
Attribute frmMenuSettingVar.VB_VarHelpID = -1
' ���V���[�g�J�b�g�L�[�ݒ�t�H�[���i�C�x���g��M�p�ɒ�`�j
Private WithEvents frmShortcutKeyVar As frmShortcutKey
Attribute frmShortcutKeyVar.VB_VarHelpID = -1
' ���|�b�v�A�b�v���j���[�ݒ�t�H�[���i�C�x���g��M�p�ɒ�`�j
Private WithEvents frmPopupMenuVar As frmPopupMenu
Attribute frmPopupMenuVar.VB_VarHelpID = -1

' ________________________________________________________
' �C�x���g
' ________________________________________________________
' =========================================================
' ���I�v�V�����t�H�[����OK�{�^�����������ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FapplicationSetting �A�v���P�[�V�����ݒ���
'
' =========================================================
Private Sub frmOptionVar_ok(ByRef applicationSetting As ValApplicationSetting)

End Sub

' =========================================================
' ���I�v�V�����t�H�[���ŃL�����Z���{�^�����������ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub frmOptionVar_cancel()

End Sub

' =========================================================
' ���E�N���b�N���j���[�ݒ�t�H�[����OK�{�^�����������ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FapplicationSetting �A�v���P�[�V�����ݒ���
'
' =========================================================
Private Sub frmMenuSettingVar_ok(ByRef appSettingShortcut As ValApplicationSettingShortcut _
                               , ByRef selectedItemList As ValCollection _
                               , ByVal menuName As String)

    ' ���݂̏��ŉE�N���b�N���j���[����U�폜����
    appSettingShortcut.clearRClickMenu

    ' �A�v���P�[�V�����ݒ���ɉE�N���b�N���j���[�ɓo�^���鍀�ڂ�ݒ肵���W�X�g���ɓo�^����
    Set appSettingShortcut.rclickMenuItemList = selectedItemList
    appSettingShortcut.writeForDataRClick

    ' �o�^���ꂽ���ŉE�N���b�N���j���[���X�V����
    appSettingShortcut.updateRClickMenu

End Sub

' =========================================================
' ���E�N���b�N���j���[�ݒ�t�H�[���ŃL�����Z���{�^�����������ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Private Sub frmMenuSettingVar_cancel()

End Sub

' =========================================================
' ���E�N���b�N���j���[�ݒ�t�H�[���Ń��Z�b�g�{�^�����������ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Private Sub frmMenuSettingVar_reset(appSettingShortcut As ValApplicationSettingShortcut _
                                  , ByRef Cancel As Boolean)

    Dim result As Long
    result = VBUtil.showMessageBoxForYesNo("�E�N���b�N���j���[��������ԂɃ��Z�b�g����܂��B" & vbNewLine & _
                    "������ԂƂ́AExcel���C���X�g�[�����ꂽ����̎����w���܂��B" & vbNewLine & _
                    "���������āA���̃}�N���ɂ���Ēǉ����ꂽ���ڂ�����΁A���������Z�b�g�̑ΏۂɂȂ�܂�����낵���ł����H", ConstantsCommon.APPLICATION_NAME)
    
    If result = WinAPI_User.IDNO Then
    
        Cancel = True
        Exit Sub
    End If
    
    ' �E�N���b�N���j���[���f�t�H���g��Ԃɂ���
    appSettingShortcut.resetRClickMenu

    ' ��̃R���N�V������ݒ肵���W�X�g���ɓo�^����i������S�č폜����j
    Set appSettingShortcut.rclickMenuItemList = New ValCollection
    appSettingShortcut.writeForDataRClick

End Sub


' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ���I�v�V�����ݒ�
'
' �T�v�@�@�@�F�I�v�V���� �������̏���
'
' =========================================================
Public Sub settingOption()

    On Error GoTo err

    ' �t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmOption) Then Unload frmOption
    Load frmOption
    ' �C�x���g��M�p�̃t�H�[��������������
    Set frmOptionVar = frmOption
    
    ' �I�v�V�����t�H�[����\������
    frmOptionVar.ShowExt vbModal, Main.getApplicationSetting, Main.getApplicationSettingColFormat

    ' �C�x���g��M�p�̃I�v�V�����t�H�[����j������
    Set frmOptionVar = Nothing
    
    Exit Sub

err:

    Main.ShowErrorMessage

    ' �C�x���g��M�p�̃I�v�V�����t�H�[����j������
    Set frmOptionVar = Nothing
    
End Sub

' =========================================================
' ���E�N���b�N���j���[�̐ݒ�
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub settingRClickMenu()

    On Error GoTo err

    Dim appSettingShortcut As ValApplicationSettingShortcut
    Set appSettingShortcut = Main.getApplicationSettingShortcut

    ' �t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmMenuSetting) Then Unload frmMenuSetting
    Load frmMenuSetting
    ' �C�x���g��M�p�̃t�H�[��������������
    Set frmMenuSettingVar = frmMenuSetting

    ' ���j���[�ݒ�t�H�[����\������
    frmMenuSettingVar.ShowExt Nothing _
                            , vbModal _
                            , appSettingShortcut _
                            , appSettingShortcut.cloneRClickMenuList _
                            , "�E�N���b�N���j���[�̐ݒ�" _
                            , "�V�[�g��ŉE�N���b�N�����ۂɕ\�������|�b�v�A�b�v���J�X�^�}�C�Y���܂��B" _
                            , "�E�N���b�N���j���[" _
                            , True

    ' �C�x���g��M�p�̃��j���[�ݒ�t�H�[����j������
    Set frmMenuSettingVar = Nothing

    Exit Sub

err:

    Main.ShowErrorMessage

    ' �C�x���g��M�p�̃��j���[�ݒ�t�H�[����j������
    Set frmMenuSettingVar = Nothing

End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�̐ݒ�
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub settingShortCutKey()

    On Error GoTo err

    ' �t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmShortcutKey) Then Unload frmShortcutKey
    Load frmShortcutKey
    ' �C�x���g��M�p�̃t�H�[��������������
    Set frmShortcutKeyVar = frmShortcutKey
    ' �V���[�g�J�b�g�L�[�t�H�[����\������
    frmShortcutKeyVar.ShowExt vbModal, Main.getApplicationSettingShortcut

    ' �C�x���g��M�p�̃V���[�g�J�b�g�L�[�t�H�[����j������
    Set frmShortcutKeyVar = Nothing
    
    Exit Sub

err:

    Main.ShowErrorMessage

    ' �C�x���g��M�p�̃V���[�g�J�b�g�L�[�t�H�[����j������
    Set frmShortcutKeyVar = Nothing
    
End Sub

' =========================================================
' ���|�b�v�A�b�v���j���[�̐ݒ�
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub settingPopupMenu()

    On Error GoTo err

    ' �t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmPopupMenu) Then Unload frmPopupMenu
    Load frmPopupMenu
    ' �C�x���g��M�p�̃t�H�[��������������
    Set frmPopupMenuVar = frmPopupMenu
    ' �V���[�g�J�b�g�L�[�t�H�[����\������
    frmPopupMenuVar.ShowExt vbModal, Main.getApplicationSettingShortcut

    ' �C�x���g��M�p�̃|�b�v�A�b�v���j���[�t�H�[����j������
    Set frmPopupMenuVar = Nothing

    Exit Sub

err:

    Main.ShowErrorMessage

    ' �C�x���g��M�p�̃|�b�v�A�b�v���j���[�t�H�[����j������
    Set frmPopupMenuVar = Nothing
    
End Sub
