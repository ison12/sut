VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShortcutKey 
   Caption         =   "�V���[�g�J�b�g�L�[�̐ݒ�"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6435
   OleObjectBlob   =   "frmShortcutKey.frx":0000
End
Attribute VB_Name = "frmShortcutKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' �V���[�g�J�b�g�L�[�̐ݒ�
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/06/02�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FapplicationSetting �A�v���P�[�V�����ݒ���
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

' �V���[�g�J�b�g�L�[�ݒ���
Private WithEvents frmShortcutKeySettingVar As frmShortcutKeySetting
Attribute frmShortcutKeySettingVar.VB_VarHelpID = -1

' �A�v���P�[�V�����ݒ���i�V���[�g�J�b�g�L�[�j
Private applicationSetting As ValApplicationSettingShortcut

' �@�\���X�g �R���g���[��
Private appMenuList As CntListBox

' �@�\���X�g�ł̑I�����ڃC���f�b�N�X
Private appMenuListSelectedIndex As Long
' �@�\���X�g�ł̑I�����ڃI�u�W�F�N�g
Private appMenuListSelectedItem As ValShortcutKey

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
    
    Main.storeFormPosition Me.name, Me
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

    If VBUtil.unloadFormIfChangeActiveBook(frmShortcutKeySetting) Then Unload frmShortcutKeySetting
    Load frmShortcutKeySetting

    restoreShortcut
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

    ' �t�H�[���N���[�Y��ɃC�x���g����M���Ȃ��悤�Ƀt�H�[���ϐ����N���A���Ă���
    Set frmShortcutKeySettingVar = Nothing
    
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
' ���t�H�[���̕��鎞�̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' �{�����ł͏������̂��L�����Z������
        Cancel = True
        ' �ȉ��̃C�x���g�o�R�ŕ���
        cmdCancel_Click
    End If
    
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
' �����Z�b�g�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdReset_Click()

    On Error GoTo err
    
    resetShortcut
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���@�\���X�g�{�b�N�X�_�u���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstAppList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    editAppShortcutKey
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

    editAppShortcutKey
End Sub

' =========================================================
' �������{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDelete_Click()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    appMenuListSelectedIndex = lstAppList.ListIndex

    ' ���I���̏ꍇ
    If appMenuListSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' �V���[�g�J�b�g���̎擾
    Set appMenuListSelectedItem = appMenuList.getItem(appMenuListSelectedIndex)

    appMenuListSelectedItem.shortcutKeyCode = ""
    appMenuListSelectedItem.shortcutKeyLabel = ""
    
    lstAppList.list(appMenuListSelectedIndex, 1) = ""
    
    lstAppList.SetFocus

End Sub

Private Sub editAppShortcutKey()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    appMenuListSelectedIndex = lstAppList.ListIndex

    ' ���I���̏ꍇ
    If appMenuListSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' �V���[�g�J�b�g���̎擾
    Set appMenuListSelectedItem = appMenuList.getItem(appMenuListSelectedIndex)

    Set frmShortcutKeySettingVar = frmShortcutKeySetting
    ' �V���[�g�J�b�g�L�[�ݒ�p�̃t�H�[�����J��
    frmShortcutKeySettingVar.ShowExt vbModal, appMenuListSelectedItem.shortcutKeyCode
    Set frmShortcutKeySettingVar = Nothing

End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�̐ݒ�_�C�A���O��OK�{�^�����������ꂽ�ꍇ�̃C�x���g
' =========================================================
Private Sub frmShortcutKeySettingVar_ok(ByVal KeyCode As String, ByVal keyLabel As String)

    appMenuListSelectedItem.shortcutKeyCode = KeyCode
    appMenuListSelectedItem.shortcutKeyLabel = keyLabel
    
    lstAppList.list(appMenuListSelectedIndex, 1) = keyLabel
    
    lstAppList.SetFocus
End Sub

' =========================================================
' ���V���[�g�J�b�g�L�[�̐ݒ�_�C�A���O�ŃL�����Z���{�^�����������ꂽ�ꍇ�̃C�x���g
' =========================================================
Private Sub frmShortcutKeySettingVar_cancel()

    lstAppList.SetFocus
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

    ' �����̃V���[�g�J�b�g���폜����
    applicationSetting.clearShortcutKey
    
    ' �����Őݒ肳�ꂽ�V���[�g�J�b�g�����A�v���P�[�V�����I�u�W�F�N�g�ɐݒ�
    Dim shortCut     As ValShortcutKey
    Dim shortCutList As ValCollection
    Set shortCutList = appMenuList.collection

    Dim shortCutApp  As ValShortcutKey

    For Each shortCut In shortCutList.col
        
        Set shortCutApp = applicationSetting.shortcutAppList.getItem(shortCut.commandBarControl.Tag)
        If Not shortCutApp Is Nothing Then
            shortCutApp.shortcutKeyCode = shortCut.shortcutKeyCode
            shortCutApp.shortcutKeyLabel = shortCut.shortcutKeyLabel
        End If
    Next
    
    ' �o�^����
    applicationSetting.writeForDataShortcut
    
    ' �V���ɐݒ肳�ꂽ�V���[�g�J�b�g��o�^����
    applicationSetting.updateShortcutKey
    
End Sub

' =========================================================
' ���I�v�V��������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�FisResetShortcutKey �V���[�g�J�b�g�L�[�̃��Z�b�g�����{���邩�̃t���O
' �߂�l�@�@�F
'
' =========================================================
Private Sub loadShortcut(ByVal isResetShortcutKey As Boolean)

    ' �@�\���X�g�����Z�b�g����
    lstAppList.clear
    
    ' �@�\���X�g�̏�����
    Set appMenuList = New CntListBox: appMenuList.init lstAppList
    
    ' �V���[�g�J�b�g���X�g���擾����
    ' ��Clone���\�b�h���g�p���ď����R�s�[����B
    ' �@�����ł́AApplicationSetting#ShortcutAppList�Ɋi�[����Ă���ValShortCut�v�f�𒼐ڕύX������
    ' �@�N���[���𐶐����ҏW���s���B
    Dim shortCut     As ValShortcutKey
    Dim shortCutList As ValCollection
    Set shortCutList = applicationSetting.cloneShortcutAppList
    
    If isResetShortcutKey Then
        For Each shortCut In shortCutList.col
            shortCut.shortcutKeyCode = ""
            shortCut.shortcutKeyLabel = ""
        Next
    End If

    ' �@�\���X�g�ɔ��f����
    appMenuList.addAll shortCutList, "commandName", "shortcutKeyLabel"

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

    loadShortcut False
End Sub

' =========================================================
' ���I�v�V�������̃��Z�b�g
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub resetShortcut()

    loadShortcut True
End Sub

