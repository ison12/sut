VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShortcutKeySetting 
   Caption         =   "�L�[�ݒ�"
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
' �V���[�g�J�b�g�L�[�̐ݒ�i�q��ʁj
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
Public Event ok(ByVal KeyCode As String, ByVal keyLabel As String)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event cancel()

' �V���[�g�J�b�g�L�[���X�g
Private shortcutKeyList As CntListBox

' �L�[�R�[�h�i�t�H�[���\�����_�ł̃L�[�R�[�h�j
Private keyCodeBefore As String

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
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal KeyCode As String)

    keyCodeBefore = KeyCode
    
    activate
    
    ' �f�t�H���g�t�H�[�J�X�R���g���[����ݒ肷��
    cboKey.SetFocus
    
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

    ' �L�[�R�[�h�𕪉�������Ɋi�[����ϐ�
    Dim shiftCtrl  As Boolean
    Dim shiftShift As Boolean
    Dim shiftAlt   As Boolean
    Dim keyName    As String
    
    ' �L�[�R�[�h�𕪉����A�Ή�����ϐ��Ɋi�[����
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
' ���t�H�[���f�B�A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub deactivate()
    
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
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' �{�����ł͏������̂��L�����Z������
        cancel = True
        ' �ȉ��̃C�x���g�o�R�ŕ���
        cmdCancel_Click
    End If
    
End Sub

' =========================================================
' ���폜�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDelete_Click()

    chbCtrl.value = False
    chbShift.value = False
    chbAlt.value = False
    cboKey.ListIndex = -1

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
    
    ' �t�H�[�������
    HideExt
    
    ' �L�[�R�[�h�𕪉����A�Ή�����ϐ��Ɋi�[����
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

    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(KeyCode, keyName)
    
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
    RaiseEvent cancel
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
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

    Set shortcutKeyList = New CntListBox: shortcutKeyList.init cboKey
    shortcutKeyList.addAll VBUtil.getAppOnKeyCodeList

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

    Set shortcutKeyList = Nothing
    
End Sub
