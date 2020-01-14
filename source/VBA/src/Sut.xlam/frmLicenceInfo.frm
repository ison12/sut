VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLicenceInfo 
   Caption         =   "���C�Z���X�ɂ���"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6585
   OleObjectBlob   =   "frmLicenceInfo.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "frmLicenceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���C�Z���X����\���i�o�^�j����t�H�[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2008/05/18�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' ��OK�{�^���������ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event ok()

' =========================================================
' ������{�^���������ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event closed()

' �G���[���b�Z�[�W
Private Const ERR_MSG_AUTH_FAILED               As String = "�F�؂Ɏ��s���܂����B"
' �C���t�H���b�Z�[�W
Private Const INFO_MSG_AUTH_SUCCESS             As String = "�F�؂���܂����B"
' �C���t�H���b�Z�[�W�i���p���ԓ��t�j
Private Const INFO_MSG_PROBATION_DAY            As String = "��A${date}���g�p�ł��܂��B�w������]�����ꍇ�́A���C�Z���X�o�^���s���Ă��������B�ڂ����̓}�j���A�����Q�Ƃ��Ă��������B"
' �C���t�H���b�Z�[�W�i���p���ԃI�[�o�[�j
Private Const INFO_MSG_OVER_PROBATION_DAY       As String = "���p���Ԃ��߂��Ă��܂��B�p�����ė��p����ꍇ�́A���C�Z���X�o�^�����肢���܂��B�ڂ����̓}�j���A�����Q�Ƃ��Ă��������B"
' �C���t�H���b�Z�[�W�i�F�؊����j
Private Const INFO_MSG_AUTH_LICENCE_COMPLETED   As String = "�F�؂��������Ă��܂��B"
' �C���t�H���b�Z�[�W�i���C�Z���X���͗��̃��b�Z�[�W�j
Private Const INFO_MSG_AUTH_LICENCE_AREA        As String = "���s�������[�UID�ƃ��C�Z���X�L�[����͂��Ĥ�F�؃{�^�����������Ă��������"
' �C���t�H���b�Z�[�W�i���C�Z���X���͗��̃��b�Z�[�W2�j
Private Const INFO_MSG_AUTH_LICENCE_AREA2       As String = "�{�\�t�g�E�F�A�͎��̕��Ƀ��C�Z���X����Ă��܂��B"

' ���C�Z���X�F�ؗL���t���O
Private m_authenticatedLicence As Boolean
' ���C�Z���X�F�؏��
Private m_licenceInfo As ValLicenceInfo

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal  ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@authenticatedLicence ���C�Z���X���F�ؗL���t���O
' �@�@�@�@�@�@licenceInfo ���C�Z���X���
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByVal authenticatedLicence As Boolean _
                 , ByRef licenceInfo As ValLicenceInfo)

    
    ' �����o��ݒ�
    m_authenticatedLicence = authenticatedLicence
    Set m_licenceInfo = licenceInfo
    
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
' ���t�H�[�����������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
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

    ' �őO�ʕ\���ɂ���
    ExcelUtil.setUserFormTopMost Me

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
    
    ' �R���g���[���̃`�F�b�N
    If validAllControl = False Then
    
        Exit Sub
    End If
    
    ' ----------------------------------------------------
    ' ���C�Z���X�F�؏���
    ' ----------------------------------------------------
    ' ���C�Z���X���i�ꎞ�I�ɐ����j
    Dim tmpLicenceInfo As New ValLicenceInfo
    ' ���C�Z���X���̐ݒ�
    tmpLicenceInfo.userId = txtUserId.value
    tmpLicenceInfo.password = txtPassword.value
    
    ' �F�؃I�u�W�F�N�g
    Dim author As New ExeAuthenticateLicence
    ' �F�؃I�u�W�F�N�g�Ƀ��C�Z���X����ݒ肷��
    author.init tmpLicenceInfo
    
    ' �F�؂����{����
    If author.executeAuthor = False Then
    
        ' ���s�����ꍇ
        lblErrorMessage.Caption = ERR_MSG_AUTH_FAILED
        Exit Sub
        
    End If
    ' ----------------------------------------------------
    
    ' �t�@�C���o�̓I�v�V��������������
    storeLicenceInfo

    ' �������b�Z�[�W��\��
    VBUtil.showMessageBoxForInformation INFO_MSG_AUTH_SUCCESS, ConstantsCommon.APPLICATION_NAME

    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub


' =========================================================
' ������{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdClose_Click()

    On Error GoTo err
    
    ' �t�H�[�������
    HideExt
    
    ' ����C�x���g�𑗐M����
    RaiseEvent closed

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

    ' �F�؏��I�u�W�F�N�g��j������
    Set m_licenceInfo = Nothing
    
End Sub

' =========================================================
' ���A�N�e�B�u���̏���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()
    
    ' �F�؍ς�
    If m_authenticatedLicence = True Then
    
        ' �F�؍ς݃��b�Z�[�W��ݒ肷��
        lblMessage.Caption = INFO_MSG_AUTH_LICENCE_COMPLETED
        ' �F�؍ς݃��b�Z�[�W�����C�Z���X���͗��ɐݒ肷��
        lblMessageAuthLicence.Caption = INFO_MSG_AUTH_LICENCE_AREA2
        
        ' �F�؍��ڂ��J��
        openItemOfLicence
        ' �F�؍��ڂ�F�؍ςݗp�ɕύX����
        changeItemOfLicenceAuthenticatedLicence
        
    ' ���F��
    Else
    
        ' ���F�؃��b�Z�[�W�����C�Z���X���͗��ɐݒ肷��
        lblMessageAuthLicence.Caption = INFO_MSG_AUTH_LICENCE_AREA
        
        ' �F�؃I�u�W�F�N�g
        Dim author As New ExeAuthenticateLicence
        
        ' �\�t�g�̎g�p�J�n��
        Dim fromDate As Date
        ' �g�p�J�n�����擾����
        fromDate = author.getProbationDate
    
        ' ���p���Ԃ͈͓̔�
        If author.isRangeProbation(fromDate) = True Then
        
            lblMessage.Caption = replace(INFO_MSG_PROBATION_DAY _
                                            , "${date}" _
                                            , author.getRemainderProbationDay(fromDate))
        
        
        ' ���p���Ԃ͈̔͊O
        Else
        
            lblMessage.Caption = INFO_MSG_OVER_PROBATION_DAY
        End If
        
        ' �F�؍��ڂ𖢔F�ؗp�ɕύX����
        changeItemOfLicenceNonAuthenticatedLicence
        
        ' ���C�Z���X�o�^�{�^�����[���I�ɉ�������
        tglRegistLicence_Click
    
    End If
    
    ' �G���[���b�Z�[�W����������
    lblErrorMessage.Caption = ""
    
    ' ���C�Z���X����ǂݍ���
    restoreLicenceInfo
    
End Sub

' =========================================================
' ���m���A�N�e�B�u���̏���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub deactivate()

End Sub

' =========================================================
' �����C�Z���X�F�؃g�O���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub tglRegistLicence_Click()

    If tglRegistLicence.value = True Then
    
        openItemOfLicence
        
    ElseIf tglRegistLicence.value = False Then
    
        closeItemOfLicence
    End If
    
End Sub

Private Sub openItemOfLicence()

    ' �֘A����R���g���[����\������
    fraLicenceInfo.visible = True
    cmdOk.visible = True
    
    cmdClose.Top = 183
    
    frmLicenceInfo.Height = 228.75
    
End Sub

Private Sub closeItemOfLicence()

    ' �֘A����R���g���[�����\���ɂ���
    fraLicenceInfo.visible = False
    cmdOk.visible = False
    
    cmdClose.Top = fraLicenceInfo.Top
    
    frmLicenceInfo.Height = 136.5
    
End Sub

Private Sub changeItemOfLicenceAuthenticatedLicence()

    tglRegistLicence.Enabled = False
    txtUserId.Locked = True
    txtPassword.Locked = True
    cmdOk.Enabled = False
    
End Sub

Private Sub changeItemOfLicenceNonAuthenticatedLicence()

    tglRegistLicence.Enabled = True
    txtUserId.Locked = False
    txtPassword.Locked = False
    cmdOk.Enabled = True

End Sub

' =========================================================
' ���S�R���g���[���̃`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�FTrue �`�F�b�NOK
'
' =========================================================
Private Function validAllControl() As Boolean

    ' �S�ẴR���g���[���̃`�F�b�N�����{����
    If validUserId = False Then
    
        validAllControl = False
        
    ElseIf validPassword = False Then
    
        validAllControl = False
        
    Else
    
        validAllControl = True
    
    End If
    
End Function

' =========================================================
' �����[�UID�̃`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�FTrue �`�F�b�NOK
'
' =========================================================
Private Function validUserId() As Boolean

    ' �K�{�`�F�b�N
    If txtUserId.value = "" Then
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse txtUserId
        
        validUserId = False
        
        txtUserId.SetFocus
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidTrue txtUserId
        
        validUserId = True
    
    End If

End Function

' =========================================================
' ���p�X���[�h�̃`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�FTrue �`�F�b�NOK
'
' =========================================================
Private Function validPassword() As Boolean

    ' �K�{�`�F�b�N
    If txtPassword.value = "" Then
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse txtPassword
        
        txtPassword.SetFocus
        validPassword = False
    
    ' 16�i���`�F�b�N
    ElseIf VBUtil.validHex(txtPassword.value) = False Then
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INVALID
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse txtPassword
        
        txtPassword.SetFocus
        validPassword = False
    
    ' �T�C�Y�`�F�b�N�i�T�C�Y��2�̔{���ł͂Ȃ��ꍇ�j
    ElseIf Len(txtPassword.value) Mod 2 <> 0 Then
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INVALID_SIZE
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse txtPassword

        txtPassword.SetFocus
        validPassword = False
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidTrue txtPassword
        
        validPassword = True
    
    End If
    

End Function

' =========================================================
' ���e�L�X�g�{�b�N�X�`�F�b�N�������̃R���g���[���ύX����
'
' �T�v�@�@�@�F
' �����@�@�@�Fcnt �R���g���[��
' �߂�l�@�@�F
'
' =========================================================
Public Sub changeControlPropertyByValidTrue(ByRef cnt As MSForms.control)

    With cnt
        .BackColor = &H80000005
        .ForeColor = &H80000012
    
    End With

End Sub

' =========================================================
' ���e�L�X�g�{�b�N�X�`�F�b�N���s���̃R���g���[���ύX����
'
' �T�v�@�@�@�F
' �����@�@�@�Fcnt �R���g���[��
' �߂�l�@�@�F
'
' =========================================================
Public Sub changeControlPropertyByValidFalse(ByRef cnt As MSForms.control)

    With cnt
        ' �e�L�X�g�S�̂�I������
        .SelStart = 0
        .SelLength = Len(.text)
        
        .BackColor = RGB(&HFF, &HFF, &HCC)
        .ForeColor = reverseRGB(&HFF, &HFF, &HCC)
        
    End With

End Sub

' =========================================================
' �����C�Z���X����ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeLicenceInfo()

    On Error GoTo err
    
    m_licenceInfo.userId = txtUserId.value
    m_licenceInfo.password = txtPassword.value
    
    ' ���W�X�g���ɏ���ۑ�����
    m_licenceInfo.writeForRegistry
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' �����C�Z���X����ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreLicenceInfo()

    On Error GoTo err
    
    ' ���W�X�g��������͎擾���Ȃ�
    ' �i���W�X�g������̏��ǂݍ��݂͑O�i�K�ōs���Ă��邽�߁A�����ł͍s��Ȃ��j
    txtUserId.value = m_licenceInfo.userId
    txtPassword.value = m_licenceInfo.password
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub


