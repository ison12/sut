VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "������"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6690
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
   Tag             =   "168"
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �v���O���X�t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/01/21�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' �Z�J���_���v���O���X�̕\���L��
Private enableSecProgressParam As Boolean

' �Ώۃu�b�N
Private targetBook As Workbook
' �Ώۃu�b�N���擾����
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ���^�C�g���v���p�e�B
' =========================================================
Public Property Get title() As String
    title = lblTitle.Caption
End Property

Public Property Let title(ByVal vNewValue As String)
    lblTitle.Caption = vNewValue
End Property

' =========================================================
' ���v���C�}������
' =========================================================
Public Property Get priCount() As Long
    priCount = CLng(lblPriCount.Caption)
End Property

Public Property Let priCount(ByVal vNewValue As Long)
    lblPriCount.Caption = CStr(vNewValue)
    updatePrimaryProgressBar
End Property

' =========================================================
' ���v���C�}�����v����
' =========================================================
Public Property Get priCountOfAll() As Long
    priCountOfAll = CLng(lblPriCountOfAll.Caption)
End Property

Public Property Let priCountOfAll(ByVal vNewValue As Long)
    lblPriCountOfAll.Caption = CStr(vNewValue)
    updatePrimaryProgressBar
End Property

' =========================================================
' ���v���C�}�����b�Z�[�W
' =========================================================
Public Property Get priMessage() As String
    priMessage = lblPriMessage.Caption
End Property

Public Property Let priMessage(ByVal vNewValue As String)
    lblPriMessage.Caption = vNewValue
End Property

' =========================================================
' ���Z�J���_������
' =========================================================
Public Property Get secCount() As Long
    secCount = CLng(lblSecCount.Caption)
End Property

Public Property Let secCount(ByVal vNewValue As Long)
    lblSecCount.Caption = CStr(vNewValue)
    updateSecondaryProgressBar
End Property

' =========================================================
' ���Z�J���_�����v����
' =========================================================
Public Property Get secCountOfAll() As Long
    secCountOfAll = CLng(lblSecCountOfAll.Caption)
End Property

Public Property Let secCountOfAll(ByVal vNewValue As Long)
    lblSecCountOfAll.Caption = CStr(vNewValue)
    updateSecondaryProgressBar
End Property

' =========================================================
' ���Z�J���_�����b�Z�[�W
' =========================================================
Public Property Get secMessage() As String
    secMessage = lblSecMessage.Caption
End Property

Public Property Let secMessage(ByVal vNewValue As String)
    lblSecMessage.Caption = vNewValue
End Property

' =========================================================
' ���v���C�}�����̏�����
'
' �T�v�@�@�@�F
' �����@�@�@�Fall     ���v����
' �@�@�@�@�@�@message ���b�Z�[�W
' �߂�l�@�@�F
'
' =========================================================
Public Function initPri(ByVal all As Long, ByVal message As String)

    priCount = 0
    priCountOfAll = all
    lblPriMessage.Caption = message
    
End Function

' =========================================================
' ���Z�J���_�����̏�����
'
' �T�v�@�@�@�F
' �����@�@�@�Fall     ���v����
' �@�@�@�@�@�@message ���b�Z�[�W
' �߂�l�@�@�F
'
' =========================================================
Public Function initSec(ByVal all As Long, ByVal message As String)
    
    secCount = 0
    secCountOfAll = all
    lblSecMessage.Caption = message

End Function

' =========================================================
' ���v���C�}�����̌����X�V�i���ݒl��+1�J�E���g����j
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Function inclimentPri()
    
    priCount = priCount + 1

End Function

' =========================================================
' ���Z�J���_�����̌����X�V�i���ݒl��+1�J�E���g����j
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Function inclimentSec()
    
    secCount = secCount + 1

End Function

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal               ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@enableSecProgress   �Z�J���_���v���O���X�̕\���L��
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal enableSecProgress As Boolean)

    ' �p�����[�^�̐ݒ�
    enableSecProgressParam = enableSecProgress
    
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
' ���L�����Z���m�F���b�Z�[�W�̊m�F�_�C�A���O�̕\��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�_�C�A���O�ŉ������ꂽ�{�^��
'
' =========================================================
Private Function showCancelConfDialog() As Long

    showCancelConfDialog = VBUtil.showMessageBoxForYesNo("�L�����Z�����Ă���낵���ł����H", ConstantsCommon.APPLICATION_NAME)

End Function

' =========================================================
' ���t�H�[���A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    ' �����ʒu
    Dim frmAdjustValue    As Double: frmAdjustValue = 50        ' �Z�J���_���v���O���X��\�����Ȃ��ꍇ�̍����̒���

    If enableSecProgressParam = True Then
        ' �Z�J���_���v���O���X��\������ꍇ
        lblSecMessage.visible = True
        lblSecCount.visible = True
        lblSecCountSeparator.visible = True
        lblSecCountOfAll.visible = True
        lblSecProgressBg.visible = True
        lblSecProgressFg.visible = True
        
        ' �t�H�[���̈ʒu����
        frmProgress.Height = CLng(frmProgress.Tag)
    Else
        ' �Z�J���_���v���O���X��\������ꍇ���Ȃ��ꍇ
        lblSecMessage.visible = False
        lblSecCount.visible = False
        lblSecCountSeparator.visible = False
        lblSecCountOfAll.visible = False
        lblSecProgressBg.visible = False
        lblSecProgressFg.visible = False
        
        ' �t�H�[���̈ʒu����
        frmProgress.Height = CLng(frmProgress.Tag) - frmAdjustValue
    End If
    
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
' ���v���O���X�o�[�̍X�V
'
' �T�v�@�@�@�F
' �����@�@�@�FcntCurrent ���ݒl��\������R���g���[��
'     �@�@�@  cntAll     �ő�l��\������R���g���[��
'     �@�@�@  valCurrent ���ݒl
'     �@�@�@  valAll     �ő�l
' �߂�l�@�@�F
'
' =========================================================
Private Sub updateProgressBar(ByRef cntCurrent As MSForms.label _
                            , ByRef cntAll As MSForms.label _
                            , ByVal valCurrent As Long _
                            , ByVal valAll As Long)
                         
    If valAll <= 0 Then
        ' 0���Z���Ȃ��悤�Ƀ`�F�b�N����
        cntCurrent.Width = 0
        Exit Sub
    End If
                         
    cntCurrent.Width = CDbl(cntAll.Width) * (CDbl(valCurrent) / CDbl(valAll))
    
End Sub

' =========================================================
' ���v���C�}���v���O���X�o�[�̍X�V
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub updatePrimaryProgressBar()

    updateProgressBar lblPriProgressFg, lblPriProgressBg, CLng(lblPriCount.Caption), CLng(lblPriCountOfAll.Caption)
    
End Sub

' =========================================================
' ���Z�J���_���v���O���X�o�[�̍X�V
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub updateSecondaryProgressBar()

    updateProgressBar lblSecProgressFg, lblSecProgressBg, CLng(lblSecCount.Caption), CLng(lblSecCountOfAll.Caption)
    
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
    
        If showCancelConfDialog = 6 Then
        
            ' �L�����Z���C�x���g�𑗐M����
            RaiseEvent Cancel
            
        End If
        
        Cancel = True
        
    End If
    
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

