VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetUpdate 
   Caption         =   "�e�[�u���V�[�g�X�V"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5940
   OleObjectBlob   =   "frmTableSheetUpdate.frx":0000
End
Attribute VB_Name = "frmTableSheetUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �e�[�u���V�[�g�X�V�t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/03�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����������������ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FrecFormat �s�t�H�[�}�b�g
'
' =========================================================
Public Event ok(ByVal recFormat As REC_FORMAT)

' =========================================================
' ���������L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event cancel()

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants)

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
' ���A�N�e�B�u���̏���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    ' �s�t�H�[�}�b�g
    Dim recFormat As REC_FORMAT
    
    ' �A�N�e�B�u�ȃe�[�u���V�[�g�̍s�t�H�[�}�b�g���擾��
    ' �I�v�V�����{�^���ɔ��f����
        
    ' �e�[�u���V�[�g�Ǎ��I�u�W�F�N�g
    Dim tableSheetReader As ExeTableSheetReader
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.sheet = ActiveSheet
        
    recFormat = tableSheetReader.getRowFormat
    
    If recFormat = REC_FORMAT.recFormatToUnder Then
    
        optRowFormatToUnder.value = True
    
    Else
    
        optRowFormatToRight.value = True
    End If
    
        
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
' ��OK�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdOk_Click()

    ' �s�t�H�[�}�b�g�萔�𗘗p���邽�߂ɁAValTable�𐶐�����
    Dim table As ValDbDefineTable
    ' �s�t�H�[�}�b�g
    Dim recFormat As REC_FORMAT
    
    ' �I�v�V�����{�^���őI������Ă���l��
    ' Long�^�̍s�t�H�[�}�b�g�萔�ɕϊ�����B
    If optRowFormatToUnder.value = True Then
    
        ' �s�t�H�[�}�b�gX�̃��W�I�{�^�����I������Ă���ꍇ
        recFormat = REC_FORMAT.recFormatToUnder
    
    Else
    
        ' �s�t�H�[�}�b�gY�̃��W�I�{�^�����I������Ă���ꍇ
        recFormat = REC_FORMAT.recFormatToRight
        
    End If
    
    RaiseEvent ok(recFormat)
    HideExt
End Sub

' =========================================================
' ���L�����Z���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdCancel_Click()

    RaiseEvent cancel
    HideExt
End Sub
