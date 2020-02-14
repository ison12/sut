VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBColumnFormatSetting 
   Caption         =   "DB�J���������ݒ�̕ҏW"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   OleObjectBlob   =   "frmDBColumnFormatSetting.frx":0000
End
Attribute VB_Name = "frmDBColumnFormatSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�J���������ݒ�̈ꌏ���̕ҏW�i�q��ʁj
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/08�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FdbColumnTypeColInfo DB�J���������ݒ�
'
' =========================================================
Public Event ok(ByVal dbColumnTypeColInfo As ValDbColumnTypeColInfo)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' DB�J���������ݒ���i�t�H�[���\�����_�ł̏��j
Private dbColumnTypeColInfoParam As ValDbColumnTypeColInfo

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
' �����@�@�@�Fmodal               ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@dbColumnTypeColInfo DB�J���������ݒ���
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal dbColumnTypeColInfo As ValDbColumnTypeColInfo)

    Set dbColumnTypeColInfoParam = dbColumnTypeColInfo
    
    activate
    
    ' �f�t�H���g�t�H�[�J�X�R���g���[����ݒ肷��
    txtColumnName.SetFocus
    
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

    txtColumnName.value = dbColumnTypeColInfoParam.columnName
    txtFormatUpdate.value = dbColumnTypeColInfoParam.formatUpdate
    txtFormatSelect.value = dbColumnTypeColInfoParam.formatSelect
    
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
    
    ' ���͏����A�Ή�����ϐ��Ɋi�[����
    Dim dbColumnTypeColInfo As New ValDbColumnTypeColInfo
    dbColumnTypeColInfo.columnName = txtColumnName.value
    dbColumnTypeColInfo.formatUpdate = txtFormatUpdate.value
    dbColumnTypeColInfo.formatSelect = txtFormatSelect.value

    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(dbColumnTypeColInfo)
    
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

Private Sub onPasteValue()

    Me.ActiveControl.text = "$value"

End Sub


