VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecordAppender 
   Caption         =   "�s�̒ǉ��E�폜"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4620
   OleObjectBlob   =   "frmRecordAppender.frx":0000
End
Attribute VB_Name = "frmRecordAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' ���[�N�V�[�g�̍s����ύX����t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/11/15�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' ��OK�{�^�������C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event ok(ByVal recCount As Long)

' =========================================================
' ���L�����Z���{�^�������C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' �����Ώۃ��[�N�V�[�g
Public sheet As Worksheet
' �A�v���P�[�V�����ݒ���
Private applicationSetting As ValApplicationSetting

' �����Ώۃe�[�u���I�u�W�F�N�g
Public tableSheet As ValTableWorksheet
' �����̍s��
Private recCountOrign As Long

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
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef sheet As Worksheet, ByRef aps As ValApplicationSetting)

    Set Me.sheet = sheet
    Set applicationSetting = aps
    
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
    
    ' �t�H�[�������
    HideExt
    
    ' �e�[�u���V�[�g�����I�u�W�F�N�g
    Dim tableSheetSheetCreator As New ExeTableSheetCreator
    tableSheetSheetCreator.book = sheet.parent
    tableSheetSheetCreator.applicationSetting = applicationSetting
    
    ' �ύX��̍s��
    Dim recCount    As Long
    ' �����̍s���ƕύX��̍s���̍�
    Dim recCountDiff As Long
    
    ' �e�L�X�g�{�b�N�X����ύX��̍s�����擾����
    recCount = txtRecCount.value
    ' �����̍s���ƕύX��̍s���̍������擾����
    recCountDiff = recCount - recCountOrign
    
    Dim recStart As Long
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        recStart = ConstantsTable.U_RECORD_OFFSET_ROW
    Else
    
        recStart = ConstantsTable.R_RECORD_OFFSET_COL
    End If
    
    
    ' �s�̍폜
    If recCountDiff < 0 Then
    
        tableSheetSheetCreator.deleteCellOfRecord tableSheet, recStart + recCount
    
    ' �s�̒ǉ�
    ElseIf recCountDiff > 0 Then
    
        tableSheetSheetCreator.insertEmptyCell tableSheet, recStart + recCountOrign, recCountDiff
            
    ' �������Ȃ�
    Else
    
    
    End If
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(recCount)
    
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

' =========================================================
' ���A�N�e�B�u���̏���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    ' ----------------------------------------------
    ' �e�[�u���V�[�g�����x�e�[�u������ǂݍ���
    Dim srctableSheet As ValTableWorksheet
    
    Dim tableSheetReader As ExeTableSheetReader
    
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.sheet = ActiveSheet
    
    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
    
    Set srctableSheet = tableSheetReader.readTableInfo
    ' ----------------------------------------------
    
    ' ���R�[�h����
    Dim recCount As Long
    ' ���R�[�h�����̎擾
    recCount = tableSheetReader.getRecordSize(srctableSheet)
    
    ' �e�L�X�g�{�b�N�X�Ƀ��R�[�h������ݒ肷��
    txtRecCount.value = recCount
    
    ' �e�L�X�g�{�b�N�X�Ƀt�H�[�J�X��^���S�I����Ԃɂ���
    txtRecCount.SetFocus
    txtRecCount.SelStart = 0
    txtRecCount.SelLength = Len(txtRecCount)
    
    ' �����Ώۃe�[�u���I�u�W�F�N�g��ݒ�
    Set tableSheet = srctableSheet
    ' �����̃��R�[�h����
    recCountOrign = recCount
    
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
' ���s���e�L�X�g�{�b�N�X�̃`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtRecCount_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' ��̏ꍇ�A�G���[
    If txtRecCount.text = "" Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        ' �A���[�g��\������
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        changeControlPropertyByValidTrue txtRecCount

    ' �e�L�X�g�{�b�N�X�̒l�����������`�F�b�N����
    ElseIf validInteger(txtRecCount.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        ' �A���[�g��\������
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
        
        changeControlPropertyByValidFalse txtRecCount
    
    ' ���l�͈̓`�F�b�N
    ElseIf CDec(txtRecCount.text) < 1 Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse txtRecCount
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecCount
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

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

