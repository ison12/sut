VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryParameterSetting 
   Caption         =   "�N�G���p�����[�^�̕ҏW"
   ClientHeight    =   8445.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7095
   OleObjectBlob   =   "frmQueryParameterSetting.frx":0000
End
Attribute VB_Name = "frmQueryParameterSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' �N�G���p�����[�^�̈ꌏ���̕ҏW�i�q��ʁj
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
' �����@�@�@�FqueryParameter �N�G���p�����[�^���
'
' =========================================================
Public Event ok(ByVal queryParameter As ValQueryParameter)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' �N�G���p�����[�^���i�t�H�[���\�����_�ł̏��j
Private queryParameterParam As ValQueryParameter

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
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal queryParameter As ValQueryParameter)

    Set queryParameterParam = queryParameter
    
    activate
    
    ' �f�t�H���g�t�H�[�J�X�R���g���[����ݒ肷��
    txtParameter.SetFocus
    
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

    lblErrorMessage.Caption = ""
    
    txtParameter.value = queryParameterParam.name
    txtValue.value = queryParameterParam.value
    
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
    
    ' �L�[�R�[�h�𕪉����A�Ή�����ϐ��Ɋi�[����
    Dim queryParameter As New ValQueryParameter
    queryParameter.name = txtParameter.value
    queryParameter.value = txtValue.value

    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(queryParameter)
    
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
' ��DB�ڑ��ύX�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDBConnectedChange_Click()

    On Error GoTo err
    
    Main.disconnectDB
    
    ' DB�R�l�N�V�������擾����
    Dim dbConn As Object
    Set dbConn = Main.getDBConnection
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���e�X�g�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdTest_Click()

    On Error GoTo err
    
    Const MSG_TITLE As String = "SELECT�e�X�g����"
    
    ' DB�R�l�N�V�������擾����
    Dim dbConn As Object
    Set dbConn = Main.getDBConnection

    ' �����Ԃ̏��������s�����̂Ń}�E�X�J�[�\���������v�ɂ���
    Dim cursorWait As New ExcelCursorWait: cursorWait.init
    
    Dim resultSet   As Object
    Dim resultRecs  As Variant
    Dim resultVal   As Variant

    Set resultSet = ADOUtil.querySelect(dbConn, txtValue.text, 0)
    
    ' ------------------------------------------------
    ' �߂�l

    ' ���R�[�h�Z�b�g��EOF�ł͂Ȃ��ꍇ
    If Not resultSet.EOF Then
        ' ���R�[�h�Z�b�g����S���R�[�h���擾����
        resultRecs = resultSet.getRows(1)
        resultVal = resultRecs(0, 0)
    Else
        ' ���Ԃ�
        resultVal = Empty
    End If
    ' ------------------------------------------------
    
    ADOUtil.closeRecordSet resultSet

    ' �����Ԃ̏������I�������̂Ń}�E�X�J�[�\�������ɖ߂�
    cursorWait.destroy
    
    If isNull(resultVal) Then
        VBUtil.showMessageBoxForInformation "�擾�f�[�^�F" & "NULL", MSG_TITLE
    ElseIf VBUtil.arraySize(resultRecs) <= 0 Then
        VBUtil.showMessageBoxForInformation "�擾�f�[�^�F" & "NULL (�擾���R�[�h��0��)", MSG_TITLE
    Else
        VBUtil.showMessageBoxForInformation "�擾�f�[�^�F" & CStr(resultVal), MSG_TITLE
    End If
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���p�����[�^�����͎��̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtParameter_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:


    ' ���K�\���I�u�W�F�N�g�𐶐�
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        ' �����Ώە�����
        .Pattern = "^" & "([a-z0-9_-]|[^\u0000-\u007F])+" & "$"
        ' �啶�������������t���O
        .IgnoreCase = True
        ' ������S�̂��J��Ԃ���������t���O
        .Global = False
    End With


    ' �����͎�
    If txtParameter.text = "" Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtParameter
    
    ' �s���ȕ�������
    ElseIf reg.test(txtParameter.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NOT_ALPHA_NUM_MARK_FULL
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtParameter
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue txtParameter
    End If
    
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


