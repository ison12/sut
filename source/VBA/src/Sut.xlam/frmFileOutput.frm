VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileOutput 
   Caption         =   "�t�@�C���o��"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7815
   OleObjectBlob   =   "frmFileOutput.frx":0000
End
Attribute VB_Name = "frmFileOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �t�@�C���o�͂��s���t�H�[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2008/09/06�@�V�K�쐬
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
Public Event ok(ByVal filePath As String _
              , ByVal characterCode As String _
              , ByVal newline As String)

' =========================================================
' ���L�����Z���{�^���������ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

Private Const NEW_LINE_STR_CRLF As String = "CRLF"
Private Const NEW_LINE_STR_CR As String = "CR"
Private Const NEW_LINE_STR_LF As String = "LF"

Private Const REG_SUB_KEY_FILE_OUTPUT_OPTION As String = "file_output_option"

' ���W�X�g���p�X - �����R�[�h�ꗗ
Private Const REG_PATH_CHARACTER_CODE_LIST As String = "MIME\Database\Charset"
' ���W�X�g���L�[ - �����R�[�h�̕ʖ�
Private Const REG_KEY_ALIAS_CHARSET As String = "AliasForCharset"

' �����R�[�h���X�g
Private charcterList As CntListBox

' �f�t�H���g�t�@�C����
Private defaultFileName As String

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal  ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@header �w�b�_�e�L�X�g
' �@�@�@�@�@�@defFileName �f�t�H���g�t�@�C����
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByVal header As String _
                 , ByVal defFileName As String)

    lblHeader.Caption = header
    defaultFileName = defFileName

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
' �������R�[�h���X�g�@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboChoiceCharacterCode_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' �R���N�V����
    Dim col As ValCollection
    ' �R���g���[������R���N�V�������擾����
    Set col = charcterList.collection

    ' ���X�g�Ɍ��ݓ��͂���Ă���e�L�X�g�̗v�f�����݂��Ȃ��ꍇ
    If col.exist(cboChoiceCharacterCode.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NO_LIST_ITEM
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse cboChoiceCharacterCode
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidTrue cboChoiceCharacterCode
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

    
End Sub

' =========================================================
' �����s�R�[�h���X�g�@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboChoiceNewLine_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' �R���N�V����
    Dim col As New ValCollection

    col.setItem NEW_LINE_STR_CRLF, NEW_LINE_STR_CRLF
    col.setItem NEW_LINE_STR_CR, NEW_LINE_STR_CR
    col.setItem NEW_LINE_STR_LF, NEW_LINE_STR_LF
    
    ' ���X�g�Ɍ��ݓ��͂���Ă���e�L�X�g�̗v�f�����݂��Ȃ��ꍇ
    If col.exist(cboChoiceNewLine.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NO_LIST_ITEM
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse cboChoiceNewLine
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidTrue cboChoiceNewLine
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

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
' ���t�@�C���I���{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub btnFileSelect_Click()

    Dim selectFile As String
    
    selectFile = saveFileDialog
    
    If selectFile <> "" Then
        ' �t�@�C�����J���_�C�A���O���I�[�v�����ă��[�U�Ƀt�@�C����I��������
        txtFilePath.text = selectFile
    End If
    
End Sub

' =========================================================
' ���t�@�C�����J���_�C�A���O�I�[�v��
'
' �T�v�@�@�@�F�t�@�C�����J���_�C�A���O���I�[�v������
'
' =========================================================
Private Function saveFileDialog() As String

    On Error GoTo err
        
    ' �I���t�@�C��
    Dim selectFile As String
    
    ' �J���_�C�A���O��I������
    selectFile = VBUtil.openFileSaveDialog("�ۑ��t�@�C����I�����Ă��������B" _
                                         , "SQL�t�@�C�� (*.sql),*.sql,���ׂẴt�@�C�� (*.*),*.*" _
                                         , VBUtil.extractFileName(txtFilePath.value))

    ' �t�@�C���p�X��ݒ肷��
    saveFileDialog = selectFile
    
    Exit Function
    
err:

End Function

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
    
    ' �t�@�C���p�X
    Dim filePath As String
    ' �f�B���N�g���p�X
    Dim dirPath As String
    ' �����R�[�h
    Dim characterCode As String
    ' ���s�R�[�h
    Dim newline As String

    ' �t�@�C���p�X���擾
    filePath = txtFilePath.text
    ' �����R�[�h���擾
    characterCode = cboChoiceCharacterCode.text
    ' ���s�R�[�h���擾
    newline = cboChoiceNewLine.text
    
    ' �t�@�C���p�X�̐e�f�B���N�g�����擾����
    dirPath = VBUtil.extractDirPathFromFilePath(filePath)
    ' �e�t�H���_���쐬����
    VBUtil.createDir dirPath
    
    ' �t�@�C���p�X�̃f�B���N�g�������݂��邩���m�F����
    If VBUtil.isExistDirectory(dirPath) = False Then
    
        VBUtil.showMessageBoxForWarning "�w�肳�ꂽ�t�@�C���p�X�̃t�H���_��������܂���B" _
                                      , ConstantsCommon.APPLICATION_NAME _
                                      , Nothing
        
        Exit Sub
    End If
    
    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(filePath, characterCode, convertNewLineStrToNewLineCode(cboChoiceNewLine.text))
    
    ' �t�@�C���o�̓I�v�V��������������
    storeFileOutputOption

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

    ' �����R�[�h���X�g�̃R���g���[���I�u�W�F�N�g������������
    Set charcterList = New CntListBox
    
    charcterList.init cboChoiceCharacterCode
    charcterList.addAll VBUtil.getEncodeList
    
    ' ���s�R�[�h���X�g�ɉ��s�R�[�h��ǉ�����
    Dim newLineList As ValCollection
    Set newLineList = VBUtil.getNewlineList
    
    Dim var As Variant
    
    For Each var In newLineList.col
    
        cboChoiceNewLine.addItem var
    Next
    
    cboChoiceCharacterCode.value = "shift_jis"
    cboChoiceNewLine.ListIndex = 0

    ' �t�@�C���o�̓I�v�V������ǂݍ���
    restoreFileOutputOption
    
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
    
    ' �t�@�C���p�X�Ƀf�t�H���g�̃t�@�C������ݒ肷��
    txtFilePath.value = VBUtil.concatFilePath( _
                                    VBUtil.extractDirPathFromFilePath(txtFilePath.value) _
                                  , defaultFileName)
    
    
    ' �G���[���b�Z�[�W���N���A����
    lblErrorMessage.Caption = ""
    
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

Private Function convertNewLineStrToNewLineCode(ByVal newLineStr As String) As String

    If newLineStr = NEW_LINE_STR_CRLF Then
    
        ' Windows
        convertNewLineStrToNewLineCode = vbCr & vbLf
    
    ElseIf newLineStr = NEW_LINE_STR_CR Then
    
        ' Mac
        convertNewLineStrToNewLineCode = vbCr
    
    ElseIf newLineStr = NEW_LINE_STR_LF Then
    
        ' Unix
        convertNewLineStrToNewLineCode = vbLf
        
    ' ���Ă͂܂�Ȃ��ꍇ
    Else
    
        ' Windows
        convertNewLineStrToNewLineCode = vbCr & vbLf
    
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
' ���t�@�C���I�v�V������ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeFileOutputOption()

    On Error GoTo err
    
    Dim j As Long
    
    Dim fileOutputOption(0 To 2 _
                       , 0 To 1) As Variant
    
    
    fileOutputOption(j, 0) = txtFilePath.name
    fileOutputOption(j, 1) = VBUtil.extractDirPathFromFilePath(txtFilePath.value): j = j + 1
    
    fileOutputOption(j, 0) = cboChoiceCharacterCode.name
    fileOutputOption(j, 1) = cboChoiceCharacterCode.value: j = j + 1

    fileOutputOption(j, 0) = cboChoiceNewLine.name
    fileOutputOption(j, 1) = cboChoiceNewLine.value: j = j + 1
    
    ' ���W�X�g������N���X
    Dim registry As New RegistryManipulator
    ' ���W�X�g������N���X������������
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_FILE_OUTPUT_OPTION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' ���W�X�g���ɏ���ݒ肷��
    registry.setValues fileOutputOption
    
    Set registry = Nothing
        
    ' ----------------------------------------------
    ' �u�b�N�ݒ���
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet
    
    bookProp.setValue ConstantsBookProperties.TABLE_FILE_OUTPUT_DIALOG, txtFilePath.name, VBUtil.extractDirPathFromFilePath(txtFilePath.value)
    bookProp.setValue ConstantsBookProperties.TABLE_FILE_OUTPUT_DIALOG, cboChoiceCharacterCode.name, cboChoiceCharacterCode.value
    bookProp.setValue ConstantsBookProperties.TABLE_FILE_OUTPUT_DIALOG, cboChoiceNewLine.name, cboChoiceNewLine.value
    ' ----------------------------------------------

    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���t�@�C���I�v�V������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreFileOutputOption()

    On Error GoTo err
    
    ' ----------------------------------------------
    ' �u�b�N�ݒ���
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' �ݒ���V�[�g�����݂���
        
        Set bookPropVal = bookProp.getValues(ConstantsBookProperties.TABLE_FILE_OUTPUT_DIALOG)
        If bookPropVal.count > 0 Then
            ' �ݒ��񂪑��݂���̂ŁA�t�H�[���ɔ��f����
            
            txtFilePath.value = bookPropVal.getItem(txtFilePath.name, vbString)
            cboChoiceCharacterCode.value = bookPropVal.getItem(cboChoiceCharacterCode.name, vbString)
            cboChoiceNewLine.value = bookPropVal.getItem(cboChoiceNewLine.name, vbString)

            Exit Sub
        End If
    End If
    ' ----------------------------------------------

    ' ���W�X�g������N���X
    Dim registry As New RegistryManipulator
    ' ���W�X�g������N���X������������
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_FILE_OUTPUT_OPTION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Dim retFilepath As String
    Dim retChar     As String
    Dim retNewLine  As String
    
    registry.getValue txtFilePath.name, retFilepath
    registry.getValue cboChoiceCharacterCode.name, retChar
    registry.getValue cboChoiceNewLine.name, retNewLine
    
    txtFilePath.value = retFilepath
    cboChoiceCharacterCode.value = retChar
    cboChoiceNewLine.value = retNewLine
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub
