VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileOutput 
   Caption         =   "�t�@�C���o��"
   ClientHeight    =   3630
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
' �쐬�ҁ@�FIson
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
Public Event cancel()

' �����R�[�h���X�g
Private charcterList As CntListBox

' �f�t�H���g�t�@�C����
Private defaultFileName As String

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
    ' �t�H���_�쐬�̐����L��
    Dim isSuccessCreateDir As Boolean

    ' �t�@�C���p�X���擾
    filePath = txtFilePath.text
    ' �����R�[�h���擾
    characterCode = cboChoiceCharacterCode.text
    ' ���s�R�[�h���擾
    newline = cboChoiceNewLine.text
    
    If VBUtil.isExistDirectory(filePath) Then
        ' �t�@�C���p�X���f�B���N�g���̏ꍇ�̓G���[�Ƃ���
        VBUtil.showMessageBoxForWarning "�t�H���_���w�肳��Ă��܂��B�t�@�C���p�X���w�肵�Ă��������B" _
                                      , ConstantsCommon.APPLICATION_NAME _
                                      , Nothing

        Exit Sub
    End If
    
    ' �t�@�C���p�X�̐e�f�B���N�g�����擾����
    dirPath = VBUtil.extractDirPathFromFilePath(filePath)

    ' --------------------------------------
    ' �e�t�H���_���쐬����
    On Error Resume Next
    
    isSuccessCreateDir = False
    
    VBUtil.createDir dirPath
    If err.Number = 0 Then
        ' �쐬�ɐ���
        isSuccessCreateDir = True
    End If
    
    On Error GoTo err
    ' --------------------------------------

    ' �t�H���_�ւ̃e�X�g�o�͂Ɏ��s�����ꍇ
    If isSuccessCreateDir = False Or VBUtil.touch(dirPath) = False Then
    
        VBUtil.showMessageBoxForWarning "�w�肳�ꂽ�t�@�C���p�X�Ƀt�@�C�����o�͂ł��܂���B" & vbNewLine & "�����́A�s���ȃp�X�A�܂��͌������s�����Ă���\��������܂��B" _
                                      , ConstantsCommon.APPLICATION_NAME _
                                      , Nothing
        
        Exit Sub
    End If
    
    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(filePath, characterCode, VBUtil.convertNewLineStrToNewLineCode(cboChoiceNewLine.text))
    
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
    
    cboChoiceCharacterCode.value = "Shift_JIS"
    cboChoiceNewLine.ListIndex = 0
    
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

    ' �t�@�C���o�̓I�v�V������ǂݍ���
    restoreFileOutputOption
    
    ' �t�@�C���p�X�Ƀf�t�H���g�̃t�@�C������ݒ肷��
    If txtFilePath.value = "" Then
        txtFilePath.value = VBUtil.concatFilePath( _
                                        VBUtil.extractDirPathFromFilePath(targetBook.path) _
                                      , defaultFileName)
    Else
        txtFilePath.value = VBUtil.concatFilePath( _
                                        VBUtil.extractDirPathFromFilePath(txtFilePath.value) _
                                      , defaultFileName)
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
' ���ݒ���̐���
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & Me.name & ".ini"
    appProp.initWorksheet targetBook, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & Me.name & ".ini"

    Set createApplicationProperties = appProp
    
End Function

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
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' �������݃f�[�^
    Dim values As New ValCollection
    
    values.setItem Array(txtFilePath.name, VBUtil.extractDirPathFromFilePath(txtFilePath.value))
    values.setItem Array(cboChoiceCharacterCode.name, cboChoiceCharacterCode.value)
    values.setItem Array(cboChoiceNewLine.name, cboChoiceNewLine.value)

    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

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
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties

    ' �f�[�^��ǂݍ���
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
            
    val = values.getItem(txtFilePath.name, vbVariant): If IsArray(val) Then txtFilePath.value = val(2)
    val = values.getItem(cboChoiceCharacterCode.name, vbVariant): If IsArray(val) Then cboChoiceCharacterCode.value = val(2)
    val = values.getItem(cboChoiceNewLine.name, vbVariant): If IsArray(val) Then cboChoiceNewLine.value = val(2)
    
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
Private Sub changeControlPropertyByValidTrue(ByRef cnt As MSForms.control)

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
Private Sub changeControlPropertyByValidFalse(ByRef cnt As MSForms.control)

    With cnt
        ' �e�L�X�g�S�̂�I������
        .SelStart = 0
        .SelLength = Len(.text)
        
        .BackColor = RGB(&HFF, &HFF, &HCC)
        .ForeColor = reverseRGB(&HFF, &HFF, &HCC)
        
    End With

End Sub

