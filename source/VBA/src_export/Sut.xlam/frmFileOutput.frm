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

Private Const REG_SUB_KEY_FILE_OUTPUT_OPTION As String = "file_output_option"

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
    
        VBUtil.showMessageBoxForWarning "�w�肳�ꂽ�t�@�C���p�X�Ƀt�@�C�����o�͂ł��܂���B" & vbNewLine & "�s���ȃp�X�A�܂��͌������s�����Ă���\��������܂��B" _
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
    
    cboChoiceCharacterCode.value = "shift_jis"
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
    txtFilePath.value = VBUtil.concatFilePath( _
                                    VBUtil.extractDirPathFromFilePath(txtFilePath.value) _
                                  , defaultFileName)
    
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
    If txtFilePath.value = "" Then txtFilePath.value = ThisWorkbook.path
    
    cboChoiceCharacterCode.value = retChar
    If cboChoiceCharacterCode.ListIndex = -1 Then cboChoiceCharacterCode.ListIndex = 0
    
    cboChoiceNewLine.value = retNewLine
    If cboChoiceNewLine.ListIndex = -1 Then cboChoiceNewLine.ListIndex = 0
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub
