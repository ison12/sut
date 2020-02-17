VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBQueryBatch 
   Caption         =   "�N�G���ꊇ���s"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13935
   OleObjectBlob   =   "frmDBQueryBatch.frx":0000
End
Attribute VB_Name = "frmDBQueryBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' �N�G���ꊇ���s�t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/01/18�@�V�K�쐬
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
Public Event ok(ByVal dbQueryBatchMode As DB_QUERY_BATCH_MODE _
              , ByVal filePath As String _
              , ByVal characterCode As String _
              , ByVal newline As String _
              , ByVal tableWorksheets As ValCollection)

' =========================================================
' ���L�����Z���{�^���������ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event cancel()

' DB�N�G���o�b�`���[�h
Public Enum DB_QUERY_BATCH_MODE

    ' �t�@�C���o��
    FileOutput
    ' �N�G�����s
    QueryExecute

End Enum

' DB�N�G���o�b�`�̃N�G����ނ̈ꌏ���̕ҏW�i�q��ʁj
Private WithEvents frmDBQueryBatchTypeSettingVar As frmDBQueryBatchTypeSetting
Attribute frmDBQueryBatchTypeSettingVar.VB_VarHelpID = -1

' �e�[�u�����X�g�ł̑I�����ڃC���f�b�N�X
Private tableSheetSelectedIndex As Long
' �e�[�u�����X�g�ł̑I�����ڃI�u�W�F�N�g
Private tableSheetSelectedItem As ValDbQueryBatchTableWorksheet

' DB�N�G���o�b�`���[�h
Private dbQueryBatchMode As DB_QUERY_BATCH_MODE
' DB�N�G���o�b�`���
Private dbQueryBatchType As DB_QUERY_BATCH_TYPE
' �����Ώۃ��[�N�u�b�N
Private book As Workbook

' �����R�[�h���X�g
Private charcterList As CntListBox
' DB�N�G���o�b�`��ޕύX�R���{�{�b�N�X���X�g
Private dbQueryBatchTypeChangeAll As CntListBox
' DB�N�G���o�b�`��ޕύX�R���{�{�b�N�X�̏�����
Private inProcessDbQueryBatchTypeChangeAll As Boolean

' �e�[�u�����X�g
Private tableSheetList  As CntListBox

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
' �@�@�@�@�@�@mode   ���[�h
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByVal dbQueryBatchMode_ As DB_QUERY_BATCH_MODE _
                 , ByVal dbQueryBatchType_ As DB_QUERY_BATCH_TYPE _
                 , ByRef book_ As Workbook)

    dbQueryBatchMode = dbQueryBatchMode_
    dbQueryBatchType = dbQueryBatchType_
    Set book = book_

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
' ���S�Ă̑I������I���ς݂ɂ���{�^���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdSelectedAll_Click()

    tableSheetList.setSelectedAll True

End Sub

' =========================================================
' ���S�Ă̑I������I�������ɂ���{�^���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdUnselectedAll_Click()

    tableSheetList.setSelectedAll False

End Sub

' =========================================================
' ���S�Ă�DB�N�G���o�b�`��ނ�ύX����R���{�{�b�N�X���X�g�̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboDbQueryBatchTypeChangeAll_Change()

    On Error GoTo err

    If inProcessDbQueryBatchTypeChangeAll = True Then
        ' ���ɏ������̏ꍇ�͏������I������
        Exit Sub
    End If

    inProcessDbQueryBatchTypeChangeAll = True

    Dim i As Long
    Dim var As ValDbQueryBatchTableWorksheet
    
    Dim selectedDbQueryBatchType As ValDbQueryBatchType
    
    i = 0
    For Each var In tableSheetList.collection.col
    
        Set selectedDbQueryBatchType = dbQueryBatchTypeChangeAll.getSelectedItem
        var.dbQueryBatchType = selectedDbQueryBatchType.dbQueryBatchType
        
        setTableSheet i, var
        
        i = i + 1
    
    Next
    
    ' �����̍Ō�ɖ��I����Ԃɖ߂�
    dbQueryBatchTypeChangeAll.setSelectedIndex 0

    inProcessDbQueryBatchTypeChangeAll = False
    
    Exit Sub
err:

    inProcessDbQueryBatchTypeChangeAll = False
    
End Sub

' =========================================================
' ��DB�N�G���o�b�`��ނ�ύX����{�^���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmbDbQueryBatchTypeChange_Click()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    tableSheetSelectedIndex = tableSheetList.getSelectedIndex

    ' ���I���̏ꍇ
    If tableSheetSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set tableSheetSelectedItem = tableSheetList.getSelectedItem

    If VBUtil.unloadFormIfChangeActiveBook(frmDBQueryBatchTypeSetting) Then Unload frmDBQueryBatchTypeSetting
    Load frmDBQueryBatchTypeSetting
    Set frmDBQueryBatchTypeSettingVar = frmDBQueryBatchTypeSetting
    
    frmDBQueryBatchTypeSettingVar.ShowExt vbModal _
                        , tableSheetSelectedItem.sheetNameOrSheetTableName _
                        , tableSheetSelectedItem.dbQueryBatchType _
                        , dbQueryBatchTypeChangeAll.collection
    
    Set frmDBQueryBatchTypeSettingVar = Nothing
    
    tableSheetList.control.SetFocus

End Sub

' =========================================================
' ��DB�N�G���o�b�`��ނ�ύX�̊m�莞�̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmDBQueryBatchTypeSettingVar_ok(ByVal dbQueryBatchType As DB_QUERY_BATCH_TYPE)

    tableSheetSelectedItem.dbQueryBatchType = dbQueryBatchType
    
    setTableSheet tableSheetSelectedIndex, tableSheetSelectedItem
    
End Sub

' =========================================================
' ��DB�N�G���o�b�`��ނ�ύX�̃L�����Z�����̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmDBQueryBatchTypeSettingVar_cancel()

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
    
    selectFile = openFolderDialog
    
    If selectFile <> "" Then
        ' �t�@�C�����J���_�C�A���O���I�[�v�����ă��[�U�Ƀt�@�C����I��������
        txtFilePath.text = selectFile
    End If
    
End Sub

' =========================================================
' ���t�H���_���J���_�C�A���O�I�[�v��
'
' �T�v�@�@�@�F�t�H���_���J���_�C�A���O���I�[�v������
'
' =========================================================
Private Function openFolderDialog() As String

    On Error GoTo err
            ' �I���t�@�C��
    Dim selectFile As String
    
    ' �J���_�C�A���O��I������
    selectFile = VBUtil.openFolderDialog("�t�@�C���o�͐�t�H���_��I�����Ă��������B" _
                                         , txtFilePath.value)

    ' �t�@�C���p�X��ݒ肷��
    openFolderDialog = selectFile
    
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
    
    Dim tableSelectedList As ValCollection
    Set tableSelectedList = tableSheetList.getSelectedList
    
    If tableSelectedList.count <= 0 Then
        err.Raise ERR_NUMBER_NOT_SELECTED_TABLE_SHEET _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_TABLE_SHEET _
                , err.HelpFile _
                , err.HelpContext
    End If

    ' �t�@�C���o�͎��݂̂̏���
    If dbQueryBatchMode = FileOutput Then
        ' �t�@�C���p�X���擾
        filePath = txtFilePath.text
        ' �����R�[�h���擾
        characterCode = cboChoiceCharacterCode.text
        ' ���s�R�[�h���擾
        newline = cboChoiceNewLine.text
        
        ' �t�@�C���p�X�̐e�f�B���N�g�����擾����
        dirPath = filePath
        
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
        If isSuccessCreateDir = False Or VBUtil.touch(filePath) = False Then
        
            VBUtil.showMessageBoxForWarning "�w�肳�ꂽ�t�H���_�p�X�Ƀt�@�C�����o�͂ł��܂���B" & vbNewLine & "�s���ȃp�X�A�܂��͌������s�����Ă���\��������܂��B" _
                                          , ConstantsCommon.APPLICATION_NAME _
                                          , Nothing
            
            Exit Sub
        End If
        
    End If
    
    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(dbQueryBatchMode, filePath, characterCode, VBUtil.convertNewLineStrToNewLineCode(cboChoiceNewLine.text), tableSelectedList)
    
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

    ' �R���g���[���̏�Ԃ𐧌䂷��
    If dbQueryBatchMode = FileOutput Then
        ' �t�@�C���o�͎�
        lblFilePath.visible = True
        txtFilePath.visible = True
        lblChoiceCharacterCode.visible = True
        cboChoiceCharacterCode.visible = True
        lblChoiceNewLine.visible = True
        cboChoiceNewLine.visible = True
        btnFileSelect.visible = True
    Else
        ' DB���s��
        lblFilePath.visible = False
        txtFilePath.visible = False
        lblChoiceCharacterCode.visible = False
        cboChoiceCharacterCode.visible = False
        lblChoiceNewLine.visible = False
        cboChoiceNewLine.visible = False
        btnFileSelect.visible = False
    End If
    
    ' DB�o�b�`�N�G����ރ��X�g�ɑI������ǉ�����
    Set dbQueryBatchTypeChangeAll = New CntListBox
    dbQueryBatchTypeChangeAll.init cboDbQueryBatchTypeChangeAll
    
    Dim dbBatchQueryTypeRawList As New ValCollection
    Dim dbBatchQueryType As ValDbQueryBatchType
    
    If dbQueryBatchMode = FileOutput Then
        ' �t�@�C���o��
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.none: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.insert: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.update: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.deleteOnSheet: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
    
    Else
        ' �N�G�����s
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.none: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.insertUpdate: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.insert: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.update: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.deleteOnSheet: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.deleteAll: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.selectAll: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.selectCondition: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.selectReExec: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
    End If
    
    dbQueryBatchTypeChangeAll.addAll dbBatchQueryTypeRawList, "dbQueryBatchTypeName"
    dbQueryBatchTypeChangeAll.setSelectedIndex 0
    
    ' �t�@�C���o�̓I�v�V������ǂݍ���
    restoreFileOutputOption
    
    ' �t�@�C���p�X�Ƀf�t�H���g�̃t�@�C������ݒ肷��
    If txtFilePath.value = "" Then
        txtFilePath.value = VBUtil.extractDirPathFromFilePath(targetBook.path)
    End If
    
    ' �e�[�u���V�[�g��ǂݍ���
    readTableSheet
    
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
    
    If dbQueryBatchMode <> FileOutput Then
        ' �t�@�C���o�̓��[�h�ł͂Ȃ��ꍇ
        Exit Sub
    End If
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' �������݃f�[�^
    Dim values As New ValCollection
    
    values.setItem Array(txtFilePath.name, txtFilePath.value)
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
        
    If dbQueryBatchMode <> FileOutput Then
        ' �t�@�C���o�̓��[�h�ł͂Ȃ��ꍇ
        Exit Sub
    End If
    
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
' ���e�[�u���V�[�g��ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub readTableSheet()

    ' �e�[�u�����X�g
    Dim tableList As ValCollection
    Dim tableWorksheet As ValTableWorksheet
    
    Dim dbQueryBatchTableWorksheet As ValDbQueryBatchTableWorksheet
    
    ' �e�[�u���V�[�g�Ǎ��I�u�W�F�N�g
    Dim tableSheetReader As ExeTableSheetReader
    Set tableSheetReader = New ExeTableSheetReader
        
    ' �V�[�g
    Dim sheet As Worksheet
    
    ' �e�[�u�����X�g������������
    Set tableList = New ValCollection
    
    ' �u�b�N�Ɋ܂܂�Ă���V�[�g��1������������
    For Each sheet In book.Worksheets
    
        Set tableSheetReader.sheet = sheet
        
        ' �ΏۃV�[�g���e�[�u���V�[�g�̏ꍇ
        If tableSheetReader.isTableSheet = True Then
        
            ' �e�[�u���V�[�g��ǂݍ���Ń��X�g�ɐݒ肷��i�e�[�u�����̂ݎ擾����j
            Set tableWorksheet = tableSheetReader.readTableInfo(True)
            
            Set dbQueryBatchTableWorksheet = New ValDbQueryBatchTableWorksheet
            dbQueryBatchTableWorksheet.dbQueryBatchType = dbQueryBatchTypeChangeAll.getItem(1).dbQueryBatchType
            Set dbQueryBatchTableWorksheet.tableWorksheet = tableWorksheet
            
            tableList.setItem dbQueryBatchTableWorksheet
        End If
    
    Next
    
    ' ���X�g�R���g���[���Ƀe�[�u���V�[�g����ǉ�����
    Set tableSheetList = New CntListBox: tableSheetList.init lstTableSheet
    tableSheetList.removeAll
    addTableSheetList tableList
    
End Sub

' =========================================================
' ���e�[�u���V�[�g���X�g��ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvalTableSheetList �e�[�u���V�[�g���X�g
'     �@�@�@  isAppend              �ǉ��L���t���O
' �߂�l�@�@�F
'
' =========================================================
Private Sub addTableSheetList(ByVal valTableSheetList As ValCollection, Optional ByVal isAppend As Boolean = True)
    
    tableSheetList.addAll valTableSheetList _
                       , "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ���e�[�u���V�[�g��ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �߂�l�@�@�F
'
' =========================================================
Private Sub addTableSheet(ByVal tableSheet As ValDbQueryBatchTableWorksheet)
    
    tableSheetList.addItemByProp tableSheet, "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName"
    
End Sub

' =========================================================
' ���e�[�u���V�[�g��ύX
'
' �T�v�@�@�@�F
' �����@�@�@�Findex �C���f�b�N�X
'     �@�@�@  rec   �e�[�u���V�[�g
' �߂�l�@�@�F
'
' =========================================================
Private Sub setTableSheet(ByVal index As Long, ByVal rec As ValDbQueryBatchTableWorksheet)
    
    tableSheetList.setItem index, rec, "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName"
    
End Sub
