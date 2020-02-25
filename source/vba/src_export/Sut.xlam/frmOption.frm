VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOption 
   Caption         =   "�I�v�V����"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9030.001
   OleObjectBlob   =   "frmOption.frx":0000
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �I�v�V�����ݒ���s���t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/03/14�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FapplicationSetting �A�v���P�[�V�����ݒ���
'
' =========================================================
Public Event ok(ByRef applicationSetting As ValApplicationSetting)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' �J���������ݒ�
Private WithEvents frmDBColumnFormatVar As frmDBColumnFormat
Attribute frmDBColumnFormatVar.VB_VarHelpID = -1

' �A�v���P�[�V�����ݒ���
Private applicationSetting As ValApplicationSetting
' �A�v���P�[�V�����ݒ���i�J���������j
Private applicationSettingColFmt As ValApplicationSettingColFormat

' �t�H���g���X�g �R���g���[��
Private fontList As CntListBox
' �t�H���g�T�C�Y���X�g �R���g���[��
Private fontSizeList As CntListBox

' �J����������ݒ蒆��DB
Private settingColFormatDb As DbmsType

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
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByRef var As ValApplicationSetting _
                 , ByRef var2 As ValApplicationSettingColFormat)

    ' �����o�ϐ��ɃA�v���P�[�V�����ݒ����ݒ肷��
    Set applicationSetting = var
    Set applicationSettingColFmt = var2
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
' ���t�H�[���A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    ' �O��Ō�ɐݒ肵�������t�H�[����̊e�R���g���[���ɕ���������
    ' �ǂݍ��݂��s��
    applicationSetting.readForData
    restoreOptionInfo applicationSetting
    
    ' �G���[���b�Z�[�W���N���A����
    lblErrorMessage.Caption = ""
    ' �u�b�N�^�C�g����ݒ肷��
    lblBookTitle.Caption = replace(lblBookTitle.Tag, "${book}", targetBook.name)

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
    
    ' �����L�^����
    storeOptionInfo
    
    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(applicationSetting)
    
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
' ���f�t�H���g�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDefault_Click()

    On Error GoTo err
    
    Dim applicationSetting As New ValApplicationSetting
    ' �f�t�H���g�l�𔽉f�������A�v���P�[�V�����f�[�^�ŃR���g���[���ɔ��f����
    restoreOptionInfo applicationSetting

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

    ' �J���������ݒ������������
    If VBUtil.unloadFormIfChangeActiveBook(frmDBColumnFormat) Then Unload frmDBColumnFormat
    Load frmDBColumnFormat
    Set frmDBColumnFormatVar = frmDBColumnFormat
    
    ' �t�H���g���X�g������������
    Set fontList = New CntListBox: fontList.init cboFontList
    ' �t�H���g�T�C�Y���X�g������������
    Set fontSizeList = New CntListBox: fontSizeList.init cboFontSizeList

    ' �t�H���g���X�g��Excel�ŗ��p�\�ȃt�H���g���i�[����
    fontList.addAll WinAPI_GDI.getFontNameList
    ' �t�H���g�T�C�Y���X�g��Excel�̃t�H���g�T�C�Y�̋K��l���i�[����
    fontSizeList.addAll ExcelUtil.getFontSizeList
    
    MultiPageGlobalOrBook.value = 0
    MultiPageGlobalOrBook.Pages.item("PageGlobalSetting").MultiPageAll.value = 0
    MultiPageGlobalOrBook.Pages.item("PageBookSetting").MultiPageBook.value = 0

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

    ' �J���������ݒ��j������
    Set frmDBColumnFormatVar = Nothing
    
End Sub

' =========================================================
' ���I�v�V��������ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeOptionInfo()

    With applicationSetting
    
        ' ���R�[�h�����P��
        If optRecProcessCountAll.value = True Then
            .recProcessCount = .REC_PROCESS_COUNT_ALL
            .recProcessCountCustom = txtRecProcessCountCustom.value
        Else
            .recProcessCount = .REC_PROCESS_COUNT_COSTOM
            .recProcessCountCustom = txtRecProcessCountCustom.value
        End If
        
        ' �R�~�b�g�m�F
        If optCommitConfirmNo.value = True Then
            .commitConfirm = .COMMIT_CONFIRM_NO
        Else
            .commitConfirm = .COMMIT_CONFIRM_YES
        End If
        
        ' SQL�G���[���̋���
        If optSqlErrorHandlingSuspend.value = True Then
            .sqlErrorHandling = .SQL_ERROR_HANDLING_SUSPEND
        Else
            .sqlErrorHandling = .SQL_ERROR_HANDLING_RESUME
        End If
        
        ' �X�L�[�}�̎w��
        If optSchemaUseOne.value = True Then
            .schemaUse = .SCHEMA_USE_ONE
        Else
            .schemaUse = .SCHEMA_USE_MULTIPLE
        End If
        
        ' ���펞�̃N�G�����ʕ\���L��
        If optQueryResultShowWhenNormalNo.value = True Then
            .queryResultShowWhenNormal = False
        Else
            .queryResultShowWhenNormal = True
        End If
        
        ' �t�H���g��
        .cellFontName = cboFontList.value
        ' �t�H���g�T�C�Y
        .cellFontSize = cboFontSizeList.value
        ' �܂�Ԃ��L��
        If optWordWrapYes.value = True Then
            .cellWordwrap = True
        Else
            .cellWordwrap = False
        End If
        ' �Z����
        .cellWidth = CDec(txtCellWidth.value)
        ' �Z������
        .cellHeight = CDec(txtCellHeight.value)
        ' �s���̎�������
        If optLineHeightAutoAdjustNo.value = True Then
            .lineHeightAutoAdjust = False
        Else
            .lineHeightAutoAdjust = True
        End If
        
        ' �󔒃Z���ǂݎ�����
        If optEmptyCellReadingDel.value = True Then
            .emptyCellReading = .EMPTY_CELL_READING_DEL
        ElseIf optEmptyCellReadingNonDel.value = True Then
            .emptyCellReading = .EMPTY_CELL_READING_NON_DEL
        Else
            .emptyCellReading = .EMPTY_CELL_READING_NON_DEL_STR_EMPTY
        End If
        
        ' ���ړ��͕���
        If optDirectInputCharDisable.value = True Then
            .directInputChar = .DIRECT_INPUT_CHAR_DISABLE
        Else
            .directInputChar = .DIRECT_INPUT_CHAR_ENABLE_CUSTOM
        End If
        .directInputCharCustomPrefix = txtDirectInputCharEnableCustomPrefix.value
        .directInputCharCustomSuffix = txtDirectInputCharEnableCustomSuffix.value
        
        ' �N�G���p�����[�^�̈͂ݕ���
        .queryParameterEncloseCustomPrefix = txtQueryParameterEncloseEnableCustomPrefix.value
        .queryParameterEncloseCustomSuffix = txtQueryParameterEncloseEnableCustomSuffix.value
        
        ' NULL���͕���
        If optNullInputCharDisable.value = True Then
            .nullInputChar = .NULL_INPUT_CHAR_DISABLE
        Else
            .nullInputChar = .NULL_INPUT_CHAR_ENABLE_CUSTOM
        End If
        .nullInputCharCustom = txtNullInputCharEnableCustom.text
        
        ' SELECT���̃Z���̍ő啶�����`�F�b�N
        If optSelectCheckCellMaxLengthDisable.value = True Then
            .selectCheckCellMaxLength = False
        Else
            .selectCheckCellMaxLength = True
        End If
        
        ' �e�[�u���E�J�������G�X�P�[�v
        .tableColumnEscapeOracle = chkTableColumnEscapeOracle.value
        .tableColumnEscapeMysql = chkTableColumnEscapeMysql.value
        .tableColumnEscapePostgresql = chkTableColumnEscapePostgresql.value
        .tableColumnEscapeSqlserver = chkTableColumnEscapeSqlserver.value
        .tableColumnEscapeAccess = chkTableColumnEscapeAccess.value
        .tableColumnEscapeSymfoware = chkTableColumnEscapeSymfoware.value
        
        ' �������݂��s��
        .writeForData
    
    End With
End Sub

' =========================================================
' ���I�v�V��������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�Fapplicationsetting �A�v���P�[�V�����f�[�^
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreOptionInfo(ByRef applicationSetting As ValApplicationSetting)

    ' �A�v���P�[�V�����ݒ�I�u�W�F�N�g
    With applicationSetting
        
        ' ���R�[�h�����P��
        If .recProcessCount = .REC_PROCESS_COUNT_ALL Then
            optRecProcessCountAll.value = True
        Else
            optRecProcessCountCustom.value = True
        End If
        txtRecProcessCountCustom.value = .recProcessCountCustom
        
        ' �R�~�b�g�m�F
        If .commitConfirm = .COMMIT_CONFIRM_NO Then
            optCommitConfirmNo.value = True
        Else
            optCommitConfirmYes.value = True
        End If
        
        ' SQL�G���[���̋���
        If .sqlErrorHandling = .SQL_ERROR_HANDLING_SUSPEND Then
            optSqlErrorHandlingSuspend.value = True
        Else
            optSqlErrorHandlingResume.value = True
        End If
        
        ' �X�L�[�}
        If .schemaUse = .SCHEMA_USE_ONE Then
            optSchemaUseOne.value = True
        Else
            optSchemaUseMultiple.value = True
        End If
        
        ' ���펞�̃N�G�����ʕ\���L��
        If .queryResultShowWhenNormal = True Then
            optQueryResultShowWhenNormalYes.value = True
        Else
            optQueryResultShowWhenNormalNo.value = True
        End If

        ' �t�H���g��
        cboFontList.value = .cellFontName
        ' �t�H���g�T�C�Y
        cboFontSizeList.value = .cellFontSize
        ' �܂�Ԃ��L��
        If .cellWordwrap = True Then
            optWordWrapYes.value = True
        Else
            optWordWrapNo.value = True
        End If
        ' �Z����
        txtCellWidth.value = .cellWidth
        ' �Z������
        txtCellHeight.value = .cellHeight
        ' �s���̎�������
        If .lineHeightAutoAdjust = True Then
            optLineHeightAutoAdjustYes.value = True
        Else
            optLineHeightAutoAdjustNo.value = True
        End If
        
        ' �󔒃Z���ǂݎ�����
        If .emptyCellReading = .EMPTY_CELL_READING_DEL Then
            optEmptyCellReadingDel.value = True
        ElseIf .emptyCellReading = .EMPTY_CELL_READING_NON_DEL Then
            optEmptyCellReadingNonDel.value = True
        Else
            optEmptyCellReadingNonDelStrEmpty.value = True
        End If
        
        ' ���ړ��͕���
        If .directInputChar = .DIRECT_INPUT_CHAR_DISABLE Then
            optDirectInputCharDisable.value = True
        Else
            optDirectInputCharEnableCustom = True
        End If
        
        txtDirectInputCharEnableCustomPrefix = .directInputCharCustomPrefix
        txtDirectInputCharEnableCustomSuffix = .directInputCharCustomSuffix
        
        ' �N�G���p�����[�^�̈͂ݕ���
        txtQueryParameterEncloseEnableCustomPrefix = .queryParameterEncloseCustomPrefix
        txtQueryParameterEncloseEnableCustomSuffix = .queryParameterEncloseCustomSuffix
        
        ' NULL���͕���
        If .nullInputChar = .NULL_INPUT_CHAR_DISABLE Then
            optNullInputCharDisable.value = True
        Else
            optNullInputCharEnableCustom = True
        End If
        txtNullInputCharEnableCustom.text = .nullInputCharCustom
        
        ' SELECT���̃Z���̍ő啶�����`�F�b�N
        If .selectCheckCellMaxLength = False Then
            optSelectCheckCellMaxLengthDisable.value = True
        Else
            optSelectCheckCellMaxLengthEnable.value = True
        End If
        
        ' �e�[�u���E�J�������G�X�P�[�v
        chkTableColumnEscapeOracle.value = .tableColumnEscapeOracle
        chkTableColumnEscapeMysql.value = .tableColumnEscapeMysql
        chkTableColumnEscapePostgresql.value = .tableColumnEscapePostgresql
        chkTableColumnEscapeSqlserver.value = .tableColumnEscapeSqlserver
        chkTableColumnEscapeAccess.value = .tableColumnEscapeAccess
        chkTableColumnEscapeSymfoware.value = .tableColumnEscapeSymfoware
        
    End With
    
    applicationSettingColFmt.readForData
    
End Sub

' =========================================================
' �������P�ʃe�L�X�g�@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtRecProcessCountCustom_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' ���������`�F�b�N����
    If validInteger(txtRecProcessCountCustom.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtRecProcessCountCustom
    
    ' ���l�͈̓`�F�b�N
    ElseIf CDec(txtRecProcessCountCustom.text) < 1 Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtRecProcessCountCustom

    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue txtRecProcessCountCustom
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���Z�����e�L�X�g�@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtCellWidth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' ���l�`�F�b�N
    If validUnsignedNumeric(txtCellWidth.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NUMERIC
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtCellWidth
    
    ' ���l�͈̓`�F�b�N
    ElseIf CDec(txtCellWidth.text) < applicationSetting.CELL_WIDTH_DEFAULT Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", applicationSetting.CELL_WIDTH_DEFAULT)
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtCellWidth
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue txtCellWidth
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���Z�������e�L�X�g�@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtCellHeight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' ���l�`�F�b�N
    If validUnsignedNumeric(txtCellHeight.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NUMERIC
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtCellHeight
    
    ' ���l�͈̓`�F�b�N
    ElseIf CDec(txtCellHeight.text) < applicationSetting.CELL_HEIGHT_DEFAULT Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", applicationSetting.CELL_HEIGHT_DEFAULT)
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtCellHeight
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue txtCellHeight
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���t�H���g���X�g�@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboFontList_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' ���X�g�Ɍ��ݓ��͂���Ă���e�L�X�g�̗v�f�����݂��Ȃ��ꍇ
    If fontList.exist(cboFontList.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NO_LIST_ITEM
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse cboFontList
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue cboFontList
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

    
End Sub

' =========================================================
' ���t�H���g�T�C�Y���X�g�@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboFontSizeList_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:
    
    ' �t�H���g�T�C�Y�ɐ��l�����͂���Ă��Ȃ��ꍇ
    If validUnsignedNumeric(cboFontSizeList.value) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NUMERIC
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse cboFontSizeList
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue cboFontSizeList
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

    
End Sub

' =========================================================
' �����ړ��͕����ړ����@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtDirectInputCharEnableCustomPrefix_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' �e�L�X�g�{�b�N�X�ɓ��͂��Ȃ��ꍇ
    If txtDirectInputCharEnableCustomPrefix.text = "" Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtDirectInputCharEnableCustomPrefix
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue txtDirectInputCharEnableCustomPrefix
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���N�G���p�����[�^�͂ݕ����ړ����@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtQueryParameterEncloseEnableCustomPrefix_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' �e�L�X�g�{�b�N�X�ɓ��͂��Ȃ��ꍇ
    If txtQueryParameterEncloseEnableCustomPrefix.text = "" Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtQueryParameterEncloseEnableCustomPrefix
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue txtQueryParameterEncloseEnableCustomPrefix
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���N�G���p�����[�^�͂ݕ����ڔ����@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtQueryParameterEncloseEnableCustomSuffix_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' �e�L�X�g�{�b�N�X�ɓ��͂��Ȃ��ꍇ
    If txtQueryParameterEncloseEnableCustomSuffix.text = "" Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtQueryParameterEncloseEnableCustomSuffix
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue txtQueryParameterEncloseEnableCustomSuffix
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ��NULL���ړ��͕����@�X�V���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtNullInputCharEnableCustom_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' �e�L�X�g�{�b�N�X�ɓ��͂��Ȃ��ꍇ
    If txtNullInputCharEnableCustom.text = "" Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidFalse txtNullInputCharEnableCustom
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        ' �R���g���[���̃v���p�e�B��ύX����
        VBUtil.changeControlPropertyByValidTrue txtNullInputCharEnableCustom
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���J���������ݒ�iOracle�j�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdColumnTypeFormatOracle_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.Oracle)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ���J���������ݒ�iMySQL�j�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdColumnTypeFormatMySQL_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.MySQL)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ���J���������ݒ�iPostgreSQL�j�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdColumnTypeFormatPostgreSQL_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.PostgreSQL)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ���J���������ݒ�iSQLServer�j�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdColumnTypeFormatSQLServer_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.MicrosoftSqlServer)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ���J���������ݒ�iAccess�j�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdColumnTypeFormatAccess_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.MicrosoftAccess)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ���J���������ݒ�iSymfoware�j�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdColumnTypeFormatSymfoware_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.Symfoware)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ���J���������ݒ�E�B���h�E��OK�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�FdbColumnFormatInfo �J���������ݒ�E�B���h�E�Őݒ肳�ꂽ���
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmDBColumnFormatVar_ok(ByVal dbColumnFormatInfo As ValDbColumnFormatInfo)

    ' �A�v���P�[�V�����ݒ���Ƀ��[�h���ꂽ����ݒ肷��
    applicationSettingColFmt.setDbColFormatInfo dbColumnFormatInfo
    
    ' ������������
    applicationSettingColFmt.writeForDataDbInfo dbColumnFormatInfo

End Sub

' =========================================================
' ���J���������ݒ�E�B���h�E�̃L�����Z���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmDBColumnFormatVar_cancel()

End Sub
