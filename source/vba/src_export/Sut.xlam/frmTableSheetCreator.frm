VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetCreator 
   Caption         =   "�e�[�u���V�[�g�̍쐬"
   ClientHeight    =   8790.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535.001
   OleObjectBlob   =   "frmTableSheetCreator.frx":0000
End
Attribute VB_Name = "frmTableSheetCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' �e�[�u���V�[�g�쐬�t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/01/25�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����������������ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event complete(ByRef createTargetTable As ValCollection)

' =========================================================
' ���������L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

Private Const MULTIPAGE_MIN_PAGE As Long = 0
Private Const MULTIPAGE_MAX_PAGE As Long = 4

Private Const MULTIPAGE_WELCOME            As Long = 0
Private Const MULTIPAGE_CHOICE_SCHEMA      As Long = 1
Private Const MULTIPAGE_CHOICE_TABLE       As Long = 2
Private Const MULTIPAGE_SETTING_ROW_FORMAT As Long = 3
Private Const MULTIPAGE_COMPLETE           As Long = 4

Private Const ROW_FORMAT_STR_TO_UNDER As String = "��"
Private Const ROW_FORMAT_STR_TO_RIGHT As String = "��"

' �A�v���P�[�V�����ݒ���
Private applicationSetting As ValApplicationSetting

' DB�R�l�N�V�����I�u�W�F�N�g
Private dbConn As Object

' -------------------------------------------------------------
' �X�L�[�}���X�g
Private schemaInfoList As CntListBox
' �e�[�u�����X�g
Private tableInfoList  As CntListBox
' �e�[�u�����X�g�i�s�t�H�[�}�b�g�j
Private tableInfoListRowFormat As CntListBox

' �I�����ꂽ�e�[�u�����X�g
Private selectedTableList As ValCollection
' -------------------------------------------------------------

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
' �@�@�@�@�@�@conn  DB�R�l�N�V����
' �@�@�@�@�@�@aps   �A�v���P�[�V�����ݒ���
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef aps As ValApplicationSetting, ByRef conn As Object)

    ' �A�v���P�[�V��������ݒ肷��
    Set applicationSetting = aps
    ' DB�R�l�N�V������ݒ肷��
    Set dbConn = conn
    ' �A�N�e�B�u����
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

    ' �f�B�A�N�e�B�u����
    deactivate

    Main.storeFormPosition Me.name, Me
    Me.Hide
End Sub

' =========================================================
' ���A�N�e�B�u
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    ' �X�L�[�}���X�g�I�u�W�F�N�g������������
    Set schemaInfoList = New CntListBox: schemaInfoList.init lstSchemaList
    ' �e�[�u�����X�g�I�u�W�F�N�g������������
    Set tableInfoList = New CntListBox: tableInfoList.init lstTableList1
    ' �e�[�u�����X�g�i�s�t�H�[�}�b�g�j�I�u�W�F�N�g������������
    Set tableInfoListRowFormat = New CntListBox: tableInfoListRowFormat.init lstTableListRowFormat
    
    ' �����̃X�L�[�}�𗘗p���Ȃ��i�P�̂̃X�L�[�}�̂ݎQ�Ɓj
    If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
    
        ' �X�L�[�}��������I���ł��Ȃ��悤�ɂ���
        lstSchemaList.multiSelect = fmMultiSelectSingle
        
    ' �����̃X�L�[�}�𗘗p����
    Else
        
        ' �X�L�[�}�𕡐��I���ł���悤�ɂ���
        lstSchemaList.multiSelect = fmMultiSelectMulti
    End If

    ' �}���`�y�[�W�̃y�[�W�ԍ�����Ԏn�߂̃y�[�W�ɐݒ肷��
    multiPage.value = MULTIPAGE_MIN_PAGE
    
    ' �E�B�U�[�h�`���̃E�B���h�E�𑀍삷�邽�߂̊e�{�^����enable�v���p�e�B��ݒ肷��
    ' 1�y�[�W�ڂȂ̂Ŗ߂�Ȃ�
    btnBack.Enabled = False
    ' 1�y�[�W�ڂȂ̂Ŗ߂��
    btnNext.Enabled = True
    ' �L�����Z���͉����\
    btnCancel.Enabled = True
    ' �����͉����s��
    btnFinish.Enabled = False
    
    ' �y�[�W���̏��������s��
    ' �E�F���J���y�[�W
    initPageWelcome
    ' �X�L�[�}�I���y�[�W
    initPageChoiceSchema
    ' �e�[�u���I���y�[�W
    initPageChoiceTable
    ' �����y�[�W
    initPageComplete
    
End Sub

' =========================================================
' ���f�B�A�N�e�B�u
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
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���t�H�[���N���[�Y���̃C�x���g�v���V�[�W��
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
        btnCancel_Click
    End If
    
End Sub

' =========================================================
' ���߂�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnBack_Click()

    On Error GoTo err

    Dim page As Long

    ' 1�y�[�W�ڂ��O�ɖ߂낤�Ƃ����ꍇ
    If multiPage.value - 1 <= MULTIPAGE_MIN_PAGE Then
    
        ' �}���`�y�[�W��1�y�[�W�ڂɐݒ�
        page = MULTIPAGE_MIN_PAGE
        
        ' �y�[�W�؂�ւ�����
        changePage page
        
        ' �e�{�^����enable�v���p�e�B��1�y�[�W�̏�Ԃɐݒ�
        btnBack.Enabled = False
        btnNext.Enabled = True
        btnCancel.Enabled = True
        btnFinish.Enabled = False
        
    ' 1�y�[�W�ȊO
    Else
    
        ' �}���`�y�[�W�����݂̃y�[�W����1�y�[�W�O�ɐݒ�
        page = multiPage.value - 1
        
        ' �y�[�W�؂�ւ�����
        changePage page
        
        ' �e�{�^����enable�v���p�e�B��1�y�[�W�ȊO�̏�Ԃɐݒ�
        btnBack.Enabled = True
        btnNext.Enabled = True
        btnCancel.Enabled = True
        btnFinish.Enabled = False
        
        ' �}���`�y�[�W�����݂̃y�[�W����1�y�[�W�O�ɐݒ�
        page = multiPage.value - 1

    End If

    ' �y�[�W��؂�ւ���
    multiPage.value = page

    Exit Sub
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �����փ{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnNext_Click()

    On Error GoTo err

    Dim page As Long

    ' �y�[�W�؂�ւ��O�̃`�F�b�N����
    changePageBefore multiPage.value
    
    ' �ŏI�y�[�W�ڈȍ~�ɐi�����Ƃ����ꍇ
    If multiPage.value + 1 >= MULTIPAGE_MAX_PAGE Then
    
        ' �}���`�y�[�W���ŏI�y�[�W�ڂɐݒ�
        page = MULTIPAGE_MAX_PAGE
        
        ' �y�[�W�؂�ւ�����
        changePage page
        
        ' �e�{�^����enable�v���p�e�B���ŏI�y�[�W�̏�Ԃɐݒ�
        btnBack.Enabled = True
        btnNext.Enabled = False
        btnCancel.Enabled = True
        btnFinish.Enabled = True
        
    ' �ŏI�y�[�W�ȊO
    Else
    
        ' �}���`�y�[�W�����݂̃y�[�W����1�y�[�W��ɐݒ�
        page = multiPage.value + 1
        
        ' �y�[�W�؂�ւ�����
        changePage page
        
        ' �e�{�^����enable�v���p�e�B���ŏI�y�[�W�ȊO�̏�Ԃɐݒ�
        btnBack.Enabled = True
        btnNext.Enabled = True
        btnCancel.Enabled = True
        btnFinish.Enabled = False
        
    End If
    
    ' �y�[�W��؂�ւ���
    multiPage.value = page

    Exit Sub
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���L�����Z���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnCancel_Click()

    On Error GoTo err
    
    ' �L�����Z������
    If checkCancel = True Then
    
        ' �t�H�[�����\���ɂ���
        HideExt
    
        ' �C�x���g�𔭍s����
        RaiseEvent Cancel
    End If
    
    
    Exit Sub
err:

    ' �G���[���b�Z�[�W��\������
    Main.ShowErrorMessage
    
    ' �t�H�[�����\���ɂ���
    HideExt
    
End Sub

' =========================================================
' ���L�����Z������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�FTrue �L�����Z������ꍇ
'
' =========================================================
Private Function checkCancel() As Boolean

    ' ���b�Z�[�W�{�b�N�X�̖߂�l
    Dim result As Long
    
    ' �L�����Z���m�F�p�̃��b�Z�[�W�{�b�N�X��\������
    result = VBUtil.showMessageBoxForYesNo("�I�����Ă���낵���ł����H", ConstantsCommon.APPLICATION_NAME)

    If result = WinAPI_User.IDYES Then
    
        checkCancel = True
    Else
    
        checkCancel = False
    End If
    
End Function

' =========================================================
' �������{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnFinish_Click()

    On Error GoTo err
    
    fixPageComplete
    
    ' �㑱�̃C�x���g���s���ɁA�V�����t�H�[�����J���Ă��܂��̂ł����Ő�Ƀt�H�[������Ă���
    Me.Hide
    ' �C�x���g�𔭍s����
    RaiseEvent complete(selectedTableList)
        
    HideExt
        
    Exit Sub
err:

    Main.ShowErrorMessage
        
    HideExt
    
End Sub

' =========================================================
' ���}���`�y�[�W�̃y�[�W�ړ����̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub multiPage_Change()

    On Error GoTo err

    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���e�[�u�����X�g�̃`�F�b�N��Ԃ�S��ON�ɂ���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnTableAll_Click()

    Dim i As Long
    
    ' �S�Ė��I���ɂ���
    For i = 0 To lstTableList1.ListCount - 1
    
        lstTableList1.selected(i) = True
    Next
    
End Sub

' =========================================================
' ���e�[�u�����X�g�̃`�F�b�N��Ԃ�S��OFF�ɂ���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnTableNon_Click()

    Dim i As Long
    
    ' �S�Ė��I���ɂ���
    For i = 0 To lstTableList1.ListCount - 1
    
        lstTableList1.selected(i) = False
    Next
    
End Sub

' =========================================================
' ���s�t�H�[�}�b�g���X�g�I�����̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstTableListRowFormat_Change()
    
    ' �e�[�u�����X�g�̃C���f�b�N�X
    Dim i    As Long
    ' �e�[�u�����X�g�̃T�C�Y
    Dim size As Long
    
    ' �e�[�u�����X�g�̃T�C�Y���擾����
    size = lstTableListRowFormat.ListCount
        
    ' ���X�g��őI������Ă���v�f���擾����
    i = lstTableListRowFormat.ListIndex

    ' ���X�g�R���g���[���ɂđI������Ă��邩�𔻒肷��
    If lstTableListRowFormat.selected(i) = True Then
    
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_RIGHT
    Else
    
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_UNDER
    End If
    
End Sub

' =========================================================
' ���s�t�H�[�}�b�g���X�g�̑S�v�f�����ɐݒ肷��{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnRowFormatToUnder_Click()

    ' �e�[�u�����X�g�̃C���f�b�N�X
    Dim i    As Long
    ' �e�[�u�����X�g�̃T�C�Y
    Dim size As Long
    
    ' �e�[�u�����X�g�̃T�C�Y���擾����
    size = lstTableListRowFormat.ListCount
        
    ' ���X�g�R���g���[�������[�v������
    For i = 0 To size - 1
    
        lstTableListRowFormat.selected(i) = False
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_UNDER
    Next
    
End Sub

' =========================================================
' ���s�t�H�[�}�b�g���X�g�̑S�v�f�����ɐݒ肷��{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnRowFormatToRight_Click()

    ' �e�[�u�����X�g�̃C���f�b�N�X
    Dim i    As Long
    ' �e�[�u�����X�g�̃T�C�Y
    Dim size As Long
    
    ' �e�[�u�����X�g�̃T�C�Y���擾����
    size = lstTableListRowFormat.ListCount
        
    ' ���X�g�R���g���[�������[�v������
    For i = 0 To size - 1
    
        lstTableListRowFormat.selected(i) = True
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_RIGHT
    Next
    
End Sub


' =========================================================
' ���}���`�y�[�W��؂�ւ���O�ɌĂяo������
'
' �T�v�@�@�@�F
' �����@�@�@�FpageIndex �y�[�W�ԍ�
' �߂�l�@�@�F
'
' =========================================================
Private Sub changePageBefore(ByVal pageIndex As Long)

    If pageIndex = MULTIPAGE_WELCOME Then
    
        ' ���؏������s��
        validPageWelcome
    
    ElseIf pageIndex = MULTIPAGE_CHOICE_SCHEMA Then
    
        ' ���؏������s��
        validPageChoiceSchema
        
    ElseIf pageIndex = MULTIPAGE_CHOICE_TABLE Then
    
        ' ���؏������s��
        validPageChoiceTable
        
    ElseIf pageIndex = MULTIPAGE_SETTING_ROW_FORMAT Then
    
        ' ���؏������s��
        validPageSettingRowFormat
    
    ElseIf pageIndex = MULTIPAGE_COMPLETE Then
    
        ' ���؏������s��
        validPageComplete
        
    End If

End Sub

' =========================================================
' ���}���`�y�[�W��؂�ւ��鏈��
'
' �T�v�@�@�@�F
' �����@�@�@�FpageIndex �y�[�W�ԍ�
' �߂�l�@�@�F
'
' =========================================================
Private Sub changePage(ByVal pageIndex As Long)

    If pageIndex = MULTIPAGE_WELCOME Then
    
        ' �\���������s��
        activatePageWelcome
        
    ElseIf pageIndex = MULTIPAGE_CHOICE_SCHEMA Then
    
        ' �\���������s��
        activatePageChoiceSchema
        ' �����������s��
        fixPageWelcome
        
    ElseIf pageIndex = MULTIPAGE_CHOICE_TABLE Then
    
        ' �����������s��
        fixPageChoiceSchema
        ' �\���������s��
        activatePageChoiceTable
        
    ElseIf pageIndex = MULTIPAGE_SETTING_ROW_FORMAT Then
    
        ' �����������s��
        fixPageChoiceTable
        ' �\���������s��
        activatePageSettingRowFormat
    
    ElseIf pageIndex = MULTIPAGE_COMPLETE Then
    
        ' �����������s��
        fixPageSettingRowFormat
        ' �\���������s��
        activatePageComplete
        
    End If

End Sub

Private Sub initPageWelcome()

End Sub

Private Sub initPageChoiceSchema()

    On Error GoTo err

    ' �����Ԃ̏��������s�����̂Ń}�E�X�J�[�\���������v�ɂ���
    Dim cursorWait As New ExcelCursorWait: cursorWait.init

    ' �ꎞ�ϐ�
    Dim var As ValCollection
    
    Dim dbObjFactory As New DbObjectFactory
    
    Dim dbInfo As IDbMetaInfoGetter
    Set dbInfo = dbObjFactory.createMetaInfoGetterObject(dbConn)
    
    ' �����Ԃ̏������I�������̂Ń}�E�X�J�[�\�������ɖ߂�
    cursorWait.destroy
    
    Set var = dbInfo.getSchemaList
    
    ' �X�L�[�}���X�g�{�b�N�X�Ƀ��X�g��ǉ�����
    schemaInfoList.addAll var, "SchemaName", "SchemaComment"
        
    Exit Sub
    
err:

    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

Private Sub initPageChoiceTable()

End Sub

Private Sub initPageComplete()

End Sub

Private Sub validPageWelcome()

End Sub

Private Sub validPageChoiceSchema()

    Dim cnt As Long
    cnt = schemaInfoList.getSelectedList().count
    
    ' �X�L�[�}���X�g�ł̑I���������m�F����
    If cnt <= 0 Then
    
        ' 0���̏ꍇ�G���[�𔭍s����
        err.Raise ERR_NUMBER_NOT_SELECTED_SCHEMA _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_SCHEMA _
                , err.HelpFile _
                , err.HelpContext
    End If

End Sub

Private Sub validPageChoiceTable()

    Dim cnt As Long
    cnt = tableInfoList.getSelectedList().count
    
    ' �e�[�u�����X�g�ł̑I���������m�F����
    If cnt <= 0 Then
    
        ' 0���̏ꍇ�G���[�𔭍s����
        err.Raise ERR_NUMBER_NOT_SELECTED_TABLE _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_TABLE _
                , err.HelpFile _
                , err.HelpContext
    End If

End Sub

Private Sub validPageSettingRowFormat()

End Sub

Private Sub validPageComplete()

End Sub

Private Sub activatePageWelcome()

End Sub

Private Sub activatePageChoiceSchema()

End Sub

Private Sub activatePageChoiceTable()

    On Error GoTo err

    ' �����Ԃ̏��������s�����̂Ń}�E�X�J�[�\���������v�ɂ���
    Dim cursorWait As New ExcelCursorWait: cursorWait.init

    ' �ꎞ�ϐ�
    Dim var  As ValCollection
    
    Dim selectedSchema As ValCollection
    Set selectedSchema = schemaInfoList.getSelectedList(vbObject)
    
    Dim dbObjFactory As New DbObjectFactory
    
    Dim dbInfo As IDbMetaInfoGetter
    Set dbInfo = dbObjFactory.createMetaInfoGetterObject(dbConn)
    
    Set var = dbInfo.getTableList(selectedSchema)
    
    ' �e�[�u�����X�g�{�b�N�X�Ƀ��X�g��ǉ�����
    If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
        tableInfoList.addAll var, "TableName", "TableComment"
    Else
        tableInfoList.addAll var, "SchemaTableName", "TableComment"
    End If
    
    ' �����Ԃ̏������I�������̂Ń}�E�X�J�[�\�������ɖ߂�
    cursorWait.destroy
    
    Exit Sub
    
err:

    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

Private Sub activatePageSettingRowFormat()

    ' �I���ς݃e�[�u�����X�g
    Dim selectedTableList As ValCollection
    ' �I���ς݃e�[�u��
    Dim selectedTable     As ValDbDefineTable
    
    ' �I���ς݃e�[�u�����X�g���擾����
    Set selectedTableList = tableInfoList.getSelectedList
    
    ' �e�[�u�����X�g�{�b�N�X�Ƀ��X�g��ǉ�����
    If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
        tableInfoListRowFormat.addAll selectedTableList _
                                    , "TableName"
    Else
        tableInfoListRowFormat.addAll selectedTableList _
                                    , "SchemaTableName"
    End If
    
    ' �e�[�u����`�̍s�t�H�[�}�b�g�̏�Ԃ��R���g���[���ɔ��f����
    ' ���̏ꍇ�A���X�g�𖢑I��
    ' ���̏ꍇ�A���X�g��I��
    
    ' �e�[�u�����X�g�̃C���f�b�N�X
    Dim i    As Long
    ' �e�[�u�����X�g�̃T�C�Y
    Dim size As Long
    
    ' �e�[�u�����X�g�̃T�C�Y���擾����
    size = lstTableListRowFormat.ListCount
        
    ' ���X�g�R���g���[�������[�v������
    For i = 0 To size - 1
    
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_UNDER
        lstTableListRowFormat.selected(i) = False
        
    Next
    

End Sub

Private Sub activatePageComplete()

End Sub

Private Sub fixPageWelcome()

End Sub

Private Sub fixPageChoiceSchema()

End Sub

Private Sub fixPageChoiceTable()

End Sub

Private Sub fixPageSettingRowFormat()

End Sub

Private Sub fixPageComplete()

    On Error GoTo err

    ' �e�[�u�����X�g�̃C���f�b�N�X
    Dim i    As Long
    
    ' �ꎞ�ϐ�
    Dim var    As ValCollection
    Dim varObj As ValDbDefineTable
    
    Dim tableSheetList As New ValCollection
    Dim tableSheet     As ValTableWorksheet
    
    Set var = tableInfoListRowFormat.collection
    
    ' Table�ɐݒ肳��Ă���X�L�[�}�����N���A����
    For Each varObj In var.col
    
        Set tableSheet = New ValTableWorksheet
        Set tableSheet.table = varObj
        tableSheet.recFormat = recFormatToUnder
        
        ' �����̃X�L�[�}�𗘗p���Ȃ��i�P�̂̃X�L�[�}�̂ݎQ�Ɓj
        If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
        
            tableSheet.omitsSchema = True
        Else
            tableSheet.omitsSchema = False
        End If
        
        ' ���X�g�R���g���[���ɂđI������Ă��邩�𔻒肷��
        If lstTableListRowFormat.selected(i) = True Then
        
            tableSheet.recFormat = REC_FORMAT.recFormatToRight
        Else
        
            tableSheet.recFormat = REC_FORMAT.recFormatToUnder
        End If
        
        tableSheetList.setItem tableSheet, varObj.schemaTableName
        
        i = i + 1
    Next
    
    Set selectedTableList = tableSheetList
    
    Exit Sub
    
err:

    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub
