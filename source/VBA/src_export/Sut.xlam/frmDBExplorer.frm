VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBExplorer 
   Caption         =   "DB�G�N�X�v���[��"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7950
   OleObjectBlob   =   "frmDBExplorer.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmDBExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�G�N�X�v���[��
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
' �����@�@�@�FtableList  �e�[�u�����X�g
'             recFormat  ���R�[�h�t�H�[�}�b�g
' =========================================================
Public Event export(ByVal tableList As ValCollection _
                  , ByVal recFormat As REC_FORMAT)

' =========================================================
' ������{�^���������ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event closed()

Private Const REG_SUB_KEY_DB_EXPLORER_OPTION As String = "db_explorer"

' DB�R�l�N�V�����I�u�W�F�N�g
Private dbConn As Object
' �X�L�[�}���X�g
Private schemaInfoList  As CntListBox
' �e�[�u�����X�g
Private tableInfoList   As CntListBox
' �e�[�u�����X�g�̃t�B���^�����Ȃ��̃��X�g
Private tableWithoutFilterList As ValCollection

Private inFilterProcess As Boolean

' �Ώۃu�b�N
Private targetBook As Workbook
' �Ώۃu�b�N���擾����
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ��DB�R�l�N�V�����ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvNewValue DB�R�l�N�V����
' �߂�l�@�@�F
'
' =========================================================
Public Property Let DbConnection(ByVal vNewValue As Variant)

    Set dbConn = vNewValue
    
    ' �X�L�[�}�V�[�g��ǂݍ���
    readSchemaInfo
    ' �e�[�u���V�[�g��ǂݍ���
    readTableInfo
    
End Property

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal  ���[�_���܂��̓��[�h���X�\���w��
'             conn   DB�R�l�N�V����
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef conn As Object)

    ' DB�R�l�N�V������ݒ肷��
    Set dbConn = conn
    
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
' ���X�L�[�}�R���{�{�b�N�X�ύX���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboSchema_Change()

    On Error GoTo err

    inFilterProcess = True
    
    clearFilterCondition False
    readTableInfo
    
    inFilterProcess = False
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ���t�B���^�R���{�{�b�N�X�ύX���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboFilter_Change()

    On Error GoTo err

    Dim currentFilterText As String

    ' �{�C�x���g�v���V�[�W�������ŁA���R���g���[����ύX���邱�Ƃɂ��ύX�C�x���g��
    ' �ċA�I�ɔ������Ă��ǂ��悤��
    ' �t���O���Q�Ƃ��čĎ��s����Ȃ��悤�ɂ��锻������{
    If inFilterProcess = False Then

        inFilterProcess = True
    
        currentFilterText = cboFilter.text
        
        filterTableInfoList "*" & currentFilterText & "*"
        
        clearFilterCondition True
    
        inFilterProcess = False

    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub


' =========================================================
' ���t�B���^�g�O���S�ʂ̕ύX���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub tglFilterA_Click()
    filterToggle tglFilterA, "A"
End Sub
Private Sub tglFilterB_Click()
    filterToggle tglFilterB, "B"
End Sub
Private Sub tglFilterC_Click()
    filterToggle tglFilterC, "C"
End Sub
Private Sub tglFilterD_Click()
    filterToggle tglFilterD, "D"
End Sub
Private Sub tglFilterE_Click()
    filterToggle tglFilterE, "E"
End Sub
Private Sub tglFilterF_Click()
    filterToggle tglFilterF, "F"
End Sub
Private Sub tglFilterG_Click()
    filterToggle tglFilterG, "G"
End Sub
Private Sub tglFilterH_Click()
    filterToggle tglFilterH, "H"
End Sub
Private Sub tglFilterI_Click()
    filterToggle tglFilterI, "I"
End Sub
Private Sub tglFilterJ_Click()
    filterToggle tglFilterJ, "J"
End Sub
Private Sub tglFilterK_Click()
    filterToggle tglFilterK, "K"
End Sub
Private Sub tglFilterL_Click()
    filterToggle tglFilterL, "L"
End Sub
Private Sub tglFilterM_Click()
    filterToggle tglFilterM, "M"
End Sub
Private Sub tglFilterN_Click()
    filterToggle tglFilterN, "N"
End Sub
Private Sub tglFilterO_Click()
    filterToggle tglFilterO, "O"
End Sub
Private Sub tglFilterP_Click()
    filterToggle tglFilterP, "P"
End Sub
Private Sub tglFilterQ_Click()
    filterToggle tglFilterQ, "Q"
End Sub
Private Sub tglFilterR_Click()
    filterToggle tglFilterR, "R"
End Sub
Private Sub tglFilterS_Click()
    filterToggle tglFilterS, "S"
End Sub
Private Sub tglFilterT_Click()
    filterToggle tglFilterT, "T"
End Sub
Private Sub tglFilterU_Click()
    filterToggle tglFilterU, "U"
End Sub
Private Sub tglFilterV_Click()
    filterToggle tglFilterV, "V"
End Sub
Private Sub tglFilterW_Click()
    filterToggle tglFilterW, "W"
End Sub
Private Sub tglFilterX_Click()
    filterToggle tglFilterX, "X"
End Sub
Private Sub tglFilterY_Click()
    filterToggle tglFilterY, "Y"
End Sub
Private Sub tglFilterZ_Click()
    filterToggle tglFilterZ, "Z"
End Sub
Private Sub tglFilterOther_Click()
    
    ' Other�̏��������u�`�ȊO�v�Ƃ��������Ȃ̂ŕʂ̏����Ƃ��Ē�`
    
    On Error GoTo err

    ' �{�C�x���g�v���V�[�W�������ŁA���R���g���[����ύX���邱�Ƃɂ��ύX�C�x���g��
    ' �ċA�I�ɔ������Ă��ǂ��悤��
    ' �t���O���Q�Ƃ��čĎ��s����Ȃ��悤�ɂ��锻������{
    If inFilterProcess = False Then

        inFilterProcess = True
        
        If tglFilterOther.value = True Then
            ' �A���t�@�x�b�g�ȊO�̕����Ŏn�܂���Ō���
            filterTableInfoListForRegExp "[^a-zA-Z]*"
            
            clearFilterCondition
            tglFilterOther.value = True
        Else
            filterTableInfoListForRegExp ""
        End If
        
        inFilterProcess = False
        
    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ���g�O���n�t�B���^�����̋��ʏ���
'
' �T�v�@�@�@�F
' �����@�@�@�Fstate   �g�O���{�^��
'     �@�@�@  keyword �L�[���[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub filterToggle(ByVal state As ToggleButton, ByVal keyword As String)

    On Error GoTo err

    If inFilterProcess = False Then

        inFilterProcess = True
        
        If state.value = True Then
            filterTableInfoList keyword & "*"
            
            clearFilterCondition
            state.value = True
        Else
            filterTableInfoList ""
        End If
        
        inFilterProcess = False
        
    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ���t�B���^�����̃N���A����
'
' �T�v�@�@�@�F
' �����@�@�@�FisNotClearComboFilter �R���{�{�b�N�X�̃t�B���^���N���A���邩�ǂ����̃t���O
' �߂�l�@�@�F
'
' =========================================================
Private Sub clearFilterCondition(Optional ByVal isNotClearComboFilter As Boolean = False)

    tglFilterA.value = False
    tglFilterB.value = False
    tglFilterC.value = False
    tglFilterD.value = False
    tglFilterE.value = False
    tglFilterF.value = False
    tglFilterG.value = False
    tglFilterH.value = False
    tglFilterI.value = False
    tglFilterJ.value = False
    tglFilterK.value = False
    tglFilterL.value = False
    tglFilterM.value = False
    tglFilterN.value = False
    tglFilterO.value = False
    tglFilterP.value = False
    tglFilterQ.value = False
    tglFilterR.value = False
    tglFilterS.value = False
    tglFilterT.value = False
    tglFilterU.value = False
    tglFilterV.value = False
    tglFilterW.value = False
    tglFilterX.value = False
    tglFilterY.value = False
    tglFilterZ.value = False
    tglFilterOther.value = False
    
    If isNotClearComboFilter = False Then
        cboFilter.text = ""
    End If
    
End Sub

' =========================================================
' ���t�B���^�����̓K�p����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub applyFilterCondition()

    If cboFilter.text <> "" Then
        cboFilter_Change
        Exit Sub
    End If
    
    If tglFilterA.value = True Then
        tglFilterA_Click
    ElseIf tglFilterB.value = True Then
        tglFilterB_Click
    ElseIf tglFilterC.value = True Then
        tglFilterC_Click
    ElseIf tglFilterD.value = True Then
        tglFilterD_Click
    ElseIf tglFilterE.value = True Then
        tglFilterE_Click
    ElseIf tglFilterF.value = True Then
        tglFilterF_Click
    ElseIf tglFilterG.value = True Then
        tglFilterG_Click
    ElseIf tglFilterH.value = True Then
        tglFilterH_Click
    ElseIf tglFilterI.value = True Then
        tglFilterI_Click
    ElseIf tglFilterJ.value = True Then
        tglFilterJ_Click
    ElseIf tglFilterK.value = True Then
        tglFilterK_Click
    ElseIf tglFilterL.value = True Then
        tglFilterL_Click
    ElseIf tglFilterM.value = True Then
        tglFilterM_Click
    ElseIf tglFilterN.value = True Then
        tglFilterN_Click
    ElseIf tglFilterO.value = True Then
        tglFilterO_Click
    ElseIf tglFilterP.value = True Then
        tglFilterP_Click
    ElseIf tglFilterQ.value = True Then
        tglFilterQ_Click
    ElseIf tglFilterR.value = True Then
        tglFilterR_Click
    ElseIf tglFilterS.value = True Then
        tglFilterS_Click
    ElseIf tglFilterT.value = True Then
        tglFilterT_Click
    ElseIf tglFilterU.value = True Then
        tglFilterU_Click
    ElseIf tglFilterV.value = True Then
        tglFilterV_Click
    ElseIf tglFilterW.value = True Then
        tglFilterW_Click
    ElseIf tglFilterX.value = True Then
        tglFilterX_Click
    ElseIf tglFilterY.value = True Then
        tglFilterY_Click
    ElseIf tglFilterZ.value = True Then
        tglFilterZ_Click
    ElseIf tglFilterOther.value = True Then
        tglFilterOther_Click
    End If
    
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

    tableInfoList.setSelectedAll True

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

    tableInfoList.setSelectedAll False

End Sub

' =========================================================
' ���G�N�X�|�[�g�i���j�{�^���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdExportToUnder_Click()
    
    exportProcess recFormatToUnder
End Sub

' =========================================================
' ���G�N�X�|�[�g�i���j�{�^���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdExportToRight_Click()

    exportProcess recFormatToRight
End Sub

' =========================================================
' ���G�N�X�|�[�g����
'
' �T�v�@�@�@�F
' �����@�@�@�FrecFormat �s�t�H�[�}�b�g
' �߂�l�@�@�F
'
' =========================================================
Private Sub exportProcess(ByVal recFormat As REC_FORMAT)

    On Error GoTo err
    
    Dim exportTargets As ValCollection
    Set exportTargets = tableInfoList.selectedList
    
    If exportTargets.count <= 0 Then
        err.Raise ERR_NUMBER_NOT_SELECTED_TABLE _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_TABLE _
                , err.HelpFile _
                , err.HelpContext
        Exit Sub
    End If
    
    RaiseEvent export(exportTargets, recFormat)

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ������{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdClose_Click()

    On Error GoTo err
    
    ' �t�H�[�������
    HideExt
    
    ' ����C�x���g�𑗐M����
    RaiseEvent closed

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
    
    ' ���X�g�n�R���g���[���̏�����
    Set schemaInfoList = New CntListBox: schemaInfoList.init cboSchema
    Set tableInfoList = New CntListBox: tableInfoList.init lstTable

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

    ' DB�G�N�X�v���[���I�v�V������ǂݍ���
    restoreDBExplorerOption

    ' �X�L�[�}�V�[�g��ǂݍ���
    readSchemaInfo
    ' �e�[�u���V�[�g��ǂݍ���
    readTableInfo
    
    ' �t�B���^������K�p����
    applyFilterCondition
    
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
' ��DB�G�N�X�v���[���I�v�V������ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeDBExplorerOption()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' �������݃f�[�^
    Dim values As New ValCollection
    
    values.setItem Array(cboSchema.name, cboSchema.value)
    values.setItem Array(cboFilter.name, cboFilter.value)

    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DB�G�N�X�v���[���I�v�V������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreDBExplorerOption()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties

    ' �f�[�^��ǂݍ���
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
            
    inFilterProcess = True
        
    val = values.getItem(cboSchema.name, vbVariant): If IsArray(val) Then cboSchema.value = val(2)
    val = values.getItem(cboFilter.name, vbVariant): If IsArray(val) Then cboFilter.value = val(2)

    inFilterProcess = False
    
    Exit Sub
    
err:

    inFilterProcess = False
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ���X�L�[�}����ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub readSchemaInfo()

    On Error GoTo err
    
    Dim var As ValCollection
    
    If ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        ' �ؒf���
        
        Set var = New ValCollection
        addSchemaInfoList var
        
    Else
        ' �ڑ����
    
        ' �����Ԃ̏��������s�����̂Ń}�E�X�J�[�\���������v�ɂ���
        Dim cursorWait As New ExcelCursorWait: cursorWait.init
    
        ' �X�L�[�}��`���擾����
        Dim dbObjFactory As New DbObjectFactory
        
        Dim dbInfo As IDbMetaInfoGetter
        Set dbInfo = dbObjFactory.createMetaInfoGetterObject(dbConn)
           
        Set var = dbInfo.getSchemaList
        
        ' �X�L�[�}���X�g�{�b�N�X�Ƀ��X�g��ǉ�����
        addSchemaInfoList var
        
        ' �����Ԃ̏������I�������̂Ń}�E�X�J�[�\�������ɖ߂�
        cursorWait.destroy
        
    End If

    Exit Sub
    
err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ���e�[�u������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub readTableInfo()

    On Error GoTo err

    Dim var  As ValCollection

    If ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        ' �ؒf���
        
        Set var = New ValCollection
        addTableInfoList var
        
        Set tableWithoutFilterList = var.copy
        
    Else
        ' �ڑ����

        ' �I���ς݂̃X�L�[�}�����擾
        If schemaInfoList.count > 0 Then
        
            ' �����Ԃ̏��������s�����̂Ń}�E�X�J�[�\���������v�ɂ���
            Dim cursorWait As New ExcelCursorWait: cursorWait.init
        
            If schemaInfoList.getSelectedIndex = -1 Then
                ' �I�����Ȃ��ꍇ�́A�擪��I����Ԃɂ���
                schemaInfoList.setSelectedIndex 0
            End If
            
            Dim selectedSchemaList As New ValCollection
            Dim selectedSchema As ValDbDefineSchema
            Set selectedSchema = schemaInfoList.getSelectedItem(vbObject)
            selectedSchemaList.setItem selectedSchema
            
            ' �e�[�u����`���擾����
            Dim dbObjFactory As New DbObjectFactory
            
            Dim dbInfo As IDbMetaInfoGetter
            Set dbInfo = dbObjFactory.createMetaInfoGetterObject(dbConn)
            
            Set var = dbInfo.getTableList(selectedSchemaList)
            
            ' �e�[�u�����X�g�{�b�N�X�Ƀ��X�g��ǉ�����
            addTableInfoList var
            
            Set tableWithoutFilterList = var.copy
            
            ' �����Ԃ̏������I�������̂Ń}�E�X�J�[�\�������ɖ߂�
            cursorWait.destroy
            
        Else
            ' �X�L�[�}�����݂��Ȃ��ꍇ
            Set var = New ValCollection
            addTableInfoList var
        
            Set tableWithoutFilterList = var.copy
        End If
    End If

    Exit Sub
    
err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ���e�[�u�����X�g���t�B���^���鏈��
'
' �T�v�@�@�@�F�e�[�u�����X�g���t�B���^���鏈��
' �����@�@�@�FfilterKeyword         �t�B���^�L�[���[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub filterTableInfoList(ByVal filterKeyword As String)

    Dim filterTableInfoList As ValCollection
    Set filterTableInfoList = VBUtil.filterWildcard(tableWithoutFilterList, "tableName", filterKeyword)
    
    addTableInfoList filterTableInfoList, False

End Sub

' =========================================================
' ���e�[�u�����X�g���t�B���^���鏈���i���K�\���Łj
'
' �T�v�@�@�@�F�e�[�u�����X�g���t�B���^���鏈��
' �����@�@�@�FfilterKeyword         �t�B���^�L�[���[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub filterTableInfoListForRegExp(ByVal filterKeyword As String)

    Dim filterTableInfoList As ValCollection
    Set filterTableInfoList = VBUtil.filterRegExp(tableWithoutFilterList, "tableName", filterKeyword)
    
    addTableInfoList filterTableInfoList, False

End Sub

' =========================================================
' ���X�L�[�}���X�g��ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvalSchemaInfoList �X�L�[�}���X�g
'     �@�@�@  isAppend          �ǉ��L���t���O
' �߂�l�@�@�F
'
' =========================================================
Private Sub addSchemaInfoList(ByVal valSchemaInfoList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    schemaInfoList.addAll valSchemaInfoList _
                       , "schemaName" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ���e�[�u�����X�g��ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvaltableInfoList �e�[�u�����X�g
'     �@�@�@  isAppend     �ǉ��L���t���O
' �߂�l�@�@�F
'
' =========================================================
Private Sub addTableInfoList(ByVal valTableInfoList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    tableInfoList.addAll valTableInfoList _
                       , "tableName", "tableComment" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ���e�[�u����ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�Ftable �e�[�u��
' �߂�l�@�@�F
'
' =========================================================
Private Sub addTable(ByVal table As ValDbDefineTable)
    
    tableInfoList.addItemByProp table, "tableName", "tableComment"
    
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
Private Sub setTable(ByVal index As Long, ByVal rec As ValDbDefineTable)
    
    tableInfoList.setItem index, rec, "tableName", "tableComment"
    
End Sub
