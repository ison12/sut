VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetList 
   Caption         =   "�e�[�u���V�[�g�ꗗ"
   ClientHeight    =   8415.001
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10905
   OleObjectBlob   =   "frmTableSheetList.frx":0000
End
Attribute VB_Name = "frmTableSheetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �e�[�u���V�[�g�ꗗ�t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/03�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' ���e�[�u����I�������ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FtableSheet �e�[�u���V�[�g
'
' =========================================================
Public Event selected(ByRef tableSheet As ValTableWorksheet)

' =========================================================
' ������{�^���������ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event closed()

' �t�B���^�Ȃ���Ԃ̃e�[�u�����X�g
Private tableSheetWithoutFilterList As ValCollection
' �e�[�u�����X�g
Private tableSheetList  As CntListBox

Private inFilterProcess As Boolean

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@conn  DB�R�l�N�V����
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants)

    activate

    ' �f�t�H���g�t�H�[�J�X�R���g���[����ݒ肷��
    lstTableSheet.SetFocus

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
' ������������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub initial()

    ' �e�[�u���V�[�g���X�g������������
    Set tableSheetList = New CntListBox: tableSheetList.init lstTableSheet
    
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

    ' �e�[�u���V�[�g���X�g��j������
    Set tableSheetList = Nothing
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

    ' �e�[�u�����X�g
    Dim tableDistinctList As ValCollection
    Dim tableList As ValCollection
    Dim tableWorksheet As ValTableWorksheet
    
    ' �e�[�u���V�[�g�Ǎ��I�u�W�F�N�g
    Dim tableSheetReader As ExeTableSheetReader
    Set tableSheetReader = New ExeTableSheetReader
        
    ' �u�b�N
    Dim book  As Workbook
    ' �V�[�g
    Dim sheet As Worksheet
    
    ' �A�N�e�B�u�u�b�N��book�ϐ��Ɋi�[����
    Set book = ActiveWorkbook
    
    ' �e�[�u�����X�g������������
    Set tableList = New ValCollection
    Set tableSheetWithoutFilterList = New ValCollection
    
    Dim i As Long: i = 0
    Dim selectedIndex As Long: selectedIndex = -1
    
    ' �u�b�N�Ɋ܂܂�Ă���V�[�g��1������������
    For Each sheet In book.Worksheets
    
        Set tableSheetReader.sheet = sheet
        
        ' �ΏۃV�[�g���e�[�u���V�[�g�̏ꍇ
        If tableSheetReader.isTableSheet = True Then
        
            ' �e�[�u���V�[�g��ǂݍ���Ń��X�g�ɐݒ肷��i�e�[�u�����̂ݎ擾����j
            Set tableWorksheet = tableSheetReader.readTableInfo(True)
            
            tableList.setItem tableWorksheet
            tableSheetWithoutFilterList.setItem tableWorksheet
            
            If tableWorksheet.sheetName = ActiveSheet.name Then
                selectedIndex = i
            End If
        
            i = i + 1
        End If
    
    Next
    
    ' ���X�g�R���g���[���Ƀe�[�u���V�[�g����ǉ�����
    addTableSheetList tableList, False
    
    If selectedIndex <> -1 Then
        ' �A�N�e�B�u�V�[�g��I����Ԃɂ���
        tableSheetList.setSelectedIndex selectedIndex
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
' ���e�[�u���V�[�g���X�g�@�I�����ύX���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstTableSheet_Change()

    selectedTable
End Sub

' =========================================================
' ���e�[�u���V�[�g���X�g�X�V�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdTableSheetListUpdate_Click()

    activate
End Sub

' =========================================================
' ������{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub btnClose_Click()

    RaiseEvent closed

    Me.HideExt
    
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
        
        filterTableSheetList "*" & currentFilterText & "*"
        
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
            filterTableSheetListForRegExp "[^a-zA-Z]*"
            
            clearFilterCondition
            tglFilterOther.value = True
        Else
            filterTableSheetListForRegExp ""
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
            filterTableSheetList keyword & "*"
            
            clearFilterCondition
            state.value = True
        Else
            filterTableSheetList ""
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
' ���e�[�u���I�����̏���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub selectedTable()

    Dim selectedList As ValCollection
    
    Dim tableSheet As ValTableWorksheet

    Set selectedList = tableSheetList.selectedList

    If selectedList.count >= 1 Then
    
        Set tableSheet = selectedList.getItemByIndex(1)
        
        RaiseEvent selected(tableSheet)
    End If

End Sub

' =========================================================
' ���e�[�u���V�[�g���X�g���t�B���^���鏈��
'
' �T�v�@�@�@�F�e�[�u���V�[�g���X�g���t�B���^���鏈��
' �����@�@�@�FfilterKeyword         �t�B���^�L�[���[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub filterTableSheetList(ByVal filterKeyword As String)

    If filterKeyword = "" Then
        ' �t�B���^�������Ȃ��ꍇ�́A�S�Ă̏���\������
        tableSheetList.addAll tableSheetWithoutFilterList, "sheetNameOrSheetTableName", "TableComment"
        Exit Sub
    End If

    Dim filterTableSheetList As ValCollection
    Set filterTableSheetList = VBUtil.filterWildcard(tableSheetWithoutFilterList, "table.tableName", filterKeyword)
    
    addTableSheetList filterTableSheetList, False

End Sub


' =========================================================
' ���e�[�u���V�[�g���X�g���t�B���^���鏈���i���K�\���Łj
'
' �T�v�@�@�@�F�e�[�u���V�[�g���X�g���t�B���^���鏈��
' �����@�@�@�FfilterKeyword         �t�B���^�L�[���[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub filterTableSheetListForRegExp(ByVal filterKeyword As String)

    Dim filterTableSheetList As ValCollection
    Set filterTableSheetList = VBUtil.filterRegExp(tableSheetWithoutFilterList, "table.tableName", filterKeyword)
    
    addTableSheetList filterTableSheetList, False

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
                       , "sheetNameOrSheetTableName", "tableComment" _
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
    
    tableSheetList.addItemByProp tableSheet, "sheetNameOrSheetTableName", "tableComment"
    
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
    
    tableSheetList.setItem index, rec, "sheetNameOrSheetTableName", "tableComment"
    
End Sub
