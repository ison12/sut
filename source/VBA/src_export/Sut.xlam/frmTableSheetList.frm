VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetList 
   Caption         =   "�e�[�u���V�[�g�ꗗ"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10935
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
        btnClose_Click
    End If
    
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

    ' ����{�^�����\���ɂ���
    btnClose.Width = 0
    
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
    
    ' �t�B���^������K�p����
    cboFilter.text = ""
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
        
        'filterTableSheetList currentFilterText ' ���S��v
        filterTableSheetList "*" & currentFilterText & "*" ' ���Ԉ�v
        
        clearFilterCondition True
    
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

    Set selectedList = tableSheetList.getSelectedList

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
