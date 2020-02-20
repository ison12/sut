VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryResult 
   Caption         =   "�N�G������"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15900
   OleObjectBlob   =   "frmQueryResult.frx":0000
End
Attribute VB_Name = "frmQueryResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' �N�G�����ʃt�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/01/18�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' ���e�[�u����I�������ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FtableSheet �e�[�u���V�[�g
'           : row        �s�ԍ�
'
' =========================================================
Public Event selectedDetail(ByRef tableSheet As ValTableWorksheet, ByVal cell As String)

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

' �N�G�����ʏڍ׏��̈ꌏ���̕\���i�q��ʁj
Private WithEvents frmQueryResultDetailVar As frmQueryResultDetail
Attribute frmQueryResultDetailVar.VB_VarHelpID = -1

' �e�[�u�����X�g�ł̑I�����ڃC���f�b�N�X
Private tableSheetSelectedIndex As Long
' �e�[�u�����X�g�ł̑I�����ڃI�u�W�F�N�g
Private tableSheetSelectedItem As ValDbQueryBatchTableWorksheet

' �N�G�����ʏ�񃊃X�g
Private queryResultSetInfoParam As ValQueryResultSetInfo
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
' �����@�@�@�Fmodal              ���[�_���܂��̓��[�h���X�\���w��
'             queryResultSetInfo �N�G�����ʃZ�b�g���
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal queryResultSetInfo As ValQueryResultSetInfo)

    ' �p�����[�^�ݒ�
    Set queryResultSetInfoParam = queryResultSetInfo

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
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' �{�����ł͏������̂��L�����Z������
        Cancel = True
        ' �ȉ��̃C�x���g�o�R�ŕ���
        cmdClose_Click
    End If
    
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
' ���ڍ׃{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDetail_Click()


    Dim selectedList As ValCollection
    Set selectedList = tableSheetList.getSelectedList
    
    If selectedList.count <= 0 Then
    
        ' �I������
        Exit Sub
    End If

    Dim queryResultInfo As ValQueryResultInfo
    Set queryResultInfo = selectedList.getItemByIndex(1)

    If VBUtil.unloadFormIfChangeActiveBook(frmQueryResultDetail) Then Unload frmQueryResultDetail
    Load frmQueryResultDetail
    Set frmQueryResultDetailVar = frmQueryResultDetail
    frmQueryResultDetail.ShowExt vbModal, queryResultInfo
                            
    Set frmQueryResultDetailVar = Nothing
    
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
' ���N�G�����ʏڍׂ̑I�����̃C�x���g����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmQueryResultDetailVar_selected(tableSheet As ValTableWorksheet, ByVal cell As String)

    RaiseEvent selectedDetail(tableSheet, cell)
End Sub

' =========================================================
' ���N�G�����ʏڍׂ̕��鏈��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmQueryResultDetailVar_closed()

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

    ' ����{�^�����\���ɂ���
    cmdClose.Width = 0
    
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
    
    lblErrorMessage.visible = False
    
    Dim queryResultInfo As ValQueryResultInfo
    
    ' �e�[�u���V�[�g���X�g�ɕ\�����𔽉f����
    Set tableSheetList = New CntListBox: tableSheetList.init lstTableSheet
    addTableSheetList queryResultSetInfoParam.queryResultInfoList

    Dim i As Long: i = 0
    Dim selectedIndex As Long: selectedIndex = -1
    
    For Each queryResultInfo In tableSheetList.collection.col
    
        If queryResultInfo.sheetName = ActiveSheet.name Then
            selectedIndex = i
        End If
    
        i = i + 1
    Next
    
    If selectedIndex <> -1 Then
        ' �A�N�e�B�u�V�[�g��I����Ԃɂ���
        tableSheetList.setSelectedIndex selectedIndex
    End If

    ' �G���[������ꍇ�ɁA�G���[���b�Z�[�W��\������
    Dim erroredResultInfoCount As Long
    
    erroredResultInfoCount = 0
    For Each queryResultInfo In tableSheetList.collection.col
    
        If queryResultInfo.errorCount > 0 Then
        
            erroredResultInfoCount = erroredResultInfoCount + 1
        End If
    
    Next
    
    If erroredResultInfoCount > 0 Then
    
        lblErrorMessage.visible = True
        lblErrorMessage.Caption = "�������ʂɃG���[������܂��B�Ώۂ̃V�[�g��I�����ăG���[���e���m�F���Ă��������B"
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
' ���e�[�u���I�����̏���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub selectedTable()

    Dim selectedList As ValCollection
    
    Dim queryResultInfo As ValQueryResultInfo
    Dim tableSheet      As ValTableWorksheet

    Set selectedList = tableSheetList.getSelectedList

    If selectedList.count >= 1 Then
    
        Set queryResultInfo = selectedList.getItemByIndex(1)
        
        If Not queryResultInfo.tableWorksheet Is Nothing Then
            Set tableSheet = queryResultInfo.tableWorksheet
            RaiseEvent selected(tableSheet)
        End If
        
    End If

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
Private Sub addTableSheetList(ByVal valTableSheetList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    tableSheetList.addAll valTableSheetList _
                       , "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName", "processErrorCount" _
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
    
    tableSheetList.addItemByProp tableSheet, "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName", "processErrorCount"
    
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
    
    tableSheetList.setItem index, rec, "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName", "processErrorCount"
    
End Sub


