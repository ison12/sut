VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryResultDetail 
   Caption         =   "�N�G�����ʏڍ�"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15900
   OleObjectBlob   =   "frmQueryResultDetail.frx":0000
End
Attribute VB_Name = "frmQueryResultDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �N�G�����ʏڍ׃t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/02/19�@�V�K�쐬
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
Public Event selected(ByRef tableSheet As ValTableWorksheet, ByVal cell As String)

' =========================================================
' ������{�^���������ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event closed()

' �ڍ׏�񃊃X�g�ł̑I�����ڃC���f�b�N�X
Private detailInfoSelectedIndex As Long
' �ڍ׏�񃊃X�g�ł̑I�����ڃI�u�W�F�N�g
Private detailInfoSelectedItem As ValQueryResultDetailInfo

' �N�G�����ʏ��
Private queryResultInfoParam As ValQueryResultInfo
' �ڍ׏�񃊃X�g
Private detailInfoList  As CntListBox

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
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal queryResultInfo As ValQueryResultInfo)

    ' �p�����[�^�ݒ�
    Set queryResultInfoParam = queryResultInfo

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
' ���ڍ׏�񃊃X�g�@�I�����ύX���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstDetailInfo_Change()

    selectedTable
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
' ���ڍ׏��̃R�s�[�N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDetailInfoCopy_Click()

    Dim selectedIndex As Long
    Dim selectedItem As ValQueryResultDetailInfo
    
    ' ���ݑI������Ă���C���f�b�N�X���擾
    selectedIndex = detailInfoList.getSelectedIndex

    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    Set selectedItem = detailInfoList.getSelectedItem
    
    WinAPI_Clipboard.SetClipboard selectedItem.tabbedInfoHeader & vbNewLine & getDetailInfoForClipboardFormat(selectedItem)
    
End Sub

' =========================================================
' ���ڍ׏��̑S�ăR�s�[�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdAllDetailInfoCopy_Click()

    Dim data As New StringBuilder
    Dim var As Variant
    
    Dim i As Long
    
    For Each var In detailInfoList.collection.col
        If i <= 0 Then
            data.append var.tabbedInfoHeader & vbNewLine
        End If
        data.append getDetailInfoForClipboardFormat(var)
        i = i + 1
    Next
    
    WinAPI_Clipboard.SetClipboard data.str

End Sub

' =========================================================
' ���ڍ׏��̃N���b�v�{�[�h�t�H�[�}�b�g�`��������擾
'
' �T�v�@�@�@�F�ڍ׏��̃N���b�v�{�[�h�t�H�[�}�b�g�`����������擾����B
' �����@�@�@�Fvar �ڍ׏��
' �߂�l�@�@�F�ڍ׏��̃N���b�v�{�[�h�t�H�[�}�b�g�`��������擾
'
' =========================================================
Private Function getDetailInfoForClipboardFormat(ByVal var As ValQueryResultDetailInfo) As String

    getDetailInfoForClipboardFormat = var.tabbedInfo & vbNewLine

End Function

' =========================================================
' ������������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub initial()

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
    
    Dim queryResultDetailInfo As ValQueryResultDetailInfo
    
    ' �ڍ׏�񃊃X�g�ɕ\�����𔽉f����
    Set detailInfoList = New CntListBox: detailInfoList.init lstDetailInfo
    addDetailInfoList queryResultInfoParam.detailList

    detailInfoList.setSelectedIndex 0
    
    txtSheetName.value = queryResultInfoParam.sheetNameOrSheetTableName

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

    Dim selected As ValQueryResultDetailInfo
    Set selected = detailInfoList.getSelectedItem

    If Not selected Is Nothing Then
        RaiseEvent selected(queryResultInfoParam.tableWorksheet, selected.cell)
    End If

End Sub

' =========================================================
' ���ڍ׏�񃊃X�g��ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvalDetailInfoList     �ڍ׏�񃊃X�g
'     �@�@�@  isAppend              �ǉ��L���t���O
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDetailInfoList(ByVal valDetailInfoList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    detailInfoList.addAll valDetailInfoList _
                       , "cell", "messageWithSqlState", "queryWithoutNewLine" _
                       , isAppend:=isAppend
    
End Sub


