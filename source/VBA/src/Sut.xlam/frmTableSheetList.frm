VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetList 
   Caption         =   "�e�[�u���V�[�g�ꗗ"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5940
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
' �쐬�ҁ@�FHideki Isobe
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

' �e�[�u�����X�g
Private tableSheetList  As CntListBox

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

    ' �őO�ʕ\���ɂ���
    ExcelUtil.setUserFormTopMost Me

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
    Set tableSheetList = New CntListBox: tableSheetList.control = lstTableSheet
    
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
    Dim tableList As ValCollection
    
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
    
    ' �u�b�N�Ɋ܂܂�Ă���V�[�g��1������������
    For Each sheet In book.Worksheets
    
        Set tableSheetReader.sheet = sheet
        
        ' �ΏۃV�[�g���e�[�u���V�[�g�̏ꍇ
        If tableSheetReader.isTableSheet = True Then
        
            ' �e�[�u���V�[�g��ǂݍ���Ń��X�g�ɐݒ肷��i�e�[�u�����̂ݎ擾����j
            tableList.setItem tableSheetReader.readTableInfo(True)
        
        End If
    
    Next
    
    ' ���X�g�R���g���[���Ƀe�[�u���V�[�g����ǉ�����
    tableSheetList.addNestedProperty tableList, "Table", "SchemaTableName", "TableComment"
    
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
' ���e�[�u���V�[�g���X�g�@�_�u���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstTableSheet_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    selectedTable
End Sub

' =========================================================
' ���e�[�u���V�[�g���X�g�@�L�[�������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstTableSheet_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
    
        selectedTable
    End If
    
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
