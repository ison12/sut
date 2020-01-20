VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBQueryBatchTypeSetting 
   Caption         =   "�N�G���ꊇ���s�̃N�G����ޕύX"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7065
   OleObjectBlob   =   "frmDBQueryBatchTypeSetting.frx":0000
End
Attribute VB_Name = "frmDBQueryBatchTypeSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�N�G���o�b�`�̃N�G����ނ̈ꌏ���̕ҏW�i�q��ʁj
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2019/12/08�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FdbQueryBatchType DB�N�G���o�b�`���
'
' =========================================================
Public Event ok(ByVal dbQueryBatchType As DB_QUERY_BATCH_TYPE)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event cancel()

' �V�[�g���i�t�H�[���\�����_�ł̏��j
Private sheetNameParam As String
' DB�N�G���o�b�`��ށi�t�H�[���\�����_�ł̏��j
Private dbQueryBatchTypeParam As DB_QUERY_BATCH_TYPE
' DB�N�G���o�b�`��ނ̑I�������X�g
Private dbQueryBatchTypeSelectList As ValCollection
' DB�N�G���o�b�`��ރR���{�{�b�N�X���X�g
Private dbQueryBatchTypeComboList As CntListBox

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@sheetName                     �V�[�g��
' �@�@�@�@�@�@dbQueryBatchType              DB�N�G���o�b�`��ނ̏����l
' �@�@�@�@�@�@valDbQueryBatchTypeSelectList DB�N�G���o�b�`��ނ̑I�������X�g
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                    , ByVal sheetName As String _
                    , ByVal dbQueryBatchType As DB_QUERY_BATCH_TYPE _
                    , ByVal valDbQueryBatchTypeSelectList As ValCollection)

    ' �p�����[�^��ݒ�
    sheetNameParam = sheetName
    dbQueryBatchTypeParam = dbQueryBatchType
    Set dbQueryBatchTypeSelectList = valDbQueryBatchTypeSelectList

    activate
    
    ' �f�t�H���g�t�H�[�J�X�R���g���[����ݒ肷��
    cboDbQueryBatchType.SetFocus
    
    Main.restoreFormPosition Me.name, Me
    Me.Show vbModal

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

    ' �V�[�g����ݒ肷��
    txtSheetName.text = sheetNameParam
    
    ' DB�N�G���o�b�`��ނ̃R���{�{�b�N�X������������
    Set dbQueryBatchTypeComboList = New CntListBox: dbQueryBatchTypeComboList.init cboDbQueryBatchType
    dbQueryBatchTypeComboList.addAll dbQueryBatchTypeSelectList, "dbQueryBatchTypeName"
    
    ' DB�N�G���o�b�`��ރR���{�{�b�N�X�̃A�N�e�B�u�ȑI�����ڂ�ݒ肷��
    Dim v As ValDbQueryBatchType
    Dim i As Long
    
    i = 0
    For Each v In dbQueryBatchTypeComboList.collection.col
        
        If v.dbQueryBatchType = dbQueryBatchTypeParam Then
            dbQueryBatchTypeComboList.setSelectedIndex i
            Exit For
        End If
        
        i = i + 1
    Next
    
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
' ��OK�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err
    
    ' ���I���̏ꍇ
    If dbQueryBatchTypeComboList.getSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If
    
    ' �I�����������Ȃ��̏ꍇ
    If dbQueryBatchTypeComboList.getSelectedItem.dbQueryBatchType = DB_QUERY_BATCH_TYPE.none Then
    
        ' �I������
        Exit Sub
    End If
    
    ' �t�H�[�������
    HideExt

    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(dbQueryBatchTypeComboList.getSelectedItem.dbQueryBatchType)
    
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



