VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBColumnFormat 
   Caption         =   "DB�J���������ݒ�"
   ClientHeight    =   8550.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15105
   OleObjectBlob   =   "frmDBColumnFormat.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmDBColumnFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�J���������ݒ�t�H�[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2020/01/14�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfo DB�ڑ����
'
' =========================================================
Public Event ok(ByVal dbColumnFormatInfo As ValDbColumnFormatInfo)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event cancel()

' DB�J���������ҏW�t�H�[��
Private WithEvents frmDBColumnFormatSettingVar As frmDBColumnFormatSetting
Attribute frmDBColumnFormatSettingVar.VB_VarHelpID = -1

' DB�J���������ݒ��񃊃X�g�i�t�H�[���\�����_�ł̏��j
Private dbColumnFormatInfoParam As ValDbColumnFormatInfo

' DB�J���������ݒ��񃊃X�g �R���g���[��
Private dbColumnFormatList As CntListBox

' DB�J���������ݒ��񃊃X�g�ł̑I�����ڃC���f�b�N�X
Private dbColumnFormatSelectedIndex As Long
' DB�J���������ݒ��񃊃X�g�ł̑I�����ڃI�u�W�F�N�g
Private dbColumnFormatSelectedItem As ValDbColumnTypeColInfo

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
'     �@�@�@�Finfo  DB�J�����������
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal info As ValDbColumnFormatInfo)

    Set dbColumnFormatInfoParam = info

    activate
    
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

    Set dbColumnFormatList = New CntListBox: dbColumnFormatList.init lstDbColumnFormatList
    addDbColumnFormatList dbColumnFormatInfoParam.columnList
    
    ' �擪��I������
    dbColumnFormatList.setSelectedIndex 0
    
    lblDbName.Caption = DBUtil.getDbmsTypeName(dbColumnFormatInfoParam.dbName)

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

    ' Nothing��ݒ肷�邱�ƂŃC�x���g����M���Ȃ��悤�ɂ���
    Set frmDBColumnFormatSettingVar = Nothing
    
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

    ' DB�I�u�W�F�N�g�����N���X
    Dim dbObjFactory As New DbObjectFactory
    ' �J�����������擾�I�u�W�F�N�g
    Dim dbColumnType As IDbColumnType
    ' �J�����������̃f�t�H���g�l�����擾����
    Set dbColumnType = dbObjFactory.createColumnType(dbColumnFormatInfoParam.dbName)
    
    ' ���f����
    addDbColumnFormatList dbColumnType.getDefaultColumnFormat

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
    
    ' �t�H�[�������
    HideExt

    ' OK�C�x���g���M���ɐݒ肷����𐶐�����
    Dim var As New ValDbColumnFormatInfo
    var.dbName = dbColumnFormatInfoParam.dbName ' DB��
    Set var.columnList = dbColumnFormatList.collection ' ���X�g�{�b�N�X���̏����擾���Đݒ�
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(var)
    
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
' ���V�K�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdAdd_Click()

    ' ���X�g�{�b�N�X�̃T�C�Y
    Dim cnt As Long
    ' ���X�g�{�b�N�X�̃T�C�Y���擾����
    cnt = dbColumnFormatList.collection.count
    
    ' �|�b�v�A�b�v���j���[�I�u�W�F�N�g�����X�g�ɒǉ�����
    Dim dbColumnFormat As ValDbColumnTypeColInfo
    Set dbColumnFormat = New ValDbColumnTypeColInfo
    
    dbColumnFormat.columnName = ConstantsCommon.DB_COLUMN_FORMAT_DEFAULT_NAME & " " & (cnt + 1)
    
    addDbColumnFormat dbColumnFormat
    
    dbColumnFormatList.setSelectedIndex cnt
    dbColumnFormatList.control.SetFocus
    
End Sub

' =========================================================
' ���ҏW�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdEdit_Click()

    editDbColumnFormat
End Sub

Private Sub editDbColumnFormat()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    dbColumnFormatSelectedIndex = dbColumnFormatList.getSelectedIndex

    ' ���I���̏ꍇ
    If dbColumnFormatSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set dbColumnFormatSelectedItem = dbColumnFormatList.getSelectedItem
    
    Load frmDBColumnFormatSetting
    Set frmDBColumnFormatSettingVar = frmDBColumnFormatSetting
    frmDBColumnFormatSettingVar.ShowExt vbModal, dbColumnFormatSelectedItem
                            
    Set frmDBColumnFormatSettingVar = Nothing

End Sub

' =========================================================
' ��DB�J���������ݒ�i�q�t�H�[���j��OK�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmDBColumnFormatSettingVar_ok(ByVal dbColumnTypeColInfo As ValDbColumnTypeColInfo)

    Dim v As ValDbColumnTypeColInfo
    Set v = dbColumnFormatList.getItem(dbColumnFormatSelectedIndex)
    
    v.columnName = dbColumnTypeColInfo.columnName
    v.formatUpdate = dbColumnTypeColInfo.formatUpdate
    v.formatSelect = dbColumnTypeColInfo.formatSelect

    setDbColumnFormat dbColumnFormatSelectedIndex, v
    
    dbColumnFormatList.control.SetFocus
End Sub

' =========================================================
' ��DB�J���������ݒ�i�q�t�H�[���j�̃L�����Z���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmDBColumnFormatSettingVar_cancel()

End Sub

' =========================================================
' ���폜�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDelete_Click()

    Dim selectedIndex As Long
    
    ' ���ݑI������Ă���C���f�b�N�X���擾
    selectedIndex = dbColumnFormatList.getSelectedIndex

    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    dbColumnFormatList.removeItem selectedIndex
    dbColumnFormatList.control.SetFocus
    
End Sub

' =========================================================
' ����փ{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdUp_Click()

    On Error GoTo err
    
    ' �I���ς݃C���f�b�N�X
    Dim selectedIndex As Long
    
    ' ���݃��X�g�őI������Ă���C���f�b�N�X���擾����
    selectedIndex = dbColumnFormatList.getSelectedIndex
    
    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
        ' �I������
        Exit Sub
    End If

    If selectedIndex > 0 Then
    
        dbColumnFormatList.swapItem _
                          selectedIndex _
                        , selectedIndex - 1 _
                        , vbObject _
                        , 1
                              
        dbColumnFormatList.setSelectedIndex selectedIndex - 1
            
    End If
    
    dbColumnFormatList.control.SetFocus
        
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
Private Sub cmdDown_Click()

    On Error GoTo err
    
    ' �I���ς݃C���f�b�N�X
    Dim selectedIndex As Long
    
    ' ���݃��X�g�őI������Ă���C���f�b�N�X���擾����
    selectedIndex = dbColumnFormatList.getSelectedIndex
    
        ' ���I���̏ꍇ
    If selectedIndex = -1 Then
        ' �I������
        Exit Sub
    End If

    If selectedIndex < dbColumnFormatList.count - 1 Then
    
        dbColumnFormatList.swapItem _
                          selectedIndex _
                        , selectedIndex + 1 _
                        , vbObject _
                        , 1
                              
        dbColumnFormatList.setSelectedIndex selectedIndex + 1
            
    End If
    
    dbColumnFormatList.control.SetFocus
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
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
' ���t�H�[���f�B�A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Deactivate()

End Sub

' =========================================================
' ���t�H�[�����鎞�̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)

    deactivate

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

' =========================================================
' ��DB�J���������ݒ����ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�Flist DB�J���������ݒ��񃊃X�g
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDbColumnFormatList(ByVal list As ValCollection)
    
    dbColumnFormatList.addAll list, "columnName", "formatUpdate", "formatSelect"
    
End Sub

' =========================================================
' ��DB�J���������ݒ����ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�Frec DB�J���������ݒ���
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDbColumnFormat(ByVal rec As ValDbColumnTypeColInfo)
    
    dbColumnFormatList.addItemByProp rec, "columnName", "formatUpdate", "formatSelect"
    
End Sub

' =========================================================
' ��DB�J���������ݒ����ύX
'
' �T�v�@�@�@�F
' �����@�@�@�Findex �C���f�b�N�X
'     �@�@�@  rec   DB�J���������ݒ���
' �߂�l�@�@�F
'
' =========================================================
Private Sub setDbColumnFormat(ByVal index As Long, ByVal rec As ValDbColumnTypeColInfo)
    
    dbColumnFormatList.setItem index, rec, "columnName", "formatUpdate", "formatSelect"
    
End Sub
