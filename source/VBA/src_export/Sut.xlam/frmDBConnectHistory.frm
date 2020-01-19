VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectHistory 
   Caption         =   "DB�ڑ��̗���"
   ClientHeight    =   9120.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   OleObjectBlob   =   "frmDBConnectHistory.frx":0000
End
Attribute VB_Name = "frmDBConnectHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�ڑ������t�H�[��
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
Public Event ok(ByVal connectInfo As ValDBConnectInfo)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event cancel()

' ���W�X�g���L�[
Private Const REG_SUB_KEY_DB_CONNECT_HISTORY As String = "db_history"

' DB�ڑ��̗�����񃊃X�g �R���g���[��
Private dbConnectHistoryList As CntListBox
'DB�ڑ��̂��C�ɓ����񃊃X�g�i�t�B���^�����K�p�Ȃ��j
Private dbConnectHistoryWithoutFilterList As ValCollection

' DB�ڑ��̗�����񃊃X�g�ł̑I�����ڃC���f�b�N�X
Private dbConnectHistorySelectedIndex As Long
' DB�ڑ��̗�����񃊃X�g�ł̑I�����ڃI�u�W�F�N�g
Private dbConnectHistorySelectedItem As ValDBConnectInfo

Private inFilterProcess As Boolean


' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants)

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

    restoreDbConnectHistory
    
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
        
        filterConnectList "*" & currentFilterText & "*"
    
        inFilterProcess = False

    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
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
    
    ' ���ݑI������Ă���C���f�b�N�X���擾
    dbConnectHistorySelectedIndex = dbConnectHistoryList.getSelectedIndex

    ' ���I���̏ꍇ
    If dbConnectHistorySelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If
    
    ' �t�H�[�������
    HideExt

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set dbConnectHistorySelectedItem = dbConnectHistoryList.getSelectedItem
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(dbConnectHistorySelectedItem)
    
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
' ��DB�ڑ��������X�g�̃_�u���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstDbConnectHistoryList_DblClick(ByVal cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

' =========================================================
' ��DB�ڑ��������X�g �L�[�������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstDbConnectHistoryList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        cmdOk_Click
    End If
    
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
' ��DB�ڑ��̂��C�ɓ��������ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeDBConnectHistory()

    On Error GoTo err
    
    Dim i, j As Long
    ' ���W�X�g������N���X
    Dim registry As RegistryManipulator
    
    ' -------------------------------------------------------
    ' �S�Ă̏������W�X�g�������U�폜����
    ' -------------------------------------------------------
    ' ���W�X�g������N���X������������
    Set registry = New RegistryManipulator
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_HISTORY) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    Dim key     As Variant
    Dim keyList As ValCollection
    
    Set keyList = registry.getKeyList
    
    For Each key In keyList.col
        registry.delete key
    Next
    
    ' -------------------------------------------------------
    ' �S�Ă̏������W�X�g���ɕۑ�����
    ' -------------------------------------------------------
    Dim dbConnectInfo As ValDBConnectInfo
    Dim dbConnectFavoriteArray(0 To 9 _
                             , 0 To 1) As Variant
    
    i = 0
     For Each dbConnectInfo In dbConnectHistoryList.collection.col
        
        ' ���W�X�g������N���X������������
        Set registry = New RegistryManipulator
        registry.init RegKeyConstants.HKEY_CURRENT_USER _
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_HISTORY & "\" & i) _
                    , RegAccessConstants.KEY_ALL_ACCESS _
                    , True

        j = 0
        dbConnectFavoriteArray(j, 0) = "no"
        dbConnectFavoriteArray(j, 1) = j: j = j + 1
        dbConnectFavoriteArray(j, 0) = "name"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.name: j = j + 1
        dbConnectFavoriteArray(j, 0) = "type"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.type_: j = j + 1
        dbConnectFavoriteArray(j, 0) = "dsn"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.dsn: j = j + 1
        dbConnectFavoriteArray(j, 0) = "host"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.host: j = j + 1
        dbConnectFavoriteArray(j, 0) = "port"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.port: j = j + 1
        dbConnectFavoriteArray(j, 0) = "db"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.db: j = j + 1
        dbConnectFavoriteArray(j, 0) = "user"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.user: j = j + 1
        dbConnectFavoriteArray(j, 0) = "password"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.password: j = j + 1
        dbConnectFavoriteArray(j, 0) = "option"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.option_: j = j + 1
        
        ' ���W�X�g���ɏ���ݒ肷��
        registry.setValues dbConnectFavoriteArray
    
        Set registry = Nothing

        
        i = i + 1
    Next

        
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DB�ڑ��̗���������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreDbConnectHistory()

    On Error GoTo err
    
    ' ���C�ɓ���̐ڑ����
    Dim connectInfoList As ValCollection
    Set connectInfoList = New ValCollection
    Dim connectInfo As ValDBConnectInfo
    
    ' ���W�X�g������N���X
    Dim registry As New RegistryManipulator
                
    Dim key     As Variant
    Dim keyList As ValCollection

    ' -------------------------------------------------------
    ' �S�Ă̏������W�X�g������擾����i�C���f�b�N�X�ԍ����X�g�̎擾�j
    ' -------------------------------------------------------
    ' ���W�X�g������N���X������������
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_HISTORY) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Set keyList = registry.getKeyList

    Set registry = Nothing
    
    ' -------------------------------------------------------
    ' �S�Ă̏ڍ׏������W�X�g������擾����
    ' -------------------------------------------------------
    Dim valueNameList As ValCollection
    Dim valueList As ValCollection
    
    For Each key In keyList.col
    
        ' ���W�X�g������N���X������������
        Set registry = New RegistryManipulator
        registry.init RegKeyConstants.HKEY_CURRENT_USER _
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_HISTORY & "\" & key) _
                    , RegAccessConstants.KEY_ALL_ACCESS _
                    , True
                    
        registry.getValueList valueNameList, valueList

        Set connectInfo = New ValDBConnectInfo
        connectInfo.name = valueList.getItem("name", vbVariant)
        connectInfo.type_ = valueList.getItem("type", vbVariant)
        connectInfo.dsn = valueList.getItem("dsn", vbVariant)
        connectInfo.host = valueList.getItem("host", vbVariant)
        connectInfo.port = valueList.getItem("port", vbVariant)
        connectInfo.db = valueList.getItem("db", vbVariant)
        connectInfo.user = valueList.getItem("user", vbVariant)
        connectInfo.password = valueList.getItem("password", vbVariant)
        connectInfo.option_ = valueList.getItem("option", vbVariant)
        
        connectInfoList.setItem connectInfo
                    
        Set registry = Nothing
    Next
    
    Set dbConnectHistoryList = New CntListBox: dbConnectHistoryList.init lstDbConnectHistoryList
    addDbConnectHistoryList connectInfoList
    Set dbConnectHistoryWithoutFilterList = dbConnectHistoryList.collection.copy
        
    ' �擪��I������
    dbConnectHistoryList.setSelectedIndex 0

    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ��DB�ڑ��̗���������ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfo DB�ڑ����
' �߂�l�@�@�F
'
' =========================================================
Public Function registDbConnectInfo(ByVal connectInfo As ValDBConnectInfo)

    On Error GoTo err
    
    ' -------------------------------------------------------
    ' DB�ڑ��������ă��[�h���čŐV�ɂ���
    ' -------------------------------------------------------
    restoreDbConnectHistory
    
    ' -------------------------------------------------------
    ' �d������菜�����������𐶐�����
    ' -------------------------------------------------------
    Dim dbConnectHistoryDistinctList As New ValCollection
    Dim dbConnectHistory As ValDBConnectInfo
    
    For Each dbConnectHistory In dbConnectHistoryList.collection.col
        
        If dbConnectHistory.displayName <> connectInfo.displayName Then
            ' �~���ŕ\������̂ŁA�ǉ�����v�f�͐擪�ɒǉ����Ă���
            dbConnectHistoryDistinctList.setItem dbConnectHistory
        End If
        
    Next
    
    ' -------------------------------------------------------
    ' ��������擪�ɒǉ�����
    ' -------------------------------------------------------
    dbConnectHistoryDistinctList.setItemByIndexBefore connectInfo, 1
    
    ' -------------------------------------------------------
    ' DB�ڑ������ɏd������菜�������X�g�œ���ւ���
    ' -------------------------------------------------------
    dbConnectHistoryList.removeAll
    addDbConnectHistoryList dbConnectHistoryDistinctList
    
    ' -------------------------------------------------------
    ' DB�ڑ�������ۑ�����
    ' -------------------------------------------------------
    storeDBConnectHistory

    Exit Function
    
err:

    Main.ShowErrorMessage

End Function

' =========================================================
' ��DB�ڑ��̗�������ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfoList DB�ڑ���񃊃X�g
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDbConnectHistoryList(ByVal connectInfoList As ValCollection)
    
    dbConnectHistoryList.addAll connectInfoList, "displayName"
    
End Sub

' =========================================================
' ��DB�ڑ��̗�������ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfo DB�ڑ����
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDbConnectHistory(ByVal connectInfo As ValDBConnectInfo)
    
    dbConnectHistoryList.addItemByProp connectInfo, "displayName"
    
End Sub

' =========================================================
' ���e�[�u���V�[�g���X�g���t�B���^���鏈��
'
' �T�v�@�@�@�F�e�[�u���V�[�g���X�g���t�B���^���鏈��
' �����@�@�@�FfilterKeyword         �t�B���^�L�[���[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub filterConnectList(ByVal filterKeyword As String)

    Dim filterConnectList As ValCollection
    Set filterConnectList = VBUtil.filterWildcard(dbConnectHistoryWithoutFilterList, "displayName", filterKeyword)
    
    addDbConnectHistoryList filterConnectList

End Sub
