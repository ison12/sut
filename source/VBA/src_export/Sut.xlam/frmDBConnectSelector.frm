VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectSelector 
   Caption         =   "�ڑ����̑I��"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12630
   OleObjectBlob   =   "frmDBConnectSelector.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmDBConnectSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�ڑ��I���t�H�[��
'
' �쐬�ҁ@�FIson
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
Private Const REG_SUB_KEY_DB_CONNECT_FAVORITE As String = "db_favorite"
' ���W�X�g���L�[
Private Const REG_SUB_KEY_DB_CONNECT_HISTORY  As String = "db_history"

' �t�H�[�����[�h
Private formMode As DB_CONNECT_INFO_TYPE

' DB�ڑ���񃊃X�g �R���g���[��
Private dbConnectList As CntListBox
' DB�ڑ���񃊃X�g�i�t�B���^�����K�p�Ȃ��j
Private dbConnectWithoutFilterList As ValCollection

' DB�ڑ���񃊃X�g�ł̑I�����ڃC���f�b�N�X
Private dbConnectSelectedIndex As Long
' DB�ڑ���񃊃X�g�ł̑I�����ڃI�u�W�F�N�g
Private dbConnectSelectedItem As ValDBConnectInfo

Private inFilterProcess As Boolean


' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
'     �@�@�@  fm    �t�H�[�����[�h
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal fm As DB_CONNECT_INFO_TYPE)

    ' �t�H�[�����[�h
    formMode = fm

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
' ���t�H�[���A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    If formMode = DB_CONNECT_INFO_TYPE.favorite Then
        lblFormModeName.Caption = "�ݒ���"
    Else
        lblFormModeName.Caption = "�������"
    End If

    restoreDbConnectInfo formMode
    
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
    dbConnectSelectedIndex = dbConnectList.getSelectedIndex

    ' ���I���̏ꍇ
    If dbConnectSelectedIndex = -1 Then
        err.Raise ERR_NUMBER_NOT_SELECTED_DB_CONNECT _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_DB_CONNECT _
                , err.HelpFile _
                , err.HelpContext
        ' �I������
        Exit Sub
    End If
    
    ' �t�H�[�������
    HideExt

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set dbConnectSelectedItem = dbConnectList.getSelectedItem
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(dbConnectSelectedItem)
    
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
' ��DB�ڑ����X�g�̃_�u���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstDbConnectList_DblClick(ByVal cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

' =========================================================
' ��DB�ڑ����X�g �L�[�������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstDbConnectList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
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
' ��DB�ڑ�����ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�FformMode �t�H�[�����[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeDBConnectInfo(ByVal formMode As DB_CONNECT_INFO_TYPE)

    On Error GoTo err
    
    ' ���W�X�g���̃T�u�L�[�����肷��
    Dim regSubKey As String
    
    If formMode = DB_CONNECT_INFO_TYPE.favorite Then
        regSubKey = REG_SUB_KEY_DB_CONNECT_FAVORITE
    Else
        regSubKey = REG_SUB_KEY_DB_CONNECT_HISTORY
    End If
    
    Dim i, j As Long
    ' ���W�X�g������N���X
    Dim registry As RegistryManipulator
    
    ' -------------------------------------------------------
    ' �S�Ă̏������W�X�g�������U�폜����
    ' -------------------------------------------------------
    ' ���W�X�g������N���X������������
    Set registry = New RegistryManipulator
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, regSubKey) _
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
    Dim dbConnectArray(0 To 9 _
                             , 0 To 1) As Variant
    
    i = 0
     For Each dbConnectInfo In dbConnectList.collection.col
        
        ' ���W�X�g������N���X������������
        Set registry = New RegistryManipulator
        registry.init RegKeyConstants.HKEY_CURRENT_USER _
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, regSubKey & "\" & i) _
                    , RegAccessConstants.KEY_ALL_ACCESS _
                    , True

        j = 0
        dbConnectArray(j, 0) = "no"
        dbConnectArray(j, 1) = j: j = j + 1
        dbConnectArray(j, 0) = "name"
        dbConnectArray(j, 1) = dbConnectInfo.name: j = j + 1
        dbConnectArray(j, 0) = "type"
        dbConnectArray(j, 1) = dbConnectInfo.type_: j = j + 1
        dbConnectArray(j, 0) = "dsn"
        dbConnectArray(j, 1) = dbConnectInfo.dsn: j = j + 1
        dbConnectArray(j, 0) = "host"
        dbConnectArray(j, 1) = dbConnectInfo.host: j = j + 1
        dbConnectArray(j, 0) = "port"
        dbConnectArray(j, 1) = dbConnectInfo.port: j = j + 1
        dbConnectArray(j, 0) = "db"
        dbConnectArray(j, 1) = dbConnectInfo.db: j = j + 1
        dbConnectArray(j, 0) = "user"
        dbConnectArray(j, 1) = dbConnectInfo.user: j = j + 1
        dbConnectArray(j, 0) = "password"
        dbConnectArray(j, 1) = dbConnectInfo.password: j = j + 1
        dbConnectArray(j, 0) = "option"
        dbConnectArray(j, 1) = dbConnectInfo.option_: j = j + 1
        
        ' ���W�X�g���ɏ���ݒ肷��
        registry.setValues dbConnectArray
    
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
' �����@�@�@�FformMode �t�H�[�����[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreDbConnectInfo(ByVal formMode As DB_CONNECT_INFO_TYPE)

    On Error GoTo err
    
    ' ���W�X�g���̃T�u�L�[�����肷��
    Dim regSubKey As String
    
    If formMode = DB_CONNECT_INFO_TYPE.favorite Then
        regSubKey = REG_SUB_KEY_DB_CONNECT_FAVORITE
    Else
        regSubKey = REG_SUB_KEY_DB_CONNECT_HISTORY
    End If
    
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
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, regSubKey) _
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
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, regSubKey & "\" & key) _
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
    
    Set dbConnectList = New CntListBox: dbConnectList.init lstDbConnectList
    addDbConnectList connectInfoList
    Set dbConnectWithoutFilterList = dbConnectList.collection.copy
        
    ' �擪��I������
    dbConnectList.setSelectedIndex 0

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
'     �@      formMode    �t�H�[�����[�h
' �߂�l�@�@�F
'
' =========================================================
Public Function registDbConnectInfo(ByVal connectInfo As ValDBConnectInfo, ByVal formMode As DB_CONNECT_INFO_TYPE)

    On Error GoTo err
    
    ' -------------------------------------------------------
    ' DB�ڑ��������ă��[�h���čŐV�ɂ���
    ' -------------------------------------------------------
    restoreDbConnectInfo formMode
    
    ' -------------------------------------------------------
    ' �d������菜�����������𐶐�����
    ' -------------------------------------------------------
    Dim dbConnectDistinctList As New ValCollection
    Dim dbConnect As ValDBConnectInfo
    
    For Each dbConnect In dbConnectList.collection.col
        
        If dbConnect.displayName <> connectInfo.displayName Then
            ' �~���ŕ\������̂ŁA�ǉ�����v�f�͐擪�ɒǉ����Ă���
            dbConnectDistinctList.setItem dbConnect
        End If
        
    Next
    
    ' -------------------------------------------------------
    ' ��������擪�ɒǉ�����
    ' -------------------------------------------------------
    dbConnectDistinctList.setItemByIndexBefore connectInfo, 1
    
    ' -------------------------------------------------------
    ' DB�ڑ������ɏd������菜�������X�g�œ���ւ���
    ' -------------------------------------------------------
    dbConnectList.removeAll
    addDbConnectList dbConnectDistinctList
    
    ' -------------------------------------------------------
    ' DB�ڑ�������ۑ�����
    ' -------------------------------------------------------
    storeDBConnectInfo formMode

    Exit Function
    
err:

    Main.ShowErrorMessage

End Function

' =========================================================
' ��DB�ڑ�����ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfoList DB�ڑ���񃊃X�g
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDbConnectList(ByVal connectInfoList As ValCollection)
    
    dbConnectList.addAll connectInfoList, "displayName"
    
End Sub

' =========================================================
' ��DB�ڑ�����ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfo DB�ڑ����
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDbConnect(ByVal connectInfo As ValDBConnectInfo)
    
    dbConnectList.addItemByProp connectInfo, "displayName"
    
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
    Set filterConnectList = VBUtil.filterWildcard(dbConnectWithoutFilterList, "displayName", filterKeyword)
    
    addDbConnectList filterConnectList

End Sub


