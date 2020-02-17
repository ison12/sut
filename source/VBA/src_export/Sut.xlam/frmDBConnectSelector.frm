VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectSelector 
   Caption         =   "�ڑ����̑I��"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12675
   OleObjectBlob   =   "frmDBConnectSelector.frx":0000
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
    
    ' �t�B���^������K�p����
    cboFilter.text = ""
    applyFilterCondition

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
        
        'filterConnectList currentFilterText ' ���S��v
        filterConnectList "*" & currentFilterText & "*" ' ���Ԉ�v
    
        inFilterProcess = False

    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
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
        cmdCancel_Click
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
' ���ݒ���̐���
' =========================================================
Private Function createApplicationProperties(ByVal formMode As DB_CONNECT_INFO_TYPE) As ApplicationProperties
    
    ' �t�H�[�������擾����
    Dim subName As String
    
    If formMode = DB_CONNECT_INFO_TYPE.favorite Then
        subName = "frmDBConnectFavorite"
    Else
        subName = Me.name & "History"
    End If
    
    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & subName & ".ini"

    Set createApplicationProperties = appProp
    
End Function

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
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties(formMode)
    
    
    ' �������݃f�[�^
    Dim val As New ValDBConnectInfo
    Dim values As New ValCollection
    
    Dim i As Long: i = 1
    For Each val In dbConnectList.collection.col
        
        values.setItem Array(i & "_" & "no", i)
        values.setItem Array(i & "_" & "name", val.name)
        values.setItem Array(i & "_" & "type", val.type_)
        values.setItem Array(i & "_" & "dsn", val.dsn)
        values.setItem Array(i & "_" & "host", val.host)
        values.setItem Array(i & "_" & "port", val.port)
        values.setItem Array(i & "_" & "db", val.db)
        values.setItem Array(i & "_" & "user", val.user)
        values.setItem Array(i & "_" & "password", val.password)
        values.setItem Array(i & "_" & "option", val.option_)
        
        i = i + 1
    Next
        
    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
        
    Exit Sub
    
err:

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
        
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties(formMode)
    
    ' �ڑ����
    Dim connectInfoList As ValCollection
    Set connectInfoList = New ValCollection
    Dim connectInfo As ValDBConnectInfo
    

    ' �f�[�^��ǂݍ���
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    Dim i As Long: i = 1
    Do While True
    
        val = values.getItem(i & "_" & "no", vbVariant)
        If Not IsArray(val) Then
            Exit Do
        End If
        
        Set connectInfo = New ValDBConnectInfo
                    
        val = values.getItem(i & "_" & "name", vbVariant): If IsArray(val) Then connectInfo.name = val(2)
        val = values.getItem(i & "_" & "type", vbVariant): If IsArray(val) Then connectInfo.type_ = val(2)
        val = values.getItem(i & "_" & "dsn", vbVariant): If IsArray(val) Then connectInfo.dsn = val(2)
        val = values.getItem(i & "_" & "host", vbVariant): If IsArray(val) Then connectInfo.host = val(2)
        val = values.getItem(i & "_" & "port", vbVariant): If IsArray(val) Then connectInfo.port = val(2)
        val = values.getItem(i & "_" & "db", vbVariant): If IsArray(val) Then connectInfo.db = val(2)
        val = values.getItem(i & "_" & "user", vbVariant): If IsArray(val) Then connectInfo.user = val(2)
        val = values.getItem(i & "_" & "password", vbVariant): If IsArray(val) Then connectInfo.password = val(2)
        val = values.getItem(i & "_" & "option", vbVariant): If IsArray(val) Then connectInfo.option_ = val(2)
        
        connectInfoList.setItem connectInfo
    
        i = i + 1
    Loop
    
    Set dbConnectList = New CntListBox: dbConnectList.init lstDbConnectList
    addDbConnectList connectInfoList
    Set dbConnectWithoutFilterList = dbConnectList.collection.copy
        
    ' �擪��I������
    dbConnectList.setSelectedIndex 0

    Exit Sub
    
err:
    
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


