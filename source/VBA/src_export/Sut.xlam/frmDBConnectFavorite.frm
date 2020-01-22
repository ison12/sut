VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectFavorite 
   Caption         =   "DB�ڑ��̊Ǘ�"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   OleObjectBlob   =   "frmDBConnectFavorite.frx":0000
End
Attribute VB_Name = "frmDBConnectFavorite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�ڑ����C�ɓ���t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/01/14�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Implements IDbConnectListener

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

' B�ڑ��̂��C�ɓ�����̐V�K�쐬�ő吔
Private Const DB_CONNECT_FAVORITE_NEW_CREATED_OVER_SIZE As String = "DB�ڑ��̂��C�ɓ�����͍ő�${count}�܂œo�^�\�ł��B"

' DB�ڑ��t�H�[��
Private WithEvents frmDBConnectVar As frmDBConnect
Attribute frmDBConnectVar.VB_VarHelpID = -1

' DB�ڑ��̂��C�ɓ����񃊃X�g �R���g���[��
Private dbConnectFavoriteList As CntListBox
'DB�ڑ��̂��C�ɓ����񃊃X�g�i�t�B���^�����K�p�Ȃ��j
Private dbConnectFavoriteWithoutFilterList As ValCollection

' DB�ڑ��̂��C�ɓ����񃊃X�g�ł̑I�����ڃC���f�b�N�X
Private dbConnectFavoriteSelectedIndex As Long
' DB�ڑ��̂��C�ɓ����񃊃X�g�ł̑I�����ڃI�u�W�F�N�g
Private dbConnectFavoriteSelectedItem As ValDBConnectInfo

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

    restoredbConnectFavorite
    
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
    Set frmDBConnectVar = Nothing
    
    ' �t�B���^����������
    cboFilter.text = ""
    
    ' �����L�^����
    storeDBConnectFavorite

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
        
        If currentFilterText <> "" Then
            changeEnabledListManipulationControl False
        Else
            changeEnabledListManipulationControl True
        End If
        
        filterConnectList "*" & currentFilterText & "*"
    
        inFilterProcess = False

    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' �����X�g����֘A�̃R���g���[���ނ�Enabled�t���O�𐧌䂷�鏈��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub changeEnabledListManipulationControl(ByVal enabled As Boolean)

    cmdAdd.enabled = enabled
    cmdDelete.enabled = enabled
    cmdUp.enabled = enabled
    cmdDown.enabled = enabled
    cmdDbConnectFavoritePaste.enabled = enabled
    
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
    dbConnectFavoriteSelectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' ���I���̏ꍇ
    If dbConnectFavoriteSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If
    
    ' �t�H�[�������
    HideExt

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set dbConnectFavoriteSelectedItem = dbConnectFavoriteList.getSelectedItem
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(dbConnectFavoriteSelectedItem)
    
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
' ��DB�ڑ����C�ɓ��胊�X�g�̃_�u���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstDbConnectFavoriteList_DblClick(ByVal cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

' =========================================================
' ��DB�ڑ����C�ɓ��胊�X�g �L�[�������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstDbConnectFavoriteList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        cmdOk_Click
    End If
    
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
    cnt = dbConnectFavoriteList.collection.count
    
    ' �|�b�v�A�b�v�̐����ő�o�^���𒴂��Ă��邩�`�F�b�N����
    If cnt >= ConstantsCommon.DB_CONNECT_FAVORITE_NEW_CREATED_MAX_SIZE Then
    
        ' ���b�Z�[�W��\������
        Dim mess As String
        mess = replace(DB_CONNECT_FAVORITE_NEW_CREATED_OVER_SIZE, "${count}", ConstantsCommon.DB_CONNECT_FAVORITE_NEW_CREATED_MAX_SIZE)
        
        VBUtil.showMessageBoxForInformation mess _
                                          , ConstantsCommon.APPLICATION_NAME
        Exit Sub
    End If
    
    ' �|�b�v�A�b�v���j���[�I�u�W�F�N�g�����X�g�ɒǉ�����
    Dim dbConnectFavorite As ValDBConnectInfo
    Set dbConnectFavorite = New ValDBConnectInfo
    
    dbConnectFavorite.name = ConstantsCommon.DB_CONNECT_FAVORITE_DEFAULT_NAME & " " & (cnt + 1)
    
    Dim list As New ValCollection
    list.setItem dbConnectFavorite
    
    addDbConnectFavorite dbConnectFavorite
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
    
    dbConnectFavoriteList.setSelectedIndex cnt
    dbConnectFavoriteList.control.SetFocus
    
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

    editFavorite
End Sub

' =========================================================
' �����̂̕ҏW�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdEditName_Click()

    editFavoriteName
End Sub

Private Sub editFavorite()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    dbConnectFavoriteSelectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' ���I���̏ꍇ
    If dbConnectFavoriteSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set dbConnectFavoriteSelectedItem = dbConnectFavoriteList.getSelectedItem
    
    Set frmDBConnectVar = New frmDBConnect
    frmDBConnectVar.ShowExt vbModal, dbConnectFavoriteSelectedItem, Me
                            
    Set frmDBConnectVar = Nothing

End Sub

Private Sub editFavoriteName()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    dbConnectFavoriteSelectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' ���I���̏ꍇ
    If dbConnectFavoriteSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set dbConnectFavoriteSelectedItem = dbConnectFavoriteList.getSelectedItem
    
    ' DbConnectInfo.Name�v���p�e�B�̓��͂��s���v�����v�g��\������
    Dim inputedName As String
    inputedName = InputBox("DB�ڑ����̖��O��ҏW���܂��B���O����͂��Ă��������B", "DB�ڑ��̖��̕ҏW", dbConnectFavoriteSelectedItem.name)
    If StrPtr(inputedName) = 0 Then
        ' �L�����Z���{�^�����������ꂽ�ꍇ
        Exit Sub
    End If
    
    dbConnectFavoriteSelectedItem.name = inputedName
    
    setDbConnectFavorite dbConnectFavoriteSelectedIndex, dbConnectFavoriteSelectedItem
    dbConnectFavoriteList.control.SetFocus
    
End Sub

' =========================================================
' ��DB�ڑ��̂��C�ɓ�����ݒ�t�H�[����OK�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub IDbConnectListener_connect(connectInfo As ValDBConnectInfo)

    Dim v As ValDBConnectInfo
    Set v = dbConnectFavoriteList.getItem(dbConnectFavoriteSelectedIndex)
    
    v.dsn = connectInfo.dsn
    v.type_ = connectInfo.type_
    v.host = connectInfo.host
    v.port = connectInfo.port
    v.db = connectInfo.db
    v.user = connectInfo.user
    v.password = connectInfo.password
    v.option_ = connectInfo.option_

    setDbConnectFavorite dbConnectFavoriteSelectedIndex, v
    
    dbConnectFavoriteList.control.SetFocus

End Sub

' =========================================================
' ��DB�ڑ��̂��C�ɓ�����ݒ�t�H�[���̃L�����Z���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub IDbConnectListener_cancel()

    dbConnectFavoriteList.control.SetFocus
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
    selectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    dbConnectFavoriteList.removeItem selectedIndex
    dbConnectFavoriteList.control.SetFocus
    
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy

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
    selectedIndex = dbConnectFavoriteList.getSelectedIndex
    
    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
        ' �I������
        Exit Sub
    End If

    If selectedIndex > 0 Then
    
        dbConnectFavoriteList.swapItem _
                          selectedIndex _
                        , selectedIndex - 1 _
                        , vbObject _
                        , 1
                              
        dbConnectFavoriteList.setSelectedIndex selectedIndex - 1
            
    End If
    
    dbConnectFavoriteList.control.SetFocus
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
        
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
    selectedIndex = dbConnectFavoriteList.getSelectedIndex
    
        ' ���I���̏ꍇ
    If selectedIndex = -1 Then
        ' �I������
        Exit Sub
    End If

    If selectedIndex < dbConnectFavoriteList.count - 1 Then
    
        dbConnectFavoriteList.swapItem _
                          selectedIndex _
                        , selectedIndex + 1 _
                        , vbObject _
                        , 1
                              
        dbConnectFavoriteList.setSelectedIndex selectedIndex + 1
            
    End If
    
    dbConnectFavoriteList.control.SetFocus
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���p�����[�^�R�s�[���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDbConnectFavoriteCopy_Click()

    Dim selectedIndex As Long
    Dim selectedItem As ValDBConnectInfo
    
    ' ���ݑI������Ă���C���f�b�N�X���擾
    selectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    Set selectedItem = dbConnectFavoriteList.getSelectedItem
    
    WinAPI_Clipboard.SetClipboard _
        selectedItem.tabbedInfoHeader & vbNewLine & _
        getDbConnectFavoriteForClipboardFormat(selectedItem)
    
End Sub

' =========================================================
' ���S�p�����[�^�R�s�[���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdAllDbConnectFavoriteCopy_Click()

    Dim data As New StringBuilder
    Dim var As Variant
    
    Dim i As Long
    
    For Each var In dbConnectFavoriteList.collection.col
        If i <= 0 Then
            data.append var.tabbedInfoHeader & vbNewLine
        End If
        data.append getDbConnectFavoriteForClipboardFormat(var)
        i = i + 1
    Next
    
    WinAPI_Clipboard.SetClipboard data.str

End Sub

' =========================================================
' ��DB�ڑ��̂��C�ɓ�����̃N���b�v�{�[�h�t�H�[�}�b�g�`��������擾
'
' �T�v�@�@�@�FDB�ڑ��̂��C�ɓ�����̃N���b�v�{�[�h�t�H�[�}�b�g�`����������擾����B
' �����@�@�@�Fvar DB�ڑ��̂��C�ɓ�����
' �߂�l�@�@�FDB�ڑ��̂��C�ɓ�����̃N���b�v�{�[�h�t�H�[�}�b�g�`��������擾
'
' =========================================================
Private Function getDbConnectFavoriteForClipboardFormat(ByVal var As ValDBConnectInfo) As String

    getDbConnectFavoriteForClipboardFormat = var.tabbedInfo & vbNewLine

End Function

' =========================================================
' ��DB�ڑ��̂��C�ɓ�������N���b�v�{�[�h����\�t��
'
' �T�v�@�@�@�FDB�ڑ��̂��C�ɓ�������N���b�v�{�[�h����\�t������B
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmddbConnectFavoritePaste_Click()

    Dim var As Variant
    Dim dbConnectFavoriteRawList As ValCollection
    
    Dim dbConnectFavoriteObj As ValDBConnectInfo
    Dim dbConnectFavoriteObjList As New ValCollection

    Dim clipBoard As String
    clipBoard = WinAPI_Clipboard.GetClipboard
    
    Dim CsvParser As New CsvParser: CsvParser.init vbTab
    Set dbConnectFavoriteRawList = CsvParser.parse(clipBoard)
    
    For Each var In dbConnectFavoriteRawList.col
        
        Set dbConnectFavoriteObj = New ValDBConnectInfo
    
        If var.count >= 9 Then
            dbConnectFavoriteObj.name = var.getItemByIndex(1, vbVariant)
            dbConnectFavoriteObj.type_ = var.getItemByIndex(2, vbVariant)
            dbConnectFavoriteObj.dsn = var.getItemByIndex(3, vbVariant)
            dbConnectFavoriteObj.host = var.getItemByIndex(4, vbVariant)
            dbConnectFavoriteObj.port = var.getItemByIndex(5, vbVariant)
            dbConnectFavoriteObj.db = var.getItemByIndex(6, vbVariant)
            dbConnectFavoriteObj.user = var.getItemByIndex(7, vbVariant)
            dbConnectFavoriteObj.password = var.getItemByIndex(8, vbVariant)
            dbConnectFavoriteObj.option_ = var.getItemByIndex(9, vbVariant)
            
            If dbConnectFavoriteObj.tabbedInfo <> dbConnectFavoriteObj.tabbedInfoHeader Then
                dbConnectFavoriteObjList.setItem dbConnectFavoriteObj
            End If
            
        End If
    
    Next
    
    addDbConnectFavoriteList dbConnectFavoriteObjList, isAppend:=True
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy

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
' ��DB�ڑ��̂��C�ɓ��������ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeDBConnectFavorite()

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
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_FAVORITE) _
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
     For Each dbConnectInfo In dbConnectFavoriteList.collection.col
        
        ' ���W�X�g������N���X������������
        Set registry = New RegistryManipulator
        registry.init RegKeyConstants.HKEY_CURRENT_USER _
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_FAVORITE & "\" & i) _
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
' ��DB�ڑ��̂��C�ɓ��������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoredbConnectFavorite()

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
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_FAVORITE) _
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
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_FAVORITE & "\" & key) _
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
    
    cboFilter.text = ""
    Set dbConnectFavoriteList = New CntListBox: dbConnectFavoriteList.init lstDbConnectFavoriteList
    addDbConnectFavoriteList connectInfoList
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
    
    ' �擪��I������
    dbConnectFavoriteList.setSelectedIndex 0
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ��DB�ڑ��̂��C�ɓ������ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfo DB�ڑ����
' �߂�l�@�@�F
'
' =========================================================
Public Function registDbConnectInfo(ByVal connectInfo As ValDBConnectInfo)

    On Error GoTo err
    
    ' -------------------------------------------------------
    ' DB�ڑ����C�ɓ�������ă��[�h���čŐV�ɂ���
    ' -------------------------------------------------------
    restoredbConnectFavorite
    
    ' -------------------------------------------------------
    ' DB�ڑ����C�ɓ�����̖����ɏ���ǉ�����
    ' -------------------------------------------------------
    cboFilter.text = ""
    addDbConnectFavorite connectInfo
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
    
    ' -------------------------------------------------------
    ' DB�ڑ����C�ɓ������ۑ�����
    ' -------------------------------------------------------
    storeDBConnectFavorite

    Exit Function
    
err:

    Main.ShowErrorMessage
    
End Function

' =========================================================
' ��DB�ڑ��̂��C�ɓ������ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfoList DB�ڑ���񃊃X�g
'     �@�@�@  isAppend        �ǉ��L���t���O
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDbConnectFavoriteList(ByVal connectInfoList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    dbConnectFavoriteList.addAll connectInfoList, "displayName", isAppend:=isAppend
    
End Sub

' =========================================================
' ��DB�ڑ��̂��C�ɓ������ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FconnectInfo DB�ڑ����
' �߂�l�@�@�F
'
' =========================================================
Private Sub addDbConnectFavorite(ByVal connectInfo As ValDBConnectInfo)
    
    dbConnectFavoriteList.addItemByProp connectInfo, "displayName"
    
End Sub

' =========================================================
' ��DB�J���������ݒ����ύX
'
' �T�v�@�@�@�F
' �����@�@�@�Findex �C���f�b�N�X
'     �@�@�@  rec   DB�ڑ����
' �߂�l�@�@�F
'
' =========================================================
Private Sub setDbConnectFavorite(ByVal index As Long, ByVal rec As ValDBConnectInfo)
    
    dbConnectFavoriteList.setItem index, rec, "displayName"
    
End Sub

' =========================================================
' �ڑ���񃊃X�g���t�B���^���鏈��
'
' �T�v�@�@�@�F�ڑ���񃊃X�g���t�B���^���鏈��
' �����@�@�@�FfilterKeyword         �t�B���^�L�[���[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub filterConnectList(ByVal filterKeyword As String)

    Dim filterConnectList As ValCollection
    Set filterConnectList = VBUtil.filterWildcard(dbConnectFavoriteWithoutFilterList, "displayName", filterKeyword)
    
    addDbConnectFavoriteList filterConnectList

End Sub
