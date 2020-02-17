VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectFavorite 
   Caption         =   "DB�ڑ��̊Ǘ�"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12675
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
' �����@�@�@�F
'
' =========================================================
Public Event ok()

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event cancel()

' B�ڑ��̂��C�ɓ�����̐V�K�쐬�ő吔
Private Const DB_CONNECT_FAVORITE_NEW_CREATED_OVER_SIZE As String = "DB�ڑ��̂��C�ɓ�����͍ő�${count}�܂œo�^�\�ł��B"

' DB�ڑ��t�H�[��
Private WithEvents frmDBConnectVar As frmDBConnect
Attribute frmDBConnectVar.VB_VarHelpID = -1

' DB�ڑ��̂��C�ɓ����񃊃X�g �R���g���[��
Private dbConnectFavoriteList As CntListBox

' DB�ڑ��̂��C�ɓ����񃊃X�g�ł̑I�����ڃC���f�b�N�X
Private dbConnectFavoriteSelectedIndex As Long
' DB�ڑ��̂��C�ɓ����񃊃X�g�ł̑I�����ڃI�u�W�F�N�g
Private dbConnectFavoriteSelectedItem As ValDBConnectInfo

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
' ����  �@�@�Fmodal                    ���[�_���܂��̓��[�h���X�\���w��
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
    
    ' �����L�^����
    storeDBConnectFavorite
    
    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok
    
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
    editFavorite
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
    
    If VBUtil.unloadFormIfChangeActiveBook(frmDBConnect) Then Unload frmDBConnect
    Load frmDBConnect
    Set frmDBConnectVar = frmDBConnect
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

    Dim i As Long
    
    Dim var As ValCollection
    Dim dbConnectFavoriteRawList As ValCollection
    
    Dim dbConnectFavoriteObj As ValDBConnectInfo
    Dim dbConnectFavoriteObjList As New ValCollection

    Dim clipBoard As String
    clipBoard = WinAPI_Clipboard.GetClipboard
    
    Dim CsvParser As New CsvParser: CsvParser.init vbTab
    Set dbConnectFavoriteRawList = CsvParser.parse(clipBoard)
    
    For Each var In dbConnectFavoriteRawList.col
        
        Set dbConnectFavoriteObj = New ValDBConnectInfo
    
        ' �s������⊮����i�ŏI�񂪖����͂̏ꍇ�ȂǁA���s�����邱�Ƃ����邽�߁j
        For i = 1 To 9 - var.count
            var.setItem ""
        Next
    
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
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & Me.name & ".ini"

    Set createApplicationProperties = appProp
    
End Function

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
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' �������݃f�[�^
    Dim val As ValDBConnectInfo
    Dim values As New ValCollection
    
    Dim i As Long: i = 1
    For Each val In dbConnectFavoriteList.collection.col
        
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
' ��DB�ڑ��̂��C�ɓ��������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoredbConnectFavorite()

    On Error GoTo err
            
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
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
    
    Set dbConnectFavoriteList = New CntListBox: dbConnectFavoriteList.init lstDbConnectFavoriteList
    addDbConnectFavoriteList connectInfoList
        
    ' �擪��I������
    dbConnectFavoriteList.setSelectedIndex 0

    Exit Sub
    
err:
    
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
    addDbConnectFavorite connectInfo
    
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
