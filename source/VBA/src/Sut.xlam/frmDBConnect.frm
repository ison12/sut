VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnect 
   Caption         =   "DB�ڑ�"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   OleObjectBlob   =   "frmDBConnect.frx":0000
End
Attribute VB_Name = "frmDBConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�ڑ����s���t�H�[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2008/09/06�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' ���ڑ�����DB�����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event ok(ByVal connStr As String, ByVal connSimpleStr As String)

' =========================================================
' ��DB�̐ڑ����L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' �ڑ������� �z��C���f�b�N�X�ŏ��l
Private Const CONNECT_STR_MIN As Long = 1
' �ڑ������� �z��C���f�b�N�X�ő�l
Private Const CONNECT_STR_MAX As Long = 5

' �R���g���[���L���t���O �C���f�b�N�X �f�[�^�\�[�X
Private Const CONTROL_ENABLE_IDX_DATASOURCE As Long = 1
' �R���g���[���L���t���O �C���f�b�N�X �z�X�g
Private Const CONTROL_ENABLE_IDX_HOST       As Long = 2
' �R���g���[���L���t���O �C���f�b�N�X DB
Private Const CONTROL_ENABLE_IDX_DB         As Long = 3
' �R���g���[���L���t���O �C���f�b�N�X �|�[�g
Private Const CONTROL_ENABLE_IDX_PORT       As Long = 4
' �R���g���[���L���t���O �C���f�b�N�X ���[�U
Private Const CONTROL_ENABLE_IDX_USER       As Long = 5
' �R���g���[���L���t���O �C���f�b�N�X �p�X���[�h
Private Const CONTROL_ENABLE_IDX_PASSWORD   As Long = 6
' �R���g���[���L���t���O �C���f�b�N�X �t�@�C���I���{�^��
Private Const CONTROL_ENABLE_IDX_FILE_SELECT   As Long = 7

' �ڑ�������
Private connectStr(1 To 5) As String
' �v���o�C�_���x��
Private providerLabel(1 To 5) As String
' �f�t�H���g�|�[�g�ԍ�
Private defaultPort(1 To 5) As String
' �R���g���[���L���t���O
Private controlEnable(1 To 5, 1 To 7) As Boolean

' ---------------------------------------------------------
' ���W�X�g���t�@�C���L�[
' ---------------------------------------------------------
Private Const REG_SUB_KEY_DB_CONNECT As String = "db_connect"

Private WithEvents history  As sutredlib.DbConnectHistory
Attribute history.VB_VarHelpID = -1
Private WithEvents favorite As sutredlib.DbFavorite
Attribute favorite.VB_VarHelpID = -1

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants)

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

    Main.storeFormPosition Me.name, Me
    Me.Hide
End Sub

Private Function getDbConnectInfo() As dbConnectInfo

    ' DB�ڑ����𐶐����R���g���[����������W�ߐݒ肷��
    Dim connectInfo As New sutredlib.dbConnectInfo
    connectInfo.Type = cboDBType.value
    connectInfo.dsn = cboDataSourceName.value
    connectInfo.host = txtHost.value
    connectInfo.port = txtPort.value
    connectInfo.db = txtDB.value
    connectInfo.user = txtUser.value
    connectInfo.password = txtPassword.value
    connectInfo.Option = txtOption.value
    
    Set getDbConnectInfo = connectInfo

End Function

Private Sub setDbConnectInfo(ByRef connectInfo As dbConnectInfo)

    On Error Resume Next
    
    cboDBType.value = connectInfo.Type
    cboDataSourceName.value = connectInfo.dsn
    txtHost.value = connectInfo.host
    txtPort.value = connectInfo.port
    txtDB.value = connectInfo.db
    txtUser.value = connectInfo.user
    txtPassword.value = connectInfo.password
    txtOption.value = connectInfo.Option
    
    On Error GoTo 0
End Sub

Private Sub favorite_OnLoad(ByVal connectInfo As sutredlib.IDbConnectInfo)

    setDbConnectInfo connectInfo
End Sub

Private Sub history_OnLoad(ByVal connectInfo As sutredlib.IDbConnectInfo)

    setDbConnectInfo connectInfo
End Sub

Private Sub cmdHistoryChoice_Click()

    ' �������E�B���h�E��\������
    history.Show

End Sub

Private Sub cmdFavoriteLoad_Click()

    ' ���C�ɓ�����E�B���h�E��\������
    favorite.Show

End Sub

Private Sub cmdFavoriteSave_Click()

    ' DB�ڑ����𐶐����R���g���[����������W�ߐݒ肷��
    Dim connectInfo As sutredlib.dbConnectInfo
    Set connectInfo = getDbConnectInfo
    
    ' DbConnectInfo.Name�v���p�e�B�̃f�t�H���g�l
    Dim defaultName As String
    If cboDBType.value = "�ėpODBC" Then
    
        defaultName = cboDataSourceName.value
    ElseIf cboDBType.value = "Oracle Provider for OLE DB" Then
    
        defaultName = txtUser.value & "@" & txtDB.value
        
    ElseIf cboDBType.value = "Microsoft OLE DB for SQL Server" Then
    
        defaultName = txtUser.value & "@" & txtDB.value
    End If
    
    ' DbConnectInfo.Name�v���p�e�B�̓��͂��s���v�����v�g��\������
    Dim inputedName As String
    inputedName = InputBox("DB�ڑ�����ۑ����܂��B���O����͂��Ă��������B", "DB�ڑ� - �ݒ�̕ۑ�", defaultName)
    
    ' �L�����Z���{�^�����������ꂽ�ꍇ�A�߂�l����ɂȂ�̂ŏ����𒆒f����
    ' �L�����Z���{�^���������ꂽ�̂��A���ۂ̓��͒l���󂾂����̂��̌������͕t���Ȃ����A�����ł͈ꗥ�����𒆒f����悤�ɂ���B
    If inputedName = "" Then
    
        Exit Sub
    End If
    
    connectInfo.name = inputedName
    
    ' DB�ڑ��������W�X�g���ɓo�^����
    Dim favorite As New sutredlib.DbFavorite
    favorite.RegistDbConnectInfo connectInfo
    
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
' ��DB��ރR���{�{�b�N�X�ύX���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboDBType_Change()

    On Error GoTo err
    
    ' DB��ނ̃C���f�b�N�X
    Dim index As Long
    ' �|�[�g�ԍ�
    Dim port As Long
    
    ' �R���{�{�b�N�X�őI������Ă���C���f�b�N�X���擾����
    index = cboDBType.ListIndex + 1
    
    ' �C���f�b�N�X���͈͊O�̏ꍇ
    If index < CONNECT_STR_MIN Or index > CONNECT_STR_MAX Then
    
        ' �I��
        Exit Sub
    End If
    
    ' �e�R���g���[���̐ݒ�l�����Z�b�g����
    txtHost.text = ""
    txtDB.text = ""
    txtPort.text = ""
    txtUser.text = ""
    txtPassword.text = ""
    txtOption.text = ""

    ' �e�R���g���[���̗L���E������ݒ肷��
    changeControlByEnableStatus cboDataSourceName, controlEnable(index, CONTROL_ENABLE_IDX_DATASOURCE)
    changeControlByEnableStatus txtHost, controlEnable(index, CONTROL_ENABLE_IDX_HOST)
    changeControlByEnableStatus txtDB, controlEnable(index, CONTROL_ENABLE_IDX_DB)
    changeControlByEnableStatus txtPort, controlEnable(index, CONTROL_ENABLE_IDX_PORT)
    changeControlByEnableStatus txtUser, controlEnable(index, CONTROL_ENABLE_IDX_USER)
    changeControlByEnableStatus txtPassword, controlEnable(index, CONTROL_ENABLE_IDX_PASSWORD)
    changeControlByVisibleStatus cmdFileSelection, controlEnable(index, CONTROL_ENABLE_IDX_FILE_SELECT)

    ' �f�t�H���g�|�[�g�ԍ����擾����
    txtPort.text = defaultPort(index)

    ' ���f�[�^�\�[�X���X�g���X�V����
    updateDataSourceList
    
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

Private Sub changeControlByEnableStatus(ByRef c As control, ByVal enable As Boolean)

    If enable = True Then
    
        c.Enabled = True
        c.BackColor = &H80000005
    Else
        c.Enabled = False
        c.BackColor = &H8000000F
    
    End If

End Sub

Private Sub changeControlByVisibleStatus(ByRef c As control, ByVal visible As Boolean)

    c.visible = visible
End Sub

' =========================================================
' ��ODBC�ݒ胉�x���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lblODBCSetting_Click()
    
    On Error GoTo err
    
    ' �߂�l�i�[�p�ϐ�
    Dim ret        As Long
    
    ' �V�X�e�����[�g���ϐ�
    Dim systemRoot As String
    
    ' �őO�ʂ���������
    ExcelUtil.setUserFormTopMost Me, False
    
    ' �V�X�e�����[�g���ϐ����擾
    systemRoot = WinAPI_Shell.getEnvironmentVariable("SystemRoot")
    
    ' ODBC�Ǘ��R���\�[�����N������
    ret = WinAPI_Shell.ShellExecute(0 _
                           , "open" _
                           , systemRoot & "\system32\odbcad32.exe" _
                           , "" _
                           , systemRoot & "\system32" _
                           , 1)
                           
    ' �߂�l��32�ȉ��̏ꍇ�G���[
    If ret <= 32 Then
    
        VBUtil.showMessageBoxForWarning "ODBC�Ǘ��R���\�[�����J�����Ƃ��ł��܂���ł����B", ConstantsCommon.APPLICATION_NAME, Nothing
    
    End If
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ��DSN�X�V�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDSNUpdate_Click()

    On Error GoTo err
    
    ' ���f�[�^�\�[�X���X�g���X�V����
    updateDataSourceList

    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���ڑ��e�X�g�N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdConnectTest_Click()

    On Error GoTo err
    
    SutWhite.showHourglassWindowOnCenterPt Me
    
    ' �ڑ��e�X�g���������{����
    connectDBTest
    
    SutWhite.closeHourglassWindow
    
    ' ���������ꍇ
    VBUtil.showMessageBoxForInformation "DB�̐ڑ��ɐ������܂����B", ConstantsCommon.APPLICATION_NAME
    
    Exit Sub
    
err:

    SutWhite.closeHourglassWindow

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���t�@�C���I���{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdFileSelection_Click()

    ' �t�@�C����I������
    Dim filePath As String
    filePath = VBUtil.openFileDialog("Access�t�@�C����I�����Ă�������", "")

    ' �t�@�C�����I�����ꂽ���ǂ����̔���
    If filePath <> "" Then
    
        ' DB�e�L�X�g�Ƀt�@�C���p�X��ݒ肷��
        txtDB.text = filePath
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
    
    Dim connStr As String
    
    ' �ڑ��e�X�g���{���ʂ����s�������ꍇ��
    ' �ēx�ݒ���s���������[�U�ɑI��������
    
    On Error Resume Next
    
    SutWhite.showHourglassWindowOnCenterPt Me
    
    ' DB�ɐڑ�����
    connStr = connectDBTest
    
    SutWhite.closeHourglassWindow

    If err.Number <> 0 Then
        
        showMessageBoxForError "�G���[���������܂����B", ConstantsCommon.APPLICATION_NAME, err

        ' �ݒ���I�����邩�ǂ�����I��������
        If VBUtil.showMessageBoxForYesNo("�ēx�ݒ肵�܂����H" _
                , ConstantsCommon.APPLICATION_NAME) = WinAPI_User.IDYES Then
        
            ' �����𒆒f����
            Exit Sub
            
        Else
            ' �L�����Z���{�^���������Ɠ����������s�������𒆒f����
            cmdCancel_Click
        
            Exit Sub
        End If
        
    End If
    
    ' DB�ڑ������L�^����
    storeDbConnectInfo
    
    ' �t�H�[�������
    HideExt
    
        
    ' �ڑ�������
    Dim connSimpleStr As String
    
    ' �ڑ�������𐶐�����
    connSimpleStr = createConnectSimpleString(cboDBType.text _
                                , cboDataSourceName.text _
                                , txtHost.text _
                                , txtPort.text _
                                , txtDB.text _
                                , txtUser.text _
                                , txtPassword.text _
                                , txtOption.text)
    
    ' DB�ڑ�OK�C�x���g�𑗐M����
    RaiseEvent ok(connStr, connSimpleStr)
    
    ' DB�ڑ����𗚗��Ƃ��ēo�^����
    history.RegistDbConnectInfo getDbConnectInfo
    
    Exit Sub
    
err:
    
    SutWhite.closeHourglassWindow

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
    
    ' DB�ڑ��L�����Z���C�x���g�𑗐M����
    RaiseEvent Cancel

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

    Set history = New sutredlib.DbConnectHistory
    Set favorite = New sutredlib.DbFavorite

    Dim i As Long
    Dim j As Long
    
    ' �ȉ��̔z��ϐ��́A����C���f�b�N�X�ɂ���đΉ����Ă���B
    ' �E�ڑ�������
    ' �E�v���o�C�_���x��
    ' �E�f�t�H���g�|�[�g�ԍ�
    ' �E�R���g���[���L���t���O
    
    ' ----------------------------------------------
    ' �ڑ�������@������
    ' ----------------------------------------------
    i = CONNECT_STR_MIN
    
    ' ODBC
    ' ��MSDASQL.1�́A�}�C�N���\�t�g����ODBC�pOLE DB�v���o�C�_
    connectStr(i) = "Provider=MSDASQL.1;" & _
                    "Data Source=${dsn};" & _
                    "User ID=${user};" & _
                    "Password=${password};"
                    
    i = i + 1
    
'    ' PostgreSQL�iOLEDB�j
'    connectStr(i) = "Provider=PostgreSQL OLE DB Provider;" & _
'                                                 "Data Source=${host};" & _
'                                                 "Location=${db};" & _
'                                                 "User ID=${user};" & _
'                                                 "Password=${password};"
'
'    i = i + 1
'
'    ' MySQL�iODBC�j
'    connectStr(i) = "Driver={MySQL ODBC 3.51 Driver};" & _
'                                                 "Server=${host};" & _
'                                                 "Port=${port};" & _
'                                                 "Database=${db};" & _
'                                                 "User=${user};" & _
'                                                 "Password=${password};" & _
'                                                 "Option=3;"
'
'    i = i + 1
    
    ' Oracle�iOLEDB Oracle�j
    connectStr(i) = "Provider=OraOLEDB.Oracle;" & _
                                                 "Data Source=${db};" & _
                                                 "User Id=${user};" & _
                                                 "Password=${password};"
                                                 
    i = i + 1
    
'    ' Oracle�iOLEDB Microsoft�j
'    connectStr(i) = "Provider=msdaora;" & _
'                                                 "Data Source=${db};" & _
'                                                 "User Id=${user};" & _
'                                                 "Password=${password};"
                                                 
    ' Microsoft SQL Server�iOLEDB�j
    connectStr(i) = "Provider=SQLOLEDB;" & _
                                                 "Data Source=${host};" & _
                                                 "Initial Catalog=${db};" & _
                                                 "User Id=${user};" & _
                                                 "Password=${password};"
                                                 
    i = i + 1
    
    ' Microsoft Access
    connectStr(i) = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                                 "Data Source=${db};" & _
                                                 "User Id=${user};" & _
                                                 "Password=${password};"
                                                 
    i = i + 1
    
    ' Microsoft Access for 2007
    connectStr(i) = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                                 "Data Source=${db};" & _
                                                 "User Id=${user};" & _
                                                 "Password=${password};"
                                                 
    i = i + 1
                                                 
    ' ----------------------------------------------
    ' �v���o�C�_���x���@������
    ' ----------------------------------------------
    i = CONNECT_STR_MIN

    providerLabel(i) = "�ėpODBC": i = i + 1
'    providerLabel(i) = "PostgreSQL (OLE DB)": i = i + 1
'    providerLabel(i) = "MySQL (MyODBC 3.51)": i = i + 1
    providerLabel(i) = "Oracle Provider for OLE DB": i = i + 1
'    providerLabel(i) = "Oracle Provider for OLE DB (Microsoft)": i = i + 1
    providerLabel(i) = "Microsoft OLE DB for SQL Server": i = i + 1
    providerLabel(i) = "Microsoft Access (Jet Provider)": i = i + 1
    providerLabel(i) = "Microsoft Access (Ace Provider)": i = i + 1

    ' ----------------------------------------------
    ' �f�t�H���g�|�[�g�ԍ��@������
    ' ----------------------------------------------
    i = CONNECT_STR_MIN

    defaultPort(i) = "": i = i + 1
'    defaultPort(i) = "5432": i = i + 1
'    defaultPort(i) = "3306": i = i + 1
    defaultPort(i) = "": i = i + 1
'    defaultPort(i) = "": i = i + 1
    defaultPort(i) = "": i = i + 1
    defaultPort(i) = "": i = i + 1
    defaultPort(i) = "": i = i + 1
    
    ' ----------------------------------------------
    ' �R���g���[���L���t���O�@������
    ' ���v���o�C�_���ύX���ꂽ�ꍇ�ɑΉ�����R���g���[���̗L���E���������肷��l
    ' ----------------------------------------------
    i = CONNECT_STR_MIN
    j = CONTROL_ENABLE_IDX_DATASOURCE
    
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = False: j = j + 1
    controlEnable(i, j) = False: j = j + 1
    controlEnable(i, j) = False: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = False: j = j + 1

    i = i + 1
    j = CONTROL_ENABLE_IDX_DATASOURCE
    
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'
'    i = i + 1
'    j = CONTROL_ENABLE_IDX_DATASOURCE
'
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'
'    i = i + 1
'    j = CONTROL_ENABLE_IDX_DATASOURCE
    
    controlEnable(i, j) = False: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = False: j = j + 1

    i = i + 1
    j = CONTROL_ENABLE_IDX_DATASOURCE
    
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1

    controlEnable(i, CONTROL_ENABLE_IDX_DATASOURCE) = False
    controlEnable(i, CONTROL_ENABLE_IDX_HOST) = True
    controlEnable(i, CONTROL_ENABLE_IDX_DB) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PORT) = False
    controlEnable(i, CONTROL_ENABLE_IDX_USER) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PASSWORD) = True
    controlEnable(i, CONTROL_ENABLE_IDX_FILE_SELECT) = False

    i = i + 1
    j = CONTROL_ENABLE_IDX_DATASOURCE

    controlEnable(i, CONTROL_ENABLE_IDX_DATASOURCE) = False
    controlEnable(i, CONTROL_ENABLE_IDX_HOST) = False
    controlEnable(i, CONTROL_ENABLE_IDX_DB) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PORT) = False
    controlEnable(i, CONTROL_ENABLE_IDX_USER) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PASSWORD) = True
    controlEnable(i, CONTROL_ENABLE_IDX_FILE_SELECT) = True

    i = i + 1
    j = CONTROL_ENABLE_IDX_DATASOURCE

    controlEnable(i, CONTROL_ENABLE_IDX_DATASOURCE) = False
    controlEnable(i, CONTROL_ENABLE_IDX_HOST) = False
    controlEnable(i, CONTROL_ENABLE_IDX_DB) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PORT) = False
    controlEnable(i, CONTROL_ENABLE_IDX_USER) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PASSWORD) = True
    controlEnable(i, CONTROL_ENABLE_IDX_FILE_SELECT) = True

    ' ��DB��ރR���{�{�b�N�X�Ƀ��X�g��ǉ�����
    cboDBType.list = providerLabel

    ' ���O��Ō�ɐڑ����������t�H�[����̊e�R���g���[���ɕ���������
    restoreDbConnectInfo
    
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

    Set history = Nothing
    Set favorite = Nothing

End Sub

' =========================================================
' ���f�[�^�\�[�X���X�g�̍X�V����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub updateDataSourceList()

    Dim dataSourceList As ValCollection
    Dim dataSource     As ValCollection
    
    Set dataSourceList = WinAPI_ODBC.getDataSourceList
    
    cboDataSourceName.clear
    
    For Each dataSource In dataSourceList.col
    
        cboDataSourceName.addItem dataSource.getItemByIndex(1, vbVariant)
        
    Next
End Sub

' =========================================================
' ���ڑ��e�X�g����
'
' �T�v�@�@�@�FDB�ւ̐ڑ����s��
' �����@�@�@�F
' �߂�l�@�@�FDB�ڑ�������
'
' =========================================================
Private Function connectDBTest() As String

    On Error GoTo err
    
    ' �R�l�N�V�����I�u�W�F�N�g
    Dim conn As Object
    
    ' �ڑ�������
    Dim connStr As String
    
    ' �ڑ�������𐶐�����
    connStr = createConnectString(cboDBType.text _
                                , cboDataSourceName.text _
                                , txtHost.text _
                                , txtPort.text _
                                , txtDB.text _
                                , txtUser.text _
                                , txtPassword.text _
                                , txtOption.text)
                                      
    
    ' DB�ɐڑ�����
    Set conn = ADOUtil.connectDb(connStr)
    
    ' DB�ɐڑ����Ă���ꍇ�ADB�̐ڑ���ؒf����
    If Not conn Is Nothing Then
    
        ADOUtil.closeDB conn
        Set conn = Nothing
        
    End If
    
    connectDBTest = connStr
    
    Exit Function

err:

    ' DB�ɐڑ����Ă���ꍇ�ADB�̐ڑ���ؒf����
    If Not conn Is Nothing Then
    
        ADOUtil.closeDB (conn)
        Set conn = Nothing
        
    End If
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Function

' =========================================================
' ��DB�ڑ������񐶐�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Function createConnectString(ByVal dbType As String _
                                   , ByVal dsn As String _
                                   , ByVal host As String _
                                   , ByVal port As String _
                                   , ByVal db As String _
                                   , ByVal user As String _
                                   , ByVal password As String _
                                   , ByVal options As String _
                                   ) As String

    ' �ڑ�������
    Dim connStr As String
    
    ' DB��ނ̃C���f�b�N�X
    Dim index As Long
    
    ' �R���{�{�b�N�X�őI������Ă���C���f�b�N�X���擾����
    index = cboDBType.ListIndex + 1
    
    ' �C���f�b�N�X���͈͊O�̏ꍇ
    If index < CONNECT_STR_MIN Or index > CONNECT_STR_MAX Then
    
        ' �I��
        Exit Function
    End If
    
    connStr = connectStr(index)

    ' Oracle�̏ꍇ
    If dbType = "Oracle Provider for OLE DB" Then
    
        Dim dbVar As String
        dbVar = db
        If Trim$(host) <> "" And Trim$(port) <> "" Then
            dbVar = host & ":" & port & "/" & dbVar
        ElseIf Trim$(host) <> "" And Trim$(port) = "" Then
            dbVar = host & "/" & dbVar
        End If
        
        connStr = replace(connStr, "${db}", dbVar)
        connStr = replace(connStr, "${user}", user)
        connStr = replace(connStr, "${password}", password)
        connStr = connStr & options
            
    Else
    
        connStr = replace(connStr, "${dsn}", dsn)
        connStr = replace(connStr, "${host}", host)
        connStr = replace(connStr, "${port}", port)
        connStr = replace(connStr, "${db}", db)
        connStr = replace(connStr, "${user}", user)
        connStr = replace(connStr, "${password}", password)
        connStr = connStr & options
        
    End If
        
    createConnectString = connStr
    
End Function

' =========================================================
' ��DB�ڑ�������i�P���j��������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Function createConnectSimpleString(ByVal dbType As String _
                                   , ByVal dsn As String _
                                   , ByVal host As String _
                                   , ByVal port As String _
                                   , ByVal db As String _
                                   , ByVal user As String _
                                   , ByVal password As String _
                                   , ByVal options As String _
                                   ) As String

    ' �ڑ�������
    Dim connStr As String
    Dim joinStr As String
    
    If dsn <> "" Then
        connStr = connStr & joinStr & "DSN=" & dsn: joinStr = ", "
    End If
    
    If host <> "" Then
        connStr = connStr & joinStr & "�z�X�g=" & host: joinStr = ", "
    End If
    
    If port <> "" Then
        connStr = connStr & joinStr & "�|�[�g=" & port: joinStr = ", "
    End If
    
    If db <> "" Then
        connStr = connStr & joinStr & "DB=" & db: joinStr = ", "
    End If
        
    createConnectSimpleString = connStr
    
End Function

' =========================================================
' ��DB�̐ڑ�����ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeDbConnectInfo()

    On Error GoTo err
    
    ' DB�ڑ������i�[����z��
    Dim dbConnectInfoArray(0 To 7 _
                         , 0 To 1) As Variant
    
    
    Dim i As Long
    
    i = 0
    
    dbConnectInfoArray(i, 0) = cboDBType.name
    dbConnectInfoArray(i, 1) = cboDBType.value: i = i + 1
    dbConnectInfoArray(i, 0) = cboDataSourceName.name
    dbConnectInfoArray(i, 1) = cboDataSourceName.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtHost.name
    dbConnectInfoArray(i, 1) = txtHost.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtPort.name
    dbConnectInfoArray(i, 1) = txtPort.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtDB.name
    dbConnectInfoArray(i, 1) = txtDB.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtUser.name
    dbConnectInfoArray(i, 1) = txtUser.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtPassword.name
    dbConnectInfoArray(i, 1) = txtPassword.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtOption.name
    dbConnectInfoArray(i, 1) = txtOption.value: i = i + 1
    
    
    ' ���W�X�g������N���X
    Dim registry As New RegistryManipulator
    ' ���W�X�g������N���X������������
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' ���W�X�g���ɏ���ݒ肷��
    registry.setValues dbConnectInfoArray
    
    ' ----------------------------------------------
    ' �u�b�N�ݒ���
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet
    
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, cboDBType.name, cboDBType.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, cboDataSourceName.name, cboDataSourceName.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtHost.name, txtHost.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtPort.name, txtPort.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtDB.name, txtDB.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtUser.name, txtUser.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtPassword.name, txtPassword.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtOption.name, txtOption.value
    ' ----------------------------------------------
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DB�̐ڑ�����ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreDbConnectInfo()

    On Error GoTo err
    
    ' ----------------------------------------------
    ' �u�b�N�ݒ���
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' �ݒ���V�[�g�����݂���
        
        Set bookPropVal = bookProp.getValues(ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG)
        If bookPropVal.count > 0 Then
            ' �ݒ��񂪑��݂���̂ŁA�t�H�[���ɔ��f����
            cboDBType.value = bookPropVal.getItem(cboDBType.name, vbString)
            cboDataSourceName.value = bookPropVal.getItem(cboDataSourceName.name, vbString)
            txtHost.value = bookPropVal.getItem(txtHost.name, vbString)
            txtPort.value = bookPropVal.getItem(txtPort.name, vbString)
            txtDB.value = bookPropVal.getItem(txtDB.name, vbString)
            txtUser.value = bookPropVal.getItem(txtUser.name, vbString)
            txtPassword.value = bookPropVal.getItem(txtPassword.name, vbString)
            txtOption.value = bookPropVal.getItem(txtOption.name, vbString)
        
            Exit Sub
        End If
    End If
    ' ----------------------------------------------
    
    ' ���W�X�g������N���X
    Dim registry As New RegistryManipulator
    ' ���W�X�g������N���X������������
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Dim retDbType         As String
    Dim retDataSourceName As String
    Dim retHost           As String
    Dim retPort           As String
    Dim retDB             As String
    Dim retUser           As String
    Dim retPassword       As String
    Dim retOption         As String
    
    ' ���W�X�g����������擾����
    registry.getValue cboDBType.name, retDbType
    registry.getValue cboDataSourceName.name, retDataSourceName
    registry.getValue txtHost.name, retHost
    registry.getValue txtPort.name, retPort
    registry.getValue txtDB.name, retDB
    registry.getValue txtUser.name, retUser
    registry.getValue txtPassword.name, retPassword
    registry.getValue txtOption.name, retOption

    cboDBType.value = retDbType
    cboDataSourceName.value = retDataSourceName
    txtHost.value = retHost
    txtPort.value = retPort
    txtDB.value = retDB
    txtUser.value = retUser
    txtPassword.value = retPassword
    txtOption.value = retOption
    
    Exit Sub

err:

    Main.ShowErrorMessage


End Sub
