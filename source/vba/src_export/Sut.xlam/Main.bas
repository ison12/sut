Attribute VB_Name = "Main"
Option Explicit

' ________________________________________________________
' �����o�ϐ�
' ________________________________________________________

Private menuDB    As UIMenuDB
Private menuTable As UIMenuTable
Private menuData  As UIMenuData
Private menuDiff  As UIMenuDiff
Private menuFile  As UIMenuFile
Private menuTool  As UIMenuTool
Private menuHelp  As UIMenuHelp

' ��DB�R�l�N�V����
Public dbConn As Object
' ��DB�ڑ�������
Public dbConnStr As String
' ��DB�ڑ�������i�P���Ȑڑ�������j
Public dbConnSimpleStr As String

' ���A�v���P�[�V�����ݒ���
Private applicationSetting As ValApplicationSetting

' ���A�v���P�[�V�����ݒ���i�V���[�g�J�b�g�j
Private applicationSettingShortcut As ValApplicationSettingShortcut

' ���A�v���P�[�V�����ݒ���i�J�����������j
Private applicationSettingColFormat As ValApplicationSettingColFormat

' �A�h�C���̃t�@�C���N���[�Y���̏���
Public Sub Auto_Close()

    ' �����͈ȉ��̂悤�ɁASutDestroy���Ăяo���悤�ɂ��Ă������A����̃P�[�X�ł��܂��j���������ł��Ȃ����Ƃ��킩�����̂ŃR�����g�A�E�g

    '    On Error GoTo err
    '
    '    #If (DEBUG_MODE = 1) Then
    '
    '        Debug.Print "Auto_Close"
    '    #End If
    '
    '    SutDestroy
    '
    '    Exit Sub
    '
    'err:
    '
    '    ' �G���[����
    '    Main.ShowErrorMessage
    
    ' ���j������������ɂł��Ȃ�����̃P�[�X�Ƃ�
    '   1. �{�A�h�C�����g�ݍ��܂�Ă����ԂŁA��������̃u�b�N��V�K�ŊJ��
    '   2. �V�K�u�b�N�ŉ�������ҏW�����{����
    '   3. Excel�S�̂���悤�Ƃ���
    '   4. �ۑ��m�F�_�C�A���O���\������邪�L�����Z����������āA���鏈�����̂��L�����Z�������
    '   5. ���鏈�����L�����Z�����ꂽ�ɂ��ւ�炸�A�{�A�h�C����Auto_Close�����{����Ă��܂�

End Sub

' �A�h�C���̃A���C���X�g�[�����̑΍�
Public Function Auto_Remove()

    '�������܂���
    
End Function

' =========================================================
' ��Main���W���[���ŊǗ����Ă���DB�R�l�N�V�������X�V����B
'
' �T�v�@�@�@�F
' ���L�����@�F
'
' =========================================================
Public Function SutUpdateDbConn(ByRef dbConn_ As Object, ByRef dbConnStr_ As String, ByRef dbConnSimpleStr_ As String)

    If Not dbConn_ Is Nothing Then
    
        ADOUtil.closeDB dbConn
    End If
    
    Set dbConn = dbConn_
    dbConnStr = dbConnStr_
    dbConnSimpleStr = dbConnSimpleStr_
    
    If Not menuTable Is Nothing Then
        menuTable.updateDbConn dbConn_
    End If
    
    If Not menuDiff Is Nothing Then
        menuDiff.updateDbConn dbConn_, dbConnSimpleStr_
    End If
    
    If ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        changeDbConnectStatus dbConnStr_, dbConnSimpleStr_, False
    Else
        changeDbConnectStatus dbConnStr_, dbConnSimpleStr_, True
    End If

End Function

' =========================================================
' ��Main���W���[���̃����o������������
'
' �T�v�@�@�@�F
' ���L�����@�F�c�[���o�[�̏�������ɌĂяo�����s������
'
' =========================================================
Public Function SutInit()
    
    ' �e�탁���o��Get���\�b�h���R�[�����邱�ƂŃ����o������������
    getApplicationSetting
    getApplicationSettingShortcut
    getApplicationSettingColFormat
    
    initUIObject
    
End Function

' =========================================================
' ��Main���W���[���̃����o���������
'
' �T�v�@�@�@�F
' ���L�����@�F�c�[���o�[�̍폜�O�ɌĂяo�����s������
'
' =========================================================
Public Function SutRelease()

    ADOUtil.closeDB dbConn
    Set dbConn = Nothing
    
    dbConnStr = Empty
    dbConnSimpleStr = Empty
    
    Set applicationSetting = Nothing
    Set applicationSettingShortcut = Nothing
    Set applicationSettingColFormat = Nothing
    
    Set menuDB = Nothing
    Set menuTable = Nothing
    Set menuData = Nothing
    Set menuFile = Nothing
    Set menuDiff = Nothing
    Set menuTool = Nothing
    Set menuHelp = Nothing
    
    Unload frmDBColumnFormat
    Unload frmDBColumnFormatSetting
    Unload frmDBConnect
    Unload frmDBConnectFavorite
    Unload frmDBConnectSelector
    Unload frmDBExplorer
    Unload frmDBQueryBatch
    Unload frmDBQueryBatchTypeSetting
    Unload frmFileOutput
    Unload frmMenuSetting
    Unload frmOption
    Unload frmPopupMenu
    Unload frmProgress
    Unload frmQueryParameter
    Unload frmQueryParameterSetting
    Unload frmQueryResult
    Unload frmQueryResultDetail
    Unload frmRecordAppender
    Unload frmSelectConditionCreator
    Unload frmShortcutKey
    Unload frmShortcutKeySetting
    Unload frmSnapshot
    Unload frmSnapshotDiff
    Unload frmSplash
    Unload frmTableSheetCreator
    Unload frmTableSheetList
    Unload frmTableSheetUpdate
    
End Function

' =========================================================
' ��UI�I�u�W�F�N�g�̏�����
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub initUIObject()

    If menuDB Is Nothing Then

        Set menuDB = New UIMenuDB
    End If

    If menuTable Is Nothing Then

        Set menuTable = New UIMenuTable
    End If

    If menuData Is Nothing Then

        Set menuData = New UIMenuData
    End If

    If menuFile Is Nothing Then

        Set menuFile = New UIMenuFile
    End If

    If menuDiff Is Nothing Then

        Set menuDiff = New UIMenuDiff
    End If

    If menuTool Is Nothing Then

        Set menuTool = New UIMenuTool
    End If

    If menuHelp Is Nothing Then

        Set menuHelp = New UIMenuHelp
    End If

End Sub

' =========================================================
' ���A�h�C�������[�h����B
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutPreload()

    On Error GoTo err

    initLoadingToolbar
    
    Exit Function
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Function

' =========================================================
' ��Sut�����S�ɔj������
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutDestroy()

    On Error GoTo err

    ' �c�[���o�[���폜����O�ɌĂяo��
    ' �O���[�o���̈�̃f�[�^���������
    Main.SutRelease
    
    ' �c�[���o�[���폜����
    deleteToolbar
    
    Exit Function
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Function

' =========================================================
' ���A�h�C�������[�h����B
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutLoad()

    On Error GoTo err
    
    ' �J�����g�h���C�u�ƃJ�����g�f�B���N�g����؂�ւ���
    ChDrive SutWorkbook.path
    ChDir SutWorkbook.path
    
    ' �c�[���o�[������������
    initToolbar
    
    ' �c�[���o�[�̏�������ɌĂяo��
    ' �O���[�o���̈�̃f�[�^������������
    Main.SutInit
    
    Exit Function
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Function

' =========================================================
' ���A�h�C�����A�����[�h����B
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutUnload()

    On Error GoTo err

    ' �c�[���o�[���폜����O�ɌĂяo��
    ' �O���[�o���̈�̃f�[�^���������
    Main.SutRelease
    
    ' �c�[���o�[�̈ꕔ���폜����
    deleteToolbarExcludeSomeItems
    
    Exit Function
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Function

' =========================================================
' ��DB�ڑ��ݒ�t�H�[���\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutConnectDB()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    menuDB.init
    menuDB.connectDb
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ��DB�ڑ����\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutDBConnectInfo()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    menuDB.init
    menuDB.showDBConnectInfo dbConn
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ��DB�ڑ��ؒf
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutDisConnectDB()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    menuDB.init
    menuDB.disconnectDB
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function
' =========================================================
' ��DB�G�N�X�v���[���\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutShowDbExplorer()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim conn As Object: Set conn = getDBConnection
    
    menuTable.init appSetting, conn
    menuTable.showDbExplorer
    
    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ���e�[�u���V�[�g�ꗗ�\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutShowTableSheetList()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting

    ' �I�[�g���[�V�����G���[���������Ă��܂����߃_�~�[�̃I�u�W�F�N�g������Ă���
    ' �i�����͕s���j
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    menuTable.init appSetting, conn
    menuTable.showTableSheetList
    
    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ���e�[�u���V�[�g�쐬
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutCreateTableSheet()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim conn As Object: Set conn = getDBConnection

    menuTable.init appSetting, conn
    menuTable.createTableSheet
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ���e�[�u���V�[�g�X�V
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutUpdateTableSheet()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim conn As Object: Set conn = getDBConnection

    menuTable.init appSetting, conn
    menuTable.updateTableSheet
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ��INSERT���s
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutInsertUpdateAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.insertUpdateAll
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��INSERT���s�i�I��̈�j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutInsertUpdateSelection()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.insertUpdateSelection
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function
' =========================================================
' ��INSERT���s
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutInsertAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.insertAll
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��INSERT���s�i�I��̈�j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutInsertSelection()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.insertSelection
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��UPDATE���s
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutUpdateAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.updateAll
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��UPDATE���s�i�I��̈�j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutUpdateSelection()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.updateSelection
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��DELETE���s
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutDeleteAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.deleteAll
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��DELETE���s�i�I��̈�j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutDeleteSelection()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.deleteSelection
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��DELETE���s�i�e�[�u����̑S���R�[�h�j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutDeleteAllOfTable()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.deleteAllOfTable
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��SELECT���s
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutSelectAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.selectAll
    
    doAfterProcess
    
    menuData.showQueryResultWhenSettingResult

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��SELECT���s�i�����w��j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutSelectCondition()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.selectCondition
    
    doAfterProcess
    
    menuData.showQueryResultWhenSettingResult

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ��SELECT���s�i�Ď��s�j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutSelectReExecute()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.selectReExecute
    
    doAfterProcess
    
    menuData.showQueryResultWhenSettingResult

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���N�G���G�f�B�^�i !!! ������ !!! �j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutQueryEditor()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.queryEditor
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���ꊇ�N�G��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutQueryBatch()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.queryBatch
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���N�G�����ʕ\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutShowQueryResult()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    
    menuData.init appSetting, appSettingColFmt, Nothing, False ' �N�G�����ʂ��������Ȃ�
    menuData.showQueryResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ���N�G���p�����[�^�ݒ�
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutSettingQueryParameter()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    
    menuData.init appSetting, appSettingColFmt, Nothing, False ' �N�G�����ʂ��������Ȃ�
    menuData.settingQueryParameter
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���s�̒ǉ�
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutRecordAdd()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    
    menuData.init appSetting, appSettingColFmt, Nothing
    menuData.recordAdd
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���X�i�b�v�V���b�gSQL��`�V�[�g�쐬
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutCreateNewSheetSnapSqlDefine()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    
    menuDiff.init appSetting, appSettingColFmt, Nothing
    menuDiff.createNewSheetSnapSqlDefine
    
    doAfterProcess
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���X�i�b�v�V���b�g���s�t�H�[���Ăяo��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutShowSnapshot()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuDiff.init appSetting, appSettingColFmt, conn
    menuDiff.showSnapshot
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���I�v�V�����ݒ�
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutSettingOption()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject

    menuTool.settingOption
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���E�N���b�N���j���[�̐ݒ�
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutSettingRClickMenu()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    menuTool.settingRClickMenu
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���V���[�g�J�b�g�L�[�̐ݒ�
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutSettingShortCutKey()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    menuTool.settingShortCutKey
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���o�[�W�����̕\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutSettingPopupMenu()

    On Error GoTo err
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    menuTool.settingPopupMenu

    doAfterProcess

    Exit Function
    
err:
    
    Main.ShowErrorMessage

End Function

' =========================================================
' ���t�@�C���o�� - INSERT + UPDATE�i�S�āj
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileInsertUpdateAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.insertUpdateAll
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���t�@�C���o�� - INSERT + UPDATE�i�I��͈́j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileInsertUpdateSelection()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.insertUpdateSelection
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function
' =========================================================
' ���t�@�C���o�� - INSERT�i�S�āj
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileInsertAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.insertAll
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���t�@�C���o�� - INSERT�i�I��͈́j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileInsertSelection()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.insertSelection
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���t�@�C���o�� - UPDATE�i�S�āj
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileUpdateAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.updateAll
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���t�@�C���o�� - UPDATE�i�I��͈́j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileUpdateSelection()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.updateSelection
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���t�@�C���o�� - DELETE�i�S�āj
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileDeleteAll()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.deleteAll
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���t�@�C���o�� - DELETE�i�I��͈́j
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileDeleteSelection()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.deleteSelection
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���t�@�C���o�� - �ꊇ�o��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutFileBatch()

    On Error GoTo err
    
    ' �u�b�N�̃`�F�b�N���s��
    validWorkbook
    
    ' UI�I�u�W�F�N�g�̏�����
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.batchFile
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ���w���v�t�@�C���̕\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutShowHelpFile()

    On Error GoTo err
    
    ' �߂�l
    Dim ret As Long
    
    ' �w���v�t�@�C����\������
    ret = WinAPI_Shell.ShellExecute(0 _
                           , "open" _
                           , VBUtil.concatFilePath(ThisWorkbook.path _
                                                 , ConstantsCommon.HELP_FILE) _
                           , "" _
                           , ThisWorkbook.path _
                           , 1)
    
    ' �߂�l��32�ȉ��̏ꍇ�G���[
    If ret <= 32 Then
    
        VBUtil.showMessageBoxForInformation "�w���v�t�@�C�����J�����Ƃ��ł��܂���ł����B", ConstantsCommon.APPLICATION_NAME
    
    End If

    Exit Function
    
err:
    
    Main.ShowErrorMessage

End Function

' =========================================================
' ���o�[�W�����̕\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function SutShowVersion()

    On Error GoTo err
    
    ' �t�H�[����ݒ肷��
    If VBUtil.unloadFormIfChangeActiveBook(frmSplash) Then Unload frmSplash
    frmSplash.Show vbModal

    Exit Function
    
err:
    
    Main.ShowErrorMessage

End Function

Private Function SutShowPopupCommon(ByVal index As Long)

    On Error GoTo err
    
    Dim appSetting As ValApplicationSettingShortcut
    Set appSetting = Main.getApplicationSettingShortcut
    
    Dim popupMenu As ValPopupmenu
    Set popupMenu = appSetting.popupMenuList.getItemByIndex(index)
    
    If Not popupMenu Is Nothing Then
    
        ' �|�b�v�A�b�v�R���g���[�����擾����
        Dim popup As CommandBar
        Set popup = popupMenu.commandBarPopup
        
        If Not popup Is Nothing Then
        
            ' �\������
            popup.ShowPopup
        End If
    
        
    End If
    
    Exit Function
    
err:
    
    Main.ShowErrorMessage

End Function

Public Function SutShowPopup1()

    SutShowPopupCommon 1
End Function

Public Function SutShowPopup2()

    SutShowPopupCommon 2
End Function

Public Function SutShowPopup3()
    
    SutShowPopupCommon 3
End Function

Public Function SutShowPopup4()
    
    SutShowPopupCommon 4
End Function

Public Function SutShowPopup5()
    
    SutShowPopupCommon 5
End Function

Public Function SutShowPopup6()
    
    SutShowPopupCommon 6
End Function

Public Function SutShowPopup7()
    
    SutShowPopupCommon 7
End Function

Public Function SutShowPopup8()
    
    SutShowPopupCommon 8
End Function

Public Function SutShowPopup9()
    
    SutShowPopupCommon 9
End Function

Public Function SutShowPopup10()
    
    SutShowPopupCommon 10
End Function

' =========================================================
' ��DB�R�l�N�V�����擾
'
' �T�v�@�@�@�FDB�R�l�N�V�������擾����
'
' =========================================================
Public Function getDBConnection() As Object

    ' DB�R�l�N�V����������������Ă���ꍇ
    If Not dbConn Is Nothing Then
    
        #If DEBUG_MODE = 1 Then

            Debug.Print "Connection Ver. " & dbConn.version
        #End If
    
        ' �ڑ�����Ă��邩�m�F����
        If ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        
            ' �G���[�𓊂���
            err.Raise ConstantsError.ERR_NUMBER_DISCONNECT_DB _
                    , _
                    , ConstantsError.ERR_DESC_DISCONNECT_DB
            
        End If
    
    ' DB�R�l�N�V����������������Ă��Ȃ��ꍇ
    Else
        ' DB�ڑ��t�H�[����\������
        SutConnectDB
        
        ' DB�ɐڑ�����Ă��Ȃ��ꍇ
        If dbConn Is Nothing Then
        
            ' �G���[�𓊂���
            err.Raise ConstantsError.ERR_NUMBER_DISCONNECT_DB _
                    , _
                    , ConstantsError.ERR_DESC_DISCONNECT_DB
        End If
        
    End If

    ' �߂�l��ݒ肷��
    Set getDBConnection = dbConn

End Function

' =========================================================
' ��DB�R�l�N�V�����ؒf
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub disconnectDB()

    SutDisConnectDB

End Sub

' =========================================================
' ���A�v���P�[�V�����ݒ���擾
'
' �T�v�@�@�@�F�A�v���P�[�V�����ݒ�����擾����
'
' =========================================================
Public Function getApplicationSetting() As Object

    ' ����������Ă���ꍇ
    If Not applicationSetting Is Nothing Then
    
    
    ' ����������Ă��Ȃ��ꍇ
    Else
    
        Set applicationSetting = New ValApplicationSetting
        applicationSetting.readForData
        
    End If

    ' �߂�l��ݒ肷��
    Set getApplicationSetting = applicationSetting

End Function

' =========================================================
' ���A�v���P�[�V�����ݒ���擾
'
' �T�v�@�@�@�F�A�v���P�[�V�����ݒ�����擾����
'
' =========================================================
Public Function getApplicationSettingShortcut() As Object

    ' ����������Ă���ꍇ
    If Not applicationSettingShortcut Is Nothing Then
    
    
    ' ����������Ă��Ȃ��ꍇ
    Else
    
        Set applicationSettingShortcut = New ValApplicationSettingShortcut
        applicationSettingShortcut.init
        
    End If

    ' �߂�l��ݒ肷��
    Set getApplicationSettingShortcut = applicationSettingShortcut

End Function

' =========================================================
' ���A�v���P�[�V�����ݒ���擾�i�J�����������j
'
' �T�v�@�@�@�F�A�v���P�[�V�����ݒ���i�J�����������j���擾����
'
' =========================================================
Public Function getApplicationSettingColFormat() As Object

    ' ����������Ă���ꍇ
    If Not applicationSettingColFormat Is Nothing Then
    
    
    ' ����������Ă��Ȃ��ꍇ
    Else
    
        Set applicationSettingColFormat = New ValApplicationSettingColFormat
        applicationSettingColFormat.init
        
    End If

    ' �߂�l��ݒ肷��
    Set getApplicationSettingColFormat = applicationSettingColFormat

End Function

' =========================================================
' �����[�N�u�b�N�̃`�F�b�N���s��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function validWorkbook()

    ' �u�b�N�I�u�W�F�N�g
    Dim book As Workbook
    
    ' �u�b�N�I�u�W�F�N�g���擾����
    Set book = ActiveWorkbook
    
    ' �u�b�N�I�u�W�F�N�g�̃`�F�b�N
    If book Is Nothing Then
    
        err.Raise ERR_NUMBER_NON_ACTIVE_BOOK _
                , _
                , ERR_DESC_NON_ACTIVE_BOOK
            
    End If
    
    ' �u�b�N�I�u�W�F�N�g���{�A�h�C�����̂̏ꍇ
    If book Is SutWorkbook Then
    
        err.Raise ERR_NUMBER_ACTIVE_ADDIN_BOOK _
                , _
                , ERR_DESC_ACTIVE_ADDIN_BOOK
            
    End If

End Function

' =========================================================
' ���G���[���b�Z�[�W��\������
'
' �T�v�@�@�@�F�A�v���P�[�V�����G���[���ǂ����𔻒肵��
' �@�@�@�@�@�@�K�؂Ȃh�e�ŃG���[���b�Z�[�W��\������B
'
' =========================================================
Public Function ShowErrorMessage()

    If ConstantsError.isApplicationError(err.Number) = True Then
    
        ' �A�v���P�[�V�����G���[�����������ꍇ�AvbObjectError�ƌŒ萔[512]�������āA�{���̃G���[�ԍ����Z�o����
        err.Number = err.Number - vbObjectError - 512
        ' �G���[����\������
        VBUtil.showMessageBoxForWarning "", ConstantsCommon.APPLICATION_NAME, err
    Else
    
        VBUtil.showMessageBoxForError ConstantsError.ERR_MSG_ERROR_LEVEL, ConstantsCommon.APPLICATION_NAME, err
    End If
    
End Function

' =========================================================
' ���t�H�[���|�W�V�����𕜌�����
'
' �T�v�@�@�@�FformName �t�H�[���̎��ʎq
' �@�@�@�@�@�@formObj  �t�H�[���I�u�W�F�N�g
'
' =========================================================
Public Function restoreFormPosition(ByVal formName As String _
                                  , ByRef formObj As Object)
    
    Dim formRect As New ValRectPt
    formRect.Left = formObj.Left
    formRect.Top = formObj.Top
    formRect.Width = formObj.Width
    formRect.Height = formObj.Height
    
    Dim formPosition As New ValFormPosition: formPosition.init formName
    Call formPosition.readForData(formRect)

    formObj.Top = formRect.Top
    formObj.Left = formRect.Left

End Function

' =========================================================
' ���t�H�[���|�W�V������ۑ�����
'
' �T�v�@�@�@�FformName �t�H�[���̎��ʎq
' �@�@�@�@�@�@formObj  �t�H�[���I�u�W�F�N�g
'
' =========================================================
Public Function storeFormPosition(ByVal formName As String _
                                , ByRef formObj As Object)

    Dim formRect As New ValRectPt
    formRect.Left = formObj.Left
    formRect.Top = formObj.Top
    formRect.Width = formObj.Width
    formRect.Height = formObj.Height
    
    Dim formPosition As New ValFormPosition: formPosition.init formName
    Call formPosition.writeForData(formRect)

End Function

' =========================================================
' ���c�[���o�[�̏���������
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function initLoadingToolbar()

    On Error Resume Next
    
    ' �J�����g�h���C�u�ƃJ�����g�f�B���N�g����؂�ւ���
    ChDrive SutWorkbook.path
    ChDir SutWorkbook.path

    ' �G�N�Z���̃o�[�W����
    Dim excelVer As ExcelVersion: excelVer = ExcelUtil.getExcelVersion
    
    ' �R�}���h�o�[
    Dim cb   As CommandBar
    
    Set cb = Application.CommandBars.Add( _
                            name:=ConstantsCommon.COMMANDBAR_MENU_NAME _
                          , Temporary:=True _
                          , position:=msoBarFloating)
        
    ' ���ɒǉ�����Ă���ꍇ�́A�ϐ�cb��nothing�ɂȂ�
    ' �ϐ�cb��nothing�̏ꍇ�́A�����𒆒f����
    If cb Is Nothing Then
    
        Exit Function
        
    End If
    
    ' -----------------------------------------------------------------------
    ' �A�v���P�[�V�����A�C�R����ݒ肷��
    ' -----------------------------------------------------------------------
    ' �A�v���P�[�V�����A�C�R���{�^��
    Dim appIcon As CommandBarButton
    
    ' Excel2002�ȍ~�̃v���p�e�B
    If excelVer >= Ver2002 Then
        
        Set appIcon = cb.Controls.Add(Type:=msoControlButton)
        
        With appIcon
        
            .Style = msoButtonIcon
            .OnAction = "Main.SutShowVersion"
            ' �폜�Ώۂ��珜�O
            .Tag = ConstantsCommon.COMMANDBAR_DONT_DELETE_TARGET
            
            setCommandBarControlIcon appIcon _
                                   , "Database"
            
            ' ��DescriptionText�v���p�e�B�ɖ����I�ɋ󕶎����ݒ肷��
            ' �@�V���[�g�J�b�g�L�[�̋@�\���X�g�ɖ{�R���g���[���͒ǉ����Ȃ�
            .DescriptionText = ""
            
        
        End With

    End If
    
    ' -----------------------------------------------------------------------
    ' �@�\�ʂɃR�}���h�o�[�ɃR���g���[����ǉ�����
    ' -----------------------------------------------------------------------
    
    ' ***************************************************************
    ' �A�v���P�[�V�����̋N���ƏI��
    ' ***************************************************************
    ' �t�@�C���|�b�v�A�b�v
    Dim popFile                   As commandBarPopup
    ' ���[�h�{�^��
    Dim btnLoad                   As CommandBarButton
    ' �A�����[�h�{�^��
    Dim btnUnload                 As CommandBarButton
    
    ' �t�@�C���|�b�v�A�b�v��ǉ�����
    Set popFile = cb.Controls.Add(Type:=msoControlPopup)
    
    With popFile
        ' �폜�Ώۂ��珜�O
        .Tag = ConstantsCommon.COMMANDBAR_DONT_DELETE_TARGET
        .Caption = "Sut"
    End With
        
    ' ���[�h�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnLoad = popFile.Controls.Add(Type:=msoControlButton)
    
    ' ���[�h�{�^���̃v���p�e�B��ݒ肷��
    With btnLoad
    
        .Style = msoButtonIconAndCaption
        .Caption = "�A�v���P�[�V�����N��"
        .OnAction = "Main.SutLoad"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutLoad"
        
        ' ��DescriptionText�v���p�e�B�ɖ����I�ɋ󕶎����ݒ肷��
        ' �@�V���[�g�J�b�g�L�[�̋@�\���X�g�ɖ{�R���g���[���͒ǉ����Ȃ�
        .DescriptionText = ""
        
    End With
        
    ' ���[�h�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnUnload = popFile.Controls.Add(Type:=msoControlButton)
    
    ' ���[�h�{�^���̃v���p�e�B��ݒ肷��
    With btnUnload
    
        .Style = msoButtonIconAndCaption
        .Caption = "�A�v���P�[�V�����I��"
        .OnAction = "Main.SutUnload"
        .Enabled = False
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutUnload"
        
        ' ��DescriptionText�v���p�e�B�ɖ����I�ɋ󕶎����ݒ肷��
        ' �@�V���[�g�J�b�g�L�[�̋@�\���X�g�ɖ{�R���g���[���͒ǉ����Ȃ�
        .DescriptionText = ""
        
    End With
    
    ' ***************************************************************
    
    cb.visible = True

    On Error GoTo 0
    
End Function


' =========================================================
' ���c�[���o�[�̏���������
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function initToolbar()

    On Error Resume Next
    
    ' �f�B���N�g�����ꎞ�I�ɕύX����
    ' �A�C�R���ݒ�̂��߂� SutYellow.dll ���Ăяo�����߂ɕK�v�ȏ��u
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ' �G�N�Z���̃o�[�W����
    Dim excelVer As ExcelVersion: excelVer = ExcelUtil.getExcelVersion
    
    ' �R�}���h�o�[
    Dim cb   As CommandBar
    
    Set cb = Application.CommandBars(ConstantsCommon.COMMANDBAR_MENU_NAME)
    
    ' �擾�Ɏ��s�����ꍇ�A�ϐ�cb��nothing�ɂȂ�
    ' initToolbar�Ăяo���̑O��Ƃ��āA���Ƀ��j���[���ǉ�����Ă���K�v������B
    ' �ϐ�cb��nothing�̏ꍇ�́A�����𒆒f����
    If cb Is Nothing Then
    
        Exit Function
        
    End If

    ' -----------------------------------------------------------------------
    ' �@�\�ʂɃR�}���h�o�[�ɃR���g���[����ǉ�����
    ' -----------------------------------------------------------------------
    
    ' ***************************************************************
    ' DB�ڑ�
    ' ***************************************************************
    ' DB�|�b�v�A�b�v
    Dim popDB                     As commandBarPopup
    ' DB�ڑ��{�^��
    Dim btnDBConnect              As CommandBarButton
    ' DB�ؒf�{�^��
    Dim btnDBDisConnect           As CommandBarButton
    ' DB�ڑ����
    Dim btnDBInfo                 As CommandBarButton
    
    ' DB�|�b�v�A�b�v��ǉ�����
    Set popDB = cb.Controls.Add(Type:=msoControlPopup)
    
    With popDB
    
        .Caption = "DB"
    End With
        
    ' DB�ڑ��{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnDBConnect = popDB.Controls.Add(Type:=msoControlButton)
    
    ' DB�ڑ��{�^���̃v���p�e�B��ݒ肷��
    With btnDBConnect
    
        .Style = msoButtonIconAndCaption
        .Caption = "�ڑ�"
        .DescriptionText = "DB�ڑ�"
        .OnAction = "Main.SutConnectDB"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutConnectDB"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDBConnect _
                                   , "DatabaseSetting"
        End If
        
    End With
        
    ' DB�ؒf�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnDBDisConnect = popDB.Controls.Add(Type:=msoControlButton)
    
    ' DB�ؒf�{�^���̃v���p�e�B��ݒ肷��
    With btnDBDisConnect
    
        .Style = msoButtonIconAndCaption
        .Caption = "�ؒf"
        .DescriptionText = "DB�ؒf"
        .OnAction = "Main.SutDisconnectDB"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDisconnectDB"
        .state = msoButtonDown ' DB���ؒf����Ă��邱�Ƃ�������悤�ɏ�����Ԃ̓{�^��������Ԃɂ���
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDBDisConnect _
                                   , "DeleteDatabase"
        End If
        
    End With
    
    ' ***************************************************************
    
        
    ' ***************************************************************
    ' �e�[�u��
    ' ***************************************************************
    ' �e�[�u���|�b�v�A�b�v
    Dim popTable                  As commandBarPopup
    ' DB�G�N�X�v���[��
    Dim btnDbExplorer             As CommandBarButton
    ' �e�[�u���ꗗ�{�^��
    Dim btnTableList              As CommandBarButton
    ' �e�[�u�������E�B�U�[�h�{�^��
    Dim btnTableCreateSheetWizard As CommandBarButton
    ' �e�[�u���X�V�{�^��
    Dim btnTableUpdateSheetWizard As CommandBarButton
    
    ' �e�[�u���|�b�v�A�b�v��ǉ�����
    Set popTable = cb.Controls.Add(Type:=msoControlPopup)
    
    With popTable
    
        .Caption = "�e�[�u��"
    End With
    
    ' DB�G�N�X�v���[���{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnDbExplorer = popTable.Controls.Add(Type:=msoControlButton)
    
    ' DB�G�N�X�v���[���{�^���̃v���p�e�B��ݒ肷��
    With btnDbExplorer
    
        .Style = msoButtonIconAndCaption
        .Caption = "DB�G�N�X�v���[��"
        .DescriptionText = "DB�G�N�X�v���[��"
        .OnAction = "Main.SutShowDbExplorer"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowDbExplorer"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDbExplorer _
                                   , "Search"
        End If
        
    End With
    
    ' �e�[�u���ꗗ�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnTableList = popTable.Controls.Add(Type:=msoControlButton)
    
    ' �e�[�u���ꗗ�{�^���̃v���p�e�B��ݒ肷��
    With btnTableList
    
        .Style = msoButtonIconAndCaption
        .Caption = "�e�[�u���V�[�g�ꗗ"
        .DescriptionText = "�e�[�u���V�[�g�ꗗ"
        .OnAction = "Main.SutShowTableSheetList"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowTableSheetList"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnTableList _
                                   , "SearchWindow"
        End If
        
    End With
    
    ' �e�[�u�������E�B�U�[�h�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnTableCreateSheetWizard = popTable.Controls.Add(Type:=msoControlButton)
    
    ' �e�[�u�������E�B�U�[�h�{�^���̃v���p�e�B��ݒ肷��
    With btnTableCreateSheetWizard
    
        .Style = msoButtonIconAndCaption
        .Caption = "�e�[�u���V�[�g�쐬"
        .DescriptionText = "�e�[�u���V�[�g�쐬"
        .OnAction = "Main.SutCreateTableSheet"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutCreateTableSheet"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnTableCreateSheetWizard _
                                   , "AddFolder"
        End If
        
    End With
    
    ' �e�[�u���X�V�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnTableUpdateSheetWizard = popTable.Controls.Add(Type:=msoControlButton)
    
    ' �e�[�u���X�V�{�^���̃v���p�e�B��ݒ肷��
    With btnTableUpdateSheetWizard
    
        .Style = msoButtonIconAndCaption
        .Caption = "�e�[�u���V�[�g�X�V"
        .DescriptionText = "�e�[�u���V�[�g�X�V"
        .OnAction = "Main.SutUpdateTableSheet"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutUpdateTableSheet"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnTableUpdateSheetWizard _
                                   , "WindowImport"
        End If
    End With

    ' ***************************************************************
    
    
    ' ***************************************************************
    ' �f�[�^
    ' ***************************************************************
    ' �f�[�^�|�b�v�A�b�v
    Dim popData                   As commandBarPopup
    
    ' �f�[�^�|�b�v�A�b�v��ǉ�����
    Set popData = cb.Controls.Add(Type:=msoControlPopup)
    
    With popData
    
        .Caption = "�f�[�^"
    End With
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' INSERT + UPDATE
    ' ***************************************************************
    ' INSERT + UPDATE�|�b�v�A�b�v
    Dim popInsertUpdate                 As commandBarPopup
    ' INSERT + UPDATE�{�^��
    Dim btnInsertUpdate                 As CommandBarButton
    ' INSERT + UPDATE�i�͈͑I���j�{�^��
    Dim btnInsertUpdateSelected         As CommandBarButton
    
    ' INSERT + UPDATE�|�b�v�A�b�v��ǉ�����
    Set popInsertUpdate = popData.Controls.Add(Type:=msoControlPopup)
    
    With popInsertUpdate
        
        .Caption = "INSERT + UPDATE"
    End With
    
    ' INSERT + UPDATE�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnInsertUpdate = popInsertUpdate.Controls.Add(Type:=msoControlButton)
    
    ' INSERT + UPDATE�{�^���̃v���p�e�B��ݒ肷��
    With btnInsertUpdate
    
        .Style = msoButtonIconAndCaption
        .Caption = "�S��"
        .DescriptionText = "INSERT + UPDATE - �S��"
        .OnAction = "Main.SutInsertUpdateAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutInsertUpdateAll"

        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnInsertUpdate _
                                   , "Add"
        End If

    End With
    
    ' INSERT + UPDATE�i�͈͑I���j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnInsertUpdateSelected = popInsertUpdate.Controls.Add(Type:=msoControlButton)
    
    ' INSERT + UPDATE�i�͈͑I���j�{�^���̃v���p�e�B��ݒ肷��
    With btnInsertUpdateSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "�͈͑I��"
        .DescriptionText = "INSERT + UPDATE - �͈͑I��"
        .OnAction = "Main.SutInsertUpdateSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutInsertUpdateSelection"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnInsertUpdateSelected _
                                   , "AreaAdd"
        End If
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' INSERT
    ' ***************************************************************
    ' INSERT�|�b�v�A�b�v
    Dim popInsert                 As commandBarPopup
    ' INSERT�{�^��
    Dim btnInsert                 As CommandBarButton
    ' INSERT�i�͈͑I���j�{�^��
    Dim btnInsertSelected         As CommandBarButton
    
    ' INSERT�|�b�v�A�b�v��ǉ�����
    Set popInsert = popData.Controls.Add(Type:=msoControlPopup)
    
    With popInsert
        
        .Caption = "INSERT"
    End With
    
    ' INSERT�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnInsert = popInsert.Controls.Add(Type:=msoControlButton)
    
    ' INSERT�{�^���̃v���p�e�B��ݒ肷��
    With btnInsert
    
        .Style = msoButtonIconAndCaption
        .Caption = "�S��"
        .DescriptionText = "INSERT - �S��"
        .OnAction = "Main.SutInsertAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutInsertAll"

        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnInsert _
                                   , "Add"
        End If

    End With
    
    ' INSERT�i�͈͑I���j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnInsertSelected = popInsert.Controls.Add(Type:=msoControlButton)
    
    ' INSERT�i�͈͑I���j�{�^���̃v���p�e�B��ݒ肷��
    With btnInsertSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "�͈͑I��"
        .DescriptionText = "INSERT - �͈͑I��"
        .OnAction = "Main.SutInsertSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutInsertSelection"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnInsertSelected _
                                   , "AreaAdd"
        End If
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' UPDATE
    ' ***************************************************************
    ' UPDATE�|�b�v�A�b�v
    Dim popUpdate                 As commandBarPopup
    ' UPDATE�{�^��
    Dim btnupdate                 As CommandBarButton
    ' UPDATE�i�͈͑I���j�{�^��
    Dim btnUpdateSelected         As CommandBarButton
    
    ' �e�[�u���|�b�v�A�b�v��ǉ�����
    Set popUpdate = popData.Controls.Add(Type:=msoControlPopup)
    
    With popUpdate
    
        .Caption = "UPDATE"
    End With
    
    ' UPDATE�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnupdate = popUpdate.Controls.Add(Type:=msoControlButton)
    
    ' UPDATE�{�^���̃v���p�e�B��ݒ肷��
    With btnupdate
    
        .Style = msoButtonIconAndCaption
        .Caption = "�S��"
        .DescriptionText = "UPDATE - �S��"
        .OnAction = "Main.SutUpdateAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutUpdateAll"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnupdate _
                                   , "Edit"
        End If
    End With
    
    ' UPDATE�i�͈͑I���j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnUpdateSelected = popUpdate.Controls.Add(Type:=msoControlButton)
    
    ' UPDATE�i�͈͑I���j�{�^���̃v���p�e�B��ݒ肷��
    With btnUpdateSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "�͈͑I��"
        .DescriptionText = "UPDATE - �͈͑I��"
        .OnAction = "Main.SutUpdateSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutUpdateSelection"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnUpdateSelected _
                                   , "AreaEdit"
        End If
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' DELETE
    ' ***************************************************************
    ' DELETE�|�b�v�A�b�v
    Dim popDelete                 As commandBarPopup
    ' DELETE�{�^��
    Dim btnDelete                 As CommandBarButton
    ' DELETE�i�͈͑I���j�{�^��
    Dim btnDeleteSelected         As CommandBarButton
    ' DELETE�i�e�[�u����̑S���R�[�h�j�{�^��
    Dim btnDeleteAllOfTable       As CommandBarButton
    
    ' �e�[�u���|�b�v�A�b�v��ǉ�����
    Set popDelete = popData.Controls.Add(Type:=msoControlPopup)
    
    With popDelete
    
        .Caption = "DELETE"
    End With
    
    ' DELETE�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnDelete = popDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETE�{�^���̃v���p�e�B��ݒ肷��
    With btnDelete
    
        .Style = msoButtonIconAndCaption
        .Caption = "�S��"
        .DescriptionText = "DELETE - �S��"
        .OnAction = "Main.SutDeleteAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDeleteAll"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDelete _
                                   , "Remove"
        End If
    End With
    
    ' DELETE�i�͈͑I���j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnDeleteSelected = popDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETE�i�͈͑I���j�{�^���̃v���p�e�B��ݒ肷��
    With btnDeleteSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "�͈͑I��"
        .DescriptionText = "DELETE - �͈͑I��"
        .OnAction = "Main.SutDeleteSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDeleteSelection"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDeleteSelected _
                                   , "AreaRemove"
        End If
    End With
    
    ' DELETE�i�e�[�u����̑S���R�[�h�j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnDeleteAllOfTable = popDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETE�i�e�[�u����̑S���R�[�h�j�{�^���̃v���p�e�B��ݒ肷��
    With btnDeleteAllOfTable
    
        .Style = msoButtonIconAndCaption
        .Caption = "�e�[�u����̑S���R�[�h"
        .DescriptionText = "DELETE - �e�[�u����̑S���R�[�h"
        .OnAction = "Main.SutDeleteAllOfTable"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDeleteAllOfTable"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDeleteAllOfTable _
                                   , "Bug"
        End If
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' SELECT
    ' ***************************************************************
    ' SELECT�|�b�v�A�b�v
    Dim popSelect                 As commandBarPopup
    ' SELECT�{�^��
    Dim btnSelect                 As CommandBarButton
    ' SELECT�i�����w��j�{�^��
    Dim btnSelectSelected         As CommandBarButton
    ' SELECT�i�O��̏����Ŏ��s�j�{�^��
    Dim btnSelectReExecute        As CommandBarButton
    
    ' �e�[�u���|�b�v�A�b�v��ǉ�����
    Set popSelect = popData.Controls.Add(Type:=msoControlPopup)
    
    With popSelect
    
        .Caption = "SELECT"
    End With
    
    ' SELECT�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnSelect = popSelect.Controls.Add(Type:=msoControlButton)
    
    ' SELECT�{�^���̃v���p�e�B��ݒ肷��
    With btnSelect
    
        .Style = msoButtonIconAndCaption
        .Caption = "�S��"
        .DescriptionText = "SELECT - �S��"
        .OnAction = "Main.SutSelectAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSelectAll"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnSelect _
                                   , "Search"
        End If
    End With
    
    ' SELECT�i�����w��j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnSelectSelected = popSelect.Controls.Add(Type:=msoControlButton)
    
    ' SELECT�i�����w��j�{�^���̃v���p�e�B��ݒ肷��
    With btnSelectSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "�����w��"
        .DescriptionText = "SELECT - �����w��"
        .OnAction = "Main.SutSelectCondition"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSelectCondition"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnSelectSelected _
                                   , "AreaSearch"
        End If
    End With
    
    ' SELECT�i�Ď��s�j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnSelectReExecute = popSelect.Controls.Add(Type:=msoControlButton)
    
    ' SELECT�i�Ď��s�j�{�^���̃v���p�e�B��ݒ肷��
    With btnSelectReExecute
    
        .Style = msoButtonIconAndCaption
        .Caption = "�Ď��s"
        .DescriptionText = "SELECT - �Ď��s"
        .OnAction = "Main.SutSelectReExecute"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSelectReExecute"
        
    End With
    
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' �N�G���G�f�B�^�i !!! ������ !!! �j
    ' ***************************************************************
'    ' �N�G���G�f�B�^�̒ǉ�
'    Dim btnQueryEditor As CommandBarButton
'
'    ' �N�G���G�f�B�^�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
'    Set btnQueryEditor = popData.Controls.Add(Type:=msoControlButton)
'
'    ' �N�G���G�f�B�^�{�^���̃v���p�e�B��ݒ肷��
'    With btnQueryEditor
'
'        .Style = msoButtonIconAndCaption
'        .Caption = "�N�G���G�f�B�^"
'        .DescriptionText = "�N�G���G�f�B�^"
'        .OnAction = "Main.SutQueryEditor"
'        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutQueryEditor"
'
'        ' Excel2002�ȍ~�̃v���p�e�B
'        If excelVer >= Ver2002 Then
'            setCommandBarControlIcon btnQueryEditor _
'                                   , RESOURCE_ICON.EDIT _
'                                   , RESOURCE_ICON.EDIT_MASK
'        End If
'
'    End With
    
    ' ***************************************************************
    ' �ꊇ�N�G��
    ' ***************************************************************
    ' �ꊇ�N�G���̒ǉ�
    Dim btnQueryBatch As CommandBarButton
    
    ' �s�̒ǉ��{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnQueryBatch = popData.Controls.Add(Type:=msoControlButton)
    
    ' SELECT�{�^���̃v���p�e�B��ݒ肷��
    With btnQueryBatch
    
        .Style = msoButtonIconAndCaption
        .Caption = "�N�G���ꊇ���s"
        .DescriptionText = "�N�G���ꊇ���s"
        .OnAction = "Main.SutQueryBatch"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutQueryBatch"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnQueryBatch _
                                   , "Forward"
        End If
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' �N�G������
    ' ***************************************************************
    ' �N�G������
    Dim btnQueryResult             As CommandBarButton
    
    ' �N�G�����ʃ{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnQueryResult = popData.Controls.Add(Type:=msoControlButton)
    
    ' �N�G�����ʃ{�^���̃v���p�e�B��ݒ肷��
    With btnQueryResult
    
        .Style = msoButtonIconAndCaption
        .Caption = "�Ō�Ɏ��s�����N�G������"
        .DescriptionText = "�Ō�Ɏ��s�����N�G������"
        .OnAction = "Main.SutShowQueryResult"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowQueryResult"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnQueryResult _
                                   , "AlertMessage"
        End If
        
    End With
    
    ' ***************************************************************
    ' �s�̒ǉ��E�폜
    ' ***************************************************************
    ' �s�̒ǉ�
    Dim btnRecAdd As CommandBarButton
    
    ' �s�̒ǉ��{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnRecAdd = popData.Controls.Add(Type:=msoControlButton)
    
    ' SELECT�{�^���̃v���p�e�B��ݒ肷��
    With btnRecAdd
    
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "�s�̒ǉ��E�폜"
        .DescriptionText = "�s�̒ǉ��E�폜"
        .OnAction = "Main.SutRecordAdd"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutRecordAdd"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' �N�G���p�����[�^
    ' ***************************************************************
    ' �N�G���p�����[�^�{�^��
    Dim btnQueryParameter As CommandBarButton
    
    ' �N�G���p�����[�^�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnQueryParameter = popData.Controls.Add(Type:=msoControlButton)
    
    ' �N�G���p�����[�^�{�^���̃v���p�e�B��ݒ肷��
    With btnQueryParameter
    
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "�N�G���p�����[�^"
        .DescriptionText = "�N�G���p�����[�^"
        .OnAction = "Main.SutSettingQueryParameter"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingQueryParameter"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnQueryParameter _
                                   , "Run"
        End If
    End With
    ' ***************************************************************
    
    ' ***************************************************************
    ' �t�@�C��
    ' ***************************************************************
    ' �t�@�C���|�b�v�A�b�v
    Dim popFile                   As commandBarPopup
    
    ' �t�@�C���|�b�v�A�b�v��ǉ�����
    Set popFile = cb.Controls.Add(Type:=msoControlPopup)
    
    With popFile
    
        .Caption = "�t�@�C��"
    End With
    ' ***************************************************************
    
    ' ***************************************************************
    ' INSERT�o��
    ' ***************************************************************
    ' INSERT�|�b�v�A�b�v
    Dim popFileInsert                 As commandBarPopup
    ' INSERT�{�^��
    Dim btnFileInsert                 As CommandBarButton
    ' INSERT�i�͈͑I���j�{�^��
    Dim btnFileInsertSelected         As CommandBarButton
    
    ' INSERT�|�b�v�A�b�v��ǉ�����
    Set popFileInsert = popFile.Controls.Add(Type:=msoControlPopup)
    
    With popFileInsert
        
        .Caption = "INSERT SQL"
    End With
    
    ' INSERT�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnFileInsert = popFileInsert.Controls.Add(Type:=msoControlButton)
    
    ' INSERT�{�^���̃v���p�e�B��ݒ肷��
    With btnFileInsert
    
        .Style = msoButtonIconAndCaption
        .Caption = "�S��"
        .DescriptionText = "�t�@�C���o�� INSERT SQL - �S��"
        .OnAction = "Main.SutFileInsertAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileInsertAll"
        
    End With
    
    ' INSERT�i�͈͑I���j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnFileInsertSelected = popFileInsert.Controls.Add(Type:=msoControlButton)
    
    ' INSERT�i�͈͑I���j�{�^���̃v���p�e�B��ݒ肷��
    With btnFileInsertSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "�͈͑I��"
        .DescriptionText = "�t�@�C���o�� INSERT SQL - �͈͑I��"
        .OnAction = "Main.SutFileInsertSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileInsertSelection"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' UPDATE�o��
    ' ***************************************************************
    ' UPDATE�|�b�v�A�b�v
    Dim popFileUpdate                 As commandBarPopup
    ' UPDATE�{�^��
    Dim btnFileUpdate                 As CommandBarButton
    ' UPDATE�i�͈͑I���j�{�^��
    Dim btnFileUpdateSelected         As CommandBarButton
    
    ' UPDATE�|�b�v�A�b�v��ǉ�����
    Set popFileUpdate = popFile.Controls.Add(Type:=msoControlPopup)
    
    With popFileUpdate
        
        .Caption = "UPDATE SQL"
    End With
    
    ' UPDATE�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnFileUpdate = popFileUpdate.Controls.Add(Type:=msoControlButton)
    
    ' UPDATE�{�^���̃v���p�e�B��ݒ肷��
    With btnFileUpdate
    
        .Style = msoButtonIconAndCaption
        .Caption = "�S��"
        .DescriptionText = "�t�@�C���o�� UPDATE SQL - �S��"
        .OnAction = "Main.SutFileUpdateAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileUpdateAll"
        
    End With
    
    ' UPDATE�i�͈͑I���j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnFileUpdateSelected = popFileUpdate.Controls.Add(Type:=msoControlButton)
    
    ' UPDATE�i�͈͑I���j�{�^���̃v���p�e�B��ݒ肷��
    With btnFileUpdateSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "�͈͑I��"
        .DescriptionText = "�t�@�C���o�� UPDATE SQL - �͈͑I��"
        .OnAction = "Main.SutFileUpdateSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileUpdateSelection"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' DELETE�o��
    ' ***************************************************************
    ' DELETE�|�b�v�A�b�v
    Dim popFileDelete                 As commandBarPopup
    ' DELETE�{�^��
    Dim btnFileDelete                 As CommandBarButton
    ' DELETE�i�͈͑I���j�{�^��
    Dim btnFileDeleteSelected         As CommandBarButton
    
    ' DELETE�|�b�v�A�b�v��ǉ�����
    Set popFileDelete = popFile.Controls.Add(Type:=msoControlPopup)
    
    With popFileDelete
        
        .Caption = "DELETE SQL"
    End With
    
    ' DELETE�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnFileDelete = popFileDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETE�{�^���̃v���p�e�B��ݒ肷��
    With btnFileDelete
    
        .Style = msoButtonIconAndCaption
        .Caption = "�S��"
        .DescriptionText = "�t�@�C���o�� DELETE SQL - �S��"
        .OnAction = "Main.SutFileDeleteAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileDeleteAll"
        
    End With
    
    ' DELETE�i�͈͑I���j�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnFileDeleteSelected = popFileDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETE�i�͈͑I���j�{�^���̃v���p�e�B��ݒ肷��
    With btnFileDeleteSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "�͈͑I��"
        .DescriptionText = "�t�@�C���o�� DELETE SQL - �͈͑I��"
        .OnAction = "Main.SutFileDeleteSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileDeleteSelection"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' �ꊇ�t�@�C���o��
    ' ***************************************************************
    ' DELETE�i�͈͑I���j�{�^��
    Dim btnFileBatch         As CommandBarButton
    
    ' DELETE�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnFileBatch = popFile.Controls.Add(Type:=msoControlButton)
    
    ' DELETE�{�^���̃v���p�e�B��ݒ肷��
    With btnFileBatch
    
        .Style = msoButtonIconAndCaption
        .Caption = "�ꊇ�o��"
        .DescriptionText = "�t�@�C���ꊇ�o��"
        .OnAction = "Main.SutFileBatch"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileBatch"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnFileBatch _
                                   , "Forward"
        End If
        
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' Diff
    ' ***************************************************************
    ' Diff�|�b�v�A�b�v
    Dim popDiff                   As commandBarPopup
    
    ' Diff�|�b�v�A�b�v��ǉ�����
    Set popDiff = cb.Controls.Add(Type:=msoControlPopup)
    
    With popDiff
    
        .Caption = "Diff"
    End With
    ' ***************************************************************
    
    ' ***************************************************************
    ' DB�X�i�b�v�V���b�g�擾�t�H�[���Ăяo��
    ' ***************************************************************
    ' �X�i�b�v�V���b�g�擾
    Dim btnShowDBSnapshot As CommandBarButton
    
    ' �X�i�b�v�V���b�g�擾�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnShowDBSnapshot = popDiff.Controls.Add(Type:=msoControlButton)
    
    ' �X�i�b�v�V���b�g�擾�{�^���̃v���p�e�B��ݒ肷��
    With btnShowDBSnapshot
    
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "�X�i�b�v�V���b�g�擾�E��r"
        .DescriptionText = "�X�i�b�v�V���b�g�擾�E��r"
        .OnAction = "Main.SutShowSnapshot"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowSnapshot"
        
    End With
    
    ' ***************************************************************
    ' DB�X�i�b�v�V���b�gSQL��`�V�[�g�ǉ�
    ' ***************************************************************
    ' �X�i�b�v�V���b�gSQL�V�[�g�ǉ�
    Dim btnNewSheetDataSnapshotSqlDefine As CommandBarButton
    
    ' �X�i�b�v�V���b�gSQL�V�[�g�ǉ��{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnNewSheetDataSnapshotSqlDefine = popDiff.Controls.Add(Type:=msoControlButton)
    
    ' �X�i�b�v�V���b�gSQL�V�[�g�ǉ��{�^���̃v���p�e�B��ݒ肷��
    With btnNewSheetDataSnapshotSqlDefine
    
        .BeginGroup = False
        .Style = msoButtonIconAndCaption
        .Caption = "�X�i�b�v�V���b�gSQL�V�[�g�ǉ�"
        .DescriptionText = "�X�i�b�v�V���b�gSQL�V�[�g�ǉ�"
        .OnAction = "Main.SutCreateNewSheetSnapSqlDefine"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutCreateNewSheetSnapSqlDefine"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' �c�[��
    ' ***************************************************************
    ' �c�[���|�b�v�A�b�v
    Dim popTool             As commandBarPopup
    ' �I�v�V�����{�^��
    Dim btnOption           As CommandBarButton
    ' �E�N���b�N���j���[�̃J�X�^�}�C�Y�{�^��
    Dim btnRClickMenuCustom As CommandBarButton
    ' �V���[�g�J�b�g�L�[�̊��蓖�ă{�^��
    Dim btnShortCutKey      As CommandBarButton
    ' �|�b�v�A�b�v���j���[�̃J�X�^�}�C�Y�{�^��
    Dim btnPopupKey         As CommandBarButton
    
    ' �c�[���|�b�v�A�b�v��ǉ�����
    Set popTool = cb.Controls.Add(Type:=msoControlPopup)
    
    With popTool
    
        .Caption = "�c�[��"
    End With
    
    ' �I�v�V�����{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnOption = popTool.Controls.Add(Type:=msoControlButton)
    
    ' �I�v�V�����{�^���̃v���p�e�B��ݒ肷��
    With btnOption
    
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "�I�v�V����"
        .DescriptionText = "�I�v�V����"
        .OnAction = "Main.SutSettingOption"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingOption"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnOption _
                                   , "Settings"
        End If
    End With
    
    ' �E�N���b�N���j���[�̃J�X�^�}�C�Y�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnRClickMenuCustom = popTool.Controls.Add(Type:=msoControlButton)
    
    ' �E�N���b�N���j���[�̃J�X�^�}�C�Y�{�^���̃v���p�e�B��ݒ肷��
    With btnRClickMenuCustom
    
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "�E�N���b�N���j���[�̐ݒ�"
        .DescriptionText = "�E�N���b�N���j���[�̐ݒ�"
        .OnAction = "Main.SutSettingRClickMenu"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingRClickMenu"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnRClickMenuCustom _
                                   , "FlagRed"
        End If
    End With
    
    ' �V���[�g�J�b�g�L�[�̊��蓖�ă{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnShortCutKey = popTool.Controls.Add(Type:=msoControlButton)
    
    ' �V���[�g�J�b�g�L�[�̊��蓖�ă{�^���̃v���p�e�B��ݒ肷��
    With btnShortCutKey
    
        .Style = msoButtonIconAndCaption
        .Caption = "�V���[�g�J�b�g�L�[�̐ݒ�"
        .DescriptionText = "�V���[�g�J�b�g�L�[�̐ݒ�"
        .OnAction = "Main.SutSettingShortCutKey"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingShortCutKey"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnShortCutKey _
                                   , "FlagGreen"
        End If
    End With
    
    ' �V���[�g�J�b�g�L�[�̊��蓖�ă{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnPopupKey = popTool.Controls.Add(Type:=msoControlButton)
    
    ' �V���[�g�J�b�g�L�[�̊��蓖�ă{�^���̃v���p�e�B��ݒ肷��
    With btnPopupKey
    
        .Style = msoButtonIconAndCaption
        .Caption = "�|�b�v�A�b�v���j���[�̐ݒ�"
        .DescriptionText = "�|�b�v�A�b�v���j���[�̐ݒ�"
        .OnAction = "Main.SutSettingPopupMenu"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingPopupMenu"
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnPopupKey _
                                   , "FlagBlue"
        End If
    End With
    
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' �w���v
    ' ***************************************************************
    ' �w���v�|�b�v�A�b�v
    Dim popHelp           As commandBarPopup
    ' �w���v
    Dim btnHelp           As CommandBarButton
    ' ���C�Z���X
    Dim btnLicence        As CommandBarButton
    ' �o�[�W����
    Dim btnVersion        As CommandBarButton
    
    ' �c�[���|�b�v�A�b�v��ǉ�����
    Set popHelp = cb.Controls.Add(Type:=msoControlPopup)
    
    With popHelp
    
        .Caption = "�w���v"
    End With
    
    ' �w���v�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnHelp = popHelp.Controls.Add(Type:=msoControlButton)
    
    ' �w���v�{�^���̃v���p�e�B��ݒ肷��
    With btnHelp
    
        .Style = msoButtonIconAndCaption
        .Caption = "Sut�w���v"
        .DescriptionText = "Sut�w���v"
        .OnAction = "Main.SutShowHelpFile"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowHelpFile"
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnHelp _
                                   , "Book"
        End If
        
    End With
    
    ' �o�[�W���������R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnVersion = popHelp.Controls.Add(Type:=msoControlButton)
    
    ' �o�[�W�������{�^���̃v���p�e�B��ݒ肷��
    With btnVersion
    
        .Style = msoButtonIconAndCaption
        .Caption = "�o�[�W�������"
        .DescriptionText = "�o�[�W�������"
        .OnAction = "Main.SutShowVersion"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowVersion"
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnVersion _
                                   , "AlertMessage"
        End If
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' DB�ڑ����
    ' ***************************************************************
    ' DB�ڑ����|�b�v�A�b�v
    Dim btnDBConnectInfo As CommandBarButton
    
    ' DB�ڑ����|�b�v�A�b�v��ǉ�����
    Set btnDBConnectInfo = cb.Controls.Add(Type:=msoControlButton)
    
    With btnDBConnectInfo
    
        .Style = msoButtonCaption
        .Caption = ""
        .DescriptionText = "DB�ڑ����"
        .OnAction = "Main.SutDBConnectInfo"
        .visible = False
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDBConnectInfo"
    End With
    
    ' ***************************************************************
    
    ' ���[�h�{�^���������s�ɂ���
    Dim btnLoad As CommandBarButton
    Set btnLoad = cb.FindControl(Tag:=COMMANDBAR_CONTROL_BASE_ID & "Main.SutLoad", recursive:=True)
    
    If Not btnLoad Is Nothing Then
    
        btnLoad.Enabled = False
    End If

    ' �A�����[�h�{�^���������\�ɂ���
    Dim btnUnload As CommandBarButton
    Set btnUnload = cb.FindControl(Tag:=COMMANDBAR_CONTROL_BASE_ID & "Main.SutUnload", recursive:=True)
    
    If Not btnUnload Is Nothing Then
    
        btnUnload.Enabled = True
    End If
    
    cb.visible = True

    On Error GoTo 0
    
End Function

' =========================================================
' ���c�[���o�[�̍폜����
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function deleteToolbar()

    On Error Resume Next
    
    ' �R�}���h�o�[
    Dim cb   As CommandBar
    
    Set cb = Application.CommandBars.item(ConstantsCommon.COMMANDBAR_MENU_NAME)
        
    ' �擾�Ɏ��s�����ꍇ�A�ϐ�cb��nothing�ɂȂ�
    ' �ϐ�cb��nothing�̏ꍇ�́A�����𒆒f����
    If cb Is Nothing Then
    
        Exit Function
        
    End If
    
    cb.delete
    
    On Error GoTo 0
    
End Function

' =========================================================
' ���c�[���o�[�̍폜�����i����̃��j���[�͎c���j
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function deleteToolbarExcludeSomeItems()

    On Error Resume Next
    
    ' �R�}���h�o�[
    Dim cb   As CommandBar
    
    Set cb = Application.CommandBars.item(ConstantsCommon.COMMANDBAR_MENU_NAME)
        
    ' �擾�Ɏ��s�����ꍇ�A�ϐ�cb��nothing�ɂȂ�
    ' �ϐ�cb��nothing�̏ꍇ�́A�����𒆒f����
    If cb Is Nothing Then
    
        Exit Function
        
    End If
    
    Dim ctl As commandBarControl
    
    For Each ctl In cb.Controls
    
        If ctl.Tag <> ConstantsCommon.COMMANDBAR_DONT_DELETE_TARGET Then
        
            ' �R���g���[�����폜����
            ctl.delete
        End If
    Next
    
    ' ���[�h�{�^���������\�ɂ���
    Dim btnLoad As CommandBarButton
    Set btnLoad = cb.FindControl(Tag:=COMMANDBAR_CONTROL_BASE_ID & "Main.SutLoad", recursive:=True)
    
    If Not btnLoad Is Nothing Then
    
        btnLoad.Enabled = True
    End If

    ' �A�����[�h�{�^���������s�ɂ���
    Dim btnUnload As CommandBarButton
    Set btnUnload = cb.FindControl(Tag:=COMMANDBAR_CONTROL_BASE_ID & "Main.SutUnload", recursive:=True)
    
    If Not btnUnload Is Nothing Then
    
        btnUnload.Enabled = False
    End If
    
    On Error GoTo 0
    
End Function

' =========================================================
' ���R�}���h�o�[�̃A�C�R����ݒ肷�鏈��
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function setCommandBarControlIcon(ByVal control As Object _
                                        , ByVal iconName As String)
                                   
    control.Picture = LoadPicture(SutWorkbook.path & "\resource\icon\" & iconName & "_16x16.bmp")
    control.mask = LoadPicture(SutWorkbook.path & "\resource\icon\" & iconName & "_16x16_mask.bmp")

End Function

' =========================================================
' �������S�ʂ̌㏈��
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function doAfterProcess()

    On Error Resume Next
    
    ' �����I����ɁAExcel�E�B���h�E���A�N�e�B�u�ɂȂ炸�ɁA���̃E�B���h�E���A�N�e�B�u�ɂȂ鎖�ۂ��m�F
    ' ������󂯂āA�ȉ��̂悤�ɁA���݂̃A�N�e�B�u�u�b�N���A�N�e�B�u�ɂ���悤�ɖ����I�Ɏw�肷��
    'Application.ActiveWindow.activate

    On Error GoTo 0

End Function

' =========================================================
' ��DB�ڑ��X�e�[�^�X��ݒ肷�鏈��
'
' �T�v�@�@�@�F
' �����@�@�@�FdbConnStr_        DB�ڑ�������
'     �@�@�@�FdbConnSimpleStr_  DB�ڑ�������i�P���Ȗ��O�j
'     �@�@�@�Fconn              DB�ڑ��L��
'
' =========================================================
Private Sub changeDbConnectStatus(ByRef dbConnStr_ As String, ByRef dbConnSimpleStr_ As String, ByVal conn As Boolean)

    ' �R�}���h�o�[
    Dim cb   As CommandBar
    Set cb = Application.CommandBars.item(ConstantsCommon.COMMANDBAR_MENU_NAME)
        
    ' ���ɒǉ�����Ă���ꍇ�́A�ϐ�cb��nothing�ɂȂ�
    ' �ϐ�cb��nothing�̏ꍇ�́A�����𒆒f����
    If cb Is Nothing Then
        Exit Sub
    End If
    
    Dim btnConn          As CommandBarButton
    Dim btnDisconn       As CommandBarButton
    Dim btnDBConnectInfo As CommandBarButton
    
    Set btnConn = cb.FindControl(Tag:=ConstantsCommon.COMMANDBAR_CONTROL_BASE_ID & "Main.SutConnectDB", recursive:=True)
    Set btnDisconn = cb.FindControl(Tag:=ConstantsCommon.COMMANDBAR_CONTROL_BASE_ID & "Main.SutDisconnectDB", recursive:=True)
    Set btnDBConnectInfo = cb.FindControl(Tag:=ConstantsCommon.COMMANDBAR_CONTROL_BASE_ID & "Main.SutDBConnectInfo", recursive:=True)
    
    If _
        Not btnConn Is Nothing And _
        Not btnDisconn Is Nothing Then
    
        If conn = True Then
        
            btnConn.state = msoButtonDown
            btnDisconn.state = msoButtonUp
        Else
        
            btnConn.state = msoButtonUp
            btnDisconn.state = msoButtonDown
        End If
    End If
    
    If _
        Not btnDBConnectInfo Is Nothing Then
    
        btnDBConnectInfo.Caption = "( " & dbConnSimpleStr_ & " )"
        btnDBConnectInfo.visible = conn
    End If
    
End Sub

' =========================================================
' ���o�[�W���������擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�o�[�W�������
' ���L�����@�F
'
' =========================================================
Public Function getVersionInfo() As String
    
    Dim version     As String
    Dim machineName As String
    
    version = ConstantsCommon.version
    
    #If VBA7 And Win64 Then
        machineName = "64bit"
    #Else
        machineName = "32bit"
    #End If
    
    #If DEBUG_MODE = "1" Then
        machineName = machineName & " !!! IS DEBUG MODE"
    #End If
    
    getVersionInfo = machineName & " - ver " & version
    
End Function
