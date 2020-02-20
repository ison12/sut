VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���j���[���쎞�̏��� - Diff
'
' �쐬�ҁ@�FIson
' �����@�@�F2018/12/30�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �v���O���X�t�H�[��
Private WithEvents frmProgressVar               As frmProgress
Attribute frmProgressVar.VB_VarHelpID = -1
' DB�ڑ��t�H�[���i�C�x���g��M�p�ɒ�`�j
Private WithEvents frmDBConnectVar              As frmDBConnect
Attribute frmDBConnectVar.VB_VarHelpID = -1
' �X�i�b�v�V���b�g�擾�t�H�[��
Private WithEvents frmSnapShotVar               As frmSnapshot
Attribute frmSnapShotVar.VB_VarHelpID = -1
' �X�i�b�v�V���b�g�擾�t�H�[���i��r�j
Private WithEvents frmSnapshotDiffVar           As frmSnapshotDiff
Attribute frmSnapshotDiffVar.VB_VarHelpID = -1

' �X�i�b�v�V���b�gSQL��`�V�[�g����
Private WithEvents snapSqlDefineSheetCreator    As ExeSnapSqlDefineSheetCreator
Attribute snapSqlDefineSheetCreator.VB_VarHelpID = -1
' �X�i�b�v�V���b�gSQL��`�V�[�g�ǂݎ��
Private WithEvents snapSqlDefineSheetReader     As ExeSnapSqlDefineSheetReader
Attribute snapSqlDefineSheetReader.VB_VarHelpID = -1
' �X�i�b�v�V���b�gSQL���s
Private WithEvents snapQueryExecution           As ExeSnapQueryExecution
Attribute snapQueryExecution.VB_VarHelpID = -1
' �X�i�b�v�V���b�g��r
Private WithEvents snapDiff                     As ExeSnapDiff
Attribute snapDiff.VB_VarHelpID = -1

' �A�v���P�[�V�����ݒ�
Private applicationSetting As ValApplicationSetting

' DB�R�l�N�V�����I�u�W�F�N�g
Private dbConn As Object
' SQL��`���
Private sqlDefineWorkSheet As ValSnapSqlDefineWorkSheet
' �X�i�b�v�V���b�g
Private snapshotSet As ValCollection

' ________________________________________________________
' �C�x���g��`
' ________________________________________________________
Private Sub frmProgressVar_cancel()

    ' �L�����Z���t���O��ON�ɂ���
    If Not snapSqlDefineSheetCreator Is Nothing Then
        snapSqlDefineSheetCreator.canceled = True
    End If
    
    If Not snapSqlDefineSheetReader Is Nothing Then
        snapSqlDefineSheetReader.canceled = True
    End If
    
    If Not snapQueryExecution Is Nothing Then
        snapQueryExecution.canceled = True
    End If
    
    If Not snapDiff Is Nothing Then
        snapDiff.canceled = True
    End If

End Sub

' =========================================================
' ��DB�ڑ��t�H�[����OK�{�^�����������ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub frmDBConnectVar_ok(ByVal connStr As String, ByVal connSimpleStr As String, ByVal connectInfo As ValDBConnectInfo)

    On Error GoTo err

    ' DB�ڑ�����Ă���ꍇ
    If Not Main.dbConn Is Nothing Then
    
        ' �ؒf����
        Main.SutUpdateDbConn Nothing, Empty, Empty
    End If
    
    ' DB�ڑ����ϐ��ɐݒ肷��
    Dim conn As Object
    Set conn = ADOUtil.connectDb(connStr)
    
    ' DB�ڑ��������ݒ肷��
    Main.SutUpdateDbConn conn, connStr, connSimpleStr
    
    Exit Sub
    
err:

    ' �G���[����
    Main.SutUpdateDbConn Nothing, Empty, Empty

    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DB�ڑ��t�H�[���ŏ������L�����Z�����ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub frmDBConnectVar_cancel()

    On Error GoTo err

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾�t�H�[�� - �X�i�b�v�V���b�g�擾
' =========================================================
Private Sub frmSnapShotVar_execSnapshot(sheet As Worksheet)
    
    On Error GoTo err
    
    Dim ssKey  As String
    Dim ssList As ValCollection
    Dim ss     As ValSnapRecordsSet

    Set snapSqlDefineSheetReader = New ExeSnapSqlDefineSheetReader
    Set snapSqlDefineSheetReader.sheet = sheet
    Set snapSqlDefineSheetReader.conn = dbConn
    
    ' SQL��`����ǂݎ��
    Set sqlDefineWorkSheet = snapSqlDefineSheetReader.readSqlDefineInfo
    
    ' SQL�����s���X�i�b�v�V���b�g���擾����
    
    ' Excel.Application�v���p�e�B�𑀍삷��I�u�W�F�N�g
    ' �֐��𔲂���Ǝ�����Application�v���p�e�B�����������
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' �v���O���X�E�B���h�E��\������
    frmProgressVar.ShowExt vbModeless, False
    frmProgressVar.title = "�N�G�����s"
    
    ' �G�f�B�^���J���Ă���Ԃɐڑ���Ԃ��ύX����Ă���\��������̂ŁA�R�l�N�V������ݒ肷��
    Dim conn As Object
    Set conn = Main.getDBConnection
    
    Set snapQueryExecution = New ExeSnapQueryExecution
    Set snapQueryExecution.conn = conn
    snapQueryExecution.applicationSetting = Main.getApplicationSetting ' �G�f�B�^���J���Ă���Ԃɐڑ���Ԃ��ύX����Ă���\��������̂ŁA�A�v���P�[�V�����ݒ�����擾����
    snapQueryExecution.canceled = False
    
    ' -----------------------------------------------------------
    ' �X�i�b�v�V���b�g�f�[�^�̊i�[
    Set ss = snapQueryExecution.execute(sqlDefineWorkSheet)
    
    ssKey = sheet.parent.name & " " & sheet.name
    
    Set ssList = snapshotSet.getItem(ssKey)
    If ssList Is Nothing Then
        Set ssList = New ValCollection
        snapshotSet.setItem ssList, sheet.parent.name & " " & sheet.name
    End If
    
    ssList.setItem ss
    frmSnapShotVar_changeSql sheet
    ' -----------------------------------------------------------
    
    ' ExcelLongTimeProcessing�ɂāAApplication.interactive�𖳌��ɂ�����Ԃ�
    ' �_�C�A���O�����J����������肷��ƁA�E�B���h�E��Z�I�[�_�[������Ă��܂��B�i�����I����ɁAExcel���A�N�e�B�u�ɂȂ�Ȃ��j
    ' ���̂��߁AExcelLongTimeProcessing�I�u�W�F�N�g�𖾎��I�ɔj������
    Set longTimeProcessing = Nothing

    frmProgressVar.HideExt
    frmSnapShotVar.Show vbModeless ' �A�N�e�B�u�ɂ��邽�߂ɕ\�����\�b�h���Ăяo��
    
    Exit Sub
err:

    ' �G���[����ޔ�����
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    frmProgressVar.HideExt
    frmSnapShotVar.Show vbModeless ' �A�N�e�B�u�ɂ��邽�߂ɕ\�����\�b�h���Ăяo��

    ' �ޔ������G���[����ݒ肵�Ȃ���
    VBUtil.setErr errT

    ' �G���[����
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾�t�H�[�� - �X�i�b�v�V���b�g�擾�L�����Z��
' =========================================================
Private Sub frmSnapShotVar_cancel()

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾�t�H�[�� - DB�ύX
' =========================================================
Private Sub frmSnapShotVar_changeDb()

    ' DB�ڑ��t�H�[����\������
    frmDBConnect.ShowExt vbModal

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾�t�H�[�� - SQL��`�ύX
' =========================================================
Private Sub frmSnapShotVar_changeSql(sheet As Worksheet)

    Dim ssKey  As String
    Dim ssList As ValCollection
    Dim ss     As ValSnapRecordsSet

    ssKey = sheet.parent.name & " " & sheet.name
    
    Set ssList = snapshotSet.getItem(ssKey)
    If Not ssList Is Nothing Then
    
        frmSnapShotVar.clearSnapshot
        
        For Each ss In ssList.col
            frmSnapShotVar.addSnapshot Format(ss.getDate, "yyyy/mm/dd hh:nn:ss") & " - " & ss.recordCount & "��", Empty
        Next
    
    Else
    
        frmSnapShotVar.clearSnapshot
        
    End If

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾�t�H�[�� - �X�i�b�v�V���b�g���X�g�N���A
' =========================================================
Private Sub frmSnapShotVar_clearSnapshot(sheet As Worksheet)

    Dim resultMsg As Long
    
    resultMsg = VBUtil.showMessageBoxForYesNo("�X�i�b�v�V���b�g���N���A���Ă���낵���ł����H" _
                   & vbNewLine _
                   & vbNewLine & "�E[�͂�] �� �N���A" _
                   & vbNewLine & "�E[������] �� �������Ȃ�" _
                , ConstantsCommon.APPLICATION_NAME)

    ' ����������
    If resultMsg = WinAPI_User.IDNO Then
    
        Exit Sub
    End If

    Dim ssKey  As String
    Dim ssList As ValCollection
    Dim ss     As ValSnapRecordsSet

    ssKey = sheet.parent.name & " " & sheet.name
    
    Set ssList = snapshotSet.getItem(ssKey)
    If Not ssList Is Nothing Then
        snapshotSet.remove ssKey
    End If
    
    frmSnapShotVar.clearSnapshot

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾�t�H�[�� - �X�i�b�v�V���b�g���X�g�N���A
' =========================================================
Private Sub frmSnapShotVar_showSnapshotDiff(ByRef sheet As Worksheet)

    Dim ssKey  As String
    Dim ssList As ValCollection
    Dim ss     As ValSnapRecordsSet

    ssKey = sheet.parent.name & " " & sheet.name
    
    Set ssList = snapshotSet.getItem(ssKey)
    If Not ssList Is Nothing Then
        ' ���[�_���\������̂Ń��[�h����Ɏ��s����
        If VBUtil.unloadFormIfChangeActiveBook(frmSnapshotDiff) Then Unload frmSnapshotDiff
        Load frmSnapshotDiff
        Set frmSnapshotDiffVar = frmSnapshotDiff
        frmSnapshotDiffVar.ShowExt vbModal, ssList
        
    Else
    
        VBUtil.showMessageBoxForWarning "�X�i�b�v�V���b�g������܂���B" _
                , ConstantsCommon.APPLICATION_NAME _
                , Nothing

    End If

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾��r�t�H�[�� - �L�����Z��
' =========================================================
Private Sub frmSnapshotDiffVar_Cancel()

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾��r�t�H�[�� - ��r���s
' =========================================================
Private Sub frmSnapshotDiffVar_execDiff(ByRef snapShotList As ValCollection, ByVal srcIndex As Long, ByVal desIndex As Long)

    On Error GoTo err
    
    ' ���[�_���t�H�[�����\���ɂ���
    frmSnapshotDiffVar.Hide

    ' �v���O���X�E�B���h�E��\������
    frmProgressVar.ShowExt vbModeless, False
    frmProgressVar.title = "�X�i�b�v�V���b�g��r"

    Set snapDiff = New ExeSnapDiff
    snapDiff.book = ActiveWorkbook
    snapDiff.applicationSetting = applicationSetting
    
    Dim valSnapRecordsSet1 As ValSnapRecordsSet
    Set valSnapRecordsSet1 = snapShotList.getItemByIndex(srcIndex + 1) ' ���X�g�̗v�f�́A1�n�܂�Ȃ̂ŁA+1����
    
    Dim valSnapRecordsSet2 As ValSnapRecordsSet
    Set valSnapRecordsSet2 = snapShotList.getItemByIndex(desIndex + 1) ' ���X�g�̗v�f�́A1�n�܂�Ȃ̂ŁA+1����
    
    snapDiff.execute valSnapRecordsSet1, valSnapRecordsSet2
    
    frmProgressVar.HideExt
    
    ' ���[�_���t�H�[�������ɖ߂�
    frmSnapshotDiffVar.Show vbModal

    Exit Sub
err:
    
    ' �G���[����ޔ�����
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    frmProgressVar.HideExt
    
    ' �ޔ������G���[����ݒ肵�Ȃ���
    VBUtil.setErr errT
    
    ' �G���[����
    Main.ShowErrorMessage
    
    ' ���[�_���t�H�[�������ɖ߂�
    frmSnapshotDiffVar.Show

End Sub

' =========================================================
' ���X�i�b�v�V���b�gSQL��`�V�[�g�쐬 - �V�[�g�쐬�J�n
' =========================================================
Private Sub snapSqlDefineSheetCreator_createSnapshotSqlDefineSheetBegin()

    frmProgressVar.initPri 1, "�X�i�b�v�V���b�gSQL��`�V�[�g�쐬��"
    
    DoEvents

End Sub

' =========================================================
' ���X�i�b�v�V���b�gSQL��`�V�[�g�쐬 - �V�[�g�쐬����
' =========================================================
Private Sub snapSqlDefineSheetCreator_createSnapshotSqlDefineSheet()

    frmProgressVar.inclimentPri
    
    DoEvents

End Sub

' =========================================================
' ���X�i�b�v�V���b�gSQL���s - SQL���X�g���s�J�n
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlListBegin(ByVal count As Long)

    frmProgressVar.initPri count, "SQL���s��"

End Sub

' =========================================================
' ���X�i�b�v�V���b�gSQL���s - SQL���X�g���s�I��
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlListComplete()

    frmProgressVar.priMessage = "SQL���s����"
    
    DoEvents

    ' �����Ԃ̏����̃J�[�\����������
    Dim cursorWait As New ExcelCursorWait: cursorWait.forceRestore

End Sub

' =========================================================
' ���X�i�b�v�V���b�gSQL���s - SQL���s�J�n
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlBegin(sql As String)

End Sub

' =========================================================
' ���X�i�b�v�V���b�gSQL���s - SQL���s����
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlSuccessful(sql As String)

    frmProgressVar.inclimentPri
    
    DoEvents
    
End Sub

' =========================================================
' ���X�i�b�v�V���b�gSQL���s - SQL���s�G���[
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlError(ByRef sheet As Worksheet, ByVal recordSetIndex As Long, sql As String, sqlState As String, message As String)

    frmProgressVar.priMessage = "SQL���s�G���["
    
    ' �f�o�b�O���[�h�݂̂̎��s
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If
    
    DoEvents
    
    sheet.activate
    
    ' ���b�Z�[�W�{�b�N�X�Ɏ��s������\������
    VBUtil.showMessageBoxForError _
        recordSetIndex & "�s�ڂŃG���[���������܂����B" & vbNewLine & _
        "SQLSTATE=" & sqlState & vbNewLine & _
        " MESSAGE=" & message & vbNewLine & _
        "     SQL=" & sql _
        , ConstantsCommon.APPLICATION_NAME _
        , err
    
End Sub

' =========================================================
' ���X�i�b�v�V���b�g�̔�r���s - �V�[�g�����J�n
' =========================================================
Private Sub snapDiff_createSnapshotDiffSheetBegin()

    frmProgressVar.priMessage = "��r���ʃV�[�g������"

    DoEvents

    ' �����Ԃ̏��������s�����̂Ń}�E�X�J�[�\���������v�ɂ���
    Dim cursorWait As New ExcelCursorWait: cursorWait.init True

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�̔�r���s - �V�[�g����
' =========================================================
Private Sub snapDiff_createSnapshotDiffSheet(sheet As Worksheet)

    frmProgressVar.priMessage = "��r���ʃV�[�g��������"

    DoEvents

    ' �����Ԃ̏����̃J�[�\����������
    Dim cursorWait As New ExcelCursorWait: cursorWait.forceRestore

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�̔�r���s - ���R�[�h�Z�b�g��r�J�n
' =========================================================
Private Sub snapDiff_diffRecordSetBegin(ByVal count As Long)

    frmProgressVar.initPri count, "���R�[�h�Z�b�g��r��"

    DoEvents

    ' �����Ԃ̏��������s�����̂Ń}�E�X�J�[�\���������v�ɂ���
    Dim cursorWait As New ExcelCursorWait: cursorWait.init True

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�̔�r���s - ���R�[�h�Z�b�g��r�J�n
' =========================================================
Private Sub snapDiff_diffRecordSetOne()

    frmProgressVar.inclimentPri

    DoEvents

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�̔�r���s - ���R�[�h�Z�b�g��r
' =========================================================
Private Sub snapDiff_diffRecordSet()

    frmProgressVar.priMessage = "���R�[�h�Z�b�g��r����"

    DoEvents

    ' �����Ԃ̏����̃J�[�\����������
    Dim cursorWait As New ExcelCursorWait: cursorWait.forceRestore

End Sub

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set snapshotSet = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' �v���O���X�t�H�[����j������
    Set frmProgress = Nothing
    ' �C�x���g��M�p��DB�ڑ��t�H�[����j������
    Set frmDBConnectVar = Nothing
    ' �X�i�b�v�V���b�g�t�H�[����j������
    Set frmSnapShotVar = Nothing
    Set frmSnapshotDiffVar = Nothing
    
    ' �A�v���P�[�V�����ݒ����j������
    Set applicationSetting = Nothing
    ' DB�R�l�N�V������j������
    Set dbConn = Nothing
    ' SQL��`����j������
    Set sqlDefineWorkSheet = Nothing
    
    Set snapSqlDefineSheetCreator = Nothing
    Set snapSqlDefineSheetReader = Nothing
    Set snapQueryExecution = Nothing
    Set snapDiff = Nothing
    
    Set snapshotSet = Nothing

End Sub

' =========================================================
' ��������
'
' �T�v�@�@�@�F���������������s����B
' �����@�@�@�FappSetting �A�v���P�[�V�����ݒ���
' �@�@�@�@�@�@appSettingColFmt �A�v���P�[�V�����ݒ���i�J���������j
' �@�@�@�@�@�@conn       DB�R�l�N�V����
'
' =========================================================
Public Sub init(ByRef appSetting As ValApplicationSetting _
              , ByRef appSettingColFmt As ValApplicationSettingColFormat _
              , ByRef conn As Object)

    ' DB�ڑ��t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmDBConnect) Then Unload frmDBConnect
    Load frmDBConnect
    ' �C�x���g��M�p��DB�ڑ��t�H�[��������������
    Set frmDBConnectVar = frmDBConnect
    
    ' �X�i�b�v�V���b�g�t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmSnapshot) Then Unload frmSnapshot
    Load frmSnapshot
    ' �X�i�b�v�V���b�g�t�H�[����ݒ肷��
    Set frmSnapShotVar = frmSnapshot

    ' �X�i�b�v�V���b�g��r�t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmSnapshotDiff) Then Unload frmSnapshotDiff
    Load frmSnapshotDiff
    ' �X�i�b�v�V���b�g��r�t�H�[����ݒ肷��
    Set frmSnapshotDiffVar = frmSnapshotDiff

    ' �t�H�[����ݒ肷��
    If VBUtil.unloadFormIfChangeActiveBook(frmProgress) Then Unload frmProgress
    Load frmProgress
    ' �t�H�[����ݒ肷��
    Set frmProgressVar = frmProgress
    
    ' �A�v���P�[�V�����ݒ�����擾����
    Set applicationSetting = appSetting
    
    ' DB�R�l�N�V�������擾����
    Set dbConn = conn
    
    ' SQL��`��������������
    Set sqlDefineWorkSheet = Nothing

End Sub

' =========================================================
' ���X�i�b�v�V���b�gSQL�V�[�g�ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub createNewSheetSnapSqlDefine()

    On Error GoTo err

    ' Excel.Application�v���p�e�B�𑀍삷��I�u�W�F�N�g
    ' �֐��𔲂���Ǝ�����Application�v���p�e�B�����������
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' �v���O���X�E�B���h�E��\������
    frmProgressVar.ShowExt vbModeless, False
    frmProgressVar.title = "�X�i�b�v�V���b�gSQL�V�[�g�ǉ�"

    DoEvents

    ' �u�b�N�I�u�W�F�N�g
    Dim book As Workbook
    
    ' �u�b�N�I�u�W�F�N�g���擾����
    Set book = ActiveWorkbook
    
    ' �u�b�N�I�u�W�F�N�g�̃`�F�b�N
    If book Is Nothing Then
    
        Set book = Application.Workbooks.Add
        book.activate
    End If

    Set snapSqlDefineSheetCreator = New ExeSnapSqlDefineSheetCreator
    snapSqlDefineSheetCreator.applicationSetting = applicationSetting
    snapSqlDefineSheetCreator.book = ActiveWorkbook
    snapSqlDefineSheetCreator.createSheet

    ' ExcelLongTimeProcessing�ɂāAApplication.interactive�𖳌��ɂ�����Ԃ�
    ' �_�C�A���O�����J����������肷��ƁA�E�B���h�E��Z�I�[�_�[������Ă��܂��B�i�����I����ɁAExcel���A�N�e�B�u�ɂȂ�Ȃ��j
    ' ���̂��߁AExcelLongTimeProcessing�I�u�W�F�N�g�𖾎��I�ɔj������
    Set longTimeProcessing = Nothing
    
    ' �v���O���X�E�B���h�E���\���ɂ���
    frmProgressVar.HideExt
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' �v���O���X�E�B���h�E���\���ɂ���
    frmProgressVar.HideExt
    
    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�t�H�[���̕\��
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub showSnapshot()

    On Error GoTo err

    frmSnapShotVar.ShowExt vbModeless, applicationSetting, dbConn, dbConnStr
    
    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DB�R�l�N�V�����̍X�V
'
' �T�v�@�@�@�Fconn    DB�R�l�N�V����
'             connStr DB�ڑ�������
'
' =========================================================
Public Sub updateDbConn(ByRef conn As Object, ByVal connStr As String)

    On Error GoTo err
    
    Set dbConn = conn
    dbConnStr = connStr
    
    If Not frmSnapShotVar Is Nothing Then
        frmSnapShotVar.updateDbConn dbConn, dbConnStr
    End If

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub
