VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���j���[���쎞�̏��� - File
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/03/14�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �v���O���X�t�H�[��
Private WithEvents frmProgressVar               As frmProgress
Attribute frmProgressVar.VB_VarHelpID = -1
' �t�@�C���o�̓t�H�[��
Private WithEvents frmFileOutputVar             As frmFileOutput
Attribute frmFileOutputVar.VB_VarHelpID = -1
' DB�N�G���ꊇ���s�t�H�[��
Private WithEvents frmDbQueryBatchVar           As frmDBQueryBatch
Attribute frmDbQueryBatchVar.VB_VarHelpID = -1
' ���R�[�h��񐶐�
Private WithEvents tableSheetReader             As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1

' �X�V�n�N�G���[����
Private WithEvents update                       As ExeUpdate
Attribute update.VB_VarHelpID = -1
' �X�V�n�N�G���[�����i�I���j
Private WithEvents updateSelection_             As ExeUpdate
Attribute updateSelection_.VB_VarHelpID = -1

' �t�@�C���o�͂� "�S��" �� "�͈͑I��" ��2��ނɕ������B
' �X�V�n�N�G���[�I�u�W�F�N�g(ExeUpdate)����������̂ŁA�ǂ��炪���s����邩�𔻒肷��I�u�W�F�N�g���K�v�ɂȂ�B�ȉ������̔���p�I�u�W�F�N�g�B
' �t�@�C���o�̓t�H�[���̌Ăяo���O�ɁA�ȉ���"update"�E"updateSelection_"�̉��ꂩ���i�[��
' �t�@�C���o�̓t�H�[����OK�C�x���g�ŁA���̃I�u�W�F�N�g�ɑ΂��đ�������s����B
Private updateCurrent As ExeUpdate

' �A�v���P�[�V�����ݒ�
Private applicationSetting As ValApplicationSetting

' ________________________________________________________
' �C�x���g��`
' ________________________________________________________
Private Sub frmProgressVar_cancel()

    ' �L�����Z���t���O��ON�ɂ���
    tableSheetReader.canceled = True
    update.canceled = True
    updateSelection_.canceled = True
    
End Sub

' =========================================================
' ���t�@�C���o��
'
' �T�v�@�@�@�F�t�@�C���o�͂����s����B
' �����@�@�@�Ftable      �e�[�u���I�u�W�F�N�g�P
' �@�@�@�@�@�@queryType  �N�G�����
' �@�@�@�@�@�@filePath   �t�@�C���p�X
' �@�@�@�@�@�@encode     �G���R�[�h
' �@�@�@�@�@�@newline    ���s�R�[�h
'
' =========================================================
Private Function execFileOutput(table As ValTableWorksheet _
                            , queryType As DB_QUERY_BATCH_TYPE _
                            , filePath As String _
                            , encode As String _
                            , newline As String) As Boolean

    On Error GoTo err

    ' Excel.Application�v���p�e�B�𑀍삷��I�u�W�F�N�g
    ' �֐��𔲂���Ǝ�����Application�v���p�e�B�����������
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' ���s����
    Dim resultCount As Long
    
    ' �X�V�I�u�W�F�N�g�̃t�@�C���o�͂ɕK�v�ȃI�v�V�����v���p�e�B��ݒ肷��
    ' �t�@�C���o�̓t���O
    update.isFileOutput = True
    ' �t�@�C���p�X
    update.filePath = VBUtil.concatFilePath(filePath, table.sheetName & ".sql")
    ' �����R�[�h
    update.charSet = encode
    ' ���s�R�[�h
    update.newline = newline
    
    ' �N�G����ނ̐ݒ�
    If queryType = DB_QUERY_BATCH_TYPE.insert Then

        update.queryType = update.QUERY_TYPE_INSERT

    ElseIf queryType = DB_QUERY_BATCH_TYPE.update Then

        update.queryType = update.QUERY_TYPE_UPDATE

    ElseIf queryType = DB_QUERY_BATCH_TYPE.deleteOnSheet Then

        update.queryType = update.QUERY_TYPE_DELETE
        
    ElseIf queryType = DB_QUERY_BATCH_TYPE.deleteAll Then

        update.queryType = update.QUERY_TYPE_DELETE_ALL

    End If
    
    ' �v���O���X�E�B���h�E��\������
    ' �����P�ʂ��ݒ肳��Ă���ꍇ
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgressVar.ShowExt vbModeless, True
    Else
        frmProgressVar.ShowExt vbModeless, False
    End If
    frmProgressVar.title = "�t�@�C���o�� - " & frmFileOutputVar.lblHeader.Caption
    
    resultCount = update.execute
    
    ' ExcelLongTimeProcessing�ɂāAApplication.interactive�𖳌��ɂ�����Ԃ�
    ' �_�C�A���O�����J����������肷��ƁA�E�B���h�E��Z�I�[�_�[������Ă��܂��B�i�����I����ɁAExcel���A�N�e�B�u�ɂȂ�Ȃ��j
    ' ���̂��߁AExcelLongTimeProcessing�I�u�W�F�N�g�𖾎��I�ɔj������
    Set longTimeProcessing = Nothing
    
    ' �v���O���X�E�B���h�E���\���ɂ���
    frmProgressVar.HideExt
    
    execFileOutput = True
    
    Exit Function
    
err:
    
    execFileOutput = False
    
    Set longTimeProcessing = Nothing
    
    ' �v���O���X�E�B���h�E���\���ɂ���
    frmProgressVar.HideExt

End Function

' =========================================================
' ���N�G���ꊇ�t�H�[���C�x���g - ���s
' =========================================================
Private Sub frmDbQueryBatchVar_ok(ByVal dbQueryBatchMode As DB_QUERY_BATCH_MODE _
                                , ByVal filePath As String _
                                , ByVal characterCode As String _
                                , ByVal newline As String _
                                , ByVal tableWorksheets As ValCollection)

    Dim var As ValDbQueryBatchTableWorksheet
    Dim tableSheet As ValTableWorksheet
        
    Dim ret As Boolean
        
    For Each var In tableWorksheets.col

        ' �e�[�u���V�[�g���擾����
        Set tableSheet = var.tableWorksheet
        ' �ΏۃV�[�g��ݒ肷��
        Set update.sheet = tableSheet.sheet

        ' �ΏۃV�[�g���A�N�e�B�u�ɂ���
        update.sheet.activate
        
        ' �t�@�C���o�͂����s����
        ret = execFileOutput(var.tableWorksheet, var.dbQueryBatchType, filePath, characterCode, newline)
        
        If ret = False Then
        
            err.Number = ERR_NUMBER_FILE_OUTPUT_FAILED
            err.Description = ERR_DESC_FILE_OUTPUT_FAILED
            Main.ShowErrorMessage
            
            Exit Sub
        End If
        
    Next
    
    VBUtil.showMessageBoxForInformation "�t�@�C�����o�͂���܂����B", "Sut"

End Sub

' =========================================================
' ���X�V�C�x���g - �J�������X�g�ǂݍ��݊J�n
' =========================================================
Private Sub update_readColumnListBegin(ByVal cnt As Long)

    frmProgressVar.initPri cnt, "�J������`�ǂݍ��ݒ�"
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g - �J�������X�g�ǂݍ���
' =========================================================
Private Sub update_readColumn(column As ValDbDefineColumn)

    frmProgressVar.inclimentPri
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g - �J�������X�g�ǂݍ��݊���
' =========================================================
Private Sub update_readColumnList(columnList As ValCollection)

    frmProgressVar.priMessage = "�J������`�ǂݍ��݊���"
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g - ���R�[�h���X�g�T�C�Y�ǂݍ���
' =========================================================
Private Sub update_readRecordListSize(ByVal cnt As Long)

    frmProgressVar.initPri cnt, "���R�[�h���X�g�T�C�Y�ǂݍ��݊���"
End Sub

' =========================================================
' ���X�V�C�x���g - ���R�[�h���X�g�ǂݍ��݊J�n
' =========================================================
Private Sub update_readRecordListBegin(ByVal cnt As Long)

    ' �����P�ʂ��ݒ肳��Ă���ꍇ
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        ' �������ď��������s���Ă��邱�Ƃ�������悤��
        ' �������R�[�h�͈̔͂����b�Z�[�W�i���̑��j�ɖ��ߍ���
        
        frmProgressVar.priMessage = "�E�E�E"
        frmProgressVar.initSec applicationSetting.recProcessCountCustom, "���R�[�h���X�g�ǂݍ��ݒ�"
        
    Else
    
        frmProgressVar.priMessage = "���R�[�h���X�g�ǂݍ��ݒ�"
        
    End If
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g - ���R�[�h���X�g�ǂݍ���
' =========================================================
Private Sub update_readRecordList(recordList As ValRecordList)

    ' �����P�ʂ��ݒ肳��Ă���ꍇ
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        ' �������ď��������s���Ă��邱�Ƃ�������悤��
        ' �������R�[�h�͈̔͂����b�Z�[�W�i���̑��j�ɖ��ߍ���
        
        frmProgressVar.secMessage = "���R�[�h���X�g�ǂݍ��݊���"
        frmProgressVar.secCount = VBUtil.arraySize(recordList.values)
        
    Else
    
        frmProgressVar.priMessage = "���R�[�h���X�g�ǂݍ��݊���"
        frmProgressVar.priCount = VBUtil.arraySize(recordList.values)
    End If
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g - SQL���s�J�n
' =========================================================
Private Sub update_ExecuteSqlBegin(ByVal count As Long)

    ' �����P�ʂ��ݒ肳��Ă���ꍇ
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        ' �������ď��������s���Ă��邱�Ƃ�������悤��
        ' �������R�[�h�͈̔͂����b�Z�[�W�i���̑��j�ɖ��ߍ���
        
        frmProgressVar.initSec count, "SQL�t�@�C���o�͒�"
    Else
    
        frmProgressVar.initPri count, "SQL�t�@�C���o�͒�"
        
    End If
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g - SQL���s
' =========================================================
Private Sub update_ExecuteSql(ByRef tableSheet As ValTableWorksheet _
                            , ByRef recordList As ValRecordList _
                            , ByVal recordListIndex As Long _
                            , ByRef sql As String _
                            , ByVal count As Long)

    On Error Resume Next

    ' �����P�ʂ��ݒ肳��Ă���ꍇ
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        frmProgressVar.inclimentSec
    Else
    
        frmProgressVar.inclimentPri
    End If
    
    DoEvents

    ' �f�o�b�O���[�h�݂̂̎��s
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    On Error GoTo 0

End Sub

' =========================================================
' ���X�V�C�x���g - SQL���s ����
' =========================================================
Private Sub update_ExecuteSqlSuccessful(ByVal count As Long)

    ' �����P�ʂ��ݒ肳��Ă���ꍇ
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        frmProgressVar.secMessage = "SQL�t�@�C���o�͊���"
        frmProgressVar.priCount = frmProgressVar.priCount + count
        
    Else
    
        frmProgressVar.priMessage = "SQL�t�@�C���o�͊���"
    End If
    
    DoEvents

End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - �J�������X�g�ǂݍ��݊J�n
' =========================================================
Private Sub updateSelection__readColumnListBegin(ByVal cnt As Long)

    frmProgressVar.initPri cnt, "�J������`�ǂݍ��ݒ�"
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - �J�������X�g�ǂݍ���
' =========================================================
Private Sub updateSelection__readColumn(column As ValDbDefineColumn)

    frmProgressVar.inclimentPri
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - �J�������X�g�ǂݍ��݊���
' =========================================================
Private Sub updateSelection__readColumnList(columnList As ValCollection)

    frmProgressVar.priMessage = "�J������`�ǂݍ��݊���"
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - �Z���I���� ���R�[�h���X�g�ǂݍ��݊J�n
' =========================================================
Private Sub updateSelection__readRecordListBySelectionBegin(ByVal cnt As Long)

    frmProgressVar.initPri cnt, "���R�[�h���X�g�ǂݍ��ݒ�"

    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - �Z���I���� ���R�[�h���X�g�ǂݍ���
' =========================================================
Private Sub updateSelection__readRecordListBySelectionOne(recordList As ValRecordList)

    Dim count As Long
    
    count = VBUtil.arraySize(recordList.values)
    
    frmProgressVar.priCount = frmProgressVar.priCount + count
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - �Z���I���� ���R�[�h���X�g�ǂݍ���
' =========================================================
Private Sub updateSelection__readRecordListBySelection(recordLists As ValCollection)

    frmProgressVar.priMessage = "���R�[�h���X�g�ǂݍ��݊���"
    
    ' �J�E���g������U0�N���A����i�J�E���g�����͂��̂܂܁j
    ' ExecuteSqlBegin�C�x���g�̓Z���I�����̏ꍇ�͕�����Ăяo�����̂�
    ' �J�E���g�����Z�b�g����悤�ȏ����͂����ɁA���̎�O�ŌĂяo�����{�C�x���g�Ń��Z�b�g�����{����
    frmProgressVar.priCount = 0
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - SQL���s�J�n
' =========================================================
Private Sub updateSelection__ExecuteSqlBegin(ByVal count As Long)

    frmProgressVar.priMessage = "SQL�t�@�C���o�͒�"
    
    DoEvents
End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - SQL���s
' =========================================================
Private Sub updateSelection__ExecuteSql(ByRef table As ValTableWorksheet _
                            , ByRef recordList As ValRecordList _
                            , ByVal recordListIndex As Long _
                            , ByRef sql As String _
                            , ByVal count As Long)

    frmProgressVar.inclimentPri

    ' �f�o�b�O���[�h�݂̂̎��s
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    DoEvents
    
End Sub

' =========================================================
' ���X�V�C�x���g�i�I���j - SQL���s ����
' =========================================================
Private Sub updateSelection__ExecuteSqlSuccessful(ByVal count As Long)

    frmProgressVar.priMessage = "SQL���s����"
    
    DoEvents

End Sub

' =========================================================
' ���t�@�C���o�̓t�H�[�� OK�{�^���������̃C�x���g�v���V�[�W��
' =========================================================
Private Sub frmFileOutputVar_ok(ByVal filePath As String, ByVal characterCode As String, ByVal newline As String)

    On Error GoTo err
    
    ' Excel.Application�v���p�e�B�𑀍삷��I�u�W�F�N�g
    ' �֐��𔲂���Ǝ�����Application�v���p�e�B�����������
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' ���s����
    Dim resultCount As Long
    
    ' �X�V�I�u�W�F�N�g�̃t�@�C���o�͂ɕK�v�ȃI�v�V�����v���p�e�B��ݒ肷��
    ' �t�@�C���o�̓t���O
    updateCurrent.isFileOutput = True
    ' �t�@�C���p�X
    updateCurrent.filePath = filePath
    ' �����R�[�h
    updateCurrent.charSet = characterCode
    ' ���s�R�[�h
    updateCurrent.newline = newline
    
    ' �v���O���X�E�B���h�E��\������
    ' �����P�ʂ��ݒ肳��Ă���ꍇ
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgressVar.ShowExt vbModeless, True
    Else
        frmProgressVar.ShowExt vbModeless, False
    End If
    frmProgressVar.title = "�t�@�C���o�� - " & frmFileOutputVar.lblHeader.Caption
    
    ' �X�V�����s����
    If updateCurrent Is update Then
    
        resultCount = updateCurrent.execute
    Else
    
        resultCount = updateCurrent.executeSelection
    End If
    
    ' ExcelLongTimeProcessing�ɂāAApplication.interactive�𖳌��ɂ�����Ԃ�
    ' �_�C�A���O�����J����������肷��ƁA�E�B���h�E��Z�I�[�_�[������Ă��܂��B�i�����I����ɁAExcel���A�N�e�B�u�ɂȂ�Ȃ��j
    ' ���̂��߁AExcelLongTimeProcessing�I�u�W�F�N�g�𖾎��I�ɔj������
    Set longTimeProcessing = Nothing
    
    ' �v���O���X�E�B���h�E���\���ɂ���
    frmProgressVar.HideExt
    
    ' ���b�Z�[�W�{�b�N�X�Ɏ��s������\������
    VBUtil.showMessageBoxForInformation resultCount & "���������܂����B", ConstantsCommon.APPLICATION_NAME
    
    Exit Sub
    
err:
    
    Set longTimeProcessing = Nothing
    
    ' �v���O���X�E�B���h�E���\���ɂ���
    frmProgressVar.HideExt

    ' �G���[����
    Main.ShowErrorMessage
    
End Sub


' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' �v���O���X�t�H�[����j������
    Set frmProgress = Nothing
    ' �t�@�C���o�̓t�H�[����j������
    Set frmFileOutputVar = Nothing
    ' �ꊇ�N�G���t�H�[����j������
    Set frmDbQueryBatchVar = Nothing
    
    ' ���R�[�h��񐶐��I�u�W�F�N�g��j������
    Set tableSheetReader = Nothing
    
    ' �X�V�N�G���[���s�I�u�W�F�N�g��j������
    Set update = Nothing
    Set updateSelection_ = Nothing
    Set updateCurrent = Nothing
    
    ' �A�v���P�[�V�����ݒ����j������
    Set applicationSetting = Nothing
    
End Sub

' =========================================================
' ��������
'
' �T�v�@�@�@�F���������������s����B
' �����@�@�@�FappSetting  �A�v���P�[�V�����ݒ���
' �@�@�@�@�@�@appSettingColFmt �A�v���P�[�V�����ݒ���i�J���������j
' �@�@�@�@�@�@conn        DB�R�l�N�V����
'
' =========================================================
Public Sub init(ByRef appSetting As ValApplicationSetting _
              , ByRef appSettingColFmt As ValApplicationSettingColFormat _
              , ByRef conn As Object)

    ' �t�H�[����ݒ肷��
    If VBUtil.unloadFormIfChangeActiveBook(frmProgress) Then Unload frmProgress
    Load frmProgress
    ' �t�H�[����ݒ肷��
    Set frmProgressVar = frmProgress
    
    ' �t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmFileOutput) Then Unload frmFileOutput
    Load frmFileOutput
    ' �t�H�[����ݒ肷��
    Set frmFileOutputVar = frmFileOutput
    
    ' ���R�[�h��񐶐��I�u�W�F�N�g�𐶐�����
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = ActiveSheet

    ' �X�V�N�G���[���s�I�u�W�F�N�g�𐶐�����
    Set update = New ExeUpdate
    Set update.conn = conn
    Set update.sheet = ActiveSheet
    Set update.applicationSetting = appSetting
    Set update.applicationSettingColFmt = appSettingColFmt
    
    ' �X�V�N�G���[���s�I�u�W�F�N�g�𐶐�����
    Set updateSelection_ = New ExeUpdate
    Set updateSelection_.conn = conn
    Set updateSelection_.sheet = ActiveSheet
    Set updateSelection_.applicationSetting = appSetting
    Set updateSelection_.applicationSettingColFmt = appSettingColFmt
    
    ' �A�v���P�[�V�����ݒ�����擾����
    Set applicationSetting = appSetting

End Sub

' =========================================================
' ��INSERT + UPDATE ���s�i�S���R�[�h�j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub insertUpdateAll()

    On Error GoTo err

    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
        
    ' �e�[�u���V�[�g�I�u�W�F�N�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���V�[�g�I�u�W�F�N�g���擾����
    Set tableSheet = tableSheetReader.readTableInfo

    ' �ΏۂƂȂ�X�V�n�N�G���[��ۑ����Ă���
    Set updateCurrent = update
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_INSERT_UPDATE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "INSERT + UPDATE SQL" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��INSERT + UPDATE ���s�i�I��͈́j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub insertUpdateSelection()

    On Error GoTo err

    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
        
    ' �e�[�u���I�u�W�F�N�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���I�u�W�F�N�g���擾����
    Set tableSheet = tableSheetReader.readTableInfo

    ' �ΏۂƂȂ�X�V�n�N�G���[��ۑ����Ă���
    Set updateCurrent = updateSelection_
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_INSERT_UPDATE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "INSERT + UPDATE SQL�i�͈͑I���j" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub
' =========================================================
' ��INSERT���s�i�S���R�[�h�j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub insertAll()

    On Error GoTo err

    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
        
    ' �e�[�u���V�[�g�I�u�W�F�N�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���V�[�g�I�u�W�F�N�g���擾����
    Set tableSheet = tableSheetReader.readTableInfo

    ' �ΏۂƂȂ�X�V�n�N�G���[��ۑ����Ă���
    Set updateCurrent = update
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_INSERT
    
    frmFileOutputVar.ShowExt vbModal _
                            , "INSERT SQL" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��INSERT���s�i�I��͈́j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub insertSelection()

    On Error GoTo err

    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
        
    ' �e�[�u���I�u�W�F�N�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���I�u�W�F�N�g���擾����
    Set tableSheet = tableSheetReader.readTableInfo

    ' �ΏۂƂȂ�X�V�n�N�G���[��ۑ����Ă���
    Set updateCurrent = updateSelection_
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_INSERT
    
    frmFileOutputVar.ShowExt vbModal _
                            , "INSERT SQL�i�͈͑I���j" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��UPDATE���s�i�S���R�[�h�j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub updateAll()

    On Error GoTo err

    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
        
    ' �e�[�u���I�u�W�F�N�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���I�u�W�F�N�g���擾����
    Set tableSheet = tableSheetReader.readTableInfo

    ' �ΏۂƂȂ�X�V�n�N�G���[��ۑ����Ă���
    Set updateCurrent = update
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_UPDATE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "UPDATE SQL" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��UPDATE���s�i�I��͈́j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub updateSelection()

    On Error GoTo err

    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
        
    ' �e�[�u���I�u�W�F�N�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���I�u�W�F�N�g���擾����
    Set tableSheet = tableSheetReader.readTableInfo

    ' �ΏۂƂȂ�X�V�n�N�G���[��ۑ����Ă���
    Set updateCurrent = updateSelection_
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_UPDATE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "UPDATE SQL�i�͈͑I���j" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DELETE���s�i�S���R�[�h�j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub deleteAll()

    On Error GoTo err

    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
        
    ' �e�[�u���V�[�g�I�u�W�F�N�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���I�u�W�F�N�g���擾����
    Set tableSheet = tableSheetReader.readTableInfo

    ' �ΏۂƂȂ�X�V�n�N�G���[��ۑ����Ă���
    Set updateCurrent = update
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_DELETE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "DELETE SQL" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DELETE���s�i�I��͈́j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub deleteSelection()

    On Error GoTo err

    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
        
    ' �e�[�u���V�[�g�I�u�W�F�N�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���V�[�g�I�u�W�F�N�g���擾����
    Set tableSheet = tableSheetReader.readTableInfo

    ' �ΏۂƂȂ�X�V�n�N�G���[��ۑ����Ă���
    Set updateCurrent = updateSelection_
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_DELETE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "DELETE SQL�i�͈͑I���j" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ���ꊇ�t�@�C���o��
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Sub batchFile()

    On Error GoTo err
        
    ' �u�b�N
    Dim book  As Workbook
    ' �A�N�e�B�u�u�b�N��book�ϐ��Ɋi�[����
    Set book = ActiveWorkbook
    
    ' �t�H�[����ݒ肷��
    If VBUtil.unloadFormIfChangeActiveBook(frmDBQueryBatch) Then Unload frmDBQueryBatch
    Load frmDBQueryBatch
    Set frmDbQueryBatchVar = frmDBQueryBatch
    
    frmDbQueryBatchVar.ShowExt vbModal, FileOutput, insertUpdate, book
    
    Set frmDbQueryBatchVar = Nothing

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub
