VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���R�[�h��I������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/02/01�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ________________________________________________________
' �C�x���g
' ________________________________________________________
' =========================================================
' ��ExeTableSheetReader�Œ�`���Ă���C�x���g���Ē�`
' �@��ʃ��W���[���ɔ��������C�x���g�����̂܂ܓ`�d���邽�߂ɍĒ�`���Ă���
' =========================================================
Public Event readColumnListBegin(ByVal cnt As Long)
Public Event readColumn(ByRef column As ValDbDefineColumn)
Public Event readColumnList(ByRef columnList As ValCollection)
Public Event readTableInfoBegin()
Public Event readTableInfo(ByRef table As ValTableWorksheet)
Public Event readRecordListBegin(ByVal cnt As Long)
Public Event readRecordList(ByRef recordList As ValRecordList)
Public Event readRecordListBySelectionBegin(ByVal cnt As Long)
Public Event readRecordListBySelectionOne(ByRef recordList As ValRecordList)
Public Event readRecordListBySelection(ByRef recordLists As ValCollection)

' =========================================================
' ��SQL���s�C�x���g
'
' �T�v�@�@�@�FSQL���s���ɌĂяo�����C�x���g
' �����@�@�@�Fsql        SQL�X�e�[�g�����g
'
' =========================================================
Public Event executeSql(ByRef sql As String)

' =========================================================
' ��SQL���s�����C�x���g
'
' �T�v�@�@�@�FSQL���s�ɐ��������ۂɌĂяo�����C�x���g
' �����@�@�@�Fsql        SQL�X�e�[�g�����g
'
' =========================================================
Public Event ExecuteSqlSuccessful(ByRef sql As String)

' =========================================================
' ��SQL���s���̃G���[�C�x���g
'
' �T�v�@�@�@�FSQL���s���ɃG���[�����������ۂɌĂяo�����C�x���g
' �����@�@�@�Fsql        SQL�X�e�[�g�����g
' �@�@�@�@�@�Fsqlstate   SQL�X�e�[�^�X
' �@�@�@�@�@�Fmessage    ���b�Z�[�W
'
' =========================================================
Public Event ExecuteSqlError(ByRef sql As String, ByRef sqlState As String, ByRef message As String)

' =========================================================
' ��SQL���s���ʔ��f���̃C�x���g�i���R�[�h�擾�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordGetBegin(ByVal recordListSize As Long)

' =========================================================
' ��SQL���s���ʔ��f���̃C�x���g�i���R�[�h�擾�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordGet(ByVal recordListSizeActual As Long _
                               , ByVal recordListSize As Long _
                               , ByRef recordList As Variant)

' =========================================================
' ��SQL���s���ʔ��f��̃C�x���g�i�擾���R�[�h�ҏW�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordConvertBegin(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ��SQL���s���ʔ��f��̃C�x���g�i�擾���R�[�h�ҏW�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordConvert(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ��SQL���s���ʔ��f���̃C�x���g�i�V�[�g���f�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordBegin(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ��SQL���s���ʔ��f��̃C�x���g�i�V�[�g���f�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecord(ByVal recordListSize As Long, ByRef recordList As Variant)


' ---------------------------------------------------------
' ���v���p�e�B
' ---------------------------------------------------------
' �e�[�u���V�[�g�ǂݍ���
Private WithEvents tableSheetReader As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1

' �e�[�u���V�[�g����
Private tableSheetCreator As New ExeTableSheetCreator
' �N�G���ϊ�
Private queryConverter As New DbQueryConverter
' �N�G���p�����[�^���X�g
Private queryParameterList_ As ValQueryParameterList

' DB�R�l�N�V����
Public conn          As Object
' ���[�N�V�[�g
Public sheet         As Worksheet
' �L�����Z���t���O
Public canceled      As Boolean

' ���[�N�V�[�g�ɒǉ����邩�ǂ���
Public appendSheet   As Boolean

' SQL���s���t���O
Private executingSql As Boolean

' �A�v���P�[�V�����ݒ���
Public applicationSetting As ValApplicationSetting
' �A�v���P�[�V�����ݒ�i�J�����������j
Public applicationSettingColFmt As ValApplicationSettingColFormat

' =========================================================
' ���N�G���p�����[�^���X�g�擾�v���p�e�B
' =========================================================
Public Property Get queryParameterList() As ValQueryParameterList

    If Not queryParameterList_ Is Nothing Then
        Set queryParameterList = queryParameterList_
        Exit Property
    End If

    Set queryParameterList_ = Nothing
        
    If sheet Is Nothing Then
        Exit Property
    End If
    
    Set queryParameterList_ = New ValQueryParameterList
    queryParameterList_.init sheet.parent
    queryParameterList_.readForData
    
    Set queryParameterList = queryParameterList_
    
End Property

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' �����R�[�h�擾���s�i�����w��Ȃ��j
'
' �T�v�@�@�@�F�e�[�u����`����ɃN�G���[�𐶐��A���s�����R�[�h���擾����B
' �����@�@�@�F
' �߂�l�@�@�F���R�[�h�擾����
'
' =========================================================
Public Function executeSelect() As Long

    ' �e�[�u���V�[�g
    Dim tableSheet As ValTableWorksheet

    ' �e�[�u���V�[�g�ǂݍ��݃I�u�W�F�N�g������������
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = sheet
    
    Set tableSheet = tableSheetReader.readTableInfo

    ' �N�G���[�����I�u�W�F�N�g�𐶐�����Factory�I�u�W�F�N�g
    Dim dbObjFactory As New DbObjectFactory
    ' �N�G���[�����I�u�W�F�N�g
    Dim queryCreator        As IDbQueryCreator
    
    ' �N�G���[�����I�u�W�F�N�g���擾�i��2,3�����͓K���j
    Set queryCreator = dbObjFactory.createQueryCreator(conn, 0, "", applicationSettingColFmt.getDbColFormatListByDbConn(conn), applicationSetting.schemaUse, applicationSetting.getTableColumnEscapeByDbConn(conn))

    ' ����
    Dim emptyCondition As New ValSelectCondition

    ' SELECT SQL
    Dim sql As String
    ' SELECT�𐶐�����
    sql = queryCreator.createSelect(tableSheet.table, emptyCondition)

    ' SELECT�����s����
    executeSelect = execute(tableSheet, sql)
    
End Function

' =========================================================
' �����R�[�h�擾���s�i�N�G���[�w��j
'
' �T�v�@�@�@�F�w�肳�ꂽ�N�G���[�����s�����R�[�h���擾����B
' �����@�@�@�Fsql    SELECT
' �߂�l�@�@�F���R�[�h�擾����
'
' =========================================================
Public Function executeSelectByQuery(ByRef sql As String) As Long

    ' �e�[�u���V�[�g
    Dim tableSheet As ValTableWorksheet

    ' �e�[�u���V�[�g�ǂݍ��݃I�u�W�F�N�g������������
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = sheet
    
    Set tableSheet = tableSheetReader.readTableInfo

    ' SELECT�����s����
    executeSelectByQuery = execute(tableSheet, sql)

End Function

' =========================================================
' �����R�[�h�擾���s
'
' �T�v�@�@�@�F�ΏۂƂȂ�e�[�u���̑S���R�[�h���V�[�g�ɔ��f����
' �����@�@�@�FtableSheet  �e�[�u���V�[�g
' �@�@�@�@�@�@sql    SELECT
' �߂�l�@�@�F���R�[�h�擾����
'
' =========================================================
Private Function execute(ByRef tableSheet As ValTableWorksheet, ByRef sql As String) As Long

    On Error GoTo err

    ' ���R�[�h�̑���
    Dim recordCount As Long
    ' ���R�[�h�Z�b�g
    Dim recordSet   As Object
    ' ���R�[�h�z��
    Dim recordList  As Variant
    ' ���R�[�h�z��i�s�T�C�Y�j
    Dim recordListSize As Long

    ' ���R�[�h���f�I�t�Z�b�g
    Dim recordOffset  As Long
    
    ' �Z���̍폜�E�}���p�ɗ��p����I�u�W�F�N�g
    Set tableSheetCreator = New ExeTableSheetCreator
    tableSheetCreator.applicationSetting = applicationSetting
    tableSheetCreator.book = sheet.parent
    
    ' ���R�[�h�����P��
    Dim processCount As Long: processCount = applicationSetting.getRecProcessCount

    ' �����̂͂��߂Ƀt���O���I�t�ɂ��Ă���
    executingSql = False

    ' -------------------------------------------------------------
    ' SQL�����s����
    ' -------------------------------------------------------------
    ' �L�����Z���t���O���`�F�b�N����
    checkCancel
    
    RaiseEvent executeSql(sql)

    ' SQL���s���t���O��ON�ɂ���
    executingSql = True
    
    queryConverter.initIfFirst Me.conn, queryParameterList
    sql = queryConverter.convertSql(sql)
    
    ' �N�G���[�����s����
    Set recordSet = ADOUtil.querySelect(conn, sql, 0)
    
    ' SQL���s���t���O��OFF�ɂ���
    executingSql = False

    RaiseEvent ExecuteSqlSuccessful(sql)


    ' -------------------------------------------------------------
    ' ���R�[�h���擾����
    ' -------------------------------------------------------------
    
    ' ���R�[�h���f�I�t�Z�b�g�ʒu���擾����
    recordOffset = getRecordOffset(tableSheet)
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
        
        sheet.activate
        sheet.Cells(recordOffset _
                  , ConstantsTable.U_COLUMN_OFFSET_COL).activate
    
    Else
    
        sheet.activate
        sheet.Cells(ConstantsTable.R_COLUMN_OFFSET_ROW _
                  , recordOffset).activate
    End If
    
    ' �L�����Z���t���O���`�F�b�N����
    checkCancel

    ' �擾���R�[�h��0���̏ꍇ
    If recordSet.EOF Then
    
        If appendSheet = False Then
        
            ' ���[�N�V�[�g�̓��͗̈���폜����
            tableSheetCreator.deleteCellOfRecord tableSheet, recordOffset
            ' ��̓��͗̈�𐶐�����
            tableSheetCreator.insertEmptyCell tableSheet, recordOffset, 1
            
        End If
            
        ' �߂�l�Ƃ���0��Ԃ�
        execute = 0
        Exit Function
        
    End If
        
    If appendSheet = False Then
    
        ' ���[�N�V�[�g�̓��͗̈���폜����
        tableSheetCreator.deleteCellOfRecord tableSheet, recordOffset
        
    End If
        
    ' �擾���R�[�h��0���ɂȂ�܂Ń��R�[�h�Z�b�g����������
    Do While Not recordSet.EOF
    
        ' �L�����Z���t���O���`�F�b�N����
        checkCancel
        
        ' �����R�[�h�擾�C�x���g
        RaiseEvent ReflectionRecordGetBegin(processCount)
        
        ' ���R�[�h�Z�b�g���烌�R�[�h���X�g���擾����
        recordList = getRecordList(recordSet, processCount)
        
        ' ���R�[�h���X�g�̃T�C�Y���擾����
        ' ADO�̃��R�[�h�Z�b�g����GetRows�Ŏ擾�����z���(��~�s)�Ńf�[�^���i�[����Ă���
        recordListSize = VBUtil.arraySize(recordList, 2)

        ' �����R�[�h�擾�C�x���g
        RaiseEvent ReflectionRecordGet(processCount, recordListSize, recordList)
        
        ' ���R�[�h���X�g���z��ł͂Ȃ��ꍇ
        If IsArray(recordList) = False Then
        
            GoTo loop_end
        End If
        
        ' �����R�[�h�ϊ��C�x���g
        RaiseEvent ReflectionRecordConvertBegin(recordListSize, recordList)
        
        ' ���R�[�h���X�g��ϊ�����
        convertRecordList tableSheet, recordList
        
        ' �����R�[�h�ϊ��C�x���g
        RaiseEvent ReflectionRecordConvert(recordListSize, recordList)
        
        ' �����R�[�h���f�C�x���g
        RaiseEvent ReflectionRecordBegin(recordListSize, recordList)
        
        ' ���[�N�V�[�g�Ƀ��R�[�h���X�g�̓��e�𔽉f����
        reflectSheet tableSheet, recordList, recordOffset
        
        ' �����R�[�h���f�C�x���g
        RaiseEvent ReflectionRecord(recordListSize, recordList)
        
        ' ����̃��R�[�h���X�g���f���̃I�t�Z�b�g�ʒu�̌v�Z
        ' ���R�[�h���X�g�̔��f��Ƀ��R�[�h���X�g�̃T�C�Y�����Z����
        recordOffset = recordOffset + recordListSize
        
        ' ���R�[�h�擾�������J�E���g����
        recordCount = recordCount + recordListSize
    Loop
    
loop_end:
    
    ' ���R�[�h�Z�b�g���������
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing

    ' ���R�[�h�擾������Ԃ�
    execute = recordCount

    Exit Function

err:

    ' �G���[�n���h���ŕʂ̊֐����Ăяo���ƃG���[��񂪏����Ă��܂����Ƃ�����̂�
    ' �\���̂ɃG���[����ۑ����Ă���
    Dim errT As errInfo: errT = VBUtil.swapErr

    ' ���R�[�h�Z�b�g���������
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    ' SQL���s���ɃG���[������
    If executingSql = True Then

        ' �G���[���b�Z�[�W��ݒ�
        errT.Description = ConstantsError.ERR_DESC_SQL_EXECUTE_FAILED _
                        & vbNewLine _
                        & vbNewLine _
                        & "[�ڍ׏��]" & vbNewLine _
                        & errT.Number & vbNewLine & errT.Description
                          
        ' �G���[�ԍ���ݒ�
        errT.Number = ConstantsError.ERR_NUMBER_SQL_EXECUTE_FAILED
        
        ' SQL�X�e�[�^�X
        Dim sqlState As String
        
        If conn.Errors.count > 0 Then
        
            ' �G���[�I�u�W�F�N�g���擾����
            Dim error As Object
            Set error = conn.Errors.item(0)
            sqlState = error.sqlState
        End If

        RaiseEvent ExecuteSqlError(sql, sqlState, CStr(errT.Description))

    End If

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' �����R�[�h�̃I�t�Z�b�g�ʒu�擾
'
' �T�v�@�@�@�F�擾�������R�[�h�𔽉f����I�t�Z�b�g�ʒu���擾����B
'
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �߂�l�@�@�F���R�[�h�̃I�t�Z�b�g�ʒu
'
' =========================================================
Private Function getRecordOffset(ByRef tableSheet As ValTableWorksheet) As Long

    ' �߂�l
    Dim recordOffset As Long
    
    ' �����̍s���폜�����ɁA�Ō�̈ʒu���烌�R�[�h��ǉ�����
    If appendSheet = True Then
        
        If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
        
            ' �ŏI���͍s���擾���{�P����
            recordOffset = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsTable.U_RECORD_NUM_COL) + 1
            
            ' ��L�Ŏ擾�����l���V�[�g�͈̔͂𒴂��Ă��邩���`�F�b�N����
            If ExcelUtil.checkOverMaxRow(sheet, recordOffset) = False Then
            
                ' �G���[�𔭍s����
                err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SHEET _
                        , _
                        , ConstantsError.ERR_DESC_OUT_OF_RANGE_SHEET
            End If
        Else
        
            ' �ŏI���͗���擾���{�P����
            recordOffset = ExcelUtil.getCellEndOfLastInputtedCol(sheet, ConstantsTable.R_RECORD_NUM_ROW) + 1
        
            ' ��L�Ŏ擾�����l���V�[�g�͈̔͂𒴂��Ă��邩���`�F�b�N����
            If ExcelUtil.checkOverMaxCol(sheet, recordOffset) = False Then
            
                ' �G���[�𔭍s����
                err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SHEET _
                        , _
                        , ConstantsError.ERR_DESC_OUT_OF_RANGE_SHEET
            End If
            
        End If
                
    ' �����̍s���폜���āA�ŏ��̈ʒu���烌�R�[�h��ǉ�����
    Else
    
        If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
        
            ' �萔���烌�R�[�h�I�t�Z�b�g���擾����
            recordOffset = ConstantsTable.U_RECORD_OFFSET_ROW
                        
        Else
        
            ' �萔���烌�R�[�h�I�t�Z�b�g���擾����
            recordOffset = ConstantsTable.R_RECORD_OFFSET_COL
        End If
    
    End If

    ' �߂�l��ݒ肷��
    getRecordOffset = recordOffset

End Function

' =========================================================
' �����R�[�h�擾
'
' �T�v�@�@�@�F���R�[�h��RecordSet�I�u�W�F�N�g����擾���A�z��Ɋi�[���ĕԂ��B
'
' �����@�@�@�FrecordSet ���R�[�h�Z�b�g
' �@�@�@�@�@�@fetchSize �t�F�b�`�T�C�Y
' �߂�l�@�@�F���R�[�h�z��
'
' =========================================================
Private Function getRecordList(ByRef recordSet As Object _
                             , Optional ByVal fetchSize As Long = 0) As Variant

    ' �߂�l
    Dim result         As Variant
    
    ' ���R�[�h�Z�b�g��EOF�ł͂Ȃ��ꍇ
    If Not recordSet.EOF Then
    
        ' �t�F�b�`�T�C�Y���`�F�b�N����
        If fetchSize = 0 Then
        
            ' ���R�[�h�Z�b�g����S���R�[�h���擾����
            result = recordSet.getRows
        Else
        
            ' ���R�[�h�Z�b�g���烌�R�[�h���t�F�b�`�T�C�Y���擾����
            result = recordSet.getRows(fetchSize)
        End If
    
        getRecordList = result
    Else
    
        ' ���Ԃ�
        getRecordList = Empty
    End If

End Function

' =========================================================
' �����R�[�h���X�g�ϊ�
'
' �T�v�@�@�@�F���R�[�h���X�g�̃f�[�^��ϊ�����B
' �����@�@�@�FtableSheet   �e�[�u���V�[�g
'             recordList   ���R�[�h���X�g
'
' =========================================================
Private Function convertRecordList(ByRef tableSheet As ValTableWorksheet, ByRef recordList As Variant)

    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ' ���R�[�h�z��̗v�f����ւ��ɂ���ĐV���ɐ��������z��
        Dim temp As Variant
        
        ' ���R�[�h�z��̃f�[�^�ϊ�����
        ' RecordSet.GetRows���Ԃ��ꂽ�z��́A(col,rec)�ƂȂ��Ă��邽��(rec,col)�ƂȂ�悤�ɓ���ւ������{����
        temp = convertArrayValuesToUnder(recordList)
        
        ' ���R�[�h�z����������
        Erase recordList

        ' ���R�[�h���X�g�����ւ���
        recordList = temp
        
    Else
        
        ' ���R�[�h�z��̃f�[�^�ϊ�����
        recordList = convertArrayValuesToRight(recordList)
    End If
        
    

End Function

' =========================================================
' ���z��v�f�̕ϊ�
'
' �T�v�@�@�@�F2�����z��̗v�f��(x,y)����(y,x)�ɐݒ肵�Ȃ����B
' �@�@�@�@�@�@����ɗv�f���Z���`���̒l�ɕϊ�����B
' �����@�@�@�Fv 2�����z��
'
' �߂�l�@�@�F2�����z��
' ���L�����@�FVBUtil.transposeDim�֐����J�X�^�}�C�Y�����֐��B
' �@�@�@�@�@�@�����͂قړ����B
'
' =========================================================
Private Function convertArrayValuesToUnder(ByRef v As Variant) As Variant
    
    Dim X As Long
    Dim Y As Long
    
    Dim Xlower As Long
    Dim Xupper As Long
    
    Dim Ylower As Long
    Dim Yupper As Long
    
    Dim tempArray As Variant
    
    Xlower = LBound(v, 2)
    Xupper = UBound(v, 2)
    Ylower = LBound(v, 1)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xlower To Xupper, Ylower To Yupper)
    
    Dim value As Variant
    
    For X = Xlower To Xupper
        For Y = Ylower To Yupper
        
            ' ��U�l���擾����
            value = v(Y, X)
            
            ' ������̏ꍇ�A�Z���`���̒l�ւƕϊ�����
            tempArray(X, Y) = ExcelUtil.convertCellValue(value)
        
        Next Y
    Next X
    
    convertArrayValuesToUnder = tempArray

End Function

' =========================================================
' ���z��v�f�̕ϊ�
'
' �T�v�@�@�@�F�v�f���Z���`���̒l�ɕϊ�����B
' �����@�@�@�Fv 2�����z��
'
' �߂�l�@�@�F2�����z��
' ���L�����@�F
'
' =========================================================
Private Function convertArrayValuesToRight(ByRef v As Variant) As Variant
    
    Dim X As Long
    Dim Y As Long
    
    Dim Xlower As Long
    Dim Xupper As Long
    
    Dim Ylower As Long
    Dim Yupper As Long
    
    Xlower = LBound(v, 2)
    Xupper = UBound(v, 2)
    Ylower = LBound(v, 1)
    Yupper = UBound(v, 1)
    
    Dim value As Variant
    
    For X = Xlower To Xupper
        For Y = Ylower To Yupper
        
            ' ��U�l���擾����
            value = v(Y, X)
            
            ' ������̏ꍇ�A�Z���`���̒l�ւƕϊ�����
            v(Y, X) = ExcelUtil.convertCellValue(value)
        
        Next Y
    Next X
    
    convertArrayValuesToRight = v
    
End Function

' =========================================================
' �����R�[�h���X�g�̃V�[�g���f
'
' �T�v�@�@�@�F���R�[�h���X�g���V�[�g�ɔ��f����
' �����@�@�@�FtableSheet   �e�[�u���V�[�g
'             recordList   ���R�[�h���X�g
' �@�@�@�@�@�@recordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' =========================================================
Private Sub reflectSheet(ByRef tableSheet As ValTableWorksheet, ByRef recordList As Variant, ByVal recordOffset As Long)

    Dim ret As Boolean
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ret = reflectSheetToUnder(tableSheet, recordList, recordOffset)
        
    Else
    
        ret = reflectSheetToRight(tableSheet, recordList, recordOffset)
    End If
    
    ' �߂�l�� False �̏ꍇ�A�G���[�𔭍s����
    If ret = False Then
    
        err.Raise ERR_NUMBER_OUT_OF_RANGE_SHEET _
                , _
                , ERR_DESC_OUT_OF_RANGE_SHEET
    End If

End Sub

' =========================================================
' �����R�[�h���X�g�̃V�[�g���f
'
' �T�v�@�@�@�F���R�[�h���X�g���V�[�g�ɔ��f����
' �����@�@�@�FtableSheet   �e�[�u���V�[�g
'             recordList   ���R�[�h���X�g
' �@�@�@�@�@�@recordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�FTrue  ����
' �@�@�@�@�@�@False �V�[�g�ւ̔��f���s
'
' =========================================================
Private Function reflectSheetToUnder(ByRef tableSheet As ValTableWorksheet _
                                   , ByRef recordList As Variant _
                                   , ByVal recordOffset As Long) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = True

    ' �e�[�u����`
    Dim table As ValDbDefineTable
    ' �e�[�u����`���擾����
    Set table = tableSheet.table
    
    ' ���R�[�h�T�C�Y
    Dim rowSize As Long
    ' �J�����T�C�Y
    Dim colSize As Long

    ' �z�񂩂烌�R�[�h�T�C�Y���擾����
    rowSize = VBUtil.arraySize(recordList)
    ' �z�񂩂�J�����T�C�Y���擾����
    colSize = VBUtil.arraySize(recordList, 2)

    ' ------------------------------------------------------------
    ' �V�[�g�͈̔̓`�F�b�N
    ' ------------------------------------------------------------
    If ExcelUtil.checkOverMaxRow(sheet, recordOffset, rowSize) = False Then
    
        ' �T�C�Y���I�[�o�[���Ă���̂Œ�������
        rowSize = ExcelUtil.getSizeOfSheetRow(sheet) - recordOffset + 1
    
        ' �S���R�[�h���f�ł��Ȃ��̂Ŗ߂�l��False�ɐݒ�
        ret = False
    End If
        
    ' ��̓��͗̈�𐶐�����
    tableSheetCreator.insertEmptyCell tableSheet, recordOffset, rowSize

    ' ------------------------------------------------------------
    ' �z��̓��e���V�[�g�ɔ��f����
    ' ------------------------------------------------------------
    ExcelUtil.copyArrayToCells recordList, sheet, recordOffset, ConstantsTable.U_COLUMN_OFFSET_COL, rowSize, colSize

    reflectSheetToUnder = ret
    
End Function

' =========================================================
' �����R�[�h���X�g�̃V�[�g���f
'
' �T�v�@�@�@�F���R�[�h���X�g���V�[�g�ɔ��f����
' �����@�@�@�FtableSheet   �e�[�u���V�[�g
'             recordList   ���R�[�h���X�g
' �@�@�@�@�@�@recordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�FTrue  ����
' �@�@�@�@�@�@False �V�[�g�ւ̔��f���s
'
' =========================================================
Private Function reflectSheetToRight(ByRef tableSheet As ValTableWorksheet _
                                   , ByRef recordList As Variant _
                                   , ByVal recordOffset As Long) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = True

    ' �e�[�u����`
    Dim table As ValDbDefineTable
    ' �e�[�u����`���擾����
    Set table = tableSheet.table
    
    ' ���R�[�h�T�C�Y
    Dim rowSize As Long
    ' �J�����T�C�Y
    Dim colSize As Long

    ' �z�񂩂烌�R�[�h�T�C�Y���擾����
    rowSize = VBUtil.arraySize(recordList, 2)
    ' �z�񂩂�J�����T�C�Y���擾����
    colSize = VBUtil.arraySize(recordList)

    ' ------------------------------------------------------------
    ' �V�[�g�͈̔̓`�F�b�N
    ' ------------------------------------------------------------
    If ExcelUtil.checkOverMaxCol(sheet, recordOffset, rowSize) = False Then
    
        ' �T�C�Y���I�[�o�[���Ă���̂Œ�������
        rowSize = ExcelUtil.getSizeOfSheetCol(sheet) - recordOffset + 1
    
        ' �S���R�[�h���f�ł��Ȃ��̂Ŗ߂�l��False�ɐݒ�
        ret = False
    End If

    ' ��̓��͗̈�𐶐�����
    tableSheetCreator.insertEmptyCell tableSheet, recordOffset, rowSize

    ' ------------------------------------------------------------
    ' �z��̓��e���V�[�g�ɔ��f����
    ' ------------------------------------------------------------
    ExcelUtil.copyArrayToCells recordList, sheet, ConstantsTable.R_COLUMN_OFFSET_ROW, recordOffset, colSize, rowSize

    reflectSheetToRight = ret

End Function

' =========================================================
' ���L�����Z���t���O�`�F�b�N
'
' �T�v�@�@�@�F�L�����Z���t���O���`�F�b�N���L�����Z������Ă����ꍇ��
' �@�@�@�@�@�@�G���[�𔭍s����B
'
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub checkCancel()

    If canceled = True Then

        err.Raise ConstantsError.ERR_NUMBER_PROC_CANCEL _
                , err.Source _
                , ConstantsError.ERR_DESC_PROC_CANCEL
    End If

End Sub

' --------------------������
' �e�[�u���V�[�g�ǂݍ��݃N���X�̃C�x���g��{�N���X�̃C�x���g�Ƃ��Ă��̂܂ܔ��s����
Private Sub tableSheetReader_readColumnListBegin(ByVal cnt As Long)

    RaiseEvent readColumnListBegin(cnt)
End Sub

Private Sub tableSheetReader_readColumnList(columnList As ValCollection)

    RaiseEvent readColumnList(columnList)
End Sub

Private Sub tableSheetReader_readColumn(column As ValDbDefineColumn)

    RaiseEvent readColumn(column)
End Sub

Private Sub tableSheetReader_readTableInfoBegin()

    RaiseEvent readTableInfoBegin
End Sub

Private Sub tableSheetReader_readTableInfo(table As ValTableWorksheet)

    RaiseEvent readTableInfo(table)
End Sub

Private Sub tableSheetReader_readRecordListBegin(ByVal cnt As Long)

    RaiseEvent readRecordListBegin(cnt)
End Sub

Private Sub tableSheetReader_readRecordList(recordList As ValRecordList)

    RaiseEvent readRecordList(recordList)
End Sub

Private Sub tableSheetReader_readRecordListBySelectionBegin(ByVal cnt As Long)

    RaiseEvent readRecordListBySelectionBegin(cnt)
End Sub

Private Sub tableSheetReader_readRecordListBySelectionOne(recordList As ValRecordList)

    RaiseEvent readRecordListBySelectionOne(recordList)
End Sub

Private Sub tableSheetReader_readRecordListBySelection(recordLists As ValCollection)

    RaiseEvent readRecordListBySelection(recordLists)
End Sub
' --------------------������
