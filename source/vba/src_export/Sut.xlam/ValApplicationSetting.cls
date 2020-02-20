VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValApplicationSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �A�v���P�[�V�����̃I�v�V�������
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/03/14�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' =========================================================
' �������o
' =========================================================
' ���R�[�h�����P��
Public recProcessCount        As Long
' ���R�[�h�����P�ʁi���[�U�w��j
Public recProcessCountCustom  As Long
' �R�~�b�g�m�F
Public commitConfirm          As Long
' SQL�G���[���̋���
Public sqlErrorHandling       As Long
' �󔒃Z���ǂݎ�����
Public emptyCellReading       As Long
' ���ړ��͕���
Public directInputChar        As Long
' ���ړ��͕����i���[�U�w��j
Public directInputCharCustom  As String
' ���펞�̃N�G�����ʕ\��
Public queryResultShowWhenNormal As Boolean
' �X�L�[�}���p
Public schemaUse              As Long

' �e�[�u���E�J�������̃G�X�P�[�v
Public tableColumnEscapeOracle     As Boolean
Public tableColumnEscapeMysql      As Boolean
Public tableColumnEscapePostgresql As Boolean
Public tableColumnEscapeSqlserver  As Boolean
Public tableColumnEscapeAccess  As Boolean
Public tableColumnEscapeSymfoware  As Boolean

' �Z������ �t�H���g��
Public cellFontName As String
' �Z������ �t�H���g�T�C�Y
Public cellFontSize As Double
' �Z������ �܂�Ԃ��L��
Public cellWordwrap As Boolean
' �Z������ �Z�����i�P�ʂ͕������j
Public cellWidth As Double
' �Z������ �Z�������i�P�ʂ̓|�C���g�j
Public cellHeight As Double
' �s���̎�������
Public lineHeightAutoAdjust      As Boolean

' =========================================================
' ���萔�v���p�e�B
' �@Class���W���[���ł́APublic Const����`�ł��Ȃ��̂Ńv���p�e�B�ő�p����
' =========================================================
' �L�[�l ���R�[�h�����P��
Public Property Get KEY_REC_PROCESS_COUNT() As String

    KEY_REC_PROCESS_COUNT = "recProcessCount"
End Property

' �L�[�l ���R�[�h�����P�ʁ@���[�U�w��
Public Property Get KEY_REC_PROCESS_COUNT_CUSTOM() As String

    KEY_REC_PROCESS_COUNT_CUSTOM = "recProcessCountCustom"
End Property

' �L�[�l �R�~�b�g�m�F
Public Property Get KEY_COMMIT_CONFIRM() As String

    KEY_COMMIT_CONFIRM = "commitConfirm"
End Property

' �L�[�l SQL�G���[���̋���
Public Property Get KEY_SQL_ERROR_HANDLING() As String

    KEY_SQL_ERROR_HANDLING = "sqlErrorHandling"
End Property

' �L�[�l ���ړ��͕���
Public Property Get KEY_DIRECT_INPUT_CHAR() As String

    KEY_DIRECT_INPUT_CHAR = "directInputChar"
End Property

' �L�[�l ���ړ��͕����@���[�U�w��
Public Property Get KEY_DIRECT_INPUT_CHAR_CUSTOM() As String

    KEY_DIRECT_INPUT_CHAR_CUSTOM = "directInputCharCustom"
End Property

' �L�[�l �󔒃Z���ǂݎ�����
Public Property Get KEY_EMPTY_CELL_READING() As String

    KEY_EMPTY_CELL_READING = "emptyCellReading"
End Property

' �L�[�l ���펞�̃N�G�����ʕ\���L��
Public Property Get KEY_QUERY_RESULT_SHOW_WHEN_NORMAL() As String

    KEY_QUERY_RESULT_SHOW_WHEN_NORMAL = "queryResultShowWhenNormal"
End Property

' �L�[�l �X�L�[�}
Public Property Get KEY_SCHEMA_USE() As String

    KEY_SCHEMA_USE = "schemaUse"
End Property

' �L�[�l �e�[�u���J�������G�X�P�[�v Oracle
Public Property Get KEY_TABLE_COLUMN_ESCAPE_ORACLE() As String

    KEY_TABLE_COLUMN_ESCAPE_ORACLE = "tableColumnEscapeOracle"
End Property

' �L�[�l �e�[�u���J�������G�X�P�[�v Mysql
Public Property Get KEY_TABLE_COLUMN_ESCAPE_MYSQL() As String

    KEY_TABLE_COLUMN_ESCAPE_MYSQL = "tableColumnEscapeMysql"
End Property

' �L�[�l �e�[�u���J�������G�X�P�[�v Postgresql
Public Property Get KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL() As String

    KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL = "tableColumnEscapePostgresql"
End Property

' �L�[�l �e�[�u���J�������G�X�P�[�v Sqlserver
Public Property Get KEY_TABLE_COLUMN_ESCAPE_SQLSERVER() As String

    KEY_TABLE_COLUMN_ESCAPE_SQLSERVER = "tableColumnEscapeSqlserver"
End Property

' �L�[�l �e�[�u���J�������G�X�P�[�v Access
Public Property Get KEY_TABLE_COLUMN_ESCAPE_ACCESS() As String

    KEY_TABLE_COLUMN_ESCAPE_ACCESS = "tableColumnEscapeAccess"
End Property

' �L�[�l �e�[�u���J�������G�X�P�[�v Symfoware
Public Property Get KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE() As String

    KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE = "tableColumnEscapeSymfoware"
End Property

' �Z������ �t�H���g��
Public Property Get KEY_CELL_FONT_NAME() As String

    KEY_CELL_FONT_NAME = "cellFontName"
End Property

' �Z������ �t�H���g�T�C�Y
Public Property Get KEY_CELL_FONT_SIZE() As String
    
    KEY_CELL_FONT_SIZE = "cellFontSize"
End Property

' �Z������ �܂�Ԃ�
Public Property Get KEY_CELL_WORDWRAP() As String
    
    KEY_CELL_WORDWRAP = "cellWordwrap"
End Property

' �Z������ ��
Public Property Get KEY_CELL_WIDTH() As String
    
    KEY_CELL_WIDTH = "cellWidth"
End Property

' �Z������ ����
Public Property Get KEY_CELL_HEIGHT() As String
    
    KEY_CELL_HEIGHT = "cellHeight"
End Property

' �L�[�l �s���̎�������
Public Property Get KEY_LINE_HEIGHT_AUTO_ADJUST() As String

    KEY_LINE_HEIGHT_AUTO_ADJUST = "lineHeightAutoAdjust"
End Property

' ���R�[�h�����P�ʁi�S�āj
Public Property Get REC_PROCESS_COUNT_ALL() As Long

    REC_PROCESS_COUNT_ALL = 1
End Property

' ���R�[�h�����P�ʁi���[�U�w��j
Public Property Get REC_PROCESS_COUNT_COSTOM() As Long

    REC_PROCESS_COUNT_COSTOM = 3
End Property

' ���R�[�h�����P�ʁi���[�U�w��j�f�t�H���g�l
Public Property Get REC_PROCESS_COUNT_COSTOM_DEFAULT() As Long

    REC_PROCESS_COUNT_COSTOM_DEFAULT = 1000
End Property

' �R�~�b�g�m�F�@���Ȃ�
Public Property Get COMMIT_CONFIRM_NO() As Long

    COMMIT_CONFIRM_NO = 1
End Property

' �R�~�b�g�m�F�@����
Public Property Get COMMIT_CONFIRM_YES() As Long

    COMMIT_CONFIRM_YES = 2
End Property

' SQL�G���[�@���f
Public Property Get SQL_ERROR_HANDLING_SUSPEND() As Long

    SQL_ERROR_HANDLING_SUSPEND = 1
End Property

' SQL�G���[�@�p��
Public Property Get SQL_ERROR_HANDLING_RESUME() As Long

    SQL_ERROR_HANDLING_RESUME = 2
End Property

' �󔒃Z���ǂݎ������i�폜�j
Public Property Get EMPTY_CELL_READING_DEL() As Long

    EMPTY_CELL_READING_DEL = 1
End Property

' �󔒃Z���ǂݎ������i�폜���Ȃ��j
Public Property Get EMPTY_CELL_READING_NON_DEL() As Long

    EMPTY_CELL_READING_NON_DEL = 2
End Property

' ���ړ��͕����i�����j
Public Property Get DIRECT_INPUT_CHAR_DISABLE() As Long

    DIRECT_INPUT_CHAR_DISABLE = 1
End Property

' ���ړ��͕����i���[�U�w��j
Public Property Get DIRECT_INPUT_CHAR_ENABLE_CUSTOM() As Long

    DIRECT_INPUT_CHAR_ENABLE_CUSTOM = 3
End Property

' ���ړ��͕����i�f�t�H���g�j
Public Property Get DIRECT_INPUT_CHAR_ENABLE_DEFAULT() As String

    DIRECT_INPUT_CHAR_ENABLE_DEFAULT = "$"
End Property

' �X�L�[�}�@�P��
Public Property Get SCHEMA_USE_ONE() As Long

    SCHEMA_USE_ONE = 1
End Property

' �X�L�[�}�@����
Public Property Get SCHEMA_USE_MULTIPLE() As Long

    SCHEMA_USE_MULTIPLE = 2
End Property

' �Z������ �t�H���g��
Public Property Get CELL_FONT_NAME_DEFAULT() As String

    CELL_FONT_NAME_DEFAULT = Application.StandardFont
End Property

' �Z������ �t�H���g�T�C�Y
Public Property Get CELL_FONT_SIZE_DEFAULT() As Long
    
    CELL_FONT_SIZE_DEFAULT = Application.StandardFontSize
End Property

' �Z������ �܂�Ԃ�
Public Property Get CELL_WORDWRAP_DEFAULT() As Boolean
    
    CELL_WORDWRAP_DEFAULT = True
End Property

' �Z������ ��
Public Property Get CELL_WIDTH_DEFAULT() As Double
    
    CELL_WIDTH_DEFAULT = 16
End Property

' �Z������ ����
Public Property Get CELL_HEIGHT_DEFAULT() As Double
    
    CELL_HEIGHT_DEFAULT = 15
End Property


' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    recProcessCount = REC_PROCESS_COUNT_ALL
    recProcessCountCustom = REC_PROCESS_COUNT_COSTOM_DEFAULT
    
    commitConfirm = COMMIT_CONFIRM_NO
    
    sqlErrorHandling = SQL_ERROR_HANDLING_SUSPEND
    
    emptyCellReading = EMPTY_CELL_READING_DEL
    
    directInputChar = DIRECT_INPUT_CHAR_DISABLE
    directInputCharCustom = DIRECT_INPUT_CHAR_ENABLE_DEFAULT
    
    queryResultShowWhenNormal = False
    
    schemaUse = SCHEMA_USE_ONE
    
    tableColumnEscapeOracle = True
    tableColumnEscapeMysql = True
    tableColumnEscapePostgresql = True
    tableColumnEscapeSqlserver = True
    tableColumnEscapeAccess = True
    tableColumnEscapeSymfoware = True
    
    cellFontName = CELL_FONT_NAME_DEFAULT
    cellFontSize = CELL_FONT_SIZE_DEFAULT
    cellWordwrap = CELL_WORDWRAP_DEFAULT
    cellWidth = CELL_WIDTH_DEFAULT
    cellHeight = CELL_HEIGHT_DEFAULT
    
    lineHeightAutoAdjust = True
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' �����R�[�h�����P�ʂ��擾����B
'
' �T�v�@�@�@�F�����P�ʂ��擾����B�S�Ă̏ꍇ�́A0��Ԃ��B
' �����@�@�@�F
' �߂�l�@�@�F���R�[�h�����P��
'
' =========================================================
Public Function getRecProcessCount() As Long

    ' ���R�[�h�����P��
    Select Case recProcessCount
    
        ' ���R�[�h�����P�ʁi�S�āj
        Case REC_PROCESS_COUNT_ALL
            getRecProcessCount = 0
                    
        ' ���R�[�h�����P�ʁi���[�U�w��j
        Case REC_PROCESS_COUNT_COSTOM
            getRecProcessCount = recProcessCountCustom
        
    End Select

End Function

' =========================================================
' �����ړ��͕������擾����B
'
' �T�v�@�@�@�F���ړ��͕������擾����B�����̏ꍇ�́A""��Ԃ��B
' �����@�@�@�F
' �߂�l�@�@�F���ړ��͕���
'
' =========================================================
Public Function getDirectInputChar() As String

    ' ���ړ��͕���
    Select Case directInputChar
    
        ' �����̏ꍇ
        Case DIRECT_INPUT_CHAR_DISABLE
            getDirectInputChar = ""
                    
        ' �L���i���[�U�w��j�̏ꍇ
        Case DIRECT_INPUT_CHAR_ENABLE_CUSTOM
            getDirectInputChar = directInputCharCustom
        
    End Select

End Function

' =========================================================
' ���ݒ���̐���
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSetting.ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ������ǂݍ���
' =========================================================
Public Sub readForData()

    On Error GoTo err
        
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties

    ' �ꎞ�ϐ�
    Dim varRecProcessCount          As String
    Dim varRecProcessCountCustom    As String
    Dim varCommitConfirm            As String
    Dim varSqlErrorHandling         As String
    Dim varEmptyCellReading         As String
    Dim varDirectInputChar          As String
    Dim varDirectInputCharCustom    As String
    Dim varSchemaUse                As String
    Dim varQueryResultShowWhenNormal As String
    
    Dim varCellFontName             As String
    Dim varCellFontSize             As String
    Dim varCellWordwrap             As String
    Dim varCellWidth                As String
    Dim varCellHeight               As String
    Dim varLineHeightAutoAdjust     As String
    
    Dim varTableColumnOracle        As String
    Dim varTableColumnMysql         As String
    Dim varTableColumnPostgresql    As String
    Dim varTableColumnSqlserver     As String
    Dim varTableColumnAccess        As String
    Dim varTableColumnSymfoware     As String
    
    ' ����ǂݍ���
    ' �ǂݍ��񂾏��̃`�F�b�N���s���s���ł���ꍇ�A�f�t�H���g�l�ɕϊ�����
    
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    Dim val As Variant
    
    val = values.getItem(KEY_REC_PROCESS_COUNT, vbVariant): If IsArray(val) Then varRecProcessCount = val(2)
    val = values.getItem(KEY_REC_PROCESS_COUNT, vbVariant): If IsArray(val) Then varRecProcessCount = val(2)
    val = values.getItem(KEY_REC_PROCESS_COUNT, vbVariant): If IsArray(val) Then varRecProcessCount = val(2)
    val = values.getItem(KEY_REC_PROCESS_COUNT_CUSTOM, vbVariant): If IsArray(val) Then varRecProcessCountCustom = val(2)
    val = values.getItem(KEY_COMMIT_CONFIRM, vbVariant): If IsArray(val) Then varCommitConfirm = val(2)
    val = values.getItem(KEY_SQL_ERROR_HANDLING, vbVariant): If IsArray(val) Then varSqlErrorHandling = val(2)
    val = values.getItem(KEY_EMPTY_CELL_READING, vbVariant): If IsArray(val) Then varEmptyCellReading = val(2)
    val = values.getItem(KEY_DIRECT_INPUT_CHAR, vbVariant): If IsArray(val) Then varDirectInputChar = val(2)
    val = values.getItem(KEY_DIRECT_INPUT_CHAR_CUSTOM, vbVariant): If IsArray(val) Then varDirectInputCharCustom = val(2)
    val = values.getItem(KEY_QUERY_RESULT_SHOW_WHEN_NORMAL, vbVariant): If IsArray(val) Then varQueryResultShowWhenNormal = val(2)
    val = values.getItem(KEY_SCHEMA_USE, vbVariant): If IsArray(val) Then varSchemaUse = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_ORACLE, vbVariant): If IsArray(val) Then varTableColumnOracle = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_MYSQL, vbVariant): If IsArray(val) Then varTableColumnMysql = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL, vbVariant): If IsArray(val) Then varTableColumnPostgresql = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_SQLSERVER, vbVariant): If IsArray(val) Then varTableColumnSqlserver = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_ACCESS, vbVariant): If IsArray(val) Then varTableColumnAccess = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE, vbVariant): If IsArray(val) Then varTableColumnSymfoware = val(2)
    val = values.getItem(KEY_CELL_FONT_NAME, vbVariant): If IsArray(val) Then varCellFontName = val(2)
    val = values.getItem(KEY_CELL_FONT_SIZE, vbVariant): If IsArray(val) Then varCellFontSize = val(2)
    val = values.getItem(KEY_CELL_WORDWRAP, vbVariant): If IsArray(val) Then varCellWordwrap = val(2)
    val = values.getItem(KEY_CELL_WIDTH, vbVariant): If IsArray(val) Then varCellWidth = val(2)
    val = values.getItem(KEY_CELL_HEIGHT, vbVariant): If IsArray(val) Then varCellHeight = val(2)
    val = values.getItem(KEY_LINE_HEIGHT_AUTO_ADJUST, vbVariant): If IsArray(val) Then varLineHeightAutoAdjust = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_ORACLE, vbVariant): If IsArray(val) Then varTableColumnOracle = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_MYSQL, vbVariant): If IsArray(val) Then varTableColumnMysql = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL, vbVariant): If IsArray(val) Then varTableColumnPostgresql = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_SQLSERVER, vbVariant): If IsArray(val) Then varTableColumnSqlserver = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_ACCESS, vbVariant): If IsArray(val) Then varTableColumnAccess = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE, vbVariant): If IsArray(val) Then varTableColumnSymfoware = val(2)
    
    If validCode(varRecProcessCount _
                , REC_PROCESS_COUNT_ALL _
                , REC_PROCESS_COUNT_COSTOM) = False Then
    
        varRecProcessCount = REC_PROCESS_COUNT_ALL
    End If
    
    If validUnsignedInteger(varRecProcessCountCustom) = False Then
    
        varRecProcessCountCustom = REC_PROCESS_COUNT_COSTOM_DEFAULT
    End If
    
    If validCode(varCommitConfirm _
                , COMMIT_CONFIRM_NO _
                , COMMIT_CONFIRM_YES) = False Then
    
        varCommitConfirm = COMMIT_CONFIRM_NO
    End If
    
    If validCode(varSqlErrorHandling _
                , SQL_ERROR_HANDLING_SUSPEND _
                , SQL_ERROR_HANDLING_RESUME) = False Then
    
        varSqlErrorHandling = SQL_ERROR_HANDLING_SUSPEND
    End If
    
    If validCode(varEmptyCellReading _
                , EMPTY_CELL_READING_DEL _
                , EMPTY_CELL_READING_NON_DEL) = False Then
    
        varEmptyCellReading = EMPTY_CELL_READING_DEL
    End If
    
    If validCode(varDirectInputChar _
                , DIRECT_INPUT_CHAR_DISABLE _
                , DIRECT_INPUT_CHAR_ENABLE_CUSTOM) = False Then
    
        varDirectInputChar = DIRECT_INPUT_CHAR_DISABLE
    End If
    
    If Len(varDirectInputCharCustom) <> 1 Then
    
        varDirectInputCharCustom = DIRECT_INPUT_CHAR_ENABLE_DEFAULT
    End If
 
    If validCode(varQueryResultShowWhenNormal _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varQueryResultShowWhenNormal = False
    End If
    
    If validCode(varSchemaUse _
                , SCHEMA_USE_ONE _
                , SCHEMA_USE_MULTIPLE) = False Then
    
        varSchemaUse = SCHEMA_USE_ONE
    End If
  
    If Len(varCellFontName) = 0 Then
    
        varCellFontName = CELL_FONT_NAME_DEFAULT
    End If
 
    If validUnsignedNumeric(varCellFontSize) = False Then
    
        varCellFontSize = CELL_FONT_SIZE_DEFAULT
    End If
    
    If validCode(varCellWordwrap _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varCellWordwrap = CELL_WORDWRAP_DEFAULT
    End If
    
    If validUnsignedNumeric(varCellWidth) = False Then
    
        varCellWidth = CELL_WIDTH_DEFAULT
    End If
 
    If validUnsignedNumeric(varCellHeight) = False Then
    
        varCellHeight = CELL_HEIGHT_DEFAULT
    End If
 
    If validCode(varLineHeightAutoAdjust _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varLineHeightAutoAdjust = True
    End If

    If validCode(varTableColumnOracle _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnOracle = True
    End If
    
    If validCode(varTableColumnMysql _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnMysql = True
    End If
    
    If validCode(varTableColumnPostgresql _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnPostgresql = True
    End If
    
    If validCode(varTableColumnSqlserver _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnSqlserver = True
    End If
    
    If validCode(varTableColumnAccess _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnAccess = True
    End If
    
    If validCode(varTableColumnSymfoware _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnSymfoware = True
    End If
 
    recProcessCount = varRecProcessCount
    recProcessCountCustom = CLng(varRecProcessCountCustom)
    commitConfirm = varCommitConfirm
    sqlErrorHandling = varSqlErrorHandling
    emptyCellReading = varEmptyCellReading
    directInputChar = varDirectInputChar
    directInputCharCustom = varDirectInputCharCustom
    queryResultShowWhenNormal = varQueryResultShowWhenNormal
    schemaUse = varSchemaUse
    cellFontName = varCellFontName
    cellFontSize = CDbl(varCellFontSize)
    cellWordwrap = CBool(varCellWordwrap)
    cellWidth = CDbl(varCellWidth)
    cellHeight = CDbl(varCellHeight)
    lineHeightAutoAdjust = varLineHeightAutoAdjust
    tableColumnEscapeOracle = varTableColumnOracle
    tableColumnEscapeMysql = varTableColumnMysql
    tableColumnEscapePostgresql = varTableColumnPostgresql
    tableColumnEscapeSqlserver = varTableColumnSqlserver
    tableColumnEscapeAccess = varTableColumnAccess
    tableColumnEscapeSymfoware = varTableColumnSymfoware
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ��������������
' =========================================================
Public Sub writeForData()

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' DB�ڑ������i�[����z��
    Dim values As New ValCollection
    
    values.setItem Array(KEY_REC_PROCESS_COUNT, recProcessCount)
    values.setItem Array(KEY_REC_PROCESS_COUNT_CUSTOM, recProcessCountCustom)
    values.setItem Array(KEY_COMMIT_CONFIRM, commitConfirm)
    values.setItem Array(KEY_SQL_ERROR_HANDLING, sqlErrorHandling)
    values.setItem Array(KEY_EMPTY_CELL_READING, emptyCellReading)
    values.setItem Array(KEY_DIRECT_INPUT_CHAR, directInputChar)
    values.setItem Array(KEY_DIRECT_INPUT_CHAR_CUSTOM, directInputCharCustom)
    values.setItem Array(KEY_QUERY_RESULT_SHOW_WHEN_NORMAL, queryResultShowWhenNormal)
    values.setItem Array(KEY_SCHEMA_USE, schemaUse)
    values.setItem Array(KEY_CELL_FONT_NAME, cellFontName)
    values.setItem Array(KEY_CELL_FONT_SIZE, cellFontSize)
    values.setItem Array(KEY_CELL_WORDWRAP, cellWordwrap)
    values.setItem Array(KEY_CELL_WIDTH, cellWidth)
    values.setItem Array(KEY_CELL_HEIGHT, cellHeight)
    values.setItem Array(KEY_LINE_HEIGHT_AUTO_ADJUST, lineHeightAutoAdjust)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_ORACLE, tableColumnEscapeOracle)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_MYSQL, tableColumnEscapeMysql)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL, tableColumnEscapePostgresql)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_SQLSERVER, tableColumnEscapeSqlserver)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_ACCESS, tableColumnEscapeAccess)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE, tableColumnEscapeSymfoware)
                
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DB�ɑΉ�����J�����̃G�X�P�[�v�l���擾
' =========================================================
Public Function getTableColumnEscapeByDbConn(ByRef dbConn As Object) As Boolean

    ' DBMS��ނ��擾����
    Dim dmt As DbmsType
    dmt = ADOUtil.getDBMSType(dbConn)
    
    ' �߂�l��ݒ肷��
    getTableColumnEscapeByDbConn = getTableColumnEscape(dmt)
    
End Function

' =========================================================
' ��DB�ɑΉ�����J�����̃G�X�P�[�v�l���擾
' =========================================================
Public Function getTableColumnEscape(ByRef dbmsT As DbmsType) As Boolean

    ' �߂�l��ݒ肷��
    If dbmsT = DbmsType.Oracle Then
        getTableColumnEscape = tableColumnEscapeOracle
    
    ElseIf dbmsT = DbmsType.MySQL Then
        getTableColumnEscape = tableColumnEscapeMysql
    
    ElseIf dbmsT = DbmsType.PostgreSQL Then
        getTableColumnEscape = tableColumnEscapePostgresql
    
    ElseIf dbmsT = DbmsType.MicrosoftSqlServer Then
        getTableColumnEscape = tableColumnEscapeSqlserver
    
    ElseIf dbmsT = DbmsType.MicrosoftAccess Then
        getTableColumnEscape = tableColumnEscapeAccess
    
    ElseIf dbmsT = DbmsType.Symfoware Then
        getTableColumnEscape = tableColumnEscapeSymfoware
    
    Else
        getTableColumnEscape = True
    End If
    
End Function
