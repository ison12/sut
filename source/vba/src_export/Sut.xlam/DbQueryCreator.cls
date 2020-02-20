VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbQueryCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�̃N�G���𐶐�����C���^�[�t�F�[�X�B
' �SDB�ŋ��ʂ���ėp�I�Ȏ����C���^�[�t�F�[�X�B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/03/28�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' SQL����`����Ă���V�[�g��
Private Const SHEET_NAME_SQL_DEFINE                         As String = "template_sql"
' SELECT�N�G���̃e���v���[�g
Private Const CELL_SQL_DEFINE_SELECT                        As String = "A2"
' INSERT�N�G���̃e���v���[�g
Private Const CELL_SQL_DEFINE_INSERT                        As String = "A4"
' UPDATE�N�G���̃e���v���[�g
Private Const CELL_SQL_DEFINE_UPDATE                        As String = "A6"
' DELETE�N�G���̃e���v���[�g
Private Const CELL_SQL_DEFINE_DELETE                        As String = "A8"
' DELETE ALL�N�G���̃e���v���[�g
Private Const CELL_SQL_DEFINE_DELETE_ALL                    As String = "A9"

' DB�R�l�N�V����
Private m_dbConn As Object
' DBMS���
Private m_dbms As DbmsType
' �󔒕����ǂݎ�����
Private m_emptyValueReading As Long
' ���ړ��͕���
Private m_directInputChar As String
' �J����������񃊃X�g
Private m_columnFormatList As ValCollection
' �X�L�[�}�g�p�L��
Private m_schemaUse As Long
' �e�[�u���E�J�������̃G�X�P�[�v
Private m_tableColumnEscape As Boolean


' �A�v���P�[�V�����ݒ�
' �萔�𗘗p���邽�߂ɃI�u�W�F�N�g�𐶐�
Private applicationSetting As ValApplicationSetting

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set applicationSetting = New ValApplicationSetting
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set applicationSetting = Nothing
    Set m_columnFormatList = Nothing
    Set m_dbConn = Nothing
End Sub

' =========================================================
' ��DB�R�l�N�V�����ݒ�
'
' �T�v�@�@�@�FDB�R�l�N�V������ݒ肷��B
' �����@�@�@�FvNewValue DB�R�l�N�V����
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let dbConn(ByVal RHS As Variant)

    ' DB�R�l�N�V������ݒ�
    Set m_dbConn = RHS
    
    ' DBMS��ނ�ݒ�
    m_dbms = ADOUtil.getDBMSType(m_dbConn)
    
End Property

' =========================================================
' ���󔒕����ǂݎ������ݒ�
'
' �T�v�@�@�@�F�󔒕����ǂݎ�������ݒ肷��B
' �����@�@�@�FvNewValue �󔒕����ǂݎ�����
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let emptyValueReading(ByVal RHS As Long)

    m_emptyValueReading = RHS
End Property

' =========================================================
' �����ړ��͕����ݒ�
'
' �T�v�@�@�@�F���ړ��͕�����ݒ肷��B
' �����@�@�@�FvNewValue ���ړ��͕���
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let directInputChar(ByVal RHS As String)

    m_directInputChar = RHS
End Property

' =========================================================
' ���J�����������̐ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FRHS �J�����������I�u�W�F�N�g
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let columnFormatList(ByVal RHS As Variant)

    Set m_columnFormatList = RHS
End Property

' =========================================================
' ���X�L�[�}�g�p�̐ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvNewValue �X�L�[�}�g�p�L��
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let schemaUse(ByVal vNewValue As Variant)

    m_schemaUse = vNewValue
End Property

' =========================================================
' ���e�[�u���E�J�������G�X�P�[�v�̐ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FRHS �J�����������I�u�W�F�N�g
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let tableColumnEscape(ByVal vNewValue As Variant)

    m_tableColumnEscape = vNewValue
End Property

' =========================================================
' ��INSERT������
'
' �T�v�@�@�@�FINSERT SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@record ���R�[�h
'
' �߂�l�@�@�FINSERT��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createInsert(ByRef table As ValDbDefineTable _
                           , ByRef record As Variant) As String

    ' �ڑ����� AND
    Const CONNECT_AND              As String = "AND "
    ' �ڑ����� �J���}
    Const CONNECT_COMMA            As String = "   ,"
    ' �ڑ����� �ŏ�
    Const CONNECT_FIRST            As String = "    "
        
    ' SQL�{��
    Dim statement         As New StringBuilder
    ' SQL�J������
    Dim columnStatement   As New StringBuilder
    ' SQL�o�����[��
    Dim valuesStatement   As New StringBuilder

    ' �N�G���[�l�i���ۂɃN�G���[�ɖ��ߍ��ޒl�j
    Dim queryValue        As String

    ' �J�������X�g
    Dim columnList      As ValCollection
    ' �J����
    Dim column          As ValDbDefineColumn
    
    ' �J�����������
    Dim columnFormat    As ValDbColumnTypeColInfo

    ' ���R�[�h�l
    Dim recordValue     As String
    
    ' ���s�R�[�h�i�Z���̉��s�R�[�h��LF�j
    Dim newline As String: newline = vbLf

    ' -------------------------------------------------
    ' �e���v���[�gSQL�X�e�[�g�����g���擾
    ' -------------------------------------------------
    statement.append SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_INSERT)


    ' -------------------------------------------------
    ' SQL�𐶐�����
    ' -------------------------------------------------
    statement.replace "${tableName}", DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)

    
    ' -------------------------------------------------
    ' ���R�[�h����N�G���[�𐶐�����
    ' -------------------------------------------------
    Dim i         As Long
    Dim j         As Long: j = 1
    Dim arrayMin  As Long   ' �z��̓Y�����i�ŏ��l�j
    Dim arrayMax  As Long   ' �z��̓Y�����i�ő�l�j
    Dim arraySize As Long   ' �z��̃T�C�Y
    
    Set columnList = table.columnList
    
    arrayMin = LBound(record)
    arrayMax = UBound(record)
    
    arraySize = VBUtil.arraySize(record)
    
    For i = arrayMin To arrayMax
    
        Set column = columnList.getItemByIndex(j)
    
        ' �J�������������擾����
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        ' �J����������񂪎擾�ł��Ȃ������ꍇ
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " ��`���Ȃ�"
            #End If
            
            ' �f�t�H���g�J�������������擾����
            Set columnFormat = defaultColumnFormat
            
        ' �J������񂪎擾�ł����ꍇ
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatUpdate
            #End If
        
        End If
    
        If columnFormat.formatUpdate = COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print "Excluded " & column.dataType & " from a query."
            #End If
            
            ' �X�L�b�v����
            GoTo continue
        End If
    
        ' ���R�[�h�l���擾
        recordValue = CStr(record(i))
        
        ' �󔒕����ǂݎ�������"�J�����폜"�̏ꍇ
        If m_emptyValueReading = applicationSetting.EMPTY_CELL_READING_DEL Then
        
            ' �l���ݒ肳��Ă���ꍇ�̂݃X�e�[�g�����g�ɒǉ�����
            If recordValue <> "" Then
    
                queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, recordValue, True, m_directInputChar)
                ' COLUMN��
                columnStatement.append(newline).append(CONNECT_COMMA).append (DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
                ' VALUES��
                valuesStatement.append(newline).append(CONNECT_COMMA).append (queryValue)
    
            End If
            
        ' �󔒕����ǂݎ�������"�J�������폜���Ȃ�"�̏ꍇ
        Else
        
            ' ��ɃX�e�[�g�����g��ǉ�����
            
            ' �󕶎��̏ꍇ�ANULL�ɕϊ�
            queryValue = DBUtil.convertEmptyToNull(recordValue)
            queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, queryValue, True, m_directInputChar)
    
            ' COLUMN��
            columnStatement.append(newline).append(CONNECT_COMMA).append (DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
            ' VALUES��
            valuesStatement.append(newline).append(CONNECT_COMMA).append (queryValue)
    
        End If

continue:
        j = j + 1
        
    Next
    
    ' COLUMN��Ɉ�ȏ�J�������ݒ肳��Ă���ꍇ
    If columnStatement.str <> "" Then

        ' �O���ɕt�����ꂽ�璷��","��"\r\n"����菜��
        columnStatement.remove 1, Len(newline & CONNECT_COMMA)
        columnStatement.insert 1, CONNECT_FIRST
                                
        ' �O���ɕt�����ꂽ�璷��","��"\r\n"����菜��
        valuesStatement.remove 1, Len(newline & CONNECT_COMMA)
        valuesStatement.insert 1, CONNECT_FIRST

        ' �J�����w�蕔���𖄂ߍ���
        statement.replace "${columns}", columnStatement.str
        ' �l�w�蕔���𖄂ߍ���
        statement.replace "${values}", valuesStatement.str

        createInsert = statement.str

    End If

End Function

' =========================================================
' ��UPDATE������
'
' �T�v�@�@�@�FUPDATE SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@record ���R�[�h
'
' �߂�l�@�@�FUPDATE��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createUpdate(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' �ڑ����� AND
    Const CONNECT_AND              As String = "AND "
    ' �ڑ����� �J���}
    Const CONNECT_COMMA            As String = "   ,"
    ' �ڑ����� �ŏ�
    Const CONNECT_FIRST            As String = "    "

    ' SQL�{��
    Dim statement             As New StringBuilder
    ' SQL�J�����l��
    Dim columnValuesStatement As New StringBuilder

    ' SQL������
    Dim conditionStatement    As New StringBuilder

    ' �N�G���[�l�i�o�����[���N�G���[�p�̒l�ɕϊ��j
    Dim queryValue        As String

    ' �J�������X�g
    Dim columnList      As ValCollection
    ' �J����
    Dim column          As ValDbDefineColumn
    ' �J�����������
    Dim columnFormat    As ValDbColumnTypeColInfo
    ' �e�[�u��������(PK)
    Dim tableConstPk    As New ValDbDefineTableConstraints
    ' PK�J�����ł��邩������킷�t���O
    Dim isColumnPk      As Boolean
    
    Dim tableConstTmp   As ValDbDefineTableConstraints
    ' �e�[�u�����񃊃X�g����PK������擾����
    For Each tableConstTmp In table.constraintsList.col
    
        If tableConstTmp.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk Then
        
            Set tableConstPk = tableConstTmp
            Exit For
        End If
    Next
    
    ' ���R�[�h�l
    Dim recordValue     As String

    ' ���s�R�[�h�i�Z���̉��s�R�[�h��LF�j
    Dim newline As String: newline = vbLf

    ' -------------------------------------------------
    ' �e���v���[�gSQL�X�e�[�g�����g���擾
    ' -------------------------------------------------
    statement.append SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_UPDATE)

    ' -------------------------------------------------
    ' SQL�𐶐�����
    ' -------------------------------------------------
    statement.replace "${tableName}", DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)
    
    ' -------------------------------------------------
    ' ���R�[�h����N�G���[�𐶐�����
    ' -------------------------------------------------
    Dim i         As Long
    Dim j         As Long: j = 1
    Dim arrayMin  As Long   ' �z��̓Y�����i�ŏ��l�j
    Dim arrayMax  As Long   ' �z��̓Y�����i�ő�l�j
    Dim arraySize As Long   ' �z��̃T�C�Y
    
    Set columnList = table.columnList
    
    arrayMin = LBound(record)
    arrayMax = UBound(record)
    
    arraySize = VBUtil.arraySize(record)
    
    For i = arrayMin To arrayMax
    
        Set column = columnList.getItemByIndex(j)
        
        ' �J�������������擾����
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        ' �J����������񂪎擾�ł��Ȃ������ꍇ
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print "Format " & column.dataType & " is not found."
            #End If
            
            ' �f�t�H���g�J�������������擾����
            Set columnFormat = defaultColumnFormat
            
        ' �J������񂪎擾�ł����ꍇ
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatUpdate
            #End If
        
        End If
    
        If columnFormat.formatUpdate = COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print "Excluded " & column.dataType & " from a query."
            #End If
            
            ' �X�L�b�v����
            GoTo continue
        End If
        
        ' ���R�[�h�l���擾
        recordValue = CStr(record(i))
        
        ' PK����ł��邩�ǂ����𔻒肷��
        If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
        
            isColumnPk = False
        Else
        
            isColumnPk = True
        End If
        
        ' PK�ł͂Ȃ��ꍇ
        If isColumnPk = False Then
        
            ' �󔒕����ǂݎ�������"�J�����폜"�̏ꍇ
            If m_emptyValueReading = applicationSetting.EMPTY_CELL_READING_DEL Then
            
                ' �l���ݒ肳��Ă���ꍇ�̂݃X�e�[�g�����g�ɒǉ�����
                If recordValue <> "" Then
        
                    queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, recordValue, True, m_directInputChar)
    
                    ' COLUMN��
                    columnValuesStatement.append(newline).append(CONNECT_COMMA).append(DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)).append(" = ").append (queryValue)
        
                End If
                
            ' �󔒕����ǂݎ�������"�J�������폜���Ȃ�"�̏ꍇ
            Else
            
                ' ��ɃX�e�[�g�����g��ǉ�����
                
                ' �󕶎��̏ꍇ�ANULL�ɕϊ�
                queryValue = DBUtil.convertEmptyToNull(recordValue)
                queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, queryValue, True, m_directInputChar)
        
                ' COLUMN��
                columnValuesStatement.append(newline).append(CONNECT_COMMA).append(DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)).append(" = ").append (queryValue)
        
            End If
            
        End If
        
        ' �����L�[���l������PK�̏ꍇ�́A���͒l�����낤���Ȃ��낤����ɃX�e�[�g�����g�Ƃ��Ēǉ�����B
        ' �Ⴆ�΁Acolumn1�Ecolumn2��PK�̏ꍇ�ɁAcolumn1�ɓ��͂�����column2�ɓ��͂��Ȃ��ꍇ��
        ' column1�̂ݏ����Ƃ��Ēǉ����Ă��܂��ƁA���R�[�h����ӂɓ���ł����ɕ������X�V����Ă��܂����ꂪ���邽�߁B
        If isColumnPk = True Then

            queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, recordValue, True, m_directInputChar)
            
            conditionStatement.append(newline).append(CONNECT_AND).append(DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)).append(" = ").append (queryValue)
        End If

continue:
        j = j + 1
    Next
    
    ' COLUMN��܂���WHERE��Ɉ�ȏ�J�������ݒ肳��Ă���ꍇ
    If conditionStatement.str <> "" And columnValuesStatement.str <> "" Then

        ' �O���ɕt�����ꂽ�璷��","�Ɖ��s�R�[�h����菜��
        columnValuesStatement.remove 1, Len(newline & CONNECT_COMMA)
        columnValuesStatement.insert 1, CONNECT_FIRST
                                
        ' �O���ɕt�����ꂽ�璷��"AND "�Ɖ��s�R�[�h����菜��
        conditionStatement.remove 1, Len(newline & CONNECT_AND)
        conditionStatement.insert 1, CONNECT_FIRST
                                
        ' �J�����w�蕔���𖄂ߍ���
        statement.replace "${columnValues}", columnValuesStatement.str
        ' �l�w�蕔���𖄂ߍ���
        statement.replace "${condition}", conditionStatement.str

        createUpdate = statement.str

    End If

End Function

' =========================================================
' ��DELETE������
'
' �T�v�@�@�@�FDELETE SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@record ���R�[�h
'
' �߂�l�@�@�FDELETE��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createDelete(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' �ڑ����� AND
    Const CONNECT_AND              As String = "AND "
    ' �ڑ����� �J���}
    Const CONNECT_COMMA            As String = "   ,"
    ' �ڑ����� �ŏ�
    Const CONNECT_FIRST            As String = "    "

    ' SQL�{��
    Dim statement            As New StringBuilder
    ' SQL������
    Dim conditionStatement   As New StringBuilder

    ' �N�G���[�l�i���ۂɃN�G���[�ɖ��ߍ��ޒl�j
    Dim queryValue        As String

    ' �J�������X�g
    Dim columnList      As ValCollection
    ' �J����
    Dim column          As ValDbDefineColumn
    ' �J�����������
    Dim columnFormat    As ValDbColumnTypeColInfo
    
    ' �e�[�u��������(PK)
    Dim tableConstPk    As New ValDbDefineTableConstraints
    ' PK�J�����ł��邩������킷�t���O
    Dim isColumnPk      As Boolean
    
    Dim tableConstTmp   As ValDbDefineTableConstraints
    ' �e�[�u�����񃊃X�g����PK������擾����
    For Each tableConstTmp In table.constraintsList.col
    
        If tableConstTmp.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk Then
        
            Set tableConstPk = tableConstTmp
            Exit For
        End If
    Next
    
    ' ���R�[�h�l
    Dim recordValue     As String

    ' ���s�R�[�h�i�Z���̉��s�R�[�h��LF�j
    Dim newline As String: newline = vbLf

    ' -------------------------------------------------
    ' �e���v���[�gSQL�X�e�[�g�����g���擾
    ' -------------------------------------------------
    statement.append SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_DELETE)

    ' -------------------------------------------------
    ' SQL�𐶐�����
    ' -------------------------------------------------
    statement.replace "${tableName}", DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)

    ' -------------------------------------------------
    ' ���R�[�h����N�G���[�𐶐�����
    ' -------------------------------------------------
    Dim i         As Long
    Dim j         As Long: j = 1
    Dim arrayMin  As Long   ' �z��̓Y�����i�ŏ��l�j
    Dim arrayMax  As Long   ' �z��̓Y�����i�ő�l�j
    Dim arraySize As Long   ' �z��̃T�C�Y
    
    Set columnList = table.columnList
    
    arrayMin = LBound(record)
    arrayMax = UBound(record)
    
    arraySize = VBUtil.arraySize(record)
    
    For i = arrayMin To arrayMax
    
        Set column = columnList.getItemByIndex(j)
    
        ' �J�������������擾����
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        ' �J����������񂪎擾�ł��Ȃ������ꍇ
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " ��`���Ȃ�"
            #End If
            
            ' �f�t�H���g�J�������������擾����
            Set columnFormat = defaultColumnFormat
        
        ' �J������񂪎擾�ł����ꍇ
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatUpdate
            #End If
        
        End If
        
        If columnFormat.formatUpdate = COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print "Excluded " & column.dataType & " from a query."
            #End If
            
            ' �X�L�b�v����
            GoTo continue
        End If
        
        ' ���R�[�h�l���擾
        recordValue = CStr(record(i))
        
        ' PK����ł��邩�ǂ����𔻒肷��
        If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
        
            isColumnPk = False
        Else
        
            isColumnPk = True
        End If
        

        ' �����L�[���l������PK�̏ꍇ�́A���͒l�����낤���Ȃ��낤����ɃX�e�[�g�����g�Ƃ��Ēǉ�����B
        ' �Ⴆ�΁Acolumn1�Ecolumn2��PK�̏ꍇ�ɁAcolumn1�ɓ��͂�����column2�ɓ��͂��Ȃ��ꍇ��
        ' column1�̂ݏ����Ƃ��Ēǉ����Ă��܂��ƁA���R�[�h����ӂɓ���ł����ɕ������X�V����Ă��܂����ꂪ���邽�߁B
        If isColumnPk = True Then

            queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, recordValue, True, m_directInputChar)
            
            conditionStatement.append(newline).append(CONNECT_AND).append(DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)).append(" = ").append (queryValue)
        End If
        
continue:
        j = j + 1
        
    Next
    
    ' WHERE��Ɉ�ȏ�J�������ݒ肳��Ă���ꍇ
    If conditionStatement.str <> "" Then

        ' �O���ɕt�����ꂽ�璷��"AND "��"\r\n"����菜��
        conditionStatement.remove 1, Len(newline & CONNECT_AND)
        conditionStatement.insert 1, CONNECT_FIRST
                                
        ' COLUMN = VALUE��
        statement.replace "${condition}", conditionStatement.str

        ' �S�ẴX�e�[�g�����g���������āA�߂�l�Ƃ��Ċi�[����
        createDelete = statement.str

    End If

End Function

' =========================================================
' ��DELETE�������i�e�[�u����̑S���R�[�h���폜�j
'
' �T�v�@�@�@�FDELETE SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
'
' �߂�l�@�@�FDELETE��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createDeleteOfAllTable(ByRef table As ValDbDefineTable) As String

    ' SQL�{��
    Dim statement As New StringBuilder
    
    ' -------------------------------------------------
    ' �e���v���[�gSQL�X�e�[�g�����g���擾
    ' -------------------------------------------------
    statement.append SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_DELETE_ALL)

    ' -------------------------------------------------
    ' SQL�𐶐�����
    ' -------------------------------------------------
    statement.replace "${tableName}", DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)

    ' �߂�l�Ƃ��Ċi�[����
    createDeleteOfAllTable = statement.str

End Function

' =========================================================
' ��SELECT�̃J���������̃X�e�[�g�����g�𐶐�����
'
' �T�v�@�@�@�F
' �����@�@�@�Ftable   �e�[�u����`
' �@�@�@�@�@�@ident   �C���f���g
' �@�@�@�@�@�@convert �ϊ�����
'
' �߂�l�@�@�FSELECT�̃J���������̃X�e�[�g�����g
'
' ���L�����@�F
'
'
' =========================================================
Public Function createSelectForColumn(ByRef table As ValDbDefineTable _
                                    , Optional ByVal INDENT As String = "" _
                                    , Optional convert As Boolean = True _
                                    , Optional ByVal identifier As String = "") As String

    ' �߂�l
    Dim ret As StringBuilder
    Set ret = New StringBuilder

    ' �J�������X�g
    Dim columnList      As ValCollection
    ' �J����
    Dim column          As ValDbDefineColumn
    ' �J�����������
    Dim columnFormat    As ValDbColumnTypeColInfo
    ' �J������
    Dim columnName      As New StringBuilder
    
    ' �J�������X�g���擾
    Set columnList = table.columnList
    
    ' �J�����̐��������[�v����
    For Each column In columnList.col

        ' �J�������������擾����
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        
        ' �J����������񂪎擾�ł��Ȃ������ꍇ
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " ��`���Ȃ�"
            #End If
            
            ' �f�t�H���g�J�������������擾����
            Set columnFormat = defaultColumnFormat
            
        ' �J������񂪎擾�ł����ꍇ
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatSelect
            #End If
        
        End If
        
        columnName.clear
        
        If columnFormat.formatSelect = COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE Then
        
            columnName.append("'Excluded ").append(column.dataType).append(".' AS ").append (DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
            
            #If DEBUG_MODE = 1 Then
                Debug.Print "Excluded " & column.dataType & " from a query."
            #End If
            
        ' �J�������̕ϊ��������{�L���̔���
        ElseIf convert = True Then
        
            ' �J�������̕ϊ��������s
            columnName.append convertSelectFormat(m_dbms, columnFormat.formatSelect, identifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
            
        Else
        
            ' �ϊ����Ȃ��ꍇ�A�J�����������̂܂ܐݒ�
            columnName.append identifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)
        End If
        
        ' COLUMN��
        ret.append(vbLf).append(INDENT).append("   ,").append (columnName.str)
    Next

    ' �O���ɕt�����ꂽ�璷��","��"\n"����菜��
    ret.remove 1, Len(vbLf & INDENT & "   ,")
    ret.insert 1, INDENT & "    "
                            
    ' �߂�l��ݒ肷��
    createSelectForColumn = ret.str
    
End Function

' =========================================================
' ��SELECT���̂̃X�e�[�g�����g�𐶐�����
'
' �T�v�@�@�@�F
' �����@�@�@�Ftable            ���́@�e�[�u����`
' �@�@�@�@�@�@condition        ���́@SELECT����
' �@�@�@�@�@�@statementWhere   �o�́@WHERE�X�e�[�g�����g
' �@�@�@�@�@�@statementOrderby �o�́@ORDER BY�X�e�[�g�����g
'
' ���L�����@�F
'
'
' =========================================================
Public Sub createSelectForWhereOrderby(ByVal table As ValDbDefineTable _
                                     , ByVal condition As ValSelectCondition _
                                     , ByRef statementWhere As String _
                                     , ByRef statementOrderby As String _
                                     , Optional ByVal identifier As String = "")

    ' �ڑ����� �ŏ�
    Const CONNECT_FIRST            As String = "     "
    ' �ڑ����� AND
    Const CONNECT_AND              As String = " AND "
    ' �ڑ����� �J���}
    Const CONNECT_COMMA            As String = "    ,"
    ' �ڑ����� ������i���j
    Const CONNECT_LEFT_RIGHT_EQUAL As String = " = "
    ' �ڑ����� ������iLIKE�j
    Const CONNECT_LEFT_RIGHT_LIKE  As String = " LIKE "
    ' �ڑ����� ������iIS�j
    Const CONNECT_LEFT_RIGHT_IS     As String = " IS "

    ' WHERE��̓��e
    Dim where       As New StringBuilder
    ' ORDER BY��̓��e
    Dim order       As New StringBuilder
    
    ' �J�������X�g
    Dim columnList As ValCollection
    ' �J������
    Dim columnName As Variant
    ' �J�����I�u�W�F�N�g
    Dim column     As ValDbDefineColumn
    ' �J�����������
    Dim columnFormat    As ValDbColumnTypeColInfo
    
    ' �l
    Dim value      As String
    ' ����
    Dim orderByVal As Variant
    
    Dim connectLeftRight As String
    
    
    ' �J�������X�g���擾����
    Set columnList = condition.getColumnList
    
    Dim i As Long
    
    ' �R���g���[���z���1������������
    For Each columnName In columnList.col
    
        ' �J�����I�u�W�F�N�g���擾����
        Set column = table.columnList.getItem(columnName)
        
        ' �J�����I�u�W�F�N�g���擾�ł��Ȃ������ꍇ�A���̃��[�v��
        If column Is Nothing Then
        
            GoTo continue
        End If
    
        ' �J�������������擾����
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        ' �J����������񂪎擾�ł��Ȃ������ꍇ
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " ��`���Ȃ�"
            #End If
            
            ' �f�t�H���g�J�������������擾����
            Set columnFormat = defaultColumnFormat
            
        ' �J������񂪎擾�ł����ꍇ
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatSelect
            #End If
        
        End If
        
        ' �l���擾����
        value = condition.getConditionValue(columnName)
        ' ���я����擾����
        orderByVal = condition.getConditionOrder(columnName)
    
        ' �l���ݒ肳��Ă���ꍇ
        If value <> "" Then
        
            ' NULL�܂���
            ' NOT NULL�Ɠ��͂��ꂽ�ꍇ�́AIS NULL�ň����|����悤�ɂ���
            If _
                Trim$(UCase$(value)) = "NULL" _
             Or Trim$(UCase$(value)) = "NOT NULL" Then
            
                where.append(vbLf).append(CONNECT_AND).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append(CONNECT_LEFT_RIGHT_IS).append (Trim$(UCase$(value)))
                
            ' ��L�ȊO�̏ꍇ
            Else
            
                ' �J�����̃��e������ނ�String�̏ꍇ�Ŋ��ALIKE�֐��̃p�^�[���������܂܂�Ă���ꍇ
                ' LIKE�֐��𗘗p����Ɣ��f����B
                If ( _
                     InStr(value, "%") <> 0 Or _
                    (InStr(value, "_") <> 0 And Not (InStr(value, ConstantsTable.QUERY_PARAMETER_ENCLOSE_START) = 1 And InStr(value, ConstantsTable.QUERY_PARAMETER_ENCLOSE_END) = Len(value)))) Then
                
                    ' �V���O���N�H�[�g���G�X�P�[�v
                    value = DBUtil.escapeValueForSinglequart(m_dbms, value)
                    ' ������̗��[�ɃV���O���N�H�[�g��t������
                    value = DBUtil.escapeValueForEscapeChar(m_dbms, value)
                
                    where.append(vbLf).append(CONNECT_AND).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append(CONNECT_LEFT_RIGHT_LIKE).append (DBUtil.encloseSingleQuart(value))
                    
                    If _
                        m_dbms = DbmsType.MySQL _
                     Or m_dbms = DbmsType.PostgreSQL Then
                     
                        ' MySQL��PostgreSQL�́A"\"�����ɃG�X�P�[�v�����Ȃ̂ŁA�ʂ̕������G�X�P�[�v�����ɂ��Ă�����
                        where.append (DBUtil.addLikeEscape(m_dbms, "#"))
                        
                    Else
                    
                        where.append (DBUtil.addLikeEscape(m_dbms))
                    End If
                Else
                
                    ' �l��ϊ�����B�ϊ�����ꍇ�̓G�X�P�[�v�����͂��̂܂܂ɂ���
                    ' �iDBUtil.convertUpdateFormat�̑�T������False�ɐݒ�B��\��\\ �ɕϊ������肵�Ȃ��B
                    ' �@�܂���U�����̒��ړ��͕����͗L���ɂ��Ȃ��B�j
                    ' �l��ϊ�����B�^���L�̃t�H�[�}�b�g�ɕϊ�����
                    value = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, value, False)
                    
                    where.append(vbLf).append(CONNECT_AND).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append(CONNECT_LEFT_RIGHT_EQUAL).append (value)
                
                End If
            
            End If
        
        End If
        
        ' ���я��ɏ������~�����ݒ肳��Ă���ꍇ
        If orderByVal <> condition.ORDER_NONE Then
        
            ' ORDER BY��̈ꕔ�𐶐�����
            If orderByVal = condition.ORDER_ASC Then
            
                ' ASC
                order.append(vbLf).append(CONNECT_COMMA).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append (" ASC")
            ElseIf orderByVal = condition.ORDER_DESC Then
                
                ' DESC
                order.append(vbLf).append(CONNECT_COMMA).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append (" DESC")
            End If
        
        End If
        
continue:

    Next
    
                  
    ' WHERE�傪�P�����ł���������Ă�����"WHERE"��t��
    If where.str <> "" Then
    
        ' �O���ɕt������Ă���璷�ȉ��s�R�[�h�ƘA������������O����
        where.remove 1, Len(vbLf & CONNECT_AND)
        where.insert 1, CONNECT_FIRST
        where.insert 1, "WHERE" & vbLf
    End If
                  
    ' ORDER BY�傪�P�����ł���������Ă�����"ORDER BY"��t��
    If order.str <> "" Then
    
        ' �O���ɕt������Ă���璷�ȉ��s�R�[�h�ƘA������������O����
        order.remove 1, Len(vbLf & CONNECT_COMMA)
        order.insert 1, CONNECT_FIRST
        order.insert 1, "ORDER BY" & vbLf
    End If
                  
    ' �߂�l��ݒ肷��
    statementWhere = where.str
    
    ' �߂�l��ݒ肷��
    statementOrderby = order.str
    
End Sub

' =========================================================
' ��SELECT���̂̃X�e�[�g�����g�𐶐�����
'
' �T�v�@�@�@�F
' �����@�@�@�Ftable            ���́@�e�[�u����`
' �@�@�@�@�@�@condition        ���́@SELECT����
' �߂�l�@�@�FORDER BY��
'
' ���L�����@�F
'
'
' =========================================================
Public Function createSelectForOrderby(ByVal table As ValDbDefineTable _
                                     , ByVal condition As ValSelectCondition) As String

    ' �ڑ����� �J���}
    Const CONNECT_COMMA            As String = ", "

    ' ORDER BY��̓��e
    Dim order       As New StringBuilder
    
    ' �J�������X�g
    Dim columnList As ValCollection
    ' �J������
    Dim columnName As Variant
    ' �J�����I�u�W�F�N�g
    Dim column     As ValDbDefineColumn
    
    ' ����
    Dim orderByVal As Variant
    
    ' �J�������X�g���擾����
    Set columnList = condition.getColumnList
    
    ' �R���g���[���z���1������������
    For Each columnName In columnList.col
    
        ' �J�����I�u�W�F�N�g���擾����
        Set column = table.columnList.getItem(columnName)
        
        ' �J�����I�u�W�F�N�g���擾�ł��Ȃ������ꍇ�A���̃��[�v��
        If column Is Nothing Then
        
            GoTo continue
        End If

        ' ���я����擾����
        orderByVal = condition.getConditionOrder(columnName)
    
        ' ���я��ɏ������~�����ݒ肳��Ă���ꍇ
        If orderByVal <> condition.ORDER_NONE Then
        
            ' ORDER BY��̈ꕔ�𐶐�����
            If orderByVal = condition.ORDER_ASC Then
            
                ' ASC
                order.append(CONNECT_COMMA).append(DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append (" ASC")
            ElseIf orderByVal = condition.ORDER_DESC Then
                
                ' DESC
                order.append(CONNECT_COMMA).append(DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append (" DESC")
            End If
        
        End If
        
continue:

    Next
    
    ' ORDER BY�傪�P�����ł���������Ă�����"ORDER BY"��t��
    If order.str <> "" Then
    
        ' �O���ɕt������Ă���璷�ȉ��s�R�[�h�ƘA������������O����
        order.remove 1, Len(CONNECT_COMMA)
        order.insert 1, "ORDER BY "
    End If
                  
    ' �߂�l��ݒ肷��
    createSelectForOrderby = order.str
    
End Function


' =========================================================
' ��PK�J�����̌����X�e�[�g�����g�𐶐�����
'
' �T�v�@�@�@�FPK�J�����̌����X�e�[�g�����g�𐶐�����
' �����@�@�@�Ftable           �e�[�u����`
' �@�@�@�@�@�@leftIdentifier  �������̏C���q
'             rightIdentifier �E�����̏C���q
'             rightStatementIsNull �E�����l����� IS NULL �Ƃ���
' �߂�l�@�@�FPK�J�����̌����X�e�[�g�����g
'
' ���L�����@�F
'
'
' =========================================================
Public Function createSelectForPkJoin(ByRef table As ValDbDefineTable _
                                    , ByVal leftIdentifier As String _
                                    , ByVal rightIdentifier As String _
                                    , Optional ByVal rightStatementIsNull = False) As String

    ' �ڑ����� AND
    Const CONNECT_AND              As String = "AND "
    ' �ڑ����� �ŏ�
    Const CONNECT_FIRST            As String = ""

    ' SQL������
    Dim conditionStatement    As New StringBuilder

    ' �J�������X�g
    Dim columnList      As ValCollection
    ' �J����
    Dim column          As ValDbDefineColumn
    ' �e�[�u��������(PK)
    Dim tableConstPk    As New ValDbDefineTableConstraints
    ' PK�J�����ł��邩������킷�t���O
    Dim isColumnPk      As Boolean
    
    Dim tableConstTmp   As ValDbDefineTableConstraints
    ' �e�[�u�����񃊃X�g����PK������擾����
    For Each tableConstTmp In table.constraintsList.col
    
        If tableConstTmp.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk Then
        
            Set tableConstPk = tableConstTmp
            Exit For
        End If
    Next
    
    ' ���R�[�h�l
    Dim recordValue     As String

    ' ���s�R�[�h�i�Z���̉��s�R�[�h��LF�j
    Dim newline As String: newline = vbLf

    ' -------------------------------------------------
    ' ���R�[�h����N�G���[�𐶐�����
    ' -------------------------------------------------
    Dim i         As Long
    Dim j         As Long: j = 1
    Dim arrayMin  As Long   ' �z��̓Y�����i�ŏ��l�j
    Dim arrayMax  As Long   ' �z��̓Y�����i�ő�l�j
    Dim arraySize As Long   ' �z��̃T�C�Y
    
    Set columnList = table.columnList
    
    arrayMin = 1
    arrayMax = columnList.count
    
    For i = arrayMin To arrayMax
    
        Set column = columnList.getItemByIndex(j)
        
        ' PK����ł��邩�ǂ����𔻒肷��
        If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
        
            isColumnPk = False
        Else
        
            isColumnPk = True
        End If
        
        ' �����L�[���l������PK�̏ꍇ�́A���͒l�����낤���Ȃ��낤����ɃX�e�[�g�����g�Ƃ��Ēǉ�����B
        ' �Ⴆ�΁Acolumn1�Ecolumn2��PK�̏ꍇ�ɁAcolumn1�ɓ��͂�����column2�ɓ��͂��Ȃ��ꍇ��
        ' column1�̂ݏ����Ƃ��Ēǉ����Ă��܂��ƁA���R�[�h����ӂɓ���ł����ɕ������X�V����Ă��܂����ꂪ���邽�߁B
        If isColumnPk = True Then

            
            If rightStatementIsNull Then
                conditionStatement.append (CONNECT_AND)
                conditionStatement.append (rightIdentifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
                conditionStatement.append (" IS NULL")
            Else
                conditionStatement.append (CONNECT_AND)
                conditionStatement.append (leftIdentifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
                conditionStatement.append (" = ")
                conditionStatement.append (rightIdentifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
            End If
            
            conditionStatement.append (" ")
            
        End If

continue:
        j = j + 1
    Next
    
    ' COLUMN��܂���WHERE��Ɉ�ȏ�J�������ݒ肳��Ă���ꍇ
    If conditionStatement.str <> "" Then

        ' �O���ɕt�����ꂽ�璷��"AND "�Ɖ��s�R�[�h����菜��
        conditionStatement.remove 1, Len(CONNECT_AND)
        conditionStatement.insert 1, CONNECT_FIRST
                                
    End If
    
    createSelectForPkJoin = conditionStatement.str

End Function

Private Function defaultColumnFormat() As ValDbColumnTypeColInfo

    Set defaultColumnFormat = New ValDbColumnTypeColInfo
    defaultColumnFormat.formatSelect = COLUMN_FORMAT_REPLACE_CHAR_COLUMN
    defaultColumnFormat.formatUpdate = COLUMN_FORMAT_REPLACE_CHAR_DIRECT

End Function