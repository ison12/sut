Attribute VB_Name = "DBUtil"
Option Explicit

' *********************************************************
' DB�Ɋ֘A�������[�e�B���e�B���W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/03/21�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' =========================================================
' ��DBMS���
'
' �T�v�@�@�@�FDBMS�i�f�[�^�x�[�X�}�l�[�W�����g�V�X�e���j�̎�ނ�����킷�񋓌^
'
' =========================================================
Public Enum DbmsType

    MySQL = 0
    PostgreSQL = 1
    Oracle = 2
    MicrosoftSqlServer = 4
    MicrosoftAccess = 5
    Symfoware = 10
    Other = 3

End Enum

' =========================================================
' ���N�G���[�̃��e�������
'
' �T�v�@�@�@�F�N�G���[�̃��e�����̎�ނ�����킷�񋓌^
'
' =========================================================
Public Enum QueryLiteralType
    
    Number = 0
    String_ = 1
    Date = 2
    Lob = 3
    Any_ = 4
    
End Enum

' �J�����������@�u�������@value
Public Const COLUMN_FORMAT_REPLACE_CHAR_VALUE       As String = "$value"
' �J�����������@�u�������@column
Public Const COLUMN_FORMAT_REPLACE_CHAR_COLUMN      As String = "$column"
' �J�����������@�u�������@exclude
Public Const COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE     As String = "$exclude"
' �J�����������@�u�������@direct
Public Const COLUMN_FORMAT_REPLACE_CHAR_DIRECT      As String = "$direct"

' =========================================================
' ��DBMS���擾
'
' �T�v�@�@�@�FDBMS�����擾����B
' �����@�@�@�Fdbms DBMS���
' �߂�l�@�@�FDBMS��
' ���L�����@�F
'
' =========================================================
Public Function getDbmsTypeName(ByVal dbms As DbmsType) As String

    Select Case dbms
    
        Case DbmsType.MySQL
            getDbmsTypeName = "MySQL"
    
        Case DbmsType.PostgreSQL
            getDbmsTypeName = "PostgreSQL"
        
        Case DbmsType.Oracle
            getDbmsTypeName = "Oracle"
        
        Case DbmsType.MicrosoftSqlServer
            getDbmsTypeName = "MicrosoftSqlServer"
        
        Case DbmsType.MicrosoftAccess
            getDbmsTypeName = "MicrosoftAccess"
        
        Case DbmsType.Symfoware
            getDbmsTypeName = "Symfoware"
        
        Case DbmsType.Other
            getDbmsTypeName = "Other"
        
    End Select

End Function

' =========================================================
' ���X�L�[�}���{�e�[�u�����̌���
'
' �T�v�@�@�@�F�X�L�[�}���{�e�[�u��������������
' �����@�@�@�FschemaName        �X�L�[�}��
' �@�@�@�@�@�@tableName         �e�[�u����
'             schemaUse         �X�L�[�}�g�p�L���i1�F�X�L�[�}���g�p�A2�F�X�L�[�}�g�p�j
' ���L�����@�F
'
' =========================================================
Public Function concatSchemaTable(ByVal schemaName As String _
                                , ByRef tableName As String _
                                , ByVal schemaUse As Long) As String


    If schemaUse = 1 Then
    
        concatSchemaTable = tableName
        Exit Function
    End If

    If schemaName = "" Then
        concatSchemaTable = tableName
    Else
        concatSchemaTable = schemaName & "." & tableName
    End If

End Function

' =========================================================
' ��""��NULL
'
' �T�v�@�@�@�F���� val ��""�󕶎���̏ꍇ�A"NULL"�ɕϊ�����B
' �����@�@�@�Fval �l
' �߂�l�@�@�F�ϊ����ꂽ������
'
' =========================================================
Public Function convertEmptyToNull(ByRef val As String) As String

    If val = "" Then
    
        convertEmptyToNull = "NULL"
    Else
    
        convertEmptyToNull = val
    End If

End Function

' =========================================================
' ��������̗��[���V���O���N�H�[�g�ň͂�
'
' �T�v�@�@�@�F���� val �̗��[��'(�V���O���N�H�[�g)��t�����߂�l�Ƃ��ĕԂ��B
' �����@�@�@�Fval �l
' �߂�l�@�@�F'(�V���O���N�H�[�g)���t�����ꂽ������
'
' =========================================================
Public Function encloseSingleQuart(ByRef val As String) As String

    encloseSingleQuart = "'" & val & "'"
End Function

' =========================================================
' ���G�X�P�[�v�����̃G�X�P�[�v
'
' �T�v�@�@�@�F���� val ��DB�ŗL�̃G�X�P�[�v�������܂܂�Ă���ꍇ
' �@�@�@�@�@�@�G�X�P�[�v����B
' �����@�@�@�Fdbms        DBMS���
' �@�@�@�@�@�@val         �l
' ���L�����@�FMySQL��Postgresql�̏ꍇ�A"\"���f�t�H���g��
' �@�@�@�@�@�@�@�G�X�P�[�v�����Ƃ��Ďw�肳��Ă��邪
' �@�@�@�@�@�@Oracle���́A���Ƀf�t�H���g�̃G�X�P�[�v�����͂Ȃ��B
'
' =========================================================
Public Function escapeValueForEscapeChar(ByVal dbms As DbmsType _
                                       , ByRef val As String) As String

    Select Case dbms
    
        Case DbmsType.MySQL
            escapeValueForEscapeChar = replace(val, "\", "\\")
    
        Case DbmsType.PostgreSQL
            escapeValueForEscapeChar = replace(val, "\", "\\")
        
        Case DbmsType.Oracle
            escapeValueForEscapeChar = val
        
        Case DbmsType.MicrosoftSqlServer
            escapeValueForEscapeChar = val
        
        Case DbmsType.MicrosoftAccess
            escapeValueForEscapeChar = val
        
        Case DbmsType.Symfoware
            escapeValueForEscapeChar = val
        
        Case DbmsType.Other
            escapeValueForEscapeChar = val
        
    End Select

End Function

' =========================================================
' ���V���O���N�H�[�g�̃G�X�P�[�v
'
' �T�v�@�@�@�F���� val �ɃV���O���N�H�[�g�������܂܂�Ă���ꍇ
' �@�@�@�@�@�@�G�X�P�[�v����B
' �����@�@�@�Fdbms        DBMS���
' �@�@�@�@�@�@val         �l
' ���L�����@�F
'
' =========================================================
Public Function escapeValueForSinglequart(ByVal dbms As DbmsType _
                                        , ByRef val As String) As String

    Select Case dbms
    
        Case DbmsType.MySQL
            escapeValueForSinglequart = replace(val, "'", "''")
    
        Case DbmsType.PostgreSQL
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.Oracle
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.MicrosoftSqlServer
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.MicrosoftAccess
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.Symfoware
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.Other
            escapeValueForSinglequart = replace(val, "'", "''")
        
    End Select

End Function

' =========================================================
' ���e�[�u���E�J�������̃G�X�P�[�v
'
' �T�v�@�@�@�F�e�[�u���E�J���������G�X�P�[�v����
' �����@�@�@�Fdbms        DBMS���
' �@�@�@�@�@�@val         �l
'             isEscape    �G�X�P�[�v�̎��s�L��
' ���L�����@�F
'
' =========================================================
Public Function escapeTableColumn(ByVal dbms As DbmsType _
                                , ByRef val As String _
                                , ByVal isEscape As Boolean) As String

    ' �G�X�P�[�v���Ȃ��ꍇ�́A���̂܂ܕԋp����
    If Not isEscape Then
        escapeTableColumn = val
        Exit Function
    End If

    ' �l���󕶎���̏ꍇ�A���̂܂ܕԋp����
    If val = "" Then
        escapeTableColumn = val
        Exit Function
    End If

    Select Case dbms
    
        Case DbmsType.MySQL
            escapeTableColumn = "`" & replace(val, "`", "``") & "`"
    
        Case DbmsType.PostgreSQL
            escapeTableColumn = """" & replace(val, """", """""") & """"
        
        Case DbmsType.Oracle
            escapeTableColumn = """" & replace(val, """", """""") & """"
        
        Case DbmsType.MicrosoftSqlServer
            ' �������� "]" �݂̂��G�X�P�[�v����
            escapeTableColumn = "[" & replace(val, "]", "]]") & "]"
        
        Case DbmsType.MicrosoftAccess
            ' �������� "]" �݂̂��G�X�P�[�v����
            escapeTableColumn = "[" & replace(val, "]", "]]") & "]"
        
        Case DbmsType.Symfoware
        
    End Select

End Function

' =========================================================
' ��LIKE�֐���ESCAPE�t��
'
' �T�v�@�@�@�F�N�G���[����LIKE�֐���ESCAPE��t������B
' �����@�@�@�Fdbms        DBMS���
' �@�@�@�@�@�@escapeChar  �G�X�P�[�v����
' �߂�l�@�@�F�ϊ���̃N�G���[��
' ���L�����@�F
'
' =========================================================
Public Function addLikeEscape(ByVal dbms As DbmsType _
                            , Optional ByRef escapeChar As String = "\") As String

    addLikeEscape = " ESCAPE '" & escapeChar & "'"

End Function

' =========================================================
' �������f�[�^�^���𔻒�
'
' �T�v�@�@�@�F
' �����@�@�@�Fdbms        DBMS���
' �@�@�@�@�@�@dataType    �f�[�^�^��
' �߂�l�@�@�FTrue �����f�[�^�^�AFalse ����ȊO
' ���L�����@�F
'
' =========================================================
Public Function isCharType(ByVal dbms As DbmsType, ByVal dataType As String) As Boolean

    If InStr(dataType, "CHAR") > 0 Or InStr(dataType, "TEXT") > 0 Then
        isCharType = True
    Else
        isCharType = False
    End If

End Function

' =========================================================
' ���N�G���[�l�ϊ�
'
' �T�v�@�@�@�F�N�G���[�̒l�i�J�����ɑ΂������l�j��ϊ�����B
' �@�@�@�@�@�@column = value ��value������
' �@�@�@�@�@�@������̏ꍇ�̓V���O���N�H�[�g(')�ň͂ޕK�v������
' �@�@�@�@�@�@�����������ꍇ�ɁA�ϊ������{����B
'
' �����@�@�@�Fdbms            DBMS���
' �@�@�@�@�@�@literalType     ���e�������
' �@�@�@�@�@�@value           �l
' �@�@�@�@�@�@isEscapeChar    �G�X�P�[�v�������G�X�P�[�v����t���O
' �@�@�@�@�@�@directInputChar ���ړ��͕���
' �߂�l�@�@�F�ϊ���̒l
'
' =========================================================
Public Function convertQueryLiteral(ByVal dbms As DbmsType _
                                  , ByVal literalType As QueryLiteralType _
                                  , ByVal value As String _
                                  , Optional ByVal isEscapeChar As Boolean = True _
                                  , Optional ByVal directInputChar As String = "") As String

    ' ���ړ��͕����̔���
    If directInputChar <> "" And InStr(value, directInputChar) = 1 Then
    
        ' �擪 1�����ڂ� directInputChar �ƈ�v����ꍇ�A2�����ڈȍ~���擾���߂�l�Ƃ��Đݒ�
        convertQueryLiteral = Mid$(value, 2)
    
    ' ���ꕶ��
    ElseIf isSpecialValue(dbms, value) = True Then
    
        convertQueryLiteral = value
    
    ' ������^
    ElseIf literalType = String_ Then

        ' �V���O���N�H�[�g���G�X�P�[�v����
        convertQueryLiteral = DBUtil.escapeValueForSinglequart(dbms, value)
        ' �G�X�P�[�v�������G�X�P�[�v����
        If isEscapeChar = True Then
            convertQueryLiteral = DBUtil.escapeValueForEscapeChar(dbms, convertQueryLiteral)
        End If
        ' �V���O���N�H�[�g�ň͂�
        convertQueryLiteral = DBUtil.encloseSingleQuart(convertQueryLiteral)

    ' ���Ԍ^
    ElseIf literalType = Date Then

        ' �V���O���N�H�[�g���G�X�P�[�v����
        convertQueryLiteral = DBUtil.escapeValueForSinglequart(dbms, value)
        ' �G�X�P�[�v�������G�X�P�[�v����
        If isEscapeChar = True Then
            convertQueryLiteral = DBUtil.escapeValueForEscapeChar(dbms, convertQueryLiteral)
        End If
        ' �V���O���N�H�[�g�ň͂�
        convertQueryLiteral = DBUtil.encloseSingleQuart(convertQueryLiteral)

    ' ���l�^
    ElseIf literalType = Number Then

        convertQueryLiteral = value
    
    ' ��L�ȊO
    Else

        convertQueryLiteral = value
    
    End If

End Function

' =========================================================
' ���N�G���[�l �X�V�n�ϊ�
'
' �T�v�@�@�@�F�N�G���[�̒l�i�J�����ւ̑���l�j��ϊ�����B
' �@�@�@�@�@�@column = value ��value������
' �@�@�@�@�@�@SQL�ł͕�����^�̏ꍇ�A�l�̗��[���V���O���N�H�[�g(')�ň͂ޕK�v������B
' �@�@�@�@�@�@�����������ꍇ�ɁA�{���\�b�h��p���ĕϊ������{����B
'
' �@�@�@�@�@�@�{���\�b�h�ɂ�����ϊ��̎d�g�݂�
' �@�@�@�@�@�@������� updateFormat �̒u���ϐ���u�����邱�ƂŎ�������B
' �@�@�@�@�@�@updateFormat�́ATO_DATE($value, 'xxxxx') �Ƃ������A�����ɒu���ϐ����܂񂾕�����ɂȂ��Ă���B
'
' �����@�@�@�Fdbms                  DBMS���
' �@�@�@�@�@�@updateFormat          �X�V�������
' �@�@�@�@�@�@value                 �l
' �@�@�@�@�@�@isEscapeChar          �G�X�P�[�v�������G�X�P�[�v����t���O
' �@�@�@�@�@�@directInputCharPrefix ���ړ��͕����ړ���
' �@�@�@�@�@�@directInputCharSuffix ���ړ��͕����ڔ���
' �@�@�@�@�@�@nullInputChar         NULL���͕���
' �߂�l�@�@�F�ϊ���̒l
'
' =========================================================
Public Function convertUpdateFormat(ByVal dbms As DbmsType _
                                  , ByVal updateFormat As String _
                                  , ByVal value As String _
                                  , Optional ByVal isEscapeChar As Boolean = True _
                                  , Optional ByVal directInputCharPrefix As String = "" _
                                  , Optional ByVal directInputCharSuffix As String = "" _
                                  , Optional ByVal nullInputChar As String = "") As String

    ' ���ړ��͕����̔���
    If directInputCharPrefix <> "" And _
       directInputCharSuffix <> "" And _
       InStr(value, directInputCharPrefix) = 1 And _
       InStrRev(value, directInputCharSuffix) = Len(value) Then
    
        ' �擪 1�����ڂƍŌ�̕����� directInputChar �ƈ�v����ꍇ�A�͂܂ꂽ���������o���Đݒ�
        convertUpdateFormat = Mid$(value, 2, Len(value) - 2)
    
    ElseIf directInputCharPrefix <> "" And _
           InStr(value, directInputCharPrefix) = 1 Then
    
        ' �擪 1�����ڂ� directInputChar �ƈ�v����ꍇ�A2�����ڈȍ~���擾���߂�l�Ƃ��Đݒ�
        convertUpdateFormat = Mid$(value, 2)
    
    ' NULL���͕���
    ElseIf nullInputChar <> "" And _
           UCase$(nullInputChar) = UCase$(value) Then
    
        convertUpdateFormat = "NULL"
        
    ' ���ړ��͌`��
    ElseIf updateFormat = COLUMN_FORMAT_REPLACE_CHAR_DIRECT Then
    
        convertUpdateFormat = value
    
    Else
    
        ' �G�X�P�[�v�������G�X�P�[�v����
        If isEscapeChar = True Then
            value = DBUtil.escapeValueForEscapeChar(dbms, value)
        End If
        
        ' �V���O���N�H�[�g��������Ɋ܂܂�Ă���ꍇ�A�G�X�P�[�v����
        If InStr(value, "'") <> 0 Then
        
            value = DBUtil.escapeValueForSinglequart(dbms, value)
        End If
    
        ' �������̒u���ϐ���value�l�ŕϊ�����
        convertUpdateFormat = replace(updateFormat, COLUMN_FORMAT_REPLACE_CHAR_VALUE, value)
    
    End If
    
End Function

' =========================================================
' ���J�����l �Q�ƌn�ϊ�
'
' �T�v�@�@�@�FSELECT SQL�ɂ�����A�J������̕ϊ������{����
'
' �@�@�@�@�@�@�{���\�b�h�ɂ�����ϊ��̎d�g�݂�
' �@�@�@�@�@�@������� selectFormat �̒u���ϐ���u�����邱�ƂŎ�������B
' �@�@�@�@�@�@selectFormat�́ATO_DATE($column, 'xxxxx') �Ƃ������A�����ɒu���ϐ����܂񂾕�����ɂȂ��Ă���B
'
' �����@�@�@�Fdbms            DBMS���
' �@�@�@�@�@�@selectFormat    �X�V�������
' �@�@�@�@�@�@column          �l
' �߂�l�@�@�F�ϊ���̒l
'
' =========================================================
Public Function convertSelectFormat(ByVal dbms As DbmsType _
                                  , ByVal selectFormat As String _
                                  , ByVal column As String) As String

    ' �������̒u���ϐ���column�l�ŕϊ�����
    convertSelectFormat = replace(selectFormat, COLUMN_FORMAT_REPLACE_CHAR_COLUMN, column)
    
End Function

' =========================================================
' �����ꕶ������
'
' �T�v�@�@�@�FNULL���̓���ȕ����ł��邩�𔻒肷��B
' �����@�@�@�Fdbms    DBMS���
' �@�@�@�@�@�@value   �f�[�^�l
'
' �߂�l�@�@�FTrue ���ꕶ��
'
' =========================================================
Public Function isSpecialValue(ByVal dbms As DbmsType _
                              , ByVal value As String) As Boolean

    isSpecialValue = False
    
    ' "NULL"�Ƃ���������̏ꍇ
    If UCase(value) = "NULL" Then
    
        isSpecialValue = True
    
    End If

End Function

' =========================================================
' �����X�g�̗v�f���J���}��؂�̕�����ɕϊ�����
'
' �T�v�@�@�@�F���X�g�̗v�f���J���}��؂�̕�����ɕϊ�����
'
' �����@�@�@�Fdbms DB���
' �@�@�@�@�@�@list ���X�g
'
' =========================================================
Public Function convertListToCommaStr(ByVal dbms As DbmsType _
                                    , ByRef list As collection) As String

    ' ��؂蕶��
    Const DELIM_STR As String = ", "

    ' �߂�l
    Dim ret As String
    
    ' ���X�g�̗v�f
    Dim value     As Variant
    Dim valueConv As String
    
    For Each value In list
    
        valueConv = CStr(value)
        
        ' �V���O���N�H�[�g���G�X�P�[�v����
        valueConv = DBUtil.escapeValueForSinglequart(dbms, valueConv)
        ' �G�X�P�[�v�������G�X�P�[�v����
        valueConv = DBUtil.escapeValueForEscapeChar(dbms, valueConv)
        ' �V���O���N�H�[�g�ň͂�
        valueConv = DBUtil.encloseSingleQuart(valueConv)

        ret = ret & DELIM_STR & valueConv
    
    Next
    
    ' �O���ɕt�����ꂽ�璷�ȕ��������������
    ret = replace(ret, DELIM_STR, "", , 1)
    
    convertListToCommaStr = ret

End Function

' =========================================================
' ���X�L�[�}�E�e�[�u�������o
'
' �T�v�@�@�@�F�C�ӂ̕����񂩂�X�L�[�}�E�e�[�u�����𒊏o����B
'
' �@�@�@�@�@�@������̓h�b�g(.)�ŋ�؂��邱�Ƃ�O��Ƃ��Ă���
' �@�@�@�@�@�@�h�b�g�̍����E�E�������ꂼ��X�L�[�}�E�e�[�u�����Ƃ��Ē��o�����B
' �@�@�@�@�@�@�h�b�g�����݂��Ȃ��ꍇ�A�e�[�u�����̂ݒ��o�����B
'
' �����@�@�@�Fval    ���́@�X�L�[�}�E�e�[�u����
' �@�@�@�@�@�@schema �o�́@�X�L�[�}��
' �@�@�@�@�@�@table  �o�́@�e�[�u����
' �߂�l�@�@�F
'
' =========================================================
Public Sub extractSchemaTable(ByVal val As String _
                            , ByRef schema As String _
                            , ByRef table As String)

    ' �h�b�g(.)�ŋ�؂�ꂽ��������擾����
    '
    ' ��F�X�L�[�}�ƃe�[�u�����h�b�g�ŘA�����ꂽ������𕪉�����B[schema].[table]
    ' �@schema = [schema]
    ' �@table  = [table]
    '
    Dim splitStr() As String
    
    splitStr = Split(val, ".")
    
    ' �������ꂽ�z��2�ȏ�̏ꍇ
    If VBUtil.arraySize(splitStr) >= 2 Then
    
        ' �h�b�g�̍���
        schema = splitStr(LBound(splitStr))
        ' �h�b�g�̉E��
        table = splitStr(UBound(splitStr))
        
    Else
    
        schema = ""
        table = val
        
    End If


End Sub

' =========================================================
' �����R�[�h��������擾����B
'
' �T�v�@�@�@�F
'
' �����@�@�@�Fval    �l
' �߂�l�@�@�F�l���擾����
'
' =========================================================
Public Function GetRecordValue(ByRef val As Variant) As Variant

    ' �`�����N�T�C�Y
    Const CHUNK_SIZE As Long = 1024

    Dim actualSize As Long ' ���ۂ̃T�C�Y
    Dim offset     As Long ' �I�t�Z�b�g

    If val.Attributes And &H80 Then
        ' GetChunk�ŏ����擾���ׂ��ꍇ�iadFldLong�̏ꍇ�j
        
        ' ���v�T�C�Y���擾����
        actualSize = val.actualSize
        
        ' GetChunk�ŏ���S�Ď擾����
        Do While offset < actualSize
            GetRecordValue = GetRecordValue & val.GetChunk(CHUNK_SIZE)
            offset = offset + CHUNK_SIZE
        Loop
        
    Else
        ' �ʏ�̏ꍇ
        GetRecordValue = val
    
    End If

End Function
