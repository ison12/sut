VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbQueryCreatorMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDbQueryCreator

' *********************************************************
' DB�̃N�G���𐶐�����C���^�[�t�F�[�X�B
' MySQL DB�����̎����B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/23�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************
' SQL����`����Ă���V�[�g��
Private Const SHEET_NAME_SQL_DEFINE                         As String = "template_sql"
' SELECT�N�G���̃e���v���[�g
Private Const CELL_SQL_DEFINE_SELECT                        As String = "A2"

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

' DB�N�G���������ʃC���^�[�t�F�[�X
Private queryCreatorCommon As DbQueryCreator

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    ' DB�N�G���������ʃC���^�[�t�F�[�X������������
    Set queryCreatorCommon = New DbQueryCreator
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' DB�N�G���������ʃC���^�[�t�F�[�X��j������
    Set queryCreatorCommon = Nothing
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
Private Property Let IDbQueryCreator_dbConn(ByVal RHS As Variant)

    ' DB�R�l�N�V������ݒ�
    Set m_dbConn = RHS
    
    ' DBMS��ނ�ݒ�
    m_dbms = ADOUtil.getDBMSType(m_dbConn)
    
    ' ���ʃC���^�[�t�F�[�X��DB�R�l�N�V������ݒ�
    queryCreatorCommon.dbConn = m_dbConn

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
Private Property Let IDbQueryCreator_emptyValueReading(ByVal RHS As Long)

    m_emptyValueReading = RHS
    queryCreatorCommon.emptyValueReading = RHS
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
Private Property Let IDbQueryCreator_directInputChar(ByVal RHS As String)

    m_directInputChar = RHS
    queryCreatorCommon.directInputChar = RHS

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
Private Property Let IDbQueryCreator_columnFormatList(ByVal RHS As Variant)

    Set m_columnFormatList = RHS
    queryCreatorCommon.columnFormatList = RHS
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
Public Property Let IDbQueryCreator_schemaUse(ByVal vNewValue As Variant)

    m_schemaUse = vNewValue
    queryCreatorCommon.schemaUse = vNewValue
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
Public Property Let IDbQueryCreator_tableColumnEscape(ByVal vNewValue As Variant)

    m_tableColumnEscape = vNewValue
    queryCreatorCommon.tableColumnEscape = vNewValue
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
Private Function IDbQueryCreator_createInsert(ByRef table As ValDbDefineTable _
                                            , ByRef record As Variant) As String

    ' ���ʃC���^�[�t�F�[�X���Ăяo���ăN�G���[�𐶐����Ă��炤
    IDbQueryCreator_createInsert = queryCreatorCommon.createInsert(table, record)
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
Private Function IDbQueryCreator_createUpdate(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' ���ʃC���^�[�t�F�[�X���Ăяo���ăN�G���[�𐶐����Ă��炤
    IDbQueryCreator_createUpdate = queryCreatorCommon.createUpdate(table, record)
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
Private Function IDbQueryCreator_createDelete(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' ���ʃC���^�[�t�F�[�X���Ăяo���ăN�G���[�𐶐����Ă��炤
    IDbQueryCreator_createDelete = queryCreatorCommon.createDelete(table, record)
End Function

' =========================================================
' ��DELETE�������i�e�[�u����̑S���R�[�h�𐶐��j
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
Private Function IDbQueryCreator_createDeleteOfAllTable(table As ValDbDefineTable) As String

    ' ���ʃC���^�[�t�F�[�X���Ăяo���ăN�G���[�𐶐����Ă��炤
    IDbQueryCreator_createDeleteOfAllTable = queryCreatorCommon.createDeleteOfAllTable(table)

End Function

' =========================================================
' ��SELECT������
'
' �T�v�@�@�@�FSELECT SQL�𐶐�����
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@condition ������
'
' �߂�l�@�@�FSELECT��
'
' ���L�����@�F
'
'
' =========================================================
Private Function IDbQueryCreator_createSelect(ByRef table As ValDbDefineTable _
                                            , ByRef condition As ValSelectCondition) As String

    ' �C���f���g
    Const INDENT As String = "    "
    
    ' �e���v���[�g
    Dim templateStatement As String

    ' SQL�{��
    Dim statement         As String
    ' COLUMN��
    Dim statementColumn   As String
    ' COLUMN��i�ϊ������j
    Dim statementColumnNotConverted As String
    ' FROM��
    Dim statementFrom     As String
    ' WHERE��
    Dim statementWhere    As String
    ' ORDER BY��
    Dim statementOrderby  As String
    ' OPTION
    Dim statementOption   As String
    
    Dim temp As String
    
    ' COLUMN��𐶐�����
    statementColumn = queryCreatorCommon.createSelectForColumn(table)
    ' FROM��𐶐�����
    statementFrom = INDENT & DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)

    ' WHERE���ORDER BY��𐶐�����
    queryCreatorCommon.createSelectForWhereOrderby table _
                                                 , condition _
                                                 , statementWhere _
                                                 , statementOrderby
    

    ' -------------------------------------------------
    ' �e���v���[�gSQL�X�e�[�g�����g���擾
    ' -------------------------------------------------
    templateStatement = SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_SELECT)
    ' -------------------------------------------------

    ' -------------------------------------------------
    ' SQL�𐶐�����
    ' -------------------------------------------------
    statement = templateStatement
    
    ' ���R�[�h�͈͎̔擾���w�肵�Ă���ꍇ
    If condition.recRangeStart <> condition.REC_RANGE_NONE Or condition.recRangeEnd <> condition.REC_RANGE_NONE Then
    
        ' LIMIT OFFSET��𐶐�����
        
        ' LIMIT����
        Dim limitOffsetL As String
        ' OFFSET����
        Dim limitOffsetO As String
        
        ' LIMIT�����𐶐�
        If condition.recRangeEnd = condition.REC_RANGE_NONE Then
    
            ' �w�肪�Ȃ��̂ŁA�傫��������ݒ�
            limitOffsetL = "18446744073709551615"
            
        Else
        
            ' �w�肳��Ă���̂Ő��l�����̂܂ܐݒ�
            limitOffsetL = condition.recRangeEnd
        End If
    
        ' OFFSET�����𐶐�
        If condition.recRangeStart = condition.REC_RANGE_NONE Then
    
            ' �w�肪�Ȃ��̂ŋ󕶎����ݒ�
            limitOffsetO = ""
            
        Else
        
            ' �w�肳��Ă���̂Ő��l�����̂܂ܐݒ�i�����l��1�ł͂Ȃ�0�Ȃ̂Ń}�C�i�X1���Ă��j
            limitOffsetO = condition.recRangeStart - 1
            
            ' �J�n�E�I�����Ɏw�肳��Ă���̂ŁALIMIT�������Čv�Z����
            If condition.recRangeEnd <> condition.REC_RANGE_NONE Then
                
                limitOffsetL = condition.recRangeEnd - condition.recRangeStart + 1
            End If
        End If
    
        ' OFFSET�̎w��L��
        If limitOffsetO <> "" Then
        
            statementOption = statementOption & "LIMIT " & limitOffsetO & "," & limitOffsetL
            
        ' OFFSET�̎w�薳��
        Else
        
            statementOption = statementOption & "LIMIT " & limitOffsetL
        
        End If
    
    End If

    
    ' COLUMN��̒u��
    statement = replace(statement _
                         , "${statementColumn}" _
                         , statementColumn)
                         
    ' FROM��̒u��
    statement = replace(statement _
                            , "${statementFrom}" _
                            , statementFrom)

    ' WHERE��̒u��
    statement = replace(statement _
                            , convertReplaceKeyword("${statementWhere}", statementWhere) _
                            , statementWhere)

    ' ORDER BY��̒u��
    statement = replace(statement _
                            , convertReplaceKeyword("${statementOrderby}", statementOrderby) _
                            , statementOrderby)

    ' OPTION�̒u��
    statement = replace(statement _
                            , convertReplaceKeyword("${statementOption}", statementOption) _
                            , statementOption)

    ' �߂�l��ݒ�
    IDbQueryCreator_createSelect = statement
    
End Function

Private Function convertReplaceKeyword(ByVal keyword As String _
                                     , ByVal replacedStr As String) As String

    ' �u����̕����񂪋�̏ꍇ
    If replacedStr = "" Then
    
        ' �L�[���[�h�ɉ��s�R�[�h��A������
        convertReplaceKeyword = vbLf & keyword
        
    Else
    
        ' �L�[���[�h�����̂܂ܕԂ�
        convertReplaceKeyword = keyword
    End If
    
End Function