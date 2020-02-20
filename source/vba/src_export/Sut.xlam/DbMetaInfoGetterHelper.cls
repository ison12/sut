VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbMetaInfoGetterHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB���^���擾�w���p�[�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/15�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' DB�R�l�N�V����
Private m_dbConn As Object
' DBMS���
Private m_dbms   As DbmsType
' �N�G�������I�u�W�F�N�g
Private m_queryCreator As IDbMetaInfoQueryCreator

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
    ' DB��ނ�ݒ�
    m_dbms = ADOUtil.getDBMSType(m_dbConn)
    
    ' �N�G���[�����I�u�W�F�N�g�𐶐�����Factory�I�u�W�F�N�g
    Dim dbObjFactory As New DbObjectFactory
    
    Set m_queryCreator = dbObjFactory.createMetaInfoQueryCreatorObject(m_dbConn)
    
End Property

' =========================================================
' ���X�L�[�}��񃊃X�g�擾
'
' �T�v�@�@�@�F�ڑ����Ă���DB�̑S�X�L�[�}�����擾����
' �����@�@�@�F
'
' �߂�l�@�@�F�X�L�[�}���X�g�B
'
' ���L�����@�F
'
'
' =========================================================
Public Function getSchemaList() As ValCollection

    On Error GoTo err
    
    ' �߂�l�I�u�W�F�N�g
    Dim ret As New ValCollection
    Dim retObj As ValDbDefineSchema
    
    ' ���sSQL�̐���
    Dim sql As String
    sql = m_queryCreator.createQuerySelectSchemaList

    ' ���R�[�h�Z�b�g
    Dim recordSet   As Object
    ' �N�G���[�����s����
    Set recordSet = ADOUtil.querySelect(m_dbConn, sql, 0)

    ' ���R�[�h�Z�b�g��������擾����
    Do While Not recordSet.EOF
    
        Set retObj = New ValDbDefineSchema
        retObj.schemaName = VBUtil.convertNullToEmptyStr(recordSet!SCHEMA_NAME)
        retObj.schemaComment = VBUtil.convertNullToEmptyStr(recordSet!SCHEMA_COMMENT)
        
        ret.setItem retObj, retObj.schemaName
    
        recordSet.MoveNext
    Loop
    
    Set getSchemaList = ret

    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
        
End Function

' =========================================================
' ���e�[�u�����擾�i�{�J�������j
'
' �T�v�@�@�@�F�C�ӂ̃e�[�u������S�Ď擾����B
' �����@�@�@�FschemaList �X�L�[�}���X�g
'
' �߂�l�@�@�F�e�[�u����`���B
'
' ���L�����@�F
'
'
' =========================================================
Public Function getTableList(ByVal schemaList As ValCollection) As ValCollection

    On Error GoTo err
    
    ' �߂�l�I�u�W�F�N�g
    Dim ret As New ValCollection
    Dim retObj As ValDbDefineTable
    
    ' ���sSQL�̐���
    Dim sql As String
    sql = m_queryCreator.createQuerySelectTableList(schemaList)

    ' ���R�[�h�Z�b�g
    Dim recordSet   As Object
    ' �N�G���[�����s����
    Set recordSet = ADOUtil.querySelect(m_dbConn, sql, 0)

    ' ���R�[�h�Z�b�g��������擾����
    Do While Not recordSet.EOF
    
        Set retObj = New ValDbDefineTable
        retObj.schemaName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA)
        retObj.tableName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME)
        retObj.tableComment = VBUtil.convertNullToEmptyStr(recordSet!TABLE_COMMENT)
        
        ret.setItem retObj, retObj.schemaTableName
    
        recordSet.MoveNext
    Loop
    
    Set getTableList = ret

    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
     
End Function

' =========================================================
' ���e�[�u����`�擾�i�{�J�������j
'
' �T�v�@�@�@�F�C�ӂ̃X�L�[�}�E�e�[�u�������L�[�ɂ��ăe�[�u�������擾����B
' �����@�@�@�FschemaName �X�L�[�}��
' �@�@�@�@�@�@tableName  �e�[�u����
'
' �߂�l�@�@�F�e�[�u����`�B
'
' ���L�����@�F
'
'
' =========================================================
Public Function getTable(ByVal schemaName As String, ByVal tableName As String) As ValDbDefineTable

    On Error GoTo err
    
    ' �߂�l�I�u�W�F�N�g
    Dim ret As New ValCollection
    Dim retTableObj  As ValDbDefineTable
    Dim retColumnObj As ValDbDefineColumn
    Dim retConstObj  As ValDbDefineTableConstraints
    Dim retConstColumnObj As ValDbDefineColumn
    Dim retConstFkColumnObj As ValDbDefineColumn
    
    Dim lastSchemaTableName As String

    Dim sqlColumn As String
    sqlColumn = m_queryCreator.createQuerySelectColumn(schemaName, tableName)
    
    Dim sqlConst As String
    sqlConst = m_queryCreator.createQuerySelectTableConst(schemaName, tableName)

    ' ���R�[�h�Z�b�g
    Dim recordSet   As Object
    
    ' ---------------------------------------------------------
    ' �J�������
    ' ---------------------------------------------------------
    ' �N�G���[�����s����
    Set recordSet = ADOUtil.querySelect(m_dbConn, sqlColumn, 0)

    ' ���R�[�h�Z�b�g��������擾����
    Do While Not recordSet.EOF
    
        If lastSchemaTableName <> VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA) & "_" & _
                                  VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME) Then
    
            Set retTableObj = New ValDbDefineTable
            retTableObj.schemaName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA)
            retTableObj.tableName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME)
            retTableObj.tableComment = VBUtil.convertNullToEmptyStr(recordSet!TABLE_COMMENT)
        
            ret.setItem retTableObj, VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA) & "_" & _
                                     VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME)
            
        End If
            
        Set retColumnObj = New ValDbDefineColumn
        retColumnObj.schemaName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA)
        retColumnObj.tableName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME)
        retColumnObj.columnName = VBUtil.convertNullToEmptyStr(recordSet!COLUMN_NAME)
        retColumnObj.dataType = VBUtil.convertNullToEmptyStr(recordSet!COLUMN_TYPE)
        retColumnObj.dataTypeFormal = VBUtil.convertNullToEmptyStr(recordSet!COLUMN_TYPE_FORMAL)
        If recordSet!IS_NULL = "Y" Then
            retColumnObj.isNull = True
        Else
            retColumnObj.isNull = False
        End If
        retColumnObj.default = VBUtil.convertNullToEmptyStr(recordSet!DEFAULT_VALUE)
        retColumnObj.charLength = VBUtil.convertNullToEmptyStr(recordSet!CHAR_LENGTH)
        retColumnObj.dataPrecision = VBUtil.convertNullToEmptyStr(recordSet!DATA_PRECISION)
        retColumnObj.dataScale = VBUtil.convertNullToEmptyStr(recordSet!DATA_SCALE)
        retColumnObj.dateTimePrecision = VBUtil.convertNullToEmptyStr(recordSet!DATETIME_PRECISION)
        retColumnObj.columnComment = VBUtil.convertNullToEmptyStr(recordSet!COLUMN_COMMENT)
        
        retTableObj.addColumn retColumnObj

        lastSchemaTableName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA) & "_" & _
                              VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME)
    
        recordSet.MoveNext
    Loop
    
    ' ---------------------------------------------------------
    ' ������
    ' ---------------------------------------------------------
    ' �N�G���[�����s����
    Set recordSet = ADOUtil.querySelect(m_dbConn, sqlConst, 0)

    ' ���R�[�h�Z�b�g��������擾����
    Do While Not recordSet.EOF
    
        Set retTableObj = ret.getItem( _
            VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA) & "_" & _
            VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME))
    
        If Not retTableObj Is Nothing Then
        
            Set retConstObj = retTableObj.constraintsList.getItem(convertNullToEmptyStr(recordSet!CONSTRAINT_NAME))
            If retConstObj Is Nothing Then
                Set retConstObj = New ValDbDefineTableConstraints
                retConstObj.constraintName = VBUtil.convertNullToEmptyStr(recordSet!CONSTRAINT_NAME)
                
                retTableObj.addConstraint retConstObj
            End If
      
            If VBUtil.convertNullToEmptyStr(recordSet!CONSTRAINT_TYPE) = "P" Then
                retConstObj.constraintType = tableConstPk
            ElseIf VBUtil.convertNullToEmptyStr(recordSet!CONSTRAINT_TYPE) = "U" Then
                retConstObj.constraintType = tableConstUk
            ElseIf VBUtil.convertNullToEmptyStr(recordSet!CONSTRAINT_TYPE) = "F" Then
                retConstObj.constraintType = tableConstFk
            Else
                retConstObj.constraintType = tableConstUnknown
            End If
            
            retConstObj.schemaName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA)
            retConstObj.tableName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME)
            
            If VBUtil.convertNullToEmptyStr(recordSet!COLUMN_NAME) <> Empty Then
            
                Set retConstColumnObj = New ValDbDefineColumn
                retConstColumnObj.schemaName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA)
                retConstColumnObj.tableName = VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME)
                retConstColumnObj.columnName = VBUtil.convertNullToEmptyStr(recordSet!COLUMN_NAME)
                retConstObj.addColumn retConstColumnObj
                
            End If
            
            If VBUtil.convertNullToEmptyStr(recordSet!REF_COLUMN_NAME) <> Empty Then
            
                Set retConstFkColumnObj = New ValDbDefineColumn
                retConstFkColumnObj.schemaName = VBUtil.convertNullToEmptyStr(recordSet!REF_TABLE_SCHEMA)
                retConstFkColumnObj.tableName = VBUtil.convertNullToEmptyStr(recordSet!REF_TABLE_NAME)
                retConstFkColumnObj.columnName = VBUtil.convertNullToEmptyStr(recordSet!REF_COLUMN_NAME)
                retConstObj.addFkColumn retConstFkColumnObj
                
            End If
            
        End If
            
        recordSet.MoveNext
    Loop
    
    If ret.count > 0 Then
        Set getTable = ret.getItemByIndex(1)
    Else
        Set getTable = Nothing
    End If

    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext

End Function

