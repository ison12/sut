VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbObjectFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�֘A�̃I�u�W�F�N�g�𐶐�����Factory�N���X�B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/05/01 �V�K�쐬
'
' ���L�����F
'
' *********************************************************

' =========================================================
' �����^���擾�I�u�W�F�N�g
'
' �T�v�@�@�@�F���^���擾�I�u�W�F�N�g�̐���
' �����@�@�@�FdbConn DB�R�l�N�V����
' �߂�l�@�@�F���^���擾�I�u�W�F�N�g
'
' =========================================================
Public Function createMetaInfoGetterObject(ByRef dbConn As Object) As IDbMetaInfoGetter

    ' �߂�l
    Dim ret As IDbMetaInfoGetter

    ' DB�̎��
    Dim t As DbmsType

    ' DB�̎�ނ��擾
    t = ADOUtil.getDBMSType(dbConn)

    ' MySQL�̏ꍇ
    If t = DbmsType.MySQL Then
    
        Set ret = New DbMetaInfoGetterMySQL
        ret.dbConn = dbConn
    
    ' PostgreSQL�̏ꍇ
    ElseIf t = DbmsType.PostgreSQL Then
    
        Set ret = New DbMetaInfoGetterPostgreSQL
        ret.dbConn = dbConn
        
    ' Oracle�̏ꍇ
    ElseIf t = DbmsType.Oracle Then
    
        Set ret = New DbMetaInfoGetterOracle
        ret.dbConn = dbConn
        
    ' Microsoft SQL Server�̏ꍇ
    ElseIf t = DbmsType.MicrosoftSqlServer Then
    
        Set ret = New DbMetaInfoGetterSQLServer
        ret.dbConn = dbConn
        
    ' Microsoft Access�̏ꍇ
    ElseIf t = DbmsType.MicrosoftAccess Then
    
        Set ret = New DbMetaInfoGetterAccess
        ret.dbConn = dbConn
        
    ' Symfoware�̏ꍇ
    ElseIf t = DbmsType.Symfoware Then
    
        Set ret = New DbMetaInfoGetterSymfoware
        ret.dbConn = dbConn
        
    ' ���̑���DB
    Else
    
        err.Raise ConstantsError.ERR_NUMBER_UNSUPPORT_DB _
                , _
                , ConstantsError.ERR_DESC_UNSUPPORT_DB
    
    End If
    
    ' �߂�l��ݒ肷��
    Set createMetaInfoGetterObject = ret
    
End Function

' =========================================================
' �����^���N�G�������擾�I�u�W�F�N�g
'
' �T�v�@�@�@�F���^���N�G�������擾�I�u�W�F�N�g�̐���
' �����@�@�@�FdbConn DB�R�l�N�V����
' �߂�l�@�@�F���^���N�G�������擾�I�u�W�F�N�g
'
' =========================================================
Public Function createMetaInfoQueryCreatorObject(ByRef dbConn As Object) As IDbMetaInfoQueryCreator

    ' �߂�l
    Dim ret As IDbMetaInfoQueryCreator

    ' DB�̎��
    Dim t As DbmsType

    ' DB�̎�ނ��擾
    t = ADOUtil.getDBMSType(dbConn)

    ' MySQL�̏ꍇ
    If t = DbmsType.MySQL Then
    
        Set ret = New DbMetaInfoQueryCreatorMySQL
    
    ' PostgreSQL�̏ꍇ
    ElseIf t = DbmsType.PostgreSQL Then
    
        Set ret = New DbMetaInfoQueryCreatorPostgres
        
    ' Oracle�̏ꍇ
    ElseIf t = DbmsType.Oracle Then
    
        Set ret = New DbMetaInfoQueryCreatorOracle
        
    ' Microsoft SQL Server�̏ꍇ
    ElseIf t = DbmsType.MicrosoftSqlServer Then
    
        Set ret = New DbMetaInfoQueryCreatorSQLServer
        
    ' Symfoware�̏ꍇ
    ElseIf t = DbmsType.Symfoware Then
    
        Set ret = New DbMetaInfoQueryCreatorSymfoware
        
    ' ���̑���DB
    Else
    
        err.Raise ConstantsError.ERR_NUMBER_UNSUPPORT_DB _
                , _
                , ConstantsError.ERR_DESC_UNSUPPORT_DB
    
    End If
    
    ' �߂�l��ݒ肷��
    Set createMetaInfoQueryCreatorObject = ret
    
End Function

' =========================================================
' ���N�G���[�����I�u�W�F�N�g
'
' �T�v�@�@�@�F�N�G���[�����I�u�W�F�N�g�̐���
' �����@�@�@�FdbConn DB�R�l�N�V����
' �߂�l�@�@�F�N�G���[�����I�u�W�F�N�g
'
' =========================================================
Public Function createQueryCreator(ByRef dbConn As Object _
                                 , ByVal emptyValueReading As Long _
                                 , ByVal directInputChar As String _
                                 , ByVal columnFormatList As ValCollection _
                                 , ByVal schemaUse As Long _
                                 , ByVal tableColumnEscape As Boolean) As IDbQueryCreator

    ' �߂�l
    Dim ret As IDbQueryCreator

    ' DB�̎��
    Dim t As DbmsType

    ' DB�̎�ނ��擾
    t = ADOUtil.getDBMSType(dbConn)

    ' MySQL�̏ꍇ
    If t = DbmsType.MySQL Then
    
        Set ret = New DbQueryCreatorMySQL
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputChar = directInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
    
    ' PostgreSQL�̏ꍇ
    ElseIf t = DbmsType.PostgreSQL Then
    
        Set ret = New DbQueryCreatorPostgreSQL
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputChar = directInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' Oracle�̏ꍇ
    ElseIf t = DbmsType.Oracle Then
    
        Set ret = New DbQueryCreatorOracle
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputChar = directInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' Microsoft SQL Server�̏ꍇ
    ElseIf t = DbmsType.MicrosoftSqlServer Then
    
        Set ret = New DbQueryCreatorSQLServer
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputChar = directInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' Microsoft Access�̏ꍇ
    ElseIf t = DbmsType.MicrosoftAccess Then
    
        Set ret = New DbQueryCreatorAccess
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputChar = directInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' Symfoware�̏ꍇ
    ElseIf t = DbmsType.Symfoware Then
    
        Set ret = New DbQueryCreatorSymfoware
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputChar = directInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' ���̑���DB
    Else
    
        err.Raise ConstantsError.ERR_NUMBER_UNSUPPORT_DB _
                , _
                , ConstantsError.ERR_DESC_UNSUPPORT_DB
    
    End If
    
    ' �߂�l��ݒ肷��
    Set createQueryCreator = ret
    
End Function

' =========================================================
' ���J�������I�u�W�F�N�g
'
' �T�v�@�@�@�F�J�������I�u�W�F�N�g�̐���
' �����@�@�@�FdbConn DB�R�l�N�V����
'
' �߂�l�@�@�F�N�G���[�����I�u�W�F�N�g
'
' =========================================================
Public Function createColumnType(ByVal dbt As DbmsType) As IDbColumnType

    ' �߂�l
    Dim ret As IDbColumnType

    ' MySQL�̏ꍇ
    If dbt = DbmsType.MySQL Then
    
        Set ret = New DbColumnTypeMySQL
    
    ' PostgreSQL�̏ꍇ
    ElseIf dbt = DbmsType.PostgreSQL Then
    
        Set ret = New DbColumnTypePostgreSQL
        
    ' Oracle�̏ꍇ
    ElseIf dbt = DbmsType.Oracle Then
    
        Set ret = New DbColumnTypeOracle
        
    ' Microsoft SQL Server�̏ꍇ
    ElseIf dbt = DbmsType.MicrosoftSqlServer Then
    
        Set ret = New DbColumnTypeSqlServer
        
    ' Microsoft Access�̏ꍇ
    ElseIf dbt = DbmsType.MicrosoftAccess Then
    
        Set ret = New DbColumnTypeAccess
        
    ' Symfoware�̏ꍇ
    ElseIf dbt = DbmsType.Symfoware Then
    
        Set ret = New DbColumnTypeSymfoware
        
    ' ���̑���DB
    Else
    
        err.Raise ConstantsError.ERR_NUMBER_UNSUPPORT_DB _
                , _
                , ConstantsError.ERR_DESC_UNSUPPORT_DB
    
    End If
    
    ' �߂�l��ݒ肷��
    Set createColumnType = ret
    
End Function

