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
' DB関連のオブジェクトを生成するFactoryクラス。
'
' 作成者　：Ison
' 履歴　　：2009/05/01 新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' ▽メタ情報取得オブジェクト
'
' 概要　　　：メタ情報取得オブジェクトの生成
' 引数　　　：dbConn DBコネクション
' 戻り値　　：メタ情報取得オブジェクト
'
' =========================================================
Public Function createMetaInfoGetterObject(ByRef dbConn As Object) As IDbMetaInfoGetter

    ' 戻り値
    Dim ret As IDbMetaInfoGetter

    ' DBの種類
    Dim t As DbmsType

    ' DBの種類を取得
    t = ADOUtil.getDBMSType(dbConn)

    ' MySQLの場合
    If t = DbmsType.MySQL Then
    
        Set ret = New DbMetaInfoGetterMySQL
        ret.dbConn = dbConn
    
    ' PostgreSQLの場合
    ElseIf t = DbmsType.PostgreSQL Then
    
        Set ret = New DbMetaInfoGetterPostgreSQL
        ret.dbConn = dbConn
        
    ' Oracleの場合
    ElseIf t = DbmsType.Oracle Then
    
        Set ret = New DbMetaInfoGetterOracle
        ret.dbConn = dbConn
        
    ' Microsoft SQL Serverの場合
    ElseIf t = DbmsType.MicrosoftSqlServer Then
    
        Set ret = New DbMetaInfoGetterSQLServer
        ret.dbConn = dbConn
        
    ' Microsoft Accessの場合
    ElseIf t = DbmsType.MicrosoftAccess Then
    
        Set ret = New DbMetaInfoGetterAccess
        ret.dbConn = dbConn
        
    ' Symfowareの場合
    ElseIf t = DbmsType.Symfoware Then
    
        Set ret = New DbMetaInfoGetterSymfoware
        ret.dbConn = dbConn
        
    ' その他のDB
    Else
    
        err.Raise ConstantsError.ERR_NUMBER_UNSUPPORT_DB _
                , _
                , ConstantsError.ERR_DESC_UNSUPPORT_DB
    
    End If
    
    ' 戻り値を設定する
    Set createMetaInfoGetterObject = ret
    
End Function

' =========================================================
' ▽メタ情報クエリ生成取得オブジェクト
'
' 概要　　　：メタ情報クエリ生成取得オブジェクトの生成
' 引数　　　：dbConn DBコネクション
' 戻り値　　：メタ情報クエリ生成取得オブジェクト
'
' =========================================================
Public Function createMetaInfoQueryCreatorObject(ByRef dbConn As Object) As IDbMetaInfoQueryCreator

    ' 戻り値
    Dim ret As IDbMetaInfoQueryCreator

    ' DBの種類
    Dim t As DbmsType

    ' DBの種類を取得
    t = ADOUtil.getDBMSType(dbConn)

    ' MySQLの場合
    If t = DbmsType.MySQL Then
    
        Set ret = New DbMetaInfoQueryCreatorMySQL
    
    ' PostgreSQLの場合
    ElseIf t = DbmsType.PostgreSQL Then
    
        Set ret = New DbMetaInfoQueryCreatorPostgres
        
    ' Oracleの場合
    ElseIf t = DbmsType.Oracle Then
    
        Set ret = New DbMetaInfoQueryCreatorOracle
        
    ' Microsoft SQL Serverの場合
    ElseIf t = DbmsType.MicrosoftSqlServer Then
    
        Set ret = New DbMetaInfoQueryCreatorSQLServer
        
    ' Symfowareの場合
    ElseIf t = DbmsType.Symfoware Then
    
        Set ret = New DbMetaInfoQueryCreatorSymfoware
        
    ' その他のDB
    Else
    
        err.Raise ConstantsError.ERR_NUMBER_UNSUPPORT_DB _
                , _
                , ConstantsError.ERR_DESC_UNSUPPORT_DB
    
    End If
    
    ' 戻り値を設定する
    Set createMetaInfoQueryCreatorObject = ret
    
End Function

' =========================================================
' ▽クエリー生成オブジェクト
'
' 概要　　　：クエリー生成オブジェクトの生成
' 引数　　　：dbConn DBコネクション
' 戻り値　　：クエリー生成オブジェクト
'
' =========================================================
Public Function createQueryCreator(ByRef dbConn As Object _
                                 , ByVal emptyValueReading As Long _
                                 , ByVal directInputCharPrefix As String _
                                 , ByVal directInputCharSuffix As String _
                                 , ByVal nullInputChar As String _
                                 , ByVal columnFormatList As ValCollection _
                                 , ByVal schemaUse As Long _
                                 , ByVal tableColumnEscape As Boolean) As IDbQueryCreator

    ' 戻り値
    Dim ret As IDbQueryCreator

    ' DBの種類
    Dim t As DbmsType

    ' DBの種類を取得
    t = ADOUtil.getDBMSType(dbConn)

    ' MySQLの場合
    If t = DbmsType.MySQL Then
    
        Set ret = New DbQueryCreatorMySQL
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputCharPrefix = directInputCharPrefix
        ret.directInputCharSuffix = directInputCharSuffix
        ret.nullInputChar = nullInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
    
    ' PostgreSQLの場合
    ElseIf t = DbmsType.PostgreSQL Then
    
        Set ret = New DbQueryCreatorPostgreSQL
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputCharPrefix = directInputCharPrefix
        ret.directInputCharSuffix = directInputCharSuffix
        ret.nullInputChar = nullInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' Oracleの場合
    ElseIf t = DbmsType.Oracle Then
    
        Set ret = New DbQueryCreatorOracle
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputCharPrefix = directInputCharPrefix
        ret.directInputCharSuffix = directInputCharSuffix
        ret.nullInputChar = nullInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' Microsoft SQL Serverの場合
    ElseIf t = DbmsType.MicrosoftSqlServer Then
    
        Set ret = New DbQueryCreatorSQLServer
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputCharPrefix = directInputCharPrefix
        ret.directInputCharSuffix = directInputCharSuffix
        ret.nullInputChar = nullInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' Microsoft Accessの場合
    ElseIf t = DbmsType.MicrosoftAccess Then
    
        Set ret = New DbQueryCreatorAccess
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputCharPrefix = directInputCharPrefix
        ret.directInputCharSuffix = directInputCharSuffix
        ret.nullInputChar = nullInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' Symfowareの場合
    ElseIf t = DbmsType.Symfoware Then
    
        Set ret = New DbQueryCreatorSymfoware
        ret.dbConn = dbConn
        ret.emptyValueReading = emptyValueReading
        ret.directInputCharPrefix = directInputCharPrefix
        ret.directInputCharSuffix = directInputCharSuffix
        ret.nullInputChar = nullInputChar
        ret.columnFormatList = columnFormatList
        ret.schemaUse = schemaUse
        ret.tableColumnEscape = tableColumnEscape
        
    ' その他のDB
    Else
    
        err.Raise ConstantsError.ERR_NUMBER_UNSUPPORT_DB _
                , _
                , ConstantsError.ERR_DESC_UNSUPPORT_DB
    
    End If
    
    ' 戻り値を設定する
    Set createQueryCreator = ret
    
End Function

' =========================================================
' ▽カラム情報オブジェクト
'
' 概要　　　：カラム情報オブジェクトの生成
' 引数　　　：dbConn DBコネクション
'
' 戻り値　　：クエリー生成オブジェクト
'
' =========================================================
Public Function createColumnType(ByVal dbt As DbmsType) As IDbColumnType

    ' 戻り値
    Dim ret As IDbColumnType

    ' MySQLの場合
    If dbt = DbmsType.MySQL Then
    
        Set ret = New DbColumnTypeMySQL
    
    ' PostgreSQLの場合
    ElseIf dbt = DbmsType.PostgreSQL Then
    
        Set ret = New DbColumnTypePostgreSQL
        
    ' Oracleの場合
    ElseIf dbt = DbmsType.Oracle Then
    
        Set ret = New DbColumnTypeOracle
        
    ' Microsoft SQL Serverの場合
    ElseIf dbt = DbmsType.MicrosoftSqlServer Then
    
        Set ret = New DbColumnTypeSqlServer
        
    ' Microsoft Accessの場合
    ElseIf dbt = DbmsType.MicrosoftAccess Then
    
        Set ret = New DbColumnTypeAccess
        
    ' Symfowareの場合
    ElseIf dbt = DbmsType.Symfoware Then
    
        Set ret = New DbColumnTypeSymfoware
        
    ' その他のDB
    Else
    
        err.Raise ConstantsError.ERR_NUMBER_UNSUPPORT_DB _
                , _
                , ConstantsError.ERR_DESC_UNSUPPORT_DB
    
    End If
    
    ' 戻り値を設定する
    Set createColumnType = ret
    
End Function


