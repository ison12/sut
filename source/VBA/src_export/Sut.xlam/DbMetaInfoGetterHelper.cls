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
' DBメタ情報取得ヘルパークラス
'
' 作成者　：Ison
' 履歴　　：2019/12/15　新規作成
'
' 特記事項：
' *********************************************************

' DBコネクション
Private m_dbConn As Object
' DBMS種類
Private m_dbms   As DbmsType
' クエリ生成オブジェクト
Private m_queryCreator As IDbMetaInfoQueryCreator

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
End Sub

' =========================================================
' ▽DBコネクション設定
'
' 概要　　　：DBコネクションを設定する。
' 引数　　　：vNewValue DBコネクション
'
' 特記事項　：
'
'
' =========================================================
Public Property Let dbConn(ByVal RHS As Variant)

    ' DBコネクションを設定
    Set m_dbConn = RHS
    ' DB種類を設定
    m_dbms = ADOUtil.getDBMSType(m_dbConn)
    
    ' クエリー生成オブジェクトを生成するFactoryオブジェクト
    Dim dbObjFactory As New DbObjectFactory
    
    Set m_queryCreator = dbObjFactory.createMetaInfoQueryCreatorObject(m_dbConn)
    
End Property

' =========================================================
' ▽スキーマ情報リスト取得
'
' 概要　　　：接続しているDBの全スキーマ情報を取得する
' 引数　　　：
'
' 戻り値　　：スキーマリスト。
'
' 特記事項　：
'
'
' =========================================================
Public Function getSchemaList() As ValCollection

    On Error GoTo err
    
    ' 戻り値オブジェクト
    Dim ret As New ValCollection
    Dim retObj As ValDbDefineSchema
    
    ' 実行SQLの生成
    Dim sql As String
    sql = m_queryCreator.createQuerySelectSchemaList

    ' レコードセット
    Dim recordSet   As Object
    ' クエリーを実行する
    Set recordSet = ADOUtil.querySelect(m_dbConn, sql, 0)

    ' レコードセットから情報を取得する
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
' ▽テーブル情報取得（＋カラム情報）
'
' 概要　　　：任意のテーブル情報を全て取得する。
' 引数　　　：schemaList スキーマリスト
'
' 戻り値　　：テーブル定義情報。
'
' 特記事項　：
'
'
' =========================================================
Public Function getTableList(ByVal schemaList As ValCollection) As ValCollection

    On Error GoTo err
    
    ' 戻り値オブジェクト
    Dim ret As New ValCollection
    Dim retObj As ValDbDefineTable
    
    ' 実行SQLの生成
    Dim sql As String
    sql = m_queryCreator.createQuerySelectTableList(schemaList)

    ' レコードセット
    Dim recordSet   As Object
    ' クエリーを実行する
    Set recordSet = ADOUtil.querySelect(m_dbConn, sql, 0)

    ' レコードセットから情報を取得する
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
' ▽テーブル定義取得（＋カラム情報）
'
' 概要　　　：任意のスキーマ・テーブル名をキーにしてテーブル情報を取得する。
' 引数　　　：schemaName スキーマ名
' 　　　　　　tableName  テーブル名
'
' 戻り値　　：テーブル定義。
'
' 特記事項　：
'
'
' =========================================================
Public Function getTable(ByVal schemaName As String, ByVal tableName As String) As ValDbDefineTable

    On Error GoTo err
    
    ' 戻り値オブジェクト
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

    ' レコードセット
    Dim recordSet   As Object
    
    ' ---------------------------------------------------------
    ' カラム情報
    ' ---------------------------------------------------------
    ' クエリーを実行する
    Set recordSet = ADOUtil.querySelect(m_dbConn, sqlColumn, 0)

    ' レコードセットから情報を取得する
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
    ' 制約情報
    ' ---------------------------------------------------------
    ' クエリーを実行する
    Set recordSet = ADOUtil.querySelect(m_dbConn, sqlConst, 0)

    ' レコードセットから情報を取得する
    Do While Not recordSet.EOF
    
        Set retTableObj = ret.getItem( _
            VBUtil.convertNullToEmptyStr(recordSet!TABLE_SCHEMA) & "_" & _
            VBUtil.convertNullToEmptyStr(recordSet!TABLE_NAME))
    
        If Not retTableObj Is Nothing Then
        
            Set retConstObj = retTableObj.constraintsList.getItem(convertNullToEmptyStr(recordSet!CONSTRAINT_NAME))
            If retConstObj Is Nothing Then
                Set retConstObj = New ValDbDefineTableConstraints
                retTableObj.addConstraint retConstObj
            End If
      
            retConstObj.constraintName = VBUtil.convertNullToEmptyStr(recordSet!CONSTRAINT_NAME)
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


