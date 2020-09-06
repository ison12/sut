VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbMetaInfoQueryCreatorHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBメタ情報クエリ生成ヘルパークラス
' DbMetaInfoQueryCreatorの各クラスで共通使用するメソッドなどを定義
'
' 作成者　：Ison
' 履歴　　：2019/12/15　新規作成
'
' 特記事項：
' *********************************************************

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

' スキーマリスト取得用クエリのディレクトリパスを取得
Private Function getQueryDir(ByVal dbms As DbmsType) As String

    Dim dbmsName As String
    dbmsName = DBUtil.getDbmsTypeName(dbms)

    Dim queryDir As String
    queryDir = VBUtil.concatFilePath(SutWorkbook.path, "resource\dbMetaInfoQuery\" & dbmsName)

    getQueryDir = queryDir

End Function

' 条件なしクエリの生成
Public Function getSqlForNoCondition(ByVal dbms As DbmsType _
                                   , ByVal sqlId As String) As String

    ' DBMS名を取得
    Dim dbmsName As String
    dbmsName = DBUtil.getDbmsTypeName(dbms)

    ' クエリディレクトリを取得
    Dim queryDir As String
    queryDir = getQueryDir(dbms)

    ' クエリファイルパスを取得
    Dim queryFilePath As String
    queryFilePath = VBUtil.concatFilePath(queryDir, LCase$(dbmsName) & "_" & sqlId & ".sql")

    ' ファイルの内容を読み込む
    Dim query     As String
    
    Dim fr As New FileReader
    fr.init queryFilePath, "UTF-8", vbCr & vbLf
    query = fr.readText
    
    getSqlForNoCondition = query

End Function

' 条件付きクエリの生成
Public Function getSqlForCondition(ByVal dbms As DbmsType _
                                    , ByVal sqlId As String _
                                    , ByVal sqlCondId As String _
                                    , Optional ByVal schemaList As ValCollection = Nothing _
                                    , Optional ByVal schemaName As String = Empty _
                                    , Optional ByVal tableName As String = Empty) As String

    Dim schemaVar As Variant
    Dim schemaStr As StringBuilder

    ' DBMS名を取得
    Dim dbmsName As String
    dbmsName = DBUtil.getDbmsTypeName(dbms)

    ' クエリディレクトリを取得
    Dim queryDir As String
    queryDir = getQueryDir(dbms)

    ' 本体クエリファイルパスを取得
    Dim queryFilePath As String
    queryFilePath = VBUtil.concatFilePath(queryDir, LCase$(dbmsName) & "_" & sqlId & ".sql")

    ' 条件クエリファイルパスを取得
    Dim queryCondFilePath As String
    queryCondFilePath = VBUtil.concatFilePath(queryDir, LCase$(dbmsName) & "_" & sqlCondId & ".sql")

    ' ファイルの内容を読み込む
    Dim query     As String
    Dim queryCond As String
    
    Dim fr As New FileReader
    fr.init queryFilePath, "UTF-8", vbCr & vbLf
    query = fr.readText

    Set fr = New FileReader
    fr.init queryCondFilePath, "UTF-8", vbCr & vbLf
    queryCond = fr.readText
    
    ' ファイルの内容に条件を埋め込む
    If schemaName <> Empty And _
        tableName <> Empty Then
        
        queryCond = replace(queryCond, "${table_schema}", DBUtil.encloseSingleQuart(DBUtil.escapeValueForSinglequart(dbms, schemaName)))
        queryCond = replace(queryCond, "${table_name}", DBUtil.encloseSingleQuart(DBUtil.escapeValueForSinglequart(dbms, tableName)))
        
        query = replace(query, "${condition}", queryCond)
        
    ElseIf Not schemaList Is Nothing Then
    
        Set schemaStr = New StringBuilder
        For Each schemaVar In schemaList.col
            If schemaStr.length <> 0 Then
                schemaStr.append ", "
            End If
            schemaStr.append DBUtil.encloseSingleQuart(DBUtil.escapeValueForSinglequart(dbms, schemaVar.schemaName))
        Next
    
        queryCond = replace(queryCond, "${schemaList}", schemaStr.str)
        query = replace(query, "${condition}", queryCond)

    Else
        
        queryCond = replace(queryCond, "${schemaList}", DBUtil.encloseSingleQuart(DBUtil.escapeValueForSinglequart(dbms, schemaName)))
        
        query = replace(query, "${condition}", queryCond)
    End If
    
    getSqlForCondition = query

End Function

