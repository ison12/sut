VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbMetaInfoQueryCreatorMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBメタ情報クエリ生成クラス
' MySQL用
'
' IDbMetaInfoQueryCreatorを実装したクラス
'
' 作成者　：Ison
' 履歴　　：2019/12/15　新規作成
'
' 特記事項：
' *********************************************************

Implements IDbMetaInfoQueryCreator

' DBMS
Private dbms As DbmsType
' ヘルパーオブジェクト
Private helper As DbMetaInfoQueryCreatorHelper

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    dbms = DbmsType.MySQL
    Set helper = New DbMetaInfoQueryCreatorHelper
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    Set helper = Nothing
End Sub

' スキーマリスト取得用クエリの生成
Private Function IDbMetaInfoQueryCreator_createQuerySelectSchemaList() As String

    ' SQLを取得する
    IDbMetaInfoQueryCreator_createQuerySelectSchemaList = helper.getSqlForNoCondition(dbms, "schema_list")

End Function

' テーブルリスト取得用クエリの生成
Private Function IDbMetaInfoQueryCreator_createQuerySelectTableList(ByVal schemaList As ValCollection) As String

    ' SQLを取得する
   IDbMetaInfoQueryCreator_createQuerySelectTableList = helper.getSqlForCondition(dbms, "table", "table_cond_sl", schemaList)
End Function

' テーブル制約リスト取得用クエリの生成
Private Function IDbMetaInfoQueryCreator_createQuerySelectTableConstList(ByVal schemaList As ValCollection) As String

    ' SQLを取得する
   IDbMetaInfoQueryCreator_createQuerySelectTableConstList = helper.getSqlForCondition(dbms, "table_const", "table_const_cond_sl", schemaList)
End Function

' テーブル情報取得用クエリの生成
Private Function IDbMetaInfoQueryCreator_createQuerySelectTable(ByVal schemaName As String, ByVal tableName As String) As String

    ' SQLを取得する
   IDbMetaInfoQueryCreator_createQuerySelectTable = helper.getSqlForCondition(dbms, "table", "table_cond_t", Nothing, schemaName, tableName)
End Function

' カラム情報取得用クエリの生成
Private Function IDbMetaInfoQueryCreator_createQuerySelectColumn(ByVal schemaName As String, ByVal tableName As String) As String

    ' SQLを取得する
   IDbMetaInfoQueryCreator_createQuerySelectColumn = helper.getSqlForCondition(dbms, "column", "table_cond_t", Nothing, schemaName, tableName)
End Function

' テーブル制約情報取得用クエリの生成
Private Function IDbMetaInfoQueryCreator_createQuerySelectTableConst(ByVal schemaName As String, ByVal tableName As String) As String

    ' SQLを取得する
   IDbMetaInfoQueryCreator_createQuerySelectTableConst = helper.getSqlForCondition(dbms, "table_const", "table_const_cond_t", Nothing, schemaName, tableName)
End Function

