VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbMetaInfoQueryCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのメタ情報を取得するクエリを生成するインターフェース。
' 対応するDB毎に本インターフェースをImplementsする。
'
' 作成者　：Ison
' 履歴　　：2019/12/15　新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' ▽スキーマリスト取得用クエリの生成
'
' 概要　　　：スキーマリスト取得用クエリの生成を生成して返却する。
' 引数　　　：
'
' 戻り値　　：スキーマリスト取得用クエリの生成文字列
'
' 特記事項　：
'
'
' =========================================================
Public Function createQuerySelectSchemaList() As String

End Function


' =========================================================
' ▽任意のスキーマ配下のテーブルリスト取得用クエリの生成
'
' 概要　　　：任意のスキーマ配下のテーブルリスト取得用クエリの生成を生成して返却する。
' 引数　　　：schemaName スキーマ名
'
' 戻り値　　：テーブルリスト取得用クエリの生成文字列
'
' 特記事項　：
'
'
' =========================================================
Public Function createQuerySelectTableList(ByVal schemaList As ValCollection) As String

End Function


' =========================================================
' ▽任意のスキーマ配下のテーブル制約リスト取得用クエリの生成
'
' 概要　　　：任意のスキーマ配下のテーブル制約リスト取得用クエリの生成を生成して返却する。
' 引数　　　：schemaName スキーマ名
'
' 戻り値　　：テーブル制約リスト取得用クエリの生成文字列
'
' 特記事項　：
'
'
' =========================================================
Public Function createQuerySelectTableConstList(ByVal schemaList As ValCollection) As String

End Function


' =========================================================
' ▽任意のテーブル情報取得用クエリの生成
'
' 概要　　　：任意のテーブル情報取得用クエリの生成を生成して返却する。
' 引数　　　：schemaName スキーマ名
' 　　　　　： tableName テーブル名
'
' 戻り値　　：テーブル情報取得用クエリの生成文字列
'
' 特記事項　：
'
'
' =========================================================
Public Function createQuerySelectTable(ByVal schemaName As String, ByVal tableName As String) As String

End Function


' =========================================================
' ▽任意のテーブルのカラム情報取得用クエリの生成
'
' 概要　　　：任意のテーブルのカラム情報取得用クエリの生成を生成して返却する。
' 引数　　　：schemaName スキーマ名
' 　　　　　： tableName テーブル名
'
' 戻り値　　：カラム情報取得用クエリの生成文字列
'
' 特記事項　：
'
'
' =========================================================
Public Function createQuerySelectColumn(ByVal schemaName As String, ByVal tableName As String) As String

End Function


' =========================================================
' ▽任意のテーブルのテーブル制約情報取得用クエリの生成
'
' 概要　　　：任意のテーブルのテーブル制約情報取得用クエリの生成を生成して返却する。
' 引数　　　：schemaName スキーマ名
' 　　　　　： tableName テーブル名
'
' 戻り値　　：テーブル制約情報取得用クエリの生成文字列
'
' 特記事項　：
'
'
' =========================================================
Public Function createQuerySelectTableConst(ByVal schemaName As String, ByVal tableName As String) As String

End Function


