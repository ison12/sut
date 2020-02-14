VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbMetaInfoGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのメタ情報を取得するインターフェース。
' 対応するDB毎に本インターフェースをImplementsする。
'
' 作成者　：Ison
' 履歴　　：2009/02/11　新規作成
'
' 特記事項：
'
' *********************************************************

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
Public Property Let dbConn(ByVal vNewValue As Variant)

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

End Function

' =========================================================
' ▽テーブルリスト取得
'
' 概要　　　：任意のスキーマに含まれるテーブル情報を全て取得する。
' 引数　　　：conn       DBコネクション
' 　　　　　　schemaList スキーマリスト
'
' 戻り値　　：テーブル定義リスト。
'
' 特記事項　：
'
'
' =========================================================
Public Function getTableList(ByRef schemaList As ValCollection) As ValCollection

End Function

' =========================================================
' ▽テーブル情報取得（＋カラム情報）
'
' 概要　　　：スキーマ・テーブル名をキーにしてテーブル情報を取得する。
' 引数　　　：schemaName スキーマ名
' 　　　　　　tableName  テーブル名
'
' 戻り値　　：テーブル定義情報。
'
' 特記事項　：
'
'
' =========================================================
Public Function getTable(ByRef schemaName As String, ByRef tableName As String) As ValDbDefineTable

End Function
