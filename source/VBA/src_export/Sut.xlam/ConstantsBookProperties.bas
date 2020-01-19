Attribute VB_Name = "ConstantsBookProperties"
Option Explicit

' *********************************************************
' アプリケーション設定の定数モジュール
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/03　新規作成
'
' 特記事項：
'
' *********************************************************

' シート名
Public Const BOOK_PROPERTIES_SHEET_NAME As String = "sut.properties"
' 開始行
Public Const FIRST_ROW As Long = 4
' テーブル クエリパラメータ
Public Const TABLE_QUERY_PARAMETER_DIALOG As Long = 2
' テーブル DB接続ダイアログ
Public Const TABLE_DB_CONNECT_DIALOG As Long = 5
' テーブル SELECT条件生成ダイアログ
Public Const TABLE_SELECT_CONDITION_CREATOR_DIALOG As Long = 11
' テーブル ファイル出力ダイアログ
Public Const TABLE_FILE_OUTPUT_DIALOG As Long = 8
' テーブル DBクエリバッチダイアログ
Public Const TABLE_DB_QUERY_BATCH_DIALOG As Long = 14
