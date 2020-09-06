Attribute VB_Name = "ConstantsTable"
Option Explicit

' *********************************************************
' テーブルシートに関連した定数モジュール
'
' 作成者　：Ison
' 履歴　　：2009/03/31　新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' 定数
' =========================================================
Public Const SCHEMA_NAME_ROW            As Long = 1
Public Const SCHEMA_NAME_COL            As Long = 2

Public Const TABLE_NAME_ROW            As Long = 3
Public Const TABLE_NAME_COL            As Long = 2

Public Const TABLE_NAME_LOG_ROW        As Long = 4
Public Const TABLE_NAME_LOG_COL        As Long = 2

Public Const TABLE_OVER_MAX_COL_SIZE_ROW   As Long = 1
Public Const TABLE_OVER_MAX_COL_SIZE_COL   As Long = 3

Public Const TABLE_ERROR_ICON_ROW As Long = 1
Public Const TABLE_ERROR_ICON_COL As Long = 1

' シート名
Public Const U_SHEET_NAME_TEMPLATE       As String = "template_tableSheet_U"
' シート判別用イメージ
Public Const U_SHEET_CHECK_IMAGE         As String = "SUT_WORKSHEET_MARK_TO_UNDER"
' エラー時に利用するイメージ
Public Const ERROR_ICON                  As String = "SUT_ERROR_ICON"

Public Const U_COLUMN_NAME_ROW           As Long = 6
Public Const U_COLUMN_NAME_LOG_ROW       As Long = 7
Public Const U_COLUMN_NAME_LOG_EXT_ROW   As Long = 7
Public Const U_COLUMN_TYPE_ROW           As Long = 8
Public Const U_COLUMN_NULL_ROW           As Long = 9
Public Const U_COLUMN_DEF_ROW            As Long = 10
Public Const U_COLUMN_PK_ROW             As Long = 11
Public Const U_COLUMN_UK_ROW             As Long = 12
Public Const U_COLUMN_REFER_ROW          As Long = 13

Public Const U_COLUMN_OFFSET_COL         As Long = 3

Public Const U_RECORD_NUM_COL            As Long = 2
Public Const U_RECORD_OFFSET_ROW         As Long = 14

' シート名
Public Const R_SHEET_NAME_TEMPLATE       As String = "template_tableSheet_R"
' シート判別用イメージ
Public Const R_SHEET_CHECK_IMAGE         As String = "SUT_WORKSHEET_MARK_TO_RIGHT"

Public Const R_COLUMN_NAME_COL           As Long = 2
Public Const R_COLUMN_NAME_LOG_COL       As Long = 3
Public Const R_COLUMN_NAME_EXT_LOG_COL   As Long = 3
Public Const R_COLUMN_TYPE_COL           As Long = 4
Public Const R_COLUMN_NULL_COL           As Long = 6
Public Const R_COLUMN_DEF_COL            As Long = 5
Public Const R_COLUMN_PK_COL             As Long = 7
Public Const R_COLUMN_UK_COL             As Long = 8
Public Const R_COLUMN_REFER_COL          As Long = 9

Public Const R_COLUMN_OFFSET_ROW         As Long = 7

Public Const R_RECORD_NUM_ROW            As Long = 6
Public Const R_RECORD_OFFSET_COL         As Long = 10

' シート名
Public Const QUERY_RESULT_SHEET_NAME_TEMPLATE As String = "template_queryEditorResultSheet"
' シート判別用イメージ
Public Const QUERY_RESULT_SHEET_CHECK_IMAGE   As String = "SUT_WORKSHEET_MARK_QUERY_RESULT"
' シート名　デフォルト
Public Const QUERY_RESULT_SHEET_DEFAULT_NAME  As String = "QueryResult"

' シート名
Public Const QUERY_SHEET_NAME_TEMPLATE       As String = "template_queryEditorResult"

Public Const QUERY_COLUMN_TITLE_COL      As Long = 1

Public Const QUERY_COLUMN_OFFSET_ROW     As Long = 2
Public Const QUERY_COLUMN_OFFSET_COL     As Long = 2

Public Const QUERY_HEADER_ROW     As Long = 1
Public Const QUERY_HEADER_COL     As Long = 1

Public Const QUERY_RECORD_ROW     As Long = 2
Public Const QUERY_RECORD_COL     As Long = 1

Public Const QUERY_RESULT_ROW     As Long = 3
Public Const QUERY_RESULT_COL     As Long = 1

Public Const QUERY_ERROR_ROW      As Long = 4
Public Const QUERY_ERROR_COL      As Long = 1

Public Const QUERY_TITLE_ROW     As Long = 5
Public Const QUERY_TITLE_COL     As Long = 1

Public Const QUERY_ROWNUMBER_ROW     As Long = 6
Public Const QUERY_ROWNUMBER_COL     As Long = 1

Public Const QUERY_RESULTSET_ROW     As Long = 7
Public Const QUERY_RESULTSET_COL     As Long = 1

