Attribute VB_Name = "ConstantsSnapSqlDefine"
Option Explicit

' *********************************************************
' スナップショットSQL定義シートに関連した定数モジュール
'
' 作成者　：Ison
' 履歴　　：2019/01/03　新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' 定数
' =========================================================
Public Const NAME_ROW            As Long = 2
Public Const NAME_COL            As Long = 3

Public Const PARAM_START_ROW     As Long = 6

Public Const PARAM_ROW_COL       As Long = 2
Public Const PARAM_NAME_COL      As Long = 3
Public Const PARAM_VALUE_COL     As Long = 4
Public Const PARAM_MEMO_COL      As Long = 5

Public Const SQL_DEFINE_ROW_COL          As Long = 2
Public Const SQL_DEFINE_SQL_COL          As Long = 3
Public Const SQL_DEFINE_PRIMARY_KEY_COL  As Long = 4
Public Const SQL_DEFINE_MEMO_COL         As Long = 5

' シート名
Public Const SHEET_NAME_TEMPLATE       As String = "template_snapshotSqlDefine"
' シート判別用イメージ
Public Const SHEET_CHECK_IMAGE         As String = "SUT_WORKSHEET_SNAP_SQL_DEFINE"
