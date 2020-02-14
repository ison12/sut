Attribute VB_Name = "ConstantsSnapDiff"
Option Explicit

' *********************************************************
' スナップショット比較シートに関連した定数モジュール
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
Public Const NAME_ROW            As Long = 3
Public Const NAME_COL            As Long = 3

Public Const MODIFY_ALL_ROW            As Long = 4
Public Const MODIFY_ALL_COL            As Long = 4

Public Const RESULT_START_ROW    As Long = 6
Public Const RESULT_START_OFFSET_SQL       As Long = 1
Public Const RESULT_START_OFFSET_MODIFY    As Long = 2
Public Const RESULT_START_OFFSET_HEADER    As Long = 3
Public Const RESULT_START_OFFSET_RECORD    As Long = 4

Public Const RESULT_RANGE_START_ROW As Long = 6
Public Const RESULT_RANGE_START_COL As Long = 2
Public Const RESULT_RANGE_END_ROW   As Long = 10
Public Const RESULT_RANGE_END_COL   As Long = 14

Public Const MODIFY_COL          As Long = 4
Public Const TOTAL_COUNT_COL     As Long = 6
Public Const NOCHANGE_COUNT_COL  As Long = 8
Public Const INSERT_COUNT_COL    As Long = 10
Public Const UPDATE_COUNT_COL    As Long = 12
Public Const DELETE_COUNT_COL    As Long = 14

Public Const SQL_NUM_COL         As Long = 2
Public Const SQL_COL             As Long = 3
Public Const PKEY_COL            As Long = 9
Public Const MEMO_COL            As Long = 11

Public Const HEADER_COL          As Long = 4

Public Const RECORD_NUM_COL      As Long = 2
Public Const RECORD_MODIFY_COL   As Long = 3
Public Const RECORD_COL          As Long = 4

Public Const MODIFY_INSERT As String = "Insert"
Public Const MODIFY_UPDATE As String = "Update"
Public Const MODIFY_DELETE As String = "Delete"
Public Const MODIFY_NOCHANGE As String = "No change"
Public Const MODIFY_NORECORD As String = "No record"
Public Const MODIFY_OFF As String = "なし"
Public Const MODIFY_ON  As String = "あり"

Public Const INSERT_COLOR_R As Long = 204
Public Const INSERT_COLOR_G As Long = 233
Public Const INSERT_COLOR_B As Long = 255

Public Const UPDATE_COLOR_R As Long = 255
Public Const UPDATE_COLOR_G As Long = 255
Public Const UPDATE_COLOR_B As Long = 153

Public Const DELETE_COLOR_R As Long = 255
Public Const DELETE_COLOR_G As Long = 204
Public Const DELETE_COLOR_B As Long = 255

Public Const MODIFY_NOCHANGE_COLOR_R As Long = 204
Public Const MODIFY_NOCHANGE_COLOR_G As Long = 233
Public Const MODIFY_NOCHANGE_COLOR_B As Long = 255

Public Const MODIFY_CHANGE_COLOR_R As Long = 255
Public Const MODIFY_CHANGE_COLOR_G As Long = 255
Public Const MODIFY_CHANGE_COLOR_B As Long = 153

' シート名
Public Const SHEET_NAME_TEMPLATE       As String = "template_diffResult"

