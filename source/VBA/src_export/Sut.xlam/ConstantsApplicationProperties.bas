Attribute VB_Name = "ConstantsApplicationProperties"
Option Explicit

' *********************************************************
' アプリケーション設定の定数モジュール
'
' 作成者　：Ison
' 履歴　　：2019/12/03　新規作成
'
' 特記事項：
'
' *********************************************************

' シート名
Public Const BOOK_PROPERTIES_SHEET_NAME     As String = "sut.properties"

Public Const INI_FILE_DIR_FORM_POSITION     As String = "formPosition"
Public Const INI_FILE_DIR_FORM              As String = "form"
Public Const INI_FILE_DIR_OPTION            As String = "option"
Public Const INI_FILE_DIR_OPTION_COL_FORMAT As String = "colFormat"
Public Const INI_FILE_DIR_QUERY             As String = "query"

Public Const INI_SECTION_DEFAULT            As String = "default"

Public Const INI_KEY_X As String = "x"
Public Const INI_KEY_Y As String = "y"

' SELECTの再実行のSQL本体
Public Const INI_KEY_SELECT_LATEST_SQL                As String = "sql"
' SELECTの再実行の条件指定の場合の追加有無フラグ
Public Const INI_KEY_SELECT_LATEST_TYPE               As String = "type"
' SELECTの再実行の種類（全て）
Public Const INI_KEY_SELECT_LATEST_TYPE_ALL           As String = "1"
' SELECTの再実行の種類（条件指定）
Public Const INI_KEY_SELECT_LATEST_TYPE_CONDITION     As String = "2"
' SELECTの再実行の条件指定の場合の追加有無フラグ
Public Const INI_KEY_SELECT_LATEST_APPEND             As String = "append"
' SELECTの再実行の条件指定の場合の追加有無フラグ
Public Const INI_KEY_SELECT_LATEST_APPEND_TRUE        As String = "True"

