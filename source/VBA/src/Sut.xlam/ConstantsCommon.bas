Attribute VB_Name = "ConstantsCommon"
Option Explicit

' *********************************************************
' 共通の定数モジュール
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/03/31　新規作成
'
' 特記事項：
'
' *********************************************************

' 会社名
Public Const COMPANY_NAME       As String = "ison"
' 会社名（別名）
Public Const COMPANY_NAME_ALIAS As String = "a-l6ia5s-8o11f35i789s5432_o1172_n-9873210"

' アプリケーション名
Public Const APPLICATION_NAME As String = "Sut"
' バージョン
Public Const version As String = "1.10"

' ヘルプファイル
Public Const HELP_FILE As String = "Sut.chm"

' セッション鍵のパスワード情報
Public Const SESSION_KEY_PASSWORD As String = "SUT_20090518_RaweV%@-Asdv"

' 試用期間の日付
Public Const PROBATION_DAY As Long = 14
' 試用期間のキー
Public Const PROBATION_REG_DIR As String = "Data"
' 試用期間のキー
Public Const PROBATION_REG_KEY As String = "Kernel Key"

' コマンドバーメニューの名称
Public Const COMMANDBAR_MENU_NAME       As String = "Sut by Ison"
' コマンドバーメニューの名称（表示に使われるプロパティ）
Public Const COMMANDBAR_MENU_NAME_LOCAL As String = "Sut"

'
Public Const COMMANDBAR_CONTROL_BASE_ID As String = "Sut_CommandBarControl"
' 削除対象としないコントロール
Public Const COMMANDBAR_DONT_DELETE_TARGET As String = "DontDeleteTarget"

' ポップアップメニューの最大数
Public Const POPUP_MENU_NEW_CREATED_MAX_SIZE As Long = 10
' ポップアップメニュー表示用の関数 接頭辞
Public Const POPUP_MENU_CALL_FUNC_PREFIX As String = "Main.SutShowPopup"

' クエリパラメータの最大数
Public Const QUERY_PARAMETER_NEW_CREATED_MAX_SIZE As Long = 1000
' クエリパラメータのデフォルト名
Public Const QUERY_PARAMETER_DEFAULT_NAME As String = "param"

