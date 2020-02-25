VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValApplicationSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' アプリケーションのオプション情報
'
' 作成者　：Ison
' 履歴　　：2008/03/14　新規作成
'
' 特記事項：
' *********************************************************

' =========================================================
' ▽メンバ
' =========================================================
' ブック
Private book As Workbook

' レコード処理単位
Public recProcessCount        As Long
' レコード処理単位（ユーザ指定）
Public recProcessCountCustom  As Long
' コミット確認
Public commitConfirm          As Long
' SQLエラー時の挙動
Public sqlErrorHandling       As Long
' スキーマ利用
Public schemaUse              As Long
' 正常時のクエリ結果表示
Public queryResultShowWhenNormal As Boolean

' セル書式 フォント名
Public cellFontName              As String
' セル書式 フォントサイズ
Public cellFontSize              As Double
' セル書式 折り返し有無
Public cellWordwrap              As Boolean
' セル書式 セル幅（単位は文字数）
Public cellWidth                 As Double
' セル書式 セル高さ（単位はポイント）
Public cellHeight                As Double
' 行高の自動調整
Public lineHeightAutoAdjust      As Boolean

' 空白セル読み取り方式
Public emptyCellReading             As Long
' 直接入力文字
Public directInputChar              As Long
' 直接入力文字（ユーザ指定）接頭辞
Public directInputCharCustomPrefix  As String
' 直接入力文字（ユーザ指定）接尾辞
Public directInputCharCustomSuffix  As String
' クエリパラメータの囲み文字接頭辞
Public queryParameterEncloseCustomPrefix  As String
' クエリパラメータの囲み文字接尾辞
Public queryParameterEncloseCustomSuffix  As String
' NULL入力文字
Public nullInputChar                As Long
' NULL入力文字（ユーザ指定）接頭辞
Public nullInputCharCustom          As String
' SELECT時のセルの最大文字数チェック
Public selectCheckCellMaxLength As Boolean

' テーブル・カラム名のエスケープ
Public tableColumnEscapeOracle     As Boolean
Public tableColumnEscapeMysql      As Boolean
Public tableColumnEscapePostgresql As Boolean
Public tableColumnEscapeSqlserver  As Boolean
Public tableColumnEscapeAccess  As Boolean
Public tableColumnEscapeSymfoware  As Boolean

' =========================================================
' ▽定数プロパティ
' 　Classモジュールでは、Public Constが定義できないのでプロパティで代用する
' =========================================================
' キー値 レコード処理単位
Public Property Get KEY_REC_PROCESS_COUNT() As String
    KEY_REC_PROCESS_COUNT = "recProcessCount"
End Property

' キー値 レコード処理単位　ユーザ指定
Public Property Get KEY_REC_PROCESS_COUNT_CUSTOM() As String
    KEY_REC_PROCESS_COUNT_CUSTOM = "recProcessCountCustom"
End Property

' キー値 コミット確認
Public Property Get KEY_COMMIT_CONFIRM() As String
    KEY_COMMIT_CONFIRM = "commitConfirm"
End Property

' キー値 SQLエラー時の挙動
Public Property Get KEY_SQL_ERROR_HANDLING() As String
    KEY_SQL_ERROR_HANDLING = "sqlErrorHandling"
End Property

' キー値 スキーマ
Public Property Get KEY_SCHEMA_USE() As String
    KEY_SCHEMA_USE = "schemaUse"
End Property

' キー値 正常時のクエリ結果表示有無
Public Property Get KEY_QUERY_RESULT_SHOW_WHEN_NORMAL() As String
    KEY_QUERY_RESULT_SHOW_WHEN_NORMAL = "queryResultShowWhenNormal"
End Property

' キー値 セル書式 フォント名
Public Property Get KEY_CELL_FONT_NAME() As String
    KEY_CELL_FONT_NAME = "cellFontName"
End Property

' キー値 セル書式 フォントサイズ
Public Property Get KEY_CELL_FONT_SIZE() As String
    KEY_CELL_FONT_SIZE = "cellFontSize"
End Property

' キー値 セル書式 折り返し
Public Property Get KEY_CELL_WORDWRAP() As String
    KEY_CELL_WORDWRAP = "cellWordwrap"
End Property

' キー値 セル書式 幅
Public Property Get KEY_CELL_WIDTH() As String
    KEY_CELL_WIDTH = "cellWidth"
End Property

' キー値 セル書式 高さ
Public Property Get KEY_CELL_HEIGHT() As String
    KEY_CELL_HEIGHT = "cellHeight"
End Property

' キー値 行高の自動調整
Public Property Get KEY_LINE_HEIGHT_AUTO_ADJUST() As String
    KEY_LINE_HEIGHT_AUTO_ADJUST = "lineHeightAutoAdjust"
End Property

' キー値 空白セル読み取り方式
Public Property Get KEY_EMPTY_CELL_READING() As String
    KEY_EMPTY_CELL_READING = "emptyCellReading"
End Property

' キー値 直接入力文字
Public Property Get KEY_DIRECT_INPUT_CHAR() As String
    KEY_DIRECT_INPUT_CHAR = "directInputChar"
End Property

' キー値 直接入力文字　ユーザ指定
Public Property Get KEY_DIRECT_INPUT_CHAR_CUSTOM_PREFIX() As String
    KEY_DIRECT_INPUT_CHAR_CUSTOM_PREFIX = "directInputCharCustomPrefix"
End Property
Public Property Get KEY_DIRECT_INPUT_CHAR_CUSTOM_SUFFIX() As String
    KEY_DIRECT_INPUT_CHAR_CUSTOM_SUFFIX = "directInputCharCustomSuffix"
End Property

' キー値 クエリパラメータの囲み文字　ユーザ指定
Public Property Get KEY_QUERY_PARAMETER_ENCLOSE_CUSTOM_PREFIX() As String
    KEY_QUERY_PARAMETER_ENCLOSE_CUSTOM_PREFIX = "queryParameterEncloseCustomPrefix"
End Property
Public Property Get KEY_QUERY_PARAMETER_ENCLOSE_CUSTOM_SUFFIX() As String
    KEY_QUERY_PARAMETER_ENCLOSE_CUSTOM_SUFFIX = "queryParameterEncloseCustomSuffix"
End Property

' キー値 NULL入力文字
Public Property Get KEY_NULL_INPUT_CHAR() As String
    KEY_NULL_INPUT_CHAR = "nullInputChar"
End Property

' キー値 NULL入力文字　ユーザ指定
Public Property Get KEY_NULL_INPUT_CHAR_CUSTOM() As String
    KEY_NULL_INPUT_CHAR_CUSTOM = "nullInputCharCustom"
End Property

' キー値 SELECT時のセルの最大文字数チェック
Public Property Get KEY_SELECT_CHECK_CELL_MAX_LENGTH() As String
    KEY_SELECT_CHECK_CELL_MAX_LENGTH = "selectCheckCellMaxLength"
End Property

' キー値 テーブルカラム名エスケープ Oracle
Public Property Get KEY_TABLE_COLUMN_ESCAPE_ORACLE() As String
    KEY_TABLE_COLUMN_ESCAPE_ORACLE = "tableColumnEscapeOracle"
End Property

' キー値 テーブルカラム名エスケープ Mysql
Public Property Get KEY_TABLE_COLUMN_ESCAPE_MYSQL() As String
    KEY_TABLE_COLUMN_ESCAPE_MYSQL = "tableColumnEscapeMysql"
End Property

' キー値 テーブルカラム名エスケープ Postgresql
Public Property Get KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL() As String
    KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL = "tableColumnEscapePostgresql"
End Property

' キー値 テーブルカラム名エスケープ Sqlserver
Public Property Get KEY_TABLE_COLUMN_ESCAPE_SQLSERVER() As String
    KEY_TABLE_COLUMN_ESCAPE_SQLSERVER = "tableColumnEscapeSqlserver"
End Property

' キー値 テーブルカラム名エスケープ Access
Public Property Get KEY_TABLE_COLUMN_ESCAPE_ACCESS() As String
    KEY_TABLE_COLUMN_ESCAPE_ACCESS = "tableColumnEscapeAccess"
End Property

' キー値 テーブルカラム名エスケープ Symfoware
Public Property Get KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE() As String
    KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE = "tableColumnEscapeSymfoware"
End Property

' レコード処理単位（全て）
Public Property Get REC_PROCESS_COUNT_ALL() As Long

    REC_PROCESS_COUNT_ALL = 1
End Property

' レコード処理単位（ユーザ指定）
Public Property Get REC_PROCESS_COUNT_COSTOM() As Long

    REC_PROCESS_COUNT_COSTOM = 3
End Property

' レコード処理単位（ユーザ指定）デフォルト値
Public Property Get REC_PROCESS_COUNT_COSTOM_DEFAULT() As Long

    REC_PROCESS_COUNT_COSTOM_DEFAULT = 1000
End Property

' コミット確認　しない
Public Property Get COMMIT_CONFIRM_NO() As Long

    COMMIT_CONFIRM_NO = 1
End Property

' コミット確認　する
Public Property Get COMMIT_CONFIRM_YES() As Long

    COMMIT_CONFIRM_YES = 2
End Property

' SQLエラー　中断
Public Property Get SQL_ERROR_HANDLING_SUSPEND() As Long

    SQL_ERROR_HANDLING_SUSPEND = 1
End Property

' SQLエラー　継続
Public Property Get SQL_ERROR_HANDLING_RESUME() As Long

    SQL_ERROR_HANDLING_RESUME = 2
End Property

' 空白セル読み取り方式（削除）
Public Property Get EMPTY_CELL_READING_DEL() As Long

    EMPTY_CELL_READING_DEL = 1
End Property

' 空白セル読み取り方式（削除しない）
Public Property Get EMPTY_CELL_READING_NON_DEL() As Long

    EMPTY_CELL_READING_NON_DEL = 2
End Property

' 空白セル読み取り方式（削除しない…文字型の場合は空文字列）
Public Property Get EMPTY_CELL_READING_NON_DEL_STR_EMPTY() As Long

    EMPTY_CELL_READING_NON_DEL_STR_EMPTY = 3
End Property

' 直接入力文字（無効）
Public Property Get DIRECT_INPUT_CHAR_DISABLE() As Long

    DIRECT_INPUT_CHAR_DISABLE = 1
End Property

' 直接入力文字（ユーザ指定）
Public Property Get DIRECT_INPUT_CHAR_ENABLE_CUSTOM() As Long

    DIRECT_INPUT_CHAR_ENABLE_CUSTOM = 3
End Property

' 直接入力文字（デフォルト）接頭辞
Public Property Get DIRECT_INPUT_CHAR_ENABLE_CUSTOM_PREFIX_DEFAULT() As String

    DIRECT_INPUT_CHAR_ENABLE_CUSTOM_PREFIX_DEFAULT = "#"
End Property

' 直接入力文字（デフォルト）接尾辞
Public Property Get DIRECT_INPUT_CHAR_ENABLE_CUSTOM_SUFFIX_DEFAULT() As String

    DIRECT_INPUT_CHAR_ENABLE_CUSTOM_SUFFIX_DEFAULT = "#"
End Property

' クエリパラメータの囲み文字（デフォルト）接頭辞
Public Property Get QUERY_PARAMETER_ENCLOSE_DEFAULT_PREFIX() As String

    QUERY_PARAMETER_ENCLOSE_DEFAULT_PREFIX = "{"
End Property

' クエリパラメータの囲み文字（デフォルト）接尾辞
Public Property Get QUERY_PARAMETER_ENCLOSE_DEFAULT_SUFFIX() As String

    QUERY_PARAMETER_ENCLOSE_DEFAULT_SUFFIX = "}"
End Property

' NULL入力文字（無効）
Public Property Get NULL_INPUT_CHAR_DISABLE() As Long

    NULL_INPUT_CHAR_DISABLE = 1
End Property

' NULL入力文字（ユーザ指定）
Public Property Get NULL_INPUT_CHAR_ENABLE_CUSTOM() As Long

    NULL_INPUT_CHAR_ENABLE_CUSTOM = 3
End Property

' NULL入力文字（デフォルト）
Public Property Get NULL_INPUT_CHAR_ENABLE_CUSTOM_DEFAULT() As String

    NULL_INPUT_CHAR_ENABLE_CUSTOM_DEFAULT = "<NULL>"
End Property

' スキーマ　単体
Public Property Get SCHEMA_USE_ONE() As Long

    SCHEMA_USE_ONE = 1
End Property

' スキーマ　複数
Public Property Get SCHEMA_USE_MULTIPLE() As Long

    SCHEMA_USE_MULTIPLE = 2
End Property

' セル書式 フォント名
Public Property Get CELL_FONT_NAME_DEFAULT() As String

    CELL_FONT_NAME_DEFAULT = Application.StandardFont
End Property

' セル書式 フォントサイズ
Public Property Get CELL_FONT_SIZE_DEFAULT() As Long
    
    CELL_FONT_SIZE_DEFAULT = Application.StandardFontSize
End Property

' セル書式 折り返し
Public Property Get CELL_WORDWRAP_DEFAULT() As Boolean
    
    CELL_WORDWRAP_DEFAULT = True
End Property

' セル書式 幅
Public Property Get CELL_WIDTH_DEFAULT() As Double
    
    CELL_WIDTH_DEFAULT = 16
End Property

' セル書式 高さ
Public Property Get CELL_HEIGHT_DEFAULT() As Double
    
    CELL_HEIGHT_DEFAULT = 15
End Property

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    recProcessCount = REC_PROCESS_COUNT_ALL
    recProcessCountCustom = REC_PROCESS_COUNT_COSTOM_DEFAULT
    
    commitConfirm = COMMIT_CONFIRM_NO
    
    sqlErrorHandling = SQL_ERROR_HANDLING_SUSPEND
    
    schemaUse = SCHEMA_USE_ONE
    
    queryResultShowWhenNormal = False
    
    cellFontName = CELL_FONT_NAME_DEFAULT
    cellFontSize = CELL_FONT_SIZE_DEFAULT
    cellWordwrap = CELL_WORDWRAP_DEFAULT
    cellWidth = CELL_WIDTH_DEFAULT
    cellHeight = CELL_HEIGHT_DEFAULT
    lineHeightAutoAdjust = True
    
    emptyCellReading = EMPTY_CELL_READING_DEL
    
    directInputChar = DIRECT_INPUT_CHAR_DISABLE
    directInputCharCustomPrefix = DIRECT_INPUT_CHAR_ENABLE_CUSTOM_PREFIX_DEFAULT
    directInputCharCustomSuffix = DIRECT_INPUT_CHAR_ENABLE_CUSTOM_SUFFIX_DEFAULT
    
    queryParameterEncloseCustomPrefix = QUERY_PARAMETER_ENCLOSE_DEFAULT_PREFIX
    queryParameterEncloseCustomSuffix = QUERY_PARAMETER_ENCLOSE_DEFAULT_SUFFIX
    
    nullInputChar = NULL_INPUT_CHAR_DISABLE
    nullInputCharCustom = NULL_INPUT_CHAR_ENABLE_CUSTOM_DEFAULT
    
    selectCheckCellMaxLength = True
    
    tableColumnEscapeOracle = True
    tableColumnEscapeMysql = True
    tableColumnEscapePostgresql = True
    tableColumnEscapeSqlserver = True
    tableColumnEscapeAccess = True
    tableColumnEscapeSymfoware = True
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ▽初期化メソッド
' =========================================================
Public Sub init(ByRef book_ As Workbook)

    If book Is Nothing Then
        ' 初回呼び出し
        Set book = book_
        readForData
    ElseIf Not book Is book_ Then
        ' ブックオブジェクトが異なるのでリロード
        Set book = book_
        readForData
    Else
        ' ブックオブジェクトが同じなのでリロードしない
    End If
    
End Sub

' =========================================================
' ▽レコード処理単位を取得する。
'
' 概要　　　：処理単位を取得する。全ての場合は、0を返す。
' 引数　　　：
' 戻り値　　：レコード処理単位
'
' =========================================================
Public Function getRecProcessCount() As Long

    ' レコード処理単位
    Select Case recProcessCount
    
        ' レコード処理単位（全て）
        Case REC_PROCESS_COUNT_ALL
            getRecProcessCount = 0
                    
        ' レコード処理単位（ユーザ指定）
        Case REC_PROCESS_COUNT_COSTOM
            getRecProcessCount = recProcessCountCustom
        
    End Select

End Function

' =========================================================
' ▽直接入力文字接頭辞を取得する。
'
' 概要　　　：直接入力文字接頭辞を取得する。無効の場合は、""を返す。
' 引数　　　：
' 戻り値　　：直接入力文字接頭辞
'
' =========================================================
Public Function getDirectInputCharPrefix() As String

    ' 直接入力文字
    Select Case directInputChar
    
        ' 無効の場合
        Case DIRECT_INPUT_CHAR_DISABLE
            getDirectInputCharPrefix = ""
                    
        ' 有効（ユーザ指定）の場合
        Case DIRECT_INPUT_CHAR_ENABLE_CUSTOM
            getDirectInputCharPrefix = directInputCharCustomPrefix
        
    End Select

End Function

' =========================================================
' ▽直接入力文字接尾辞を取得する。
'
' 概要　　　：直接入力文字接尾辞を取得する。無効の場合は、""を返す。
' 引数　　　：
' 戻り値　　：直接入力文字接尾辞
'
' =========================================================
Public Function getDirectInputCharSuffix() As String

    ' 直接入力文字
    Select Case directInputChar
    
        ' 無効の場合
        Case DIRECT_INPUT_CHAR_DISABLE
            getDirectInputCharSuffix = ""
                    
        ' 有効（ユーザ指定）の場合
        Case DIRECT_INPUT_CHAR_ENABLE_CUSTOM
            getDirectInputCharSuffix = directInputCharCustomSuffix
        
    End Select

End Function

' =========================================================
' ▽クエリパラメータの囲み文字接頭辞を取得する。
'
' 概要　　　：クエリパラメータの囲み文字接頭辞を取得する。
' 引数　　　：
' 戻り値　　：クエリパラメータの囲み文字接頭辞
'
' =========================================================
Public Function getQueryParameterEnclosePrefix() As String

    getQueryParameterEnclosePrefix = queryParameterEncloseCustomPrefix
End Function

' =========================================================
' ▽クエリパラメータの囲み文字接尾辞を取得する。
'
' 概要　　　：クエリパラメータの囲み文字接尾辞を取得する。
' 引数　　　：
' 戻り値　　：クエリパラメータの囲み文字接尾辞
'
' =========================================================
Public Function getQueryParameterEncloseSuffix() As String

    getQueryParameterEncloseSuffix = queryParameterEncloseCustomSuffix
End Function

' =========================================================
' ▽NULL入力文字を取得する。
'
' 概要　　　：NULL入力文字を取得する。無効の場合は、""を返す。
' 引数　　　：
' 戻り値　　：NULL入力文字
'
' =========================================================
Public Function getNullInputChar() As String

    ' NULL入力文字
    Select Case nullInputChar
    
        ' 無効の場合
        Case NULL_INPUT_CHAR_DISABLE
            getNullInputChar = ""
                    
        ' 有効（ユーザ指定）の場合
        Case NULL_INPUT_CHAR_ENABLE_CUSTOM
            getNullInputChar = nullInputCharCustom
        
    End Select

End Function

' =========================================================
' ▽設定情報の生成
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSetting.ini"

    Set createApplicationProperties = appProp
    
End Function

Private Function createApplicationPropertiesForBook() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initWorksheet book, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSetting.ini"

    Set createApplicationPropertiesForBook = appProp
    
End Function

' =========================================================
' ▽情報を読み込む
' =========================================================
Public Sub readForData()

    On Error GoTo err
        
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
        
    ' アプリケーションプロパティを生成する
    Dim appPropBook As ApplicationProperties
    Set appPropBook = createApplicationPropertiesForBook

    ' 一時変数
    Dim varRecProcessCount                      As String
    Dim varRecProcessCountCustom                As String
    Dim varCommitConfirm                        As String
    Dim varSqlErrorHandling                     As String
    Dim varSchemaUse                            As String
    Dim varQueryResultShowWhenNormal            As String
    
    Dim varCellFontName                         As String
    Dim varCellFontSize                         As String
    Dim varCellWordwrap                         As String
    Dim varCellWidth                            As String
    Dim varCellHeight                           As String
    Dim varLineHeightAutoAdjust                 As String
    
    Dim varEmptyCellReading                     As String
    Dim varDirectInputChar                      As String
    Dim varDirectInputCharCustomPrefix          As String
    Dim varDirectInputCharCustomSuffix          As String
    Dim varQueryParameterEncloseCustomPrefix    As String
    Dim varQueryParameterEncloseCustomSuffix    As String
    Dim varNullInputChar                        As String
    Dim varNullInputCharCustom                  As String
    Dim varSelectCheckCellMaxLength             As String
    
    Dim varTableColumnOracle                    As String
    Dim varTableColumnMysql                     As String
    Dim varTableColumnPostgresql                As String
    Dim varTableColumnSqlserver                 As String
    Dim varTableColumnAccess                    As String
    Dim varTableColumnSymfoware                 As String
    
    ' 情報を読み込む
    ' 読み込んだ情報のチェックを行い不正である場合、デフォルト値に変換する
    
    
    Dim values As ValCollection
    
    ' ---------------------------------------------------------
    ' グローバル
    ' ---------------------------------------------------------
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    Dim val As Variant
    
    val = values.getItem(KEY_REC_PROCESS_COUNT, vbVariant): If IsArray(val) Then varRecProcessCount = val(2)
    val = values.getItem(KEY_REC_PROCESS_COUNT_CUSTOM, vbVariant): If IsArray(val) Then varRecProcessCountCustom = val(2)
    val = values.getItem(KEY_COMMIT_CONFIRM, vbVariant): If IsArray(val) Then varCommitConfirm = val(2)
    val = values.getItem(KEY_SQL_ERROR_HANDLING, vbVariant): If IsArray(val) Then varSqlErrorHandling = val(2)
    val = values.getItem(KEY_SCHEMA_USE, vbVariant): If IsArray(val) Then varSchemaUse = val(2)
    val = values.getItem(KEY_QUERY_RESULT_SHOW_WHEN_NORMAL, vbVariant): If IsArray(val) Then varQueryResultShowWhenNormal = val(2)
    val = values.getItem(KEY_CELL_FONT_NAME, vbVariant): If IsArray(val) Then varCellFontName = val(2)
    val = values.getItem(KEY_CELL_FONT_SIZE, vbVariant): If IsArray(val) Then varCellFontSize = val(2)
    val = values.getItem(KEY_CELL_WORDWRAP, vbVariant): If IsArray(val) Then varCellWordwrap = val(2)
    val = values.getItem(KEY_CELL_WIDTH, vbVariant): If IsArray(val) Then varCellWidth = val(2)
    val = values.getItem(KEY_CELL_HEIGHT, vbVariant): If IsArray(val) Then varCellHeight = val(2)
    val = values.getItem(KEY_LINE_HEIGHT_AUTO_ADJUST, vbVariant): If IsArray(val) Then varLineHeightAutoAdjust = val(2)
    
    ' ---------------------------------------------------------
    ' ブック
    ' ---------------------------------------------------------
    Set values = appPropBook.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    val = values.getItem(KEY_EMPTY_CELL_READING, vbVariant): If IsArray(val) Then varEmptyCellReading = val(2)
    val = values.getItem(KEY_DIRECT_INPUT_CHAR, vbVariant): If IsArray(val) Then varDirectInputChar = val(2)
    val = values.getItem(KEY_DIRECT_INPUT_CHAR_CUSTOM_PREFIX, vbVariant): If IsArray(val) Then varDirectInputCharCustomPrefix = val(2)
    val = values.getItem(KEY_DIRECT_INPUT_CHAR_CUSTOM_SUFFIX, vbVariant): If IsArray(val) Then varDirectInputCharCustomSuffix = val(2)
    val = values.getItem(KEY_QUERY_PARAMETER_ENCLOSE_CUSTOM_PREFIX, vbVariant): If IsArray(val) Then varQueryParameterEncloseCustomPrefix = val(2)
    val = values.getItem(KEY_QUERY_PARAMETER_ENCLOSE_CUSTOM_SUFFIX, vbVariant): If IsArray(val) Then varQueryParameterEncloseCustomSuffix = val(2)
    val = values.getItem(KEY_NULL_INPUT_CHAR, vbVariant): If IsArray(val) Then varNullInputChar = val(2)
    val = values.getItem(KEY_NULL_INPUT_CHAR_CUSTOM, vbVariant): If IsArray(val) Then varNullInputCharCustom = val(2)
    val = values.getItem(KEY_SELECT_CHECK_CELL_MAX_LENGTH, vbVariant): If IsArray(val) Then varSelectCheckCellMaxLength = val(2)
    
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_ORACLE, vbVariant): If IsArray(val) Then varTableColumnOracle = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_MYSQL, vbVariant): If IsArray(val) Then varTableColumnMysql = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL, vbVariant): If IsArray(val) Then varTableColumnPostgresql = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_SQLSERVER, vbVariant): If IsArray(val) Then varTableColumnSqlserver = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_ACCESS, vbVariant): If IsArray(val) Then varTableColumnAccess = val(2)
    val = values.getItem(KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE, vbVariant): If IsArray(val) Then varTableColumnSymfoware = val(2)
    
    If validCode(varRecProcessCount _
                , REC_PROCESS_COUNT_ALL _
                , REC_PROCESS_COUNT_COSTOM) = False Then
    
        varRecProcessCount = REC_PROCESS_COUNT_ALL
    End If
    
    If validUnsignedInteger(varRecProcessCountCustom) = False Then
    
        varRecProcessCountCustom = REC_PROCESS_COUNT_COSTOM_DEFAULT
    End If
    
    If validCode(varCommitConfirm _
                , COMMIT_CONFIRM_NO _
                , COMMIT_CONFIRM_YES) = False Then
    
        varCommitConfirm = COMMIT_CONFIRM_NO
    End If
    
    If validCode(varSqlErrorHandling _
                , SQL_ERROR_HANDLING_SUSPEND _
                , SQL_ERROR_HANDLING_RESUME) = False Then
    
        varSqlErrorHandling = SQL_ERROR_HANDLING_SUSPEND
    End If
    
    If validCode(varSchemaUse _
                , SCHEMA_USE_ONE _
                , SCHEMA_USE_MULTIPLE) = False Then
    
        varSchemaUse = SCHEMA_USE_ONE
    End If
 
    If validCode(varQueryResultShowWhenNormal _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varQueryResultShowWhenNormal = False
    End If

    If Len(varCellFontName) = 0 Then
    
        varCellFontName = CELL_FONT_NAME_DEFAULT
    End If
 
    If validUnsignedNumeric(varCellFontSize) = False Then
    
        varCellFontSize = CELL_FONT_SIZE_DEFAULT
    End If
    
    If validCode(varCellWordwrap _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varCellWordwrap = CELL_WORDWRAP_DEFAULT
    End If
    
    If validUnsignedNumeric(varCellWidth) = False Then
    
        varCellWidth = CELL_WIDTH_DEFAULT
    End If
 
    If validUnsignedNumeric(varCellHeight) = False Then
    
        varCellHeight = CELL_HEIGHT_DEFAULT
    End If
 
    If validCode(varLineHeightAutoAdjust _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varLineHeightAutoAdjust = True
    End If
    
    If validCode(varEmptyCellReading _
                , EMPTY_CELL_READING_DEL _
                , EMPTY_CELL_READING_NON_DEL _
                , EMPTY_CELL_READING_NON_DEL_STR_EMPTY) = False Then
    
        varEmptyCellReading = EMPTY_CELL_READING_DEL
    End If
    
    If validCode(varDirectInputChar _
                , DIRECT_INPUT_CHAR_DISABLE _
                , DIRECT_INPUT_CHAR_ENABLE_CUSTOM) = False Then
    
        varDirectInputChar = DIRECT_INPUT_CHAR_DISABLE
    End If
    
    If Len(varDirectInputCharCustomPrefix) <> 1 Then
    
        varDirectInputCharCustomPrefix = DIRECT_INPUT_CHAR_ENABLE_CUSTOM_PREFIX_DEFAULT
    End If
    
    ' 接尾辞は未入力でもOK
    'If Len(varDirectInputCharCustomSuffix) <> 1 Then
    '
    '    varDirectInputCharCustomSuffix = DIRECT_INPUT_CHAR_ENABLE_CUSTOM_SUFFIX_DEFAULT
    'End If
    
    If Len(varQueryParameterEncloseCustomPrefix) <> 1 Then
    
        varQueryParameterEncloseCustomPrefix = QUERY_PARAMETER_ENCLOSE_DEFAULT_PREFIX
    End If
    
    If Len(varQueryParameterEncloseCustomSuffix) <> 1 Then
    
        varQueryParameterEncloseCustomSuffix = QUERY_PARAMETER_ENCLOSE_DEFAULT_SUFFIX
    End If
  
    If validCode(varNullInputChar _
                , NULL_INPUT_CHAR_DISABLE _
                , NULL_INPUT_CHAR_ENABLE_CUSTOM) = False Then
    
        varNullInputChar = NULL_INPUT_CHAR_DISABLE
    End If
    
    If Len(varNullInputCharCustom) <> 1 Then
    
        varNullInputCharCustom = NULL_INPUT_CHAR_ENABLE_CUSTOM_DEFAULT
    End If
    
    If validCode(varSelectCheckCellMaxLength _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varSelectCheckCellMaxLength = True
    End If
    
    If validCode(varTableColumnOracle _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnOracle = True
    End If
    
    If validCode(varTableColumnMysql _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnMysql = True
    End If
    
    If validCode(varTableColumnPostgresql _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnPostgresql = True
    End If
    
    If validCode(varTableColumnSqlserver _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnSqlserver = True
    End If
    
    If validCode(varTableColumnAccess _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnAccess = True
    End If
    
    If validCode(varTableColumnSymfoware _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnSymfoware = True
    End If
 
    recProcessCount = varRecProcessCount
    recProcessCountCustom = CLng(varRecProcessCountCustom)
    commitConfirm = varCommitConfirm
    sqlErrorHandling = varSqlErrorHandling
    schemaUse = varSchemaUse
    queryResultShowWhenNormal = varQueryResultShowWhenNormal
    
    cellFontName = varCellFontName
    cellFontSize = CDec(varCellFontSize)
    cellWordwrap = CBool(varCellWordwrap)
    cellWidth = CDec(varCellWidth)
    cellHeight = CDec(varCellHeight)
    lineHeightAutoAdjust = varLineHeightAutoAdjust
    
    emptyCellReading = varEmptyCellReading
    directInputChar = varDirectInputChar
    directInputCharCustomPrefix = varDirectInputCharCustomPrefix
    directInputCharCustomSuffix = varDirectInputCharCustomSuffix
    queryParameterEncloseCustomPrefix = varQueryParameterEncloseCustomPrefix
    queryParameterEncloseCustomSuffix = varQueryParameterEncloseCustomSuffix
    nullInputChar = varNullInputChar
    nullInputCharCustom = varNullInputCharCustom
    selectCheckCellMaxLength = CBool(varSelectCheckCellMaxLength)
    
    tableColumnEscapeOracle = varTableColumnOracle
    tableColumnEscapeMysql = varTableColumnMysql
    tableColumnEscapePostgresql = varTableColumnPostgresql
    tableColumnEscapeSqlserver = varTableColumnSqlserver
    tableColumnEscapeAccess = varTableColumnAccess
    tableColumnEscapeSymfoware = varTableColumnSymfoware
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を書き込む
' =========================================================
Public Sub writeForData()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' アプリケーションプロパティを生成する
    Dim appPropBook As ApplicationProperties
    Set appPropBook = createApplicationPropertiesForBook

    ' DB接続情報を格納する配列
    Dim values As ValCollection
    
    ' ---------------------------------------------------------
    ' グローバル
    ' ---------------------------------------------------------
    Set values = New ValCollection
    
    values.setItem Array(KEY_REC_PROCESS_COUNT, recProcessCount)
    values.setItem Array(KEY_REC_PROCESS_COUNT_CUSTOM, recProcessCountCustom)
    values.setItem Array(KEY_COMMIT_CONFIRM, commitConfirm)
    values.setItem Array(KEY_SQL_ERROR_HANDLING, sqlErrorHandling)
    values.setItem Array(KEY_SCHEMA_USE, schemaUse)
    values.setItem Array(KEY_QUERY_RESULT_SHOW_WHEN_NORMAL, queryResultShowWhenNormal)
    values.setItem Array(KEY_CELL_FONT_NAME, cellFontName)
    values.setItem Array(KEY_CELL_FONT_SIZE, cellFontSize)
    values.setItem Array(KEY_CELL_WORDWRAP, cellWordwrap)
    values.setItem Array(KEY_CELL_WIDTH, cellWidth)
    values.setItem Array(KEY_CELL_HEIGHT, cellHeight)
    values.setItem Array(KEY_LINE_HEIGHT_AUTO_ADJUST, lineHeightAutoAdjust)
                
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    ' ---------------------------------------------------------
    ' ブック
    ' ---------------------------------------------------------
    Set values = New ValCollection
    
    values.setItem Array(KEY_EMPTY_CELL_READING, emptyCellReading)
    values.setItem Array(KEY_DIRECT_INPUT_CHAR, directInputChar)
    values.setItem Array(KEY_DIRECT_INPUT_CHAR_CUSTOM_PREFIX, directInputCharCustomPrefix)
    values.setItem Array(KEY_DIRECT_INPUT_CHAR_CUSTOM_SUFFIX, directInputCharCustomSuffix)
    values.setItem Array(KEY_QUERY_PARAMETER_ENCLOSE_CUSTOM_PREFIX, queryParameterEncloseCustomPrefix)
    values.setItem Array(KEY_QUERY_PARAMETER_ENCLOSE_CUSTOM_SUFFIX, queryParameterEncloseCustomSuffix)
    values.setItem Array(KEY_NULL_INPUT_CHAR, nullInputChar)
    values.setItem Array(KEY_NULL_INPUT_CHAR_CUSTOM, nullInputCharCustom)
    values.setItem Array(KEY_SELECT_CHECK_CELL_MAX_LENGTH, selectCheckCellMaxLength)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_ORACLE, tableColumnEscapeOracle)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_MYSQL, tableColumnEscapeMysql)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL, tableColumnEscapePostgresql)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_SQLSERVER, tableColumnEscapeSqlserver)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_ACCESS, tableColumnEscapeAccess)
    values.setItem Array(KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE, tableColumnEscapeSymfoware)
                
    appPropBook.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appPropBook.writeData
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽DBに対応するカラムのエスケープ値を取得
' =========================================================
Public Function getTableColumnEscapeByDbConn(ByRef dbConn As Object) As Boolean

    ' DBMS種類を取得する
    Dim dmt As DbmsType
    dmt = ADOUtil.getDBMSType(dbConn)
    
    ' 戻り値を設定する
    getTableColumnEscapeByDbConn = getTableColumnEscape(dmt)
    
End Function

' =========================================================
' ▽DBに対応するカラムのエスケープ値を取得
' =========================================================
Public Function getTableColumnEscape(ByRef dbmsT As DbmsType) As Boolean

    ' 戻り値を設定する
    If dbmsT = DbmsType.Oracle Then
        getTableColumnEscape = tableColumnEscapeOracle
    
    ElseIf dbmsT = DbmsType.MySQL Then
        getTableColumnEscape = tableColumnEscapeMysql
    
    ElseIf dbmsT = DbmsType.PostgreSQL Then
        getTableColumnEscape = tableColumnEscapePostgresql
    
    ElseIf dbmsT = DbmsType.MicrosoftSqlServer Then
        getTableColumnEscape = tableColumnEscapeSqlserver
    
    ElseIf dbmsT = DbmsType.MicrosoftAccess Then
        getTableColumnEscape = tableColumnEscapeAccess
    
    ElseIf dbmsT = DbmsType.Symfoware Then
        getTableColumnEscape = tableColumnEscapeSymfoware
    
    Else
        getTableColumnEscape = True
    End If
    
End Function

