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
' 作成者　：Hideki Isobe
' 履歴　　：2008/03/14　新規作成
'
' 特記事項：
' *********************************************************

' =========================================================
' ▽メンバ
' =========================================================
' レコード処理単位
Public recProcessCount        As Long
' レコード処理単位（ユーザ指定）
Public recProcessCountCustom  As Long
' コミット確認
Public commitConfirm          As Long
' SQLエラー時の挙動
Public sqlErrorHandling       As Long
' 空白セル読み取り方式
Public emptyCellReading       As Long
' 直接入力文字
Public directInputChar        As Long
' 直接入力文字（ユーザ指定）
Public directInputCharCustom  As String
' 正常時のクエリ結果表示
Public queryResultShowWhenNormal As Boolean
' スキーマ利用
Public schemaUse              As Long

' テーブル・カラム名のエスケープ
Public tableColumnEscapeOracle     As Boolean
Public tableColumnEscapeMysql      As Boolean
Public tableColumnEscapePostgresql As Boolean
Public tableColumnEscapeSqlserver  As Boolean
Public tableColumnEscapeAccess  As Boolean
Public tableColumnEscapeSymfoware  As Boolean

' セル書式 フォント名
Public cellFontName As String
' セル書式 フォントサイズ
Public cellFontSize As Double
' セル書式 折り返し有無
Public cellWordwrap As Boolean
' セル書式 セル幅（単位は文字数）
Public cellWidth As Double
' セル書式 セル高さ（単位はポイント）
Public cellHeight As Double
' 行高の自動調整
Public lineHeightAutoAdjust      As Boolean

' =========================================================
' ▽定数プロパティ
' 　Classモジュールでは、Public Constが定義できないのでプロパティで代用する
' =========================================================
' キー値 レコード処理単位
Public Property Get REG_SUB_KEY_APPLICATION() As String

    REG_SUB_KEY_APPLICATION = "application_setting"
End Property

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

' キー値 直接入力文字
Public Property Get KEY_DIRECT_INPUT_CHAR() As String

    KEY_DIRECT_INPUT_CHAR = "directInputChar"
End Property

' キー値 直接入力文字　ユーザ指定
Public Property Get KEY_DIRECT_INPUT_CHAR_CUSTOM() As String

    KEY_DIRECT_INPUT_CHAR_CUSTOM = "directInputCharCustom"
End Property

' キー値 空白セル読み取り方式
Public Property Get KEY_EMPTY_CELL_READING() As String

    KEY_EMPTY_CELL_READING = "emptyCellReading"
End Property

' キー値 正常時のクエリ結果表示有無
Public Property Get KEY_QUERY_RESULT_SHOW_WHEN_NORMAL() As String

    KEY_QUERY_RESULT_SHOW_WHEN_NORMAL = "queryResultShowWhenNormal"
End Property

' キー値 スキーマ
Public Property Get KEY_SCHEMA_USE() As String

    KEY_SCHEMA_USE = "schemaUse"
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

' セル書式 フォント名
Public Property Get KEY_CELL_FONT_NAME() As String

    KEY_CELL_FONT_NAME = "cellFontName"
End Property

' セル書式 フォントサイズ
Public Property Get KEY_CELL_FONT_SIZE() As String
    
    KEY_CELL_FONT_SIZE = "cellFontSize"
End Property

' セル書式 折り返し
Public Property Get KEY_CELL_WORDWRAP() As String
    
    KEY_CELL_WORDWRAP = "cellWordwrap"
End Property

' セル書式 幅
Public Property Get KEY_CELL_WIDTH() As String
    
    KEY_CELL_WIDTH = "cellWidth"
End Property

' セル書式 高さ
Public Property Get KEY_CELL_HEIGHT() As String
    
    KEY_CELL_HEIGHT = "cellHeight"
End Property

' キー値 行高の自動調整
Public Property Get KEY_LINE_HEIGHT_AUTO_ADJUST() As String

    KEY_LINE_HEIGHT_AUTO_ADJUST = "lineHeightAutoAdjust"
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

' 直接入力文字（無効）
Public Property Get DIRECT_INPUT_CHAR_DISABLE() As Long

    DIRECT_INPUT_CHAR_DISABLE = 1
End Property

' 直接入力文字（ユーザ指定）
Public Property Get DIRECT_INPUT_CHAR_ENABLE_CUSTOM() As Long

    DIRECT_INPUT_CHAR_ENABLE_CUSTOM = 3
End Property

' 直接入力文字（デフォルト）
Public Property Get DIRECT_INPUT_CHAR_ENABLE_DEFAULT() As String

    DIRECT_INPUT_CHAR_ENABLE_DEFAULT = "$"
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
    
    emptyCellReading = EMPTY_CELL_READING_DEL
    
    directInputChar = DIRECT_INPUT_CHAR_DISABLE
    directInputCharCustom = DIRECT_INPUT_CHAR_ENABLE_DEFAULT
    
    queryResultShowWhenNormal = False
    
    schemaUse = SCHEMA_USE_ONE
    
    tableColumnEscapeOracle = True
    tableColumnEscapeMysql = True
    tableColumnEscapePostgresql = True
    tableColumnEscapeSqlserver = True
    tableColumnEscapeAccess = True
    tableColumnEscapeSymfoware = True
    
    cellFontName = CELL_FONT_NAME_DEFAULT
    cellFontSize = CELL_FONT_SIZE_DEFAULT
    cellWordwrap = CELL_WORDWRAP_DEFAULT
    cellWidth = CELL_WIDTH_DEFAULT
    cellHeight = CELL_HEIGHT_DEFAULT
    
    lineHeightAutoAdjust = True
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

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
' ▽直接入力文字を取得する。
'
' 概要　　　：直接入力文字を取得する。無効の場合は、""を返す。
' 引数　　　：
' 戻り値　　：直接入力文字
'
' =========================================================
Public Function getDirectInputChar() As String

    ' 直接入力文字
    Select Case directInputChar
    
        ' 無効の場合
        Case DIRECT_INPUT_CHAR_DISABLE
            getDirectInputChar = ""
                    
        ' 有効（ユーザ指定）の場合
        Case DIRECT_INPUT_CHAR_ENABLE_CUSTOM
            getDirectInputChar = directInputCharCustom
        
    End Select

End Function

' =========================================================
' ▽レジストリから情報を読み込む
' =========================================================
Public Sub readForRegistry()

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_APPLICATION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    ' 一時変数
    Dim varRecProcessCount          As String
    Dim varRecProcessCountCustom    As String
    Dim varCommitConfirm            As String
    Dim varSqlErrorHandling         As String
    Dim varEmptyCellReading         As String
    Dim varDirectInputChar          As String
    Dim varDirectInputCharCustom    As String
    Dim varSchemaUse                As String
    Dim varQueryResultShowWhenNormal As String
    
    Dim varTableColumnOracle        As String
    Dim varTableColumnMysql         As String
    Dim varTableColumnPostgresql    As String
    Dim varTableColumnSqlserver     As String
    Dim varTableColumnAccess        As String
    Dim varTableColumnSymfoware     As String
    
    Dim varCellFontName             As String
    Dim varCellFontSize             As String
    Dim varCellWordwrap             As String
    Dim varCellWidth                As String
    Dim varCellHeight               As String
    Dim varLineHeightAutoAdjust     As String
    
    ' レジストリから情報を読み込む
    ' 読み込んだ情報のチェックを行い不正である場合、デフォルト値に変換する
    registry.getValue KEY_REC_PROCESS_COUNT, varRecProcessCount
    If validCode(varRecProcessCount _
                , REC_PROCESS_COUNT_ALL _
                , REC_PROCESS_COUNT_COSTOM) = False Then
    
        varRecProcessCount = REC_PROCESS_COUNT_ALL
    End If
    
    registry.getValue KEY_REC_PROCESS_COUNT_CUSTOM, varRecProcessCountCustom
    If validUnsignedInteger(varRecProcessCountCustom) = False Then
    
        varRecProcessCountCustom = REC_PROCESS_COUNT_COSTOM_DEFAULT
    End If
    
    registry.getValue KEY_COMMIT_CONFIRM, varCommitConfirm
    If validCode(varCommitConfirm _
                , COMMIT_CONFIRM_NO _
                , COMMIT_CONFIRM_YES) = False Then
    
        varCommitConfirm = COMMIT_CONFIRM_NO
    End If
    
    registry.getValue KEY_SQL_ERROR_HANDLING, varSqlErrorHandling
    If validCode(varSqlErrorHandling _
                , SQL_ERROR_HANDLING_SUSPEND _
                , SQL_ERROR_HANDLING_RESUME) = False Then
    
        varSqlErrorHandling = SQL_ERROR_HANDLING_SUSPEND
    End If
    
    registry.getValue KEY_EMPTY_CELL_READING, varEmptyCellReading
    If validCode(varEmptyCellReading _
                , EMPTY_CELL_READING_DEL _
                , EMPTY_CELL_READING_NON_DEL) = False Then
    
        varEmptyCellReading = EMPTY_CELL_READING_DEL
    End If
    
    registry.getValue KEY_DIRECT_INPUT_CHAR, varDirectInputChar
    If validCode(varDirectInputChar _
                , DIRECT_INPUT_CHAR_DISABLE _
                , DIRECT_INPUT_CHAR_ENABLE_CUSTOM) = False Then
    
        varDirectInputChar = DIRECT_INPUT_CHAR_DISABLE
    End If
    
    registry.getValue KEY_DIRECT_INPUT_CHAR_CUSTOM, varDirectInputCharCustom
    If Len(varDirectInputCharCustom) <> 1 Then
    
        varDirectInputCharCustom = DIRECT_INPUT_CHAR_ENABLE_DEFAULT
    End If
 
    registry.getValue KEY_QUERY_RESULT_SHOW_WHEN_NORMAL, varQueryResultShowWhenNormal
    If validCode(varQueryResultShowWhenNormal _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varQueryResultShowWhenNormal = False
    End If
    
    registry.getValue KEY_SCHEMA_USE, varSchemaUse
    If validCode(varSchemaUse _
                , SCHEMA_USE_ONE _
                , SCHEMA_USE_MULTIPLE) = False Then
    
        varSchemaUse = SCHEMA_USE_ONE
    End If
 
    registry.getValue KEY_TABLE_COLUMN_ESCAPE_ORACLE, varTableColumnOracle
    If validCode(varTableColumnOracle _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnOracle = True
    End If
    
    registry.getValue KEY_TABLE_COLUMN_ESCAPE_MYSQL, varTableColumnMysql
    If validCode(varTableColumnMysql _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnMysql = True
    End If
    
    registry.getValue KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL, varTableColumnPostgresql
    If validCode(varTableColumnPostgresql _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnPostgresql = True
    End If
    
    registry.getValue KEY_TABLE_COLUMN_ESCAPE_SQLSERVER, varTableColumnSqlserver
    If validCode(varTableColumnSqlserver _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnSqlserver = True
    End If
    
    registry.getValue KEY_TABLE_COLUMN_ESCAPE_ACCESS, varTableColumnAccess
    If validCode(varTableColumnAccess _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnAccess = True
    End If
    
    registry.getValue KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE, varTableColumnSymfoware
    If validCode(varTableColumnSymfoware _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varTableColumnSymfoware = True
    End If
 
    registry.getValue KEY_CELL_FONT_NAME, varCellFontName
    If Len(varCellFontName) = 0 Then
    
        varCellFontName = CELL_FONT_NAME_DEFAULT
    End If
 
    registry.getValue KEY_CELL_FONT_SIZE, varCellFontSize
    If validUnsignedNumeric(varCellFontSize) = False Then
    
        varCellFontSize = CELL_FONT_SIZE_DEFAULT
    End If
    
    registry.getValue KEY_CELL_WORDWRAP, varCellWordwrap
    If validCode(varCellWordwrap _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varCellWordwrap = CELL_WORDWRAP_DEFAULT
    End If
    
    registry.getValue KEY_CELL_WIDTH, varCellWidth
    If validUnsignedNumeric(varCellWidth) = False Then
    
        varCellWidth = CELL_WIDTH_DEFAULT
    End If
 
    registry.getValue KEY_CELL_HEIGHT, varCellHeight
    If validUnsignedNumeric(varCellHeight) = False Then
    
        varCellHeight = CELL_HEIGHT_DEFAULT
    End If
 
    registry.getValue KEY_LINE_HEIGHT_AUTO_ADJUST, varLineHeightAutoAdjust
    If validCode(varLineHeightAutoAdjust _
               , CStr(True) _
               , CStr(False)) = False Then
               
        varLineHeightAutoAdjust = True
    End If
 
    recProcessCount = varRecProcessCount
    recProcessCountCustom = CLng(varRecProcessCountCustom)
    commitConfirm = varCommitConfirm
    sqlErrorHandling = varSqlErrorHandling
    emptyCellReading = varEmptyCellReading
    directInputChar = varDirectInputChar
    directInputCharCustom = varDirectInputCharCustom
    queryResultShowWhenNormal = varQueryResultShowWhenNormal
    schemaUse = varSchemaUse
    cellFontName = varCellFontName
    cellFontSize = CDbl(varCellFontSize)
    cellWordwrap = CBool(varCellWordwrap)
    cellWidth = CDbl(varCellWidth)
    cellHeight = CDbl(varCellHeight)
    lineHeightAutoAdjust = varLineHeightAutoAdjust
 
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽レジストリに情報を書き込む
' =========================================================
Public Sub writeForRegistry()

    On Error GoTo err
    
    ' DB接続情報を格納する配列
    Dim optionInfoArray(0 To 20 _
                      , 0 To 1) As Variant
    
    Dim i As Long
    
    i = 0
    
    optionInfoArray(i, 0) = KEY_REC_PROCESS_COUNT
    optionInfoArray(i, 1) = recProcessCount: i = i + 1
    optionInfoArray(i, 0) = KEY_REC_PROCESS_COUNT_CUSTOM
    optionInfoArray(i, 1) = recProcessCountCustom: i = i + 1
    
    optionInfoArray(i, 0) = KEY_COMMIT_CONFIRM
    optionInfoArray(i, 1) = commitConfirm: i = i + 1
    
    optionInfoArray(i, 0) = KEY_SQL_ERROR_HANDLING
    optionInfoArray(i, 1) = sqlErrorHandling: i = i + 1
    
    optionInfoArray(i, 0) = KEY_EMPTY_CELL_READING
    optionInfoArray(i, 1) = emptyCellReading: i = i + 1
    
    optionInfoArray(i, 0) = KEY_DIRECT_INPUT_CHAR
    optionInfoArray(i, 1) = directInputChar: i = i + 1
    optionInfoArray(i, 0) = KEY_DIRECT_INPUT_CHAR_CUSTOM
    optionInfoArray(i, 1) = directInputCharCustom: i = i + 1
    
    optionInfoArray(i, 0) = KEY_QUERY_RESULT_SHOW_WHEN_NORMAL
    optionInfoArray(i, 1) = queryResultShowWhenNormal: i = i + 1
    
    optionInfoArray(i, 0) = KEY_SCHEMA_USE
    optionInfoArray(i, 1) = schemaUse: i = i + 1
    
    optionInfoArray(i, 0) = KEY_CELL_FONT_NAME
    optionInfoArray(i, 1) = cellFontName: i = i + 1
    
    optionInfoArray(i, 0) = KEY_CELL_FONT_SIZE
    optionInfoArray(i, 1) = cellFontSize: i = i + 1
    
    optionInfoArray(i, 0) = KEY_CELL_WORDWRAP
    optionInfoArray(i, 1) = cellWordwrap: i = i + 1
    
    optionInfoArray(i, 0) = KEY_CELL_WIDTH
    optionInfoArray(i, 1) = cellWidth: i = i + 1
    
    optionInfoArray(i, 0) = KEY_CELL_HEIGHT
    optionInfoArray(i, 1) = cellHeight: i = i + 1
    
    optionInfoArray(i, 0) = KEY_LINE_HEIGHT_AUTO_ADJUST
    optionInfoArray(i, 1) = lineHeightAutoAdjust: i = i + 1
    
    optionInfoArray(i, 0) = KEY_TABLE_COLUMN_ESCAPE_ORACLE
    optionInfoArray(i, 1) = tableColumnEscapeOracle: i = i + 1
    
    optionInfoArray(i, 0) = KEY_TABLE_COLUMN_ESCAPE_MYSQL
    optionInfoArray(i, 1) = tableColumnEscapeMysql: i = i + 1
    
    optionInfoArray(i, 0) = KEY_TABLE_COLUMN_ESCAPE_POSTGRESQL
    optionInfoArray(i, 1) = tableColumnEscapePostgresql: i = i + 1
    
    optionInfoArray(i, 0) = KEY_TABLE_COLUMN_ESCAPE_SQLSERVER
    optionInfoArray(i, 1) = tableColumnEscapeSqlserver: i = i + 1
    
    optionInfoArray(i, 0) = KEY_TABLE_COLUMN_ESCAPE_ACCESS
    optionInfoArray(i, 1) = tableColumnEscapeAccess: i = i + 1
    
    optionInfoArray(i, 0) = KEY_TABLE_COLUMN_ESCAPE_SYMFOWARE
    optionInfoArray(i, 1) = tableColumnEscapeSymfoware: i = i + 1
    
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_APPLICATION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    registry.setValues optionInfoArray
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

Public Function getTableColumnEscapeByDbConn(ByRef dbConn As Object) As Boolean

    ' DBMS種類を取得する
    Dim dmt As DbmsType
    dmt = ADOUtil.getDBMSType(dbConn)
    
    ' 戻り値を設定する
    getTableColumnEscapeByDbConn = getTableColumnEscape(dmt)
    
End Function

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

