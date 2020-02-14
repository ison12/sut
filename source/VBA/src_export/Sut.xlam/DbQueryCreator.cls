VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbQueryCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのクエリを生成するインターフェース。
' 全DBで共通する汎用的な実装インターフェース。
'
' 作成者　：Ison
' 履歴　　：2009/03/28　新規作成
'
' 特記事項：
'
' *********************************************************

' SQLが定義されているシート名
Private Const SHEET_NAME_SQL_DEFINE                         As String = "template_sql"
' SELECTクエリのテンプレート
Private Const CELL_SQL_DEFINE_SELECT                        As String = "A2"
' INSERTクエリのテンプレート
Private Const CELL_SQL_DEFINE_INSERT                        As String = "A4"
' UPDATEクエリのテンプレート
Private Const CELL_SQL_DEFINE_UPDATE                        As String = "A6"
' DELETEクエリのテンプレート
Private Const CELL_SQL_DEFINE_DELETE                        As String = "A8"
' DELETE ALLクエリのテンプレート
Private Const CELL_SQL_DEFINE_DELETE_ALL                    As String = "A9"

' DBコネクション
Private m_dbConn As Object
' DBMS種類
Private m_dbms As DbmsType
' 空白文字読み取り方式
Private m_emptyValueReading As Long
' 直接入力文字
Private m_directInputChar As String
' カラム書式情報リスト
Private m_columnFormatList As ValCollection
' スキーマ使用有無
Private m_schemaUse As Long
' テーブル・カラム名のエスケープ
Private m_tableColumnEscape As Boolean


' アプリケーション設定
' 定数を利用するためにオブジェクトを生成
Private applicationSetting As ValApplicationSetting

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set applicationSetting = New ValApplicationSetting
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set applicationSetting = Nothing
    Set m_columnFormatList = Nothing
    Set m_dbConn = Nothing
End Sub

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
Public Property Let dbConn(ByVal RHS As Variant)

    ' DBコネクションを設定
    Set m_dbConn = RHS
    
    ' DBMS種類を設定
    m_dbms = ADOUtil.getDBMSType(m_dbConn)
    
End Property

' =========================================================
' ▽空白文字読み取り方式設定
'
' 概要　　　：空白文字読み取り方式を設定する。
' 引数　　　：vNewValue 空白文字読み取り方式
'
' 特記事項　：
'
'
' =========================================================
Public Property Let emptyValueReading(ByVal RHS As Long)

    m_emptyValueReading = RHS
End Property

' =========================================================
' ▽直接入力文字設定
'
' 概要　　　：直接入力文字を設定する。
' 引数　　　：vNewValue 直接入力文字
'
' 特記事項　：
'
'
' =========================================================
Public Property Let directInputChar(ByVal RHS As String)

    m_directInputChar = RHS
End Property

' =========================================================
' ▽カラム書式情報の設定
'
' 概要　　　：
' 引数　　　：RHS カラム書式情報オブジェクト
'
' 特記事項　：
'
'
' =========================================================
Public Property Let columnFormatList(ByVal RHS As Variant)

    Set m_columnFormatList = RHS
End Property

' =========================================================
' ▽スキーマ使用の設定
'
' 概要　　　：
' 引数　　　：vNewValue スキーマ使用有無
'
' 特記事項　：
'
'
' =========================================================
Public Property Let schemaUse(ByVal vNewValue As Variant)

    m_schemaUse = vNewValue
End Property

' =========================================================
' ▽テーブル・カラム名エスケープの設定
'
' 概要　　　：
' 引数　　　：RHS カラム書式情報オブジェクト
'
' 特記事項　：
'
'
' =========================================================
Public Property Let tableColumnEscape(ByVal vNewValue As Variant)

    m_tableColumnEscape = vNewValue
End Property

' =========================================================
' ▽INSERT文生成
'
' 概要　　　：INSERT SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　record レコード
'
' 戻り値　　：INSERT文
'
' 特記事項　：
'
'
' =========================================================
Public Function createInsert(ByRef table As ValDbDefineTable _
                           , ByRef record As Variant) As String

    ' 接続文字 AND
    Const CONNECT_AND              As String = "AND "
    ' 接続文字 カンマ
    Const CONNECT_COMMA            As String = "   ,"
    ' 接続文字 最初
    Const CONNECT_FIRST            As String = "    "
        
    ' SQL本体
    Dim statement         As New StringBuilder
    ' SQLカラム句
    Dim columnStatement   As New StringBuilder
    ' SQLバリュー句
    Dim valuesStatement   As New StringBuilder

    ' クエリー値（実際にクエリーに埋め込む値）
    Dim queryValue        As String

    ' カラムリスト
    Dim columnList      As ValCollection
    ' カラム
    Dim column          As ValDbDefineColumn
    
    ' カラム書式情報
    Dim columnFormat    As ValDbColumnTypeColInfo

    ' レコード値
    Dim recordValue     As String
    
    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf

    ' -------------------------------------------------
    ' テンプレートSQLステートメントを取得
    ' -------------------------------------------------
    statement.append SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_INSERT)


    ' -------------------------------------------------
    ' SQLを生成する
    ' -------------------------------------------------
    statement.replace "${tableName}", DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)

    
    ' -------------------------------------------------
    ' レコードからクエリーを生成する
    ' -------------------------------------------------
    Dim i         As Long
    Dim j         As Long: j = 1
    Dim arrayMin  As Long   ' 配列の添え字（最小値）
    Dim arrayMax  As Long   ' 配列の添え字（最大値）
    Dim arraySize As Long   ' 配列のサイズ
    
    Set columnList = table.columnList
    
    arrayMin = LBound(record)
    arrayMax = UBound(record)
    
    arraySize = VBUtil.arraySize(record)
    
    For i = arrayMin To arrayMax
    
        Set column = columnList.getItemByIndex(j)
    
        ' カラム書式情報を取得する
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        ' カラム書式情報が取得できなかった場合
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " 定義情報なし"
            #End If
            
            ' デフォルトカラム書式情報を取得する
            Set columnFormat = defaultColumnFormat
            
        ' カラム情報が取得できた場合
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatUpdate
            #End If
        
        End If
    
        If columnFormat.formatUpdate = COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print "Excluded " & column.dataType & " from a query."
            #End If
            
            ' スキップする
            GoTo continue
        End If
    
        ' レコード値を取得
        recordValue = CStr(record(i))
        
        ' 空白文字読み取り方式が"カラム削除"の場合
        If m_emptyValueReading = applicationSetting.EMPTY_CELL_READING_DEL Then
        
            ' 値が設定されている場合のみステートメントに追加する
            If recordValue <> "" Then
    
                queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, recordValue, True, m_directInputChar)
                ' COLUMN句
                columnStatement.append(newline).append(CONNECT_COMMA).append (DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
                ' VALUES句
                valuesStatement.append(newline).append(CONNECT_COMMA).append (queryValue)
    
            End If
            
        ' 空白文字読み取り方式が"カラムを削除しない"の場合
        Else
        
            ' 常にステートメントを追加する
            
            ' 空文字の場合、NULLに変換
            queryValue = DBUtil.convertEmptyToNull(recordValue)
            queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, queryValue, True, m_directInputChar)
    
            ' COLUMN句
            columnStatement.append(newline).append(CONNECT_COMMA).append (DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
            ' VALUES句
            valuesStatement.append(newline).append(CONNECT_COMMA).append (queryValue)
    
        End If

continue:
        j = j + 1
        
    Next
    
    ' COLUMN句に一つ以上カラムが設定されている場合
    If columnStatement.str <> "" Then

        ' 前方に付加された冗長な","と"\r\n"を取り除く
        columnStatement.remove 1, Len(newline & CONNECT_COMMA)
        columnStatement.insert 1, CONNECT_FIRST
                                
        ' 前方に付加された冗長な","と"\r\n"を取り除く
        valuesStatement.remove 1, Len(newline & CONNECT_COMMA)
        valuesStatement.insert 1, CONNECT_FIRST

        ' カラム指定部分を埋め込む
        statement.replace "${columns}", columnStatement.str
        ' 値指定部分を埋め込む
        statement.replace "${values}", valuesStatement.str

        createInsert = statement.str

    End If

End Function

' =========================================================
' ▽UPDATE文生成
'
' 概要　　　：UPDATE SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　record レコード
'
' 戻り値　　：UPDATE文
'
' 特記事項　：
'
'
' =========================================================
Public Function createUpdate(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' 接続文字 AND
    Const CONNECT_AND              As String = "AND "
    ' 接続文字 カンマ
    Const CONNECT_COMMA            As String = "   ,"
    ' 接続文字 最初
    Const CONNECT_FIRST            As String = "    "

    ' SQL本体
    Dim statement             As New StringBuilder
    ' SQLカラム値句
    Dim columnValuesStatement As New StringBuilder

    ' SQL条件句
    Dim conditionStatement    As New StringBuilder

    ' クエリー値（バリューをクエリー用の値に変換）
    Dim queryValue        As String

    ' カラムリスト
    Dim columnList      As ValCollection
    ' カラム
    Dim column          As ValDbDefineColumn
    ' カラム書式情報
    Dim columnFormat    As ValDbColumnTypeColInfo
    ' テーブル制約情報(PK)
    Dim tableConstPk    As New ValDbDefineTableConstraints
    ' PKカラムであるかをあらわすフラグ
    Dim isColumnPk      As Boolean
    
    Dim tableConstTmp   As ValDbDefineTableConstraints
    ' テーブル制約リストからPK制約を取得する
    For Each tableConstTmp In table.constraintsList.col
    
        If tableConstTmp.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk Then
        
            Set tableConstPk = tableConstTmp
            Exit For
        End If
    Next
    
    ' レコード値
    Dim recordValue     As String

    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf

    ' -------------------------------------------------
    ' テンプレートSQLステートメントを取得
    ' -------------------------------------------------
    statement.append SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_UPDATE)

    ' -------------------------------------------------
    ' SQLを生成する
    ' -------------------------------------------------
    statement.replace "${tableName}", DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)
    
    ' -------------------------------------------------
    ' レコードからクエリーを生成する
    ' -------------------------------------------------
    Dim i         As Long
    Dim j         As Long: j = 1
    Dim arrayMin  As Long   ' 配列の添え字（最小値）
    Dim arrayMax  As Long   ' 配列の添え字（最大値）
    Dim arraySize As Long   ' 配列のサイズ
    
    Set columnList = table.columnList
    
    arrayMin = LBound(record)
    arrayMax = UBound(record)
    
    arraySize = VBUtil.arraySize(record)
    
    For i = arrayMin To arrayMax
    
        Set column = columnList.getItemByIndex(j)
        
        ' カラム書式情報を取得する
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        ' カラム書式情報が取得できなかった場合
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print "Format " & column.dataType & " is not found."
            #End If
            
            ' デフォルトカラム書式情報を取得する
            Set columnFormat = defaultColumnFormat
            
        ' カラム情報が取得できた場合
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatUpdate
            #End If
        
        End If
    
        If columnFormat.formatUpdate = COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print "Excluded " & column.dataType & " from a query."
            #End If
            
            ' スキップする
            GoTo continue
        End If
        
        ' レコード値を取得
        recordValue = CStr(record(i))
        
        ' PK制約であるかどうかを判定する
        If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
        
            isColumnPk = False
        Else
        
            isColumnPk = True
        End If
        
        ' PKではない場合
        If isColumnPk = False Then
        
            ' 空白文字読み取り方式が"カラム削除"の場合
            If m_emptyValueReading = applicationSetting.EMPTY_CELL_READING_DEL Then
            
                ' 値が設定されている場合のみステートメントに追加する
                If recordValue <> "" Then
        
                    queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, recordValue, True, m_directInputChar)
    
                    ' COLUMN句
                    columnValuesStatement.append(newline).append(CONNECT_COMMA).append(DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)).append(" = ").append (queryValue)
        
                End If
                
            ' 空白文字読み取り方式が"カラムを削除しない"の場合
            Else
            
                ' 常にステートメントを追加する
                
                ' 空文字の場合、NULLに変換
                queryValue = DBUtil.convertEmptyToNull(recordValue)
                queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, queryValue, True, m_directInputChar)
        
                ' COLUMN句
                columnValuesStatement.append(newline).append(CONNECT_COMMA).append(DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)).append(" = ").append (queryValue)
        
            End If
            
        End If
        
        ' 複合キーを考慮してPKの場合は、入力値があろうがなかろうが常にステートメントとして追加する。
        ' 例えば、column1・column2がPKの場合に、column1に入力がありcolumn2に入力がない場合に
        ' column1のみ条件として追加してしまうと、レコードを一意に特定できずに複数件更新されてしまう恐れがあるため。
        If isColumnPk = True Then

            queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, recordValue, True, m_directInputChar)
            
            conditionStatement.append(newline).append(CONNECT_AND).append(DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)).append(" = ").append (queryValue)
        End If

continue:
        j = j + 1
    Next
    
    ' COLUMN句またはWHERE句に一つ以上カラムが設定されている場合
    If conditionStatement.str <> "" And columnValuesStatement.str <> "" Then

        ' 前方に付加された冗長な","と改行コードを取り除く
        columnValuesStatement.remove 1, Len(newline & CONNECT_COMMA)
        columnValuesStatement.insert 1, CONNECT_FIRST
                                
        ' 前方に付加された冗長な"AND "と改行コードを取り除く
        conditionStatement.remove 1, Len(newline & CONNECT_AND)
        conditionStatement.insert 1, CONNECT_FIRST
                                
        ' カラム指定部分を埋め込む
        statement.replace "${columnValues}", columnValuesStatement.str
        ' 値指定部分を埋め込む
        statement.replace "${condition}", conditionStatement.str

        createUpdate = statement.str

    End If

End Function

' =========================================================
' ▽DELETE文生成
'
' 概要　　　：DELETE SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　record レコード
'
' 戻り値　　：DELETE文
'
' 特記事項　：
'
'
' =========================================================
Public Function createDelete(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' 接続文字 AND
    Const CONNECT_AND              As String = "AND "
    ' 接続文字 カンマ
    Const CONNECT_COMMA            As String = "   ,"
    ' 接続文字 最初
    Const CONNECT_FIRST            As String = "    "

    ' SQL本体
    Dim statement            As New StringBuilder
    ' SQL条件句
    Dim conditionStatement   As New StringBuilder

    ' クエリー値（実際にクエリーに埋め込む値）
    Dim queryValue        As String

    ' カラムリスト
    Dim columnList      As ValCollection
    ' カラム
    Dim column          As ValDbDefineColumn
    ' カラム書式情報
    Dim columnFormat    As ValDbColumnTypeColInfo
    
    ' テーブル制約情報(PK)
    Dim tableConstPk    As New ValDbDefineTableConstraints
    ' PKカラムであるかをあらわすフラグ
    Dim isColumnPk      As Boolean
    
    Dim tableConstTmp   As ValDbDefineTableConstraints
    ' テーブル制約リストからPK制約を取得する
    For Each tableConstTmp In table.constraintsList.col
    
        If tableConstTmp.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk Then
        
            Set tableConstPk = tableConstTmp
            Exit For
        End If
    Next
    
    ' レコード値
    Dim recordValue     As String

    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf

    ' -------------------------------------------------
    ' テンプレートSQLステートメントを取得
    ' -------------------------------------------------
    statement.append SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_DELETE)

    ' -------------------------------------------------
    ' SQLを生成する
    ' -------------------------------------------------
    statement.replace "${tableName}", DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)

    ' -------------------------------------------------
    ' レコードからクエリーを生成する
    ' -------------------------------------------------
    Dim i         As Long
    Dim j         As Long: j = 1
    Dim arrayMin  As Long   ' 配列の添え字（最小値）
    Dim arrayMax  As Long   ' 配列の添え字（最大値）
    Dim arraySize As Long   ' 配列のサイズ
    
    Set columnList = table.columnList
    
    arrayMin = LBound(record)
    arrayMax = UBound(record)
    
    arraySize = VBUtil.arraySize(record)
    
    For i = arrayMin To arrayMax
    
        Set column = columnList.getItemByIndex(j)
    
        ' カラム書式情報を取得する
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        ' カラム書式情報が取得できなかった場合
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " 定義情報なし"
            #End If
            
            ' デフォルトカラム書式情報を取得する
            Set columnFormat = defaultColumnFormat
        
        ' カラム情報が取得できた場合
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatUpdate
            #End If
        
        End If
        
        If columnFormat.formatUpdate = COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print "Excluded " & column.dataType & " from a query."
            #End If
            
            ' スキップする
            GoTo continue
        End If
        
        ' レコード値を取得
        recordValue = CStr(record(i))
        
        ' PK制約であるかどうかを判定する
        If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
        
            isColumnPk = False
        Else
        
            isColumnPk = True
        End If
        

        ' 複合キーを考慮してPKの場合は、入力値があろうがなかろうが常にステートメントとして追加する。
        ' 例えば、column1・column2がPKの場合に、column1に入力がありcolumn2に入力がない場合に
        ' column1のみ条件として追加してしまうと、レコードを一意に特定できずに複数件更新されてしまう恐れがあるため。
        If isColumnPk = True Then

            queryValue = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, recordValue, True, m_directInputChar)
            
            conditionStatement.append(newline).append(CONNECT_AND).append(DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)).append(" = ").append (queryValue)
        End If
        
continue:
        j = j + 1
        
    Next
    
    ' WHERE句に一つ以上カラムが設定されている場合
    If conditionStatement.str <> "" Then

        ' 前方に付加された冗長な"AND "と"\r\n"を取り除く
        conditionStatement.remove 1, Len(newline & CONNECT_AND)
        conditionStatement.insert 1, CONNECT_FIRST
                                
        ' COLUMN = VALUE句
        statement.replace "${condition}", conditionStatement.str

        ' 全てのステートメントを結合して、戻り値として格納する
        createDelete = statement.str

    End If

End Function

' =========================================================
' ▽DELETE文生成（テーブル上の全レコードを削除）
'
' 概要　　　：DELETE SQLを生成する
' 引数　　　：table  テーブル定義
'
' 戻り値　　：DELETE文
'
' 特記事項　：
'
'
' =========================================================
Public Function createDeleteOfAllTable(ByRef table As ValDbDefineTable) As String

    ' SQL本体
    Dim statement As New StringBuilder
    
    ' -------------------------------------------------
    ' テンプレートSQLステートメントを取得
    ' -------------------------------------------------
    statement.append SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_DELETE_ALL)

    ' -------------------------------------------------
    ' SQLを生成する
    ' -------------------------------------------------
    statement.replace "${tableName}", DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)

    ' 戻り値として格納する
    createDeleteOfAllTable = statement.str

End Function

' =========================================================
' ▽SELECTのカラム部分のステートメントを生成する
'
' 概要　　　：
' 引数　　　：table   テーブル定義
' 　　　　　　ident   インデント
' 　　　　　　convert 変換処理
'
' 戻り値　　：SELECTのカラム部分のステートメント
'
' 特記事項　：
'
'
' =========================================================
Public Function createSelectForColumn(ByRef table As ValDbDefineTable _
                                    , Optional ByVal INDENT As String = "" _
                                    , Optional convert As Boolean = True _
                                    , Optional ByVal identifier As String = "") As String

    ' 戻り値
    Dim ret As StringBuilder
    Set ret = New StringBuilder

    ' カラムリスト
    Dim columnList      As ValCollection
    ' カラム
    Dim column          As ValDbDefineColumn
    ' カラム書式情報
    Dim columnFormat    As ValDbColumnTypeColInfo
    ' カラム名
    Dim columnName      As New StringBuilder
    
    ' カラムリストを取得
    Set columnList = table.columnList
    
    ' カラムの数だけループする
    For Each column In columnList.col

        ' カラム書式情報を取得する
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        
        ' カラム書式情報が取得できなかった場合
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " 定義情報なし"
            #End If
            
            ' デフォルトカラム書式情報を取得する
            Set columnFormat = defaultColumnFormat
            
        ' カラム情報が取得できた場合
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatSelect
            #End If
        
        End If
        
        columnName.clear
        
        If columnFormat.formatSelect = COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE Then
        
            columnName.append("'Excluded ").append(column.dataType).append(".' AS ").append (DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
            
            #If DEBUG_MODE = 1 Then
                Debug.Print "Excluded " & column.dataType & " from a query."
            #End If
            
        ' カラム名の変換処理実施有無の判定
        ElseIf convert = True Then
        
            ' カラム名の変換処理実行
            columnName.append convertSelectFormat(m_dbms, columnFormat.formatSelect, identifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
            
        Else
        
            ' 変換しない場合、カラム名をそのまま設定
            columnName.append identifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape)
        End If
        
        ' COLUMN句
        ret.append(vbLf).append(INDENT).append("   ,").append (columnName.str)
    Next

    ' 前方に付加された冗長な","と"\n"を取り除く
    ret.remove 1, Len(vbLf & INDENT & "   ,")
    ret.insert 1, INDENT & "    "
                            
    ' 戻り値を設定する
    createSelectForColumn = ret.str
    
End Function

' =========================================================
' ▽SELECT文ののステートメントを生成する
'
' 概要　　　：
' 引数　　　：table            入力　テーブル定義
' 　　　　　　condition        入力　SELECT条件
' 　　　　　　statementWhere   出力　WHEREステートメント
' 　　　　　　statementOrderby 出力　ORDER BYステートメント
'
' 特記事項　：
'
'
' =========================================================
Public Sub createSelectForWhereOrderby(ByVal table As ValDbDefineTable _
                                     , ByVal condition As ValSelectCondition _
                                     , ByRef statementWhere As String _
                                     , ByRef statementOrderby As String _
                                     , Optional ByVal identifier As String = "")

    ' 接続文字 最初
    Const CONNECT_FIRST            As String = "     "
    ' 接続文字 AND
    Const CONNECT_AND              As String = " AND "
    ' 接続文字 カンマ
    Const CONNECT_COMMA            As String = "    ,"
    ' 接続文字 代入式（＝）
    Const CONNECT_LEFT_RIGHT_EQUAL As String = " = "
    ' 接続文字 代入式（LIKE）
    Const CONNECT_LEFT_RIGHT_LIKE  As String = " LIKE "
    ' 接続文字 代入式（IS）
    Const CONNECT_LEFT_RIGHT_IS     As String = " IS "

    ' WHERE句の内容
    Dim where       As New StringBuilder
    ' ORDER BY句の内容
    Dim order       As New StringBuilder
    
    ' カラムリスト
    Dim columnList As ValCollection
    ' カラム名
    Dim columnName As Variant
    ' カラムオブジェクト
    Dim column     As ValDbDefineColumn
    ' カラム書式情報
    Dim columnFormat    As ValDbColumnTypeColInfo
    
    ' 値
    Dim value      As String
    ' 順序
    Dim orderByVal As Variant
    
    Dim connectLeftRight As String
    
    
    ' カラムリストを取得する
    Set columnList = condition.getColumnList
    
    Dim i As Long
    
    ' コントロール配列を1件ずつ処理する
    For Each columnName In columnList.col
    
        ' カラムオブジェクトを取得する
        Set column = table.columnList.getItem(columnName)
        
        ' カラムオブジェクトが取得できなかった場合、次のループへ
        If column Is Nothing Then
        
            GoTo continue
        End If
    
        ' カラム書式情報を取得する
        Set columnFormat = m_columnFormatList.getItem(UCase$(column.dataType))
        ' カラム書式情報が取得できなかった場合
        If columnFormat Is Nothing Then
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " 定義情報なし"
            #End If
            
            ' デフォルトカラム書式情報を取得する
            Set columnFormat = defaultColumnFormat
            
        ' カラム情報が取得できた場合
        Else
        
            #If DEBUG_MODE = 1 Then
                Debug.Print column.dataType & " : " & columnFormat.formatSelect
            #End If
        
        End If
        
        ' 値を取得する
        value = condition.getConditionValue(columnName)
        ' 並び順を取得する
        orderByVal = condition.getConditionOrder(columnName)
    
        ' 値が設定されている場合
        If value <> "" Then
        
            ' NULLまたは
            ' NOT NULLと入力された場合は、IS NULLで引っ掛けるようにする
            If _
                Trim$(UCase$(value)) = "NULL" _
             Or Trim$(UCase$(value)) = "NOT NULL" Then
            
                where.append(vbLf).append(CONNECT_AND).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append(CONNECT_LEFT_RIGHT_IS).append (Trim$(UCase$(value)))
                
            ' 上記以外の場合
            Else
            
                ' カラムのリテラル種類がStringの場合で且つ、LIKE関数のパターン文字が含まれている場合
                ' LIKE関数を利用すると判断する。
                If ( _
                     InStr(value, "%") <> 0 Or _
                    (InStr(value, "_") <> 0 And Not (InStr(value, ConstantsTable.QUERY_PARAMETER_ENCLOSE_START) = 1 And InStr(value, ConstantsTable.QUERY_PARAMETER_ENCLOSE_END) = Len(value)))) Then
                
                    ' シングルクォートをエスケープ
                    value = DBUtil.escapeValueForSinglequart(m_dbms, value)
                    ' 文字列の両端にシングルクォートを付加する
                    value = DBUtil.escapeValueForEscapeChar(m_dbms, value)
                
                    where.append(vbLf).append(CONNECT_AND).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append(CONNECT_LEFT_RIGHT_LIKE).append (DBUtil.encloseSingleQuart(value))
                    
                    If _
                        m_dbms = DbmsType.MySQL _
                     Or m_dbms = DbmsType.PostgreSQL Then
                     
                        ' MySQLやPostgreSQLは、"\"が既にエスケープ文字なので、別の文字をエスケープ文字にしてあげる
                        where.append (DBUtil.addLikeEscape(m_dbms, "#"))
                        
                    Else
                    
                        where.append (DBUtil.addLikeEscape(m_dbms))
                    End If
                Else
                
                    ' 値を変換する。変換する場合はエスケープ文字はそのままにする
                    ' （DBUtil.convertUpdateFormatの第５引数はFalseに設定。※\→\\ に変換したりしない。
                    ' 　また第６引数の直接入力文字は有効にしない。）
                    ' 値を変換する。型特有のフォーマットに変換する
                    value = convertUpdateFormat(m_dbms, columnFormat.formatUpdate, value, False)
                    
                    where.append(vbLf).append(CONNECT_AND).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append(CONNECT_LEFT_RIGHT_EQUAL).append (value)
                
                End If
            
            End If
        
        End If
        
        ' 並び順に昇順か降順が設定されている場合
        If orderByVal <> condition.ORDER_NONE Then
        
            ' ORDER BY句の一部を生成する
            If orderByVal = condition.ORDER_ASC Then
            
                ' ASC
                order.append(vbLf).append(CONNECT_COMMA).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append (" ASC")
            ElseIf orderByVal = condition.ORDER_DESC Then
                
                ' DESC
                order.append(vbLf).append(CONNECT_COMMA).append(identifier & DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append (" DESC")
            End If
        
        End If
        
continue:

    Next
    
                  
    ' WHERE句が１文字でも生成されていたら"WHERE"を付加
    If where.str <> "" Then
    
        ' 前方に付加されている冗長な改行コードと連結文字列を除外する
        where.remove 1, Len(vbLf & CONNECT_AND)
        where.insert 1, CONNECT_FIRST
        where.insert 1, "WHERE" & vbLf
    End If
                  
    ' ORDER BY句が１文字でも生成されていたら"ORDER BY"を付加
    If order.str <> "" Then
    
        ' 前方に付加されている冗長な改行コードと連結文字列を除外する
        order.remove 1, Len(vbLf & CONNECT_COMMA)
        order.insert 1, CONNECT_FIRST
        order.insert 1, "ORDER BY" & vbLf
    End If
                  
    ' 戻り値を設定する
    statementWhere = where.str
    
    ' 戻り値を設定する
    statementOrderby = order.str
    
End Sub

' =========================================================
' ▽SELECT文ののステートメントを生成する
'
' 概要　　　：
' 引数　　　：table            入力　テーブル定義
' 　　　　　　condition        入力　SELECT条件
' 戻り値　　：ORDER BY句
'
' 特記事項　：
'
'
' =========================================================
Public Function createSelectForOrderby(ByVal table As ValDbDefineTable _
                                     , ByVal condition As ValSelectCondition) As String

    ' 接続文字 カンマ
    Const CONNECT_COMMA            As String = ", "

    ' ORDER BY句の内容
    Dim order       As New StringBuilder
    
    ' カラムリスト
    Dim columnList As ValCollection
    ' カラム名
    Dim columnName As Variant
    ' カラムオブジェクト
    Dim column     As ValDbDefineColumn
    
    ' 順序
    Dim orderByVal As Variant
    
    ' カラムリストを取得する
    Set columnList = condition.getColumnList
    
    ' コントロール配列を1件ずつ処理する
    For Each columnName In columnList.col
    
        ' カラムオブジェクトを取得する
        Set column = table.columnList.getItem(columnName)
        
        ' カラムオブジェクトが取得できなかった場合、次のループへ
        If column Is Nothing Then
        
            GoTo continue
        End If

        ' 並び順を取得する
        orderByVal = condition.getConditionOrder(columnName)
    
        ' 並び順に昇順か降順が設定されている場合
        If orderByVal <> condition.ORDER_NONE Then
        
            ' ORDER BY句の一部を生成する
            If orderByVal = condition.ORDER_ASC Then
            
                ' ASC
                order.append(CONNECT_COMMA).append(DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append (" ASC")
            ElseIf orderByVal = condition.ORDER_DESC Then
                
                ' DESC
                order.append(CONNECT_COMMA).append(DBUtil.escapeTableColumn(m_dbms, CStr(columnName), m_tableColumnEscape)).append (" DESC")
            End If
        
        End If
        
continue:

    Next
    
    ' ORDER BY句が１文字でも生成されていたら"ORDER BY"を付加
    If order.str <> "" Then
    
        ' 前方に付加されている冗長な改行コードと連結文字列を除外する
        order.remove 1, Len(CONNECT_COMMA)
        order.insert 1, "ORDER BY "
    End If
                  
    ' 戻り値を設定する
    createSelectForOrderby = order.str
    
End Function


' =========================================================
' ▽PKカラムの結合ステートメントを生成する
'
' 概要　　　：PKカラムの結合ステートメントを生成する
' 引数　　　：table           テーブル定義
' 　　　　　　leftIdentifier  左結合の修飾子
'             rightIdentifier 右結合の修飾子
'             rightStatementIsNull 右結合値を常に IS NULL とする
' 戻り値　　：PKカラムの結合ステートメント
'
' 特記事項　：
'
'
' =========================================================
Public Function createSelectForPkJoin(ByRef table As ValDbDefineTable _
                                    , ByVal leftIdentifier As String _
                                    , ByVal rightIdentifier As String _
                                    , Optional ByVal rightStatementIsNull = False) As String

    ' 接続文字 AND
    Const CONNECT_AND              As String = "AND "
    ' 接続文字 最初
    Const CONNECT_FIRST            As String = ""

    ' SQL条件句
    Dim conditionStatement    As New StringBuilder

    ' カラムリスト
    Dim columnList      As ValCollection
    ' カラム
    Dim column          As ValDbDefineColumn
    ' テーブル制約情報(PK)
    Dim tableConstPk    As New ValDbDefineTableConstraints
    ' PKカラムであるかをあらわすフラグ
    Dim isColumnPk      As Boolean
    
    Dim tableConstTmp   As ValDbDefineTableConstraints
    ' テーブル制約リストからPK制約を取得する
    For Each tableConstTmp In table.constraintsList.col
    
        If tableConstTmp.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk Then
        
            Set tableConstPk = tableConstTmp
            Exit For
        End If
    Next
    
    ' レコード値
    Dim recordValue     As String

    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf

    ' -------------------------------------------------
    ' レコードからクエリーを生成する
    ' -------------------------------------------------
    Dim i         As Long
    Dim j         As Long: j = 1
    Dim arrayMin  As Long   ' 配列の添え字（最小値）
    Dim arrayMax  As Long   ' 配列の添え字（最大値）
    Dim arraySize As Long   ' 配列のサイズ
    
    Set columnList = table.columnList
    
    arrayMin = 1
    arrayMax = columnList.count
    
    For i = arrayMin To arrayMax
    
        Set column = columnList.getItemByIndex(j)
        
        ' PK制約であるかどうかを判定する
        If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
        
            isColumnPk = False
        Else
        
            isColumnPk = True
        End If
        
        ' 複合キーを考慮してPKの場合は、入力値があろうがなかろうが常にステートメントとして追加する。
        ' 例えば、column1・column2がPKの場合に、column1に入力がありcolumn2に入力がない場合に
        ' column1のみ条件として追加してしまうと、レコードを一意に特定できずに複数件更新されてしまう恐れがあるため。
        If isColumnPk = True Then

            
            If rightStatementIsNull Then
                conditionStatement.append (CONNECT_AND)
                conditionStatement.append (rightIdentifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
                conditionStatement.append (" IS NULL")
            Else
                conditionStatement.append (CONNECT_AND)
                conditionStatement.append (leftIdentifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
                conditionStatement.append (" = ")
                conditionStatement.append (rightIdentifier & DBUtil.escapeTableColumn(m_dbms, column.columnName, m_tableColumnEscape))
            End If
            
            conditionStatement.append (" ")
            
        End If

continue:
        j = j + 1
    Next
    
    ' COLUMN句またはWHERE句に一つ以上カラムが設定されている場合
    If conditionStatement.str <> "" Then

        ' 前方に付加された冗長な"AND "と改行コードを取り除く
        conditionStatement.remove 1, Len(CONNECT_AND)
        conditionStatement.insert 1, CONNECT_FIRST
                                
    End If
    
    createSelectForPkJoin = conditionStatement.str

End Function

Private Function defaultColumnFormat() As ValDbColumnTypeColInfo

    Set defaultColumnFormat = New ValDbColumnTypeColInfo
    defaultColumnFormat.formatSelect = COLUMN_FORMAT_REPLACE_CHAR_COLUMN
    defaultColumnFormat.formatUpdate = COLUMN_FORMAT_REPLACE_CHAR_DIRECT

End Function
