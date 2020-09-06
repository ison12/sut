VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbQueryCreatorSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDbQueryCreator

' *********************************************************
' DBのクエリを生成するインターフェース。
' Microsoft SQLServer DB向けの実装。
'
' 作成者　：Ison
' 履歴　　：2010/11/06　新規作成
'
' 特記事項：
'
' *********************************************************
' SQLが定義されているシート名
Private Const SHEET_NAME_SQL_DEFINE                         As String = "template_sql"
' SELECTクエリのテンプレート
Private Const CELL_SQL_DEFINE_SELECT                        As String = "A2"

' DBコネクション
Private m_dbConn As Object
' DBMS種類
Private m_dbms As DbmsType

' 空白文字読み取り方式
Private m_emptyValueReading As Long
' 直接入力文字
Private m_directInputCharPrefix As String
Private m_directInputCharSuffix As String
' NULL入力文字
Private m_nullInputChar As String
' カラム書式情報リスト
Private m_columnFormatList As ValCollection
' スキーマ使用有無
Private m_schemaUse As Long
' テーブル・カラム名のエスケープ
Private m_tableColumnEscape As Boolean

' DBクエリ生成共通インターフェース
Private queryCreatorCommon As DbQueryCreator

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    ' DBクエリ生成共通インターフェースを初期化する
    Set queryCreatorCommon = New DbQueryCreator
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' DBクエリ生成共通インターフェースを破棄する
    Set queryCreatorCommon = Nothing
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
Private Property Let IDbQueryCreator_dbConn(ByVal RHS As Variant)

    ' DBコネクションを設定
    Set m_dbConn = RHS
    
    ' DBMS種類を設定
    m_dbms = ADOUtil.getDBMSType(m_dbConn)
    
    ' 共通インターフェースにDBコネクションを設定
    queryCreatorCommon.dbConn = m_dbConn

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
Private Property Let IDbQueryCreator_emptyValueReading(ByVal RHS As Long)

    m_emptyValueReading = RHS
    queryCreatorCommon.emptyValueReading = RHS
End Property

' =========================================================
' ▽直接入力文字接頭辞設定
'
' 概要　　　：
' 引数　　　：vNewValue 直接入力文字接頭辞
'
' 特記事項　：
'
'
' =========================================================
Private Property Let IDbQueryCreator_directInputCharPrefix(ByVal RHS As String)

    m_directInputCharPrefix = RHS
    queryCreatorCommon.directInputCharPrefix = RHS

End Property

' =========================================================
' ▽直接入力文字接尾辞設定
'
' 概要　　　：
' 引数　　　：vNewValue 直接入力文字接尾辞
'
' 特記事項　：
'
'
' =========================================================
Private Property Let IDbQueryCreator_directInputCharSuffix(ByVal RHS As String)

    m_directInputCharSuffix = RHS
    queryCreatorCommon.directInputCharSuffix = RHS

End Property

' =========================================================
' ▽NULL入力文字設定
'
' 概要　　　：
' 引数　　　：vNewValue NULL入力文字
'
' 特記事項　：
'
'
' =========================================================
Public Property Let IDbQueryCreator_nullInputChar(ByVal vNewValue As String)

    m_nullInputChar = vNewValue
    queryCreatorCommon.nullInputChar = vNewValue

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
Private Property Let IDbQueryCreator_columnFormatList(ByVal RHS As Variant)

    Set m_columnFormatList = RHS
    queryCreatorCommon.columnFormatList = RHS
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
Public Property Let IDbQueryCreator_schemaUse(ByVal vNewValue As Variant)

    m_schemaUse = vNewValue
    queryCreatorCommon.schemaUse = vNewValue
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
Public Property Let IDbQueryCreator_tableColumnEscape(ByVal vNewValue As Variant)

    m_tableColumnEscape = vNewValue
    queryCreatorCommon.tableColumnEscape = vNewValue
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
Private Function IDbQueryCreator_createInsert(ByRef table As ValDbDefineTable _
                                            , ByRef record As Variant) As String

    ' 共通インターフェースを呼び出してクエリーを生成してもらう
    IDbQueryCreator_createInsert = queryCreatorCommon.createInsert(table, record)
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
Private Function IDbQueryCreator_createUpdate(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' 共通インターフェースを呼び出してクエリーを生成してもらう
    IDbQueryCreator_createUpdate = queryCreatorCommon.createUpdate(table, record)
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
Private Function IDbQueryCreator_createDelete(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' 共通インターフェースを呼び出してクエリーを生成してもらう
    IDbQueryCreator_createDelete = queryCreatorCommon.createDelete(table, record)
End Function

' =========================================================
' ▽DELETE文生成（テーブル上の全レコードを生成）
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
Private Function IDbQueryCreator_createDeleteOfAllTable(table As ValDbDefineTable) As String

    ' 共通インターフェースを呼び出してクエリーを生成してもらう
    IDbQueryCreator_createDeleteOfAllTable = queryCreatorCommon.createDeleteOfAllTable(table)

End Function

' =========================================================
' ▽SELECT文生成
'
' 概要　　　：SELECT SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　condition 条件句
'
' 戻り値　　：SELECT文
'
' 特記事項　：
'
'
' =========================================================
Private Function IDbQueryCreator_createSelect(ByRef table As ValDbDefineTable _
                                            , ByRef condition As ValSelectCondition) As String

    ' インデント
    Const INDENT As String = "    "
    
    ' テンプレート
    Dim templateStatement As String

    ' SQL本体
    Dim statement         As String
    ' COLUMN句
    Dim statementColumn   As String
    ' COLUMN句（変換無し）
    Dim statementColumnNotConverted As String
    Dim statementColumnRowNum As String
    ' FROM句
    Dim statementFrom     As String
    ' WHERE句
    Dim statementWhere    As String
    ' ORDER BY句
    Dim statementOrderby  As String
    ' OPTION
    Dim statementOption   As String
    
    Dim temp As String
    
    ' COLUMN句を生成する
    statementColumn = queryCreatorCommon.createSelectForColumn(table)
    ' FROM句を生成する
    statementFrom = INDENT & DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)
    
    ' WHERE句とORDER BY句を生成する
    queryCreatorCommon.createSelectForWhereOrderby table _
                                                 , condition _
                                                 , statementWhere _
                                                 , statementOrderby
    

    ' -------------------------------------------------
    ' テンプレートSQLステートメントを取得
    ' -------------------------------------------------
    templateStatement = SutWorkbook.Worksheets(SHEET_NAME_SQL_DEFINE).Range(CELL_SQL_DEFINE_SELECT)
    ' -------------------------------------------------

    ' -------------------------------------------------
    ' SQLを生成する
    ' -------------------------------------------------
    statement = templateStatement
    
    ' ORDER BY句が指定されている且つレコードの範囲取得を指定している場合
    If statementOrderby <> "" _
           And (condition.recRangeStart <> condition.REC_RANGE_NONE _
             Or condition.recRangeEnd <> condition.REC_RANGE_NONE) Then
    
        ' 1行で連なっているorder by句を取得する
        Dim orderbySingleLine As String
        orderbySingleLine = queryCreatorCommon.createSelectForOrderby(table, condition)
        
        Dim innerTable As New StringBuilder
        
        innerTable.append (INDENT)
        innerTable.append ("(SELECT *, ROW_NUMBER()")
        
        ' 並び順が指定されている場合、over句を指定する
        If statementOrderby <> "" Then
        
            innerTable.append(" OVER(").append(orderbySingleLine).append (")")
        End If
        
        innerTable.append (" AS row_number")
        innerTable.append(" FROM ").append(DBUtil.concatSchemaTable(DBUtil.escapeTableColumn(m_dbms, table.schemaName, m_tableColumnEscape), DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape), m_schemaUse)).append(") ").append (DBUtil.escapeTableColumn(m_dbms, table.tableName, m_tableColumnEscape))
    
        statementFrom = innerTable.str
                                
        Dim statementRecRangeStart As String
        Dim statementRecRangeEnd   As String
        
        If condition.recRangeStart <> condition.REC_RANGE_NONE Then
        
            statementRecRangeStart = condition.recRangeStart & " <= row_number"
        End If
        
        If condition.recRangeEnd <> condition.REC_RANGE_NONE Then
            
            statementRecRangeEnd = "row_number  <= " & condition.recRangeEnd
            
            If statementRecRangeStart <> "" Then
            
                statementRecRangeEnd = " AND " & statementRecRangeEnd
                
            End If

        End If
        
        If statementWhere = "" Then
        
            statementWhere = "WHERE" & vbLf & INDENT
        Else
        
            statementWhere = statementWhere & vbLf & "AND "
        End If
        
        statementWhere = statementWhere _
                            & statementRecRangeStart _
                            & statementRecRangeEnd
        
        
    End If
    
    ' COLUMN句の置換
    statement = replace(statement _
                         , "${statementColumn}" _
                         , statementColumn)
                         
    ' FROM句の置換
    statement = replace(statement _
                            , "${statementFrom}" _
                            , statementFrom)

    ' WHERE句の置換
    statement = replace(statement _
                            , convertReplaceKeyword("${statementWhere}", statementWhere) _
                            , statementWhere)

    ' ORDER BY句の置換
    statement = replace(statement _
                            , convertReplaceKeyword("${statementOrderby}", statementOrderby) _
                            , statementOrderby)

    ' OPTIONの置換
    statement = replace(statement _
                            , convertReplaceKeyword("${statementOption}", statementOption) _
                            , statementOption)

    ' 戻り値を設定
    IDbQueryCreator_createSelect = statement
    
End Function

Private Function convertReplaceKeyword(ByVal keyword As String _
                                     , ByVal replacedStr As String) As String

    ' 置換後の文字列が空の場合
    If replacedStr = "" Then
    
        ' キーワードに改行コードを連結する
        convertReplaceKeyword = vbLf & keyword
        
    Else
    
        ' キーワードをそのまま返す
        convertReplaceKeyword = keyword
    End If
    
End Function



