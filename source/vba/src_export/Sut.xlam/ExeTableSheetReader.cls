VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeTableSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' シートからレコード情報を読み取る処理を定義したクラス
'
' 作成者　：Ison
' 履歴　　：2008/02/01　新規作成
'
' 特記事項：
' *********************************************************

' ワークシート
Public sheet     As Worksheet
' DBコネクション
Public conn      As Object
' キャンセルフラグ
Public canceled  As Boolean

' =========================================================
' ▽カラムリスト読込開始前イベント
'
' 概要　　　：カラムリストの読込を開始する前に呼び出されるイベント
' 引数　　　：cnt 件数
'
' =========================================================
Public Event readColumnListBegin(ByVal cnt As Long)

' =========================================================
' ▽カラム情報読込イベント
'
' 概要　　　：カラム情報読込時に呼び出されるイベント
' 引数　　　：column カラム情報
'
' =========================================================
Public Event readColumn(ByRef column As ValDbDefineColumn)

' =========================================================
' ▽カラムリスト読込イベント
'
' 概要　　　：カラムリスト読込時に呼び出されるイベント
' 引数　　　：columnList カラムリスト
'
' =========================================================
Public Event readColumnList(ByRef columnList As ValCollection)

' =========================================================
' ▽テーブル定義読込開始前イベント
'
' 概要　　　：テーブル定義の読込を開始する前に呼び出されるイベント
' 引数　　　：
'
' =========================================================
Public Event readTableInfoBegin()

' =========================================================
' ▽テーブル定義読込イベント
'
' 概要　　　：テーブル定義の読込時に呼び出されるイベント
' 引数　　　：tableSheet テーブルシート
'
' =========================================================
Public Event readTableInfo(ByRef tableSheet As ValTableWorksheet)

' =========================================================
' ▽レコードリスト読込開始前イベント
'
' 概要　　　：レコードリストの生成を開始する前に呼び出されるイベント
' 引数　　　：cnt 件数
'
' =========================================================
Public Event readRecordListBegin(ByVal cnt As Long)

' =========================================================
' ▽レコードリスト読込イベント
'
' 概要　　　：レコードリスト生成時に呼び出されるイベント
' 引数　　　：recordList レコード
'
' =========================================================
Public Event readRecordList(ByRef recordList As ValRecordList)

' =========================================================
' ▽範囲選択時のレコードリスト【コレクション】読込開始前イベント
'
' 概要　　　：レコードリスト【コレクション】の生成を開始する前に呼び出されるイベント
' 引数　　　：cnt 件数
'
' =========================================================
Public Event readRecordListBySelectionBegin(ByVal cnt As Long)

' =========================================================
' ▽範囲選択時のレコードリスト読込イベント
'
' 概要　　　：レコードリスト【コレクション】読込時に呼び出されるイベント
' 引数　　　：recordList レコードリスト
'
' =========================================================
Public Event readRecordListBySelectionOne(ByRef recordList As ValRecordList)

' =========================================================
' ▽範囲選択時のレコードリスト【コレクション】読込イベント
'
' 概要　　　：レコードリスト【コレクション】読込時に呼び出されるイベント
' 引数　　　：recordLists レコードリスト【コレクション】
'
' =========================================================
Public Event readRecordListBySelection(ByRef recordLists As ValCollection)

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ▽テーブルシート判定
'
' 概要　　　：対象シート（メンバのsheet）が
' 　　　　　　テーブルシートであるかを判定する。
'
' 引数　　　：
' 戻り値　　：True テーブルシートの場合
' =========================================================
Public Function isTableSheet() As Boolean

    ' テーブルシート判定用のオートシェイプ
    Dim s As Shape

    ' 戻り値を初期化する
    isTableSheet = False
    
    ' ------------------------------------------------
    ' エラーが発生しても継続する
    On Error Resume Next

    Set s = sheet.Shapes(ConstantsTable.U_SHEET_CHECK_IMAGE)
    
    ' 取得できた場合
    If Not s Is Nothing Then
    
        isTableSheet = True
        Exit Function
    End If
    
    Set s = sheet.Shapes(ConstantsTable.R_SHEET_CHECK_IMAGE)
    
    ' 取得できた場合
    If Not s Is Nothing Then
    
        isTableSheet = True
        Exit Function
    End If
        
    On Error GoTo 0
    ' ------------------------------------------------

End Function

' =========================================================
' ▽テーブルシート判定
'
' 概要　　　：対象シート（メンバのsheet）がテーブルシートであるかを判定する。
' 　　　　　　テーブルシートではない場合、エラーを発行する。
'
' 引数　　　：
' 戻り値　　：
' =========================================================
Public Sub validTableSheet()

    ' テーブルシートではない場合
    If isTableSheet = False Then
    
        err.Raise ConstantsError.ERR_NUMBER_IS_NOT_TABLE_SHEET _
            , _
            , ConstantsError.ERR_DESC_IS_NOT_TABLE_SHEET
    End If

End Sub
' =========================================================
' ▽行フォーマット取得
'
' 概要　　　：対象シート（メンバのsheet）の行フォーマットを取得する
'
' 引数　　　：
' 戻り値　　：recFormatToUnder
' 　　　　　　recFormatToRight
' 　　　　　　-1
'
' =========================================================
Public Function getRowFormat() As REC_FORMAT

    ' 戻り値
    Dim ret As REC_FORMAT
    
    ' テーブルシート判定用のオートシェイプ
    Dim s As Shape

    ' 戻り値を初期化する
    ret = -1
    
    ' ------------------------------------------------
    ' エラーが発生しても継続する
    On Error Resume Next

    Set s = sheet.Shapes("SUT_WORKSHEET_MARK_TO_UNDER")
    
    ' 取得できた場合
    If Not s Is Nothing Then
    
        ' 戻り値変数に行フォーマットを設定する
        ret = REC_FORMAT.recFormatToUnder
        
        ' 戻り値を設定する
        getRowFormat = ret
        
        Exit Function
    End If
    
    Set s = sheet.Shapes("SUT_WORKSHEET_MARK_TO_RIGHT")
    
    ' 取得できた場合
    If Not s Is Nothing Then
    
        ' 戻り値変数に行フォーマットを設定する
        ret = REC_FORMAT.recFormatToRight
        
        ' 戻り値を設定する
        getRowFormat = ret
        
        Exit Function
    End If
        
    On Error GoTo 0
    ' ------------------------------------------------

    getRowFormat = ret

End Function

' =========================================================
' ▽テーブル定義読み込み
'
' 概要　　　：シート上のテーブル定義を読み込んで
' 　　　　　　テーブル定義シートオブジェクトを生成する。
'
' 引数　　　：tableInfoOnly テーブル情報のみ取得（カラム情報は取得しない）
' 戻り値　　：テーブルシート
'
' =========================================================
Public Function readTableInfo(Optional tableInfoOnly As Boolean = False) As ValTableWorksheet

    ' テーブル定義
    Dim table      As ValDbDefineTable
    ' テーブル情報を生成する
    Set table = New ValDbDefineTable
    ' テーブルシート
    Dim tableSheet As ValTableWorksheet
    ' テーブルシート情報を生成する
    Set tableSheet = New ValTableWorksheet
    Set tableSheet.table = table
    
    ' テーブル名
    Dim tableName As String
    
    ' テーブル生成開始イベント
    RaiseEvent readTableInfoBegin
    
    With sheet

        ' テーブル名を取得する
        tableName = .Cells(ConstantsTable.TABLE_NAME_ROW _
                         , ConstantsTable.TABLE_NAME_COL)
        
        
        ' テーブル情報取得元シートを設定する
        Set tableSheet.sheet = sheet
        tableSheet.sheetName = sheet.name
        
        If InStr(tableName, ".") > 0 Then
        
            ' スキーマ（一時変数）
            Dim tempSchema As String
            ' テーブル（一時変数）
            Dim tempTable  As String
            
            ' スキーマ・テーブル名を抽出し設定する
            DBUtil.extractSchemaTable tableName, tempSchema, tempTable
            
            table.schemaName = tempSchema
            table.tableName = tempTable
            
            ' スキーマ名が付加されている場合は、スキーマ名が省略されていないとみなす
            tableSheet.omitsSchema = False
        Else
        
            ' テーブル名を設定する
            table.tableName = tableName
            ' スキーマ名を設定する（隠しセルから取得）
            table.schemaName = .Cells(ConstantsTable.SCHEMA_NAME_ROW _
                                    , ConstantsTable.SCHEMA_NAME_COL)
            
            ' スキーマ名が付加されていない場合は、スキーマ名が省略されているとみなす
            tableSheet.omitsSchema = True
        End If
        
        ' テーブルコメントを設定する
        table.tableComment = .Cells(TABLE_NAME_LOG_ROW _
                                  , TABLE_NAME_LOG_COL)
        
        ' 行フォーマットを設定する
        tableSheet.recFormat = getRowFormat
        
        If tableInfoOnly = False Then
        
            If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
            
                readTableInfoToUnder tableSheet
            
            ElseIf tableSheet.recFormat = REC_FORMAT.recFormatToRight Then
            
                readTableInfoToRight tableSheet
                
            End If
        
        End If
        
    End With ' sheet

    ' テーブル定義生成完了イベント
    RaiseEvent readTableInfo(tableSheet)

    ' 戻り値を設定する
    Set readTableInfo = tableSheet

End Function

' =========================================================
' ▽テーブル定義読み込み（X方向）
'
' 概要　　　：シート上のテーブル定義がX方向の場合こちらを利用する。
' 　　　　　　主にカラムリストを読み込む。
'
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：
'
' =========================================================
Private Function readTableInfoToUnder(ByRef tableSheet As ValTableWorksheet)

    Dim i As Long
    
    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf

    ' シートの最大列数
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(tableSheet.sheet)

    ' テーブル定義
    Dim table      As ValDbDefineTable
    Set table = tableSheet.table
    ' カラム
    Dim column     As ValDbDefineColumn
    
    ' テーブル制約情報リスト
    Dim tableConstList As ValCollection
    Set tableConstList = tableSheet.table.constraintsList
    
    ' テーブル制約情報（PK）
    Dim tableConstPk   As New ValDbDefineTableConstraints
    
    ' テーブル制約情報（汎用的に使用）
    Dim tableConst     As ValDbDefineTableConstraints
    
    ' ループ継続フラグ
    Dim continue   As Variant
    
    ' 現在列
    Dim currentCol As Long
    
    ' カラム情報
    Dim columnName  As String    ' 名称
    Dim columnType  As String    ' 型
    Dim columnNull  As String    ' NULL可
    Dim columnDef   As String    ' デフォルト
    Dim columnPk    As String    ' PK
    Dim columnUk    As String    ' UNIQ
    Dim columnFk    As String    ' FK
    
    Dim columnUkArray()  As String    ' UNIQ配列
    Dim columnFkArray()  As String    ' FK  配列
    
    ' カラムサイズ
    Dim columnSize As Long
    ' カラムサイズを取得する
    columnSize = getColumnSize(tableSheet)
    
    ' カラムサイズが0の場合
    If columnSize <= 0 Then
    
        Exit Function
    End If
    
    ' カラム生成開始イベント
    RaiseEvent readColumnListBegin(columnSize)
    
    ' ---------------------------------------------------------------------
    With sheet
    
        ' 列の最初の位置を取得
        currentCol = ConstantsTable.U_COLUMN_OFFSET_COL
    
        ' ループを繰り返し行うかどうかを判定するための文字列を取得する
        continue = .Cells(ConstantsTable.U_COLUMN_NAME_ROW, currentCol)
        
        ' ループ継続条件（continueが空ではない場合継続）
        Do While continue <> Empty
        
            ' レコードにカラムを一つ追加する
            Set column = New ValDbDefineColumn
            
            ' カラム名を取得
            columnName = .Cells(ConstantsTable.U_COLUMN_NAME_ROW, currentCol)
            column.columnName = columnName
            
            
            ' カラム型を取得
            columnType = .Cells(ConstantsTable.U_COLUMN_TYPE_ROW, currentCol)
            
            column.dataTypeFormal = columnType
            
            ' カラム型文字列を変換（文字列中の括弧を取り除く）
            columnType = convertColumnType(columnType)

            column.dataType = columnType
            
            ' カラムがNULL可であるかを取得
            columnNull = .Cells(ConstantsTable.U_COLUMN_NULL_ROW, currentCol)
            
            If columnNull <> "" Then
            
                column.isNull = True
            Else
            
                column.isNull = False
            End If
            
            ' カラムのデフォルト値を取得
            columnDef = .Cells(ConstantsTable.U_COLUMN_DEF_ROW, currentCol)
            column.default = columnDef
            
            ' カラムがPKかどうかを取得
            columnPk = .Cells(ConstantsTable.U_COLUMN_PK_ROW, currentCol)
            
            If columnPk <> "" Then
            
                tableConstPk.addColumn column
            End If
            
            ' カラムがUKかどうかを取得
            columnUk = .Cells(ConstantsTable.U_COLUMN_UK_ROW, currentCol)
            
            If columnUk <> "" Then
            
                columnUk = replace(columnUk, vbCr, "")
                columnUkArray = Split(columnUk, newline)
                
                For i = LBound(columnUkArray) To UBound(columnUkArray)
                
                    ' テーブル制約情報がリストに格納されているかを確認する
                    Set tableConst = tableConstList.getItem(columnUkArray(i))
                    ' テーブル制約情報が存在しない場合
                    If tableConst Is Nothing Then
                    
                        ' 新しく制約情報を生成する
                        Set tableConst = New ValDbDefineTableConstraints
                        tableConst.schemaName = table.schemaName
                        tableConst.tableName = table.tableName
                        tableConst.constraintType = tableConstUk
                        tableConst.constraintName = columnUkArray(i)
                    End If
                    
                    tableConst.addColumn column
                Next
                
            End If
            
            ' カラムがFKかどうかを取得
            columnFk = .Cells(ConstantsTable.U_COLUMN_REFER_ROW, currentCol)
            If columnFk <> "" Then
            
                columnFk = replace(columnFk, vbCr, "")
                columnFkArray = Split(columnFk, newline)
                
                For i = LBound(columnFkArray) To UBound(columnFkArray)
                
                    ' テーブル制約情報がリストに格納されているかを確認する
                    Set tableConst = tableConstList.getItem(columnFkArray(i))
                    ' テーブル制約情報が存在しない場合
                    If tableConst Is Nothing Then
                    
                        ' 新しく制約情報を生成する
                        Set tableConst = New ValDbDefineTableConstraints
                        tableConst.schemaName = table.schemaName
                        tableConst.tableName = table.tableName
                        tableConst.constraintType = tableConstFk
                        tableConst.constraintName = columnFkArray(i)
                    End If
                    
                    tableConst.addColumn column
                Next
                
            End If
            
            table.addColumn column
            
            ' カラム生成完了イベント
            RaiseEvent readColumn(column)
        
            ' 列を一つ進める
            currentCol = currentCol + 1
            
            If sheetColSize < currentCol Then
            
                Exit Do
            End If
            
            ' ループを繰り返し行うかどうかを判定するための文字列を取得する
            continue = .Cells(ConstantsTable.U_COLUMN_NAME_ROW, currentCol)
        
        Loop
        ' ---------------------------------------------------------------------
    
        ' PK制約となるカラムが１件以上ある場合、テーブル制約情報として追加する
        If tableConstPk.columnList.count > 0 Then
        
            tableConstPk.schemaName = table.schemaName
            tableConstPk.tableName = table.tableName
            tableConstPk.constraintName = "PrimaryKey"
            tableConstPk.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk
            
            tableConstList.setItem tableConstPk
        End If
        
        ' カラムリスト生成完了イベント
        RaiseEvent readColumnList(table.columnList)
    
    End With ' sheet

End Function

' =========================================================
' ▽テーブル定義読み込み（Y方向）
'
' 概要　　　：シート上のテーブル定義がY方向の場合こちらを利用する。
' 　　　　　　主にカラムリストを読み込む。
'
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：
'
' =========================================================
Private Function readTableInfoToRight(ByRef tableSheet As ValTableWorksheet)

    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf
    
    Dim i As Long
    
    ' シートの最大行数
    Dim sheetRowSize As Long: sheetRowSize = ExcelUtil.getSizeOfSheetRow(tableSheet.sheet)

    ' テーブル定義
    Dim table      As ValDbDefineTable
    Set table = tableSheet.table
    ' カラム定義
    Dim columnList As ValCollection
    ' カラム
    Dim column     As ValDbDefineColumn
    
    ' テーブル制約情報リスト
    Dim tableConstList As ValCollection
    Set tableConstList = tableSheet.table.constraintsList
    
    ' テーブル制約情報（PK）
    Dim tableConstPk   As New ValDbDefineTableConstraints
    
    ' テーブル制約情報（汎用的に使用）
    Dim tableConst     As ValDbDefineTableConstraints
    
    ' ループ継続フラグ
    Dim continue   As Variant
    
    ' 現在行
    Dim currentRow As Long
    
    ' カラム情報
    Dim columnName  As String    ' 名称
    Dim columnType  As String    ' 型
    Dim columnNull  As String    ' NULL可
    Dim columnDef   As String    ' デフォルト
    Dim columnPk    As String    ' PK
    Dim columnUk    As String    ' UNIQ
    Dim columnFk    As String    ' FK
    
    Dim columnUkArray()  As String    ' UNIQ配列
    Dim columnFkArray()  As String    ' FK  配列
    
    ' カラムサイズ
    Dim columnSize As Long
    ' カラムサイズを取得する
    columnSize = getColumnSize(tableSheet)
    
    ' カラムサイズが0の場合
    If columnSize <= 0 Then
    
        Exit Function
    End If
    
    ' カラム生成開始イベント
    RaiseEvent readColumnListBegin(columnSize)
    
    ' ---------------------------------------------------------------------
    With sheet
    

        ' 列の最初の位置を取得
        currentRow = ConstantsTable.R_COLUMN_OFFSET_ROW
    
        ' ループを繰り返し行うかどうかを判定するための文字列を取得する
        continue = .Cells(currentRow, ConstantsTable.R_COLUMN_NAME_COL)
        
        ' ループ継続条件（continueが空ではない場合継続）
        Do While continue <> Empty
        
            ' レコードにカラムを一つ追加する
            Set column = New ValDbDefineColumn
            
            ' カラム名を取得
            columnName = .Cells(currentRow, ConstantsTable.R_COLUMN_NAME_COL)
            column.columnName = columnName
            
            
            ' カラム型を取得
            columnType = .Cells(currentRow, ConstantsTable.R_COLUMN_TYPE_COL)
            
            column.dataTypeFormal = columnType
            
            ' カラム型文字列を変換（文字列中の括弧を取り除く）
            columnType = convertColumnType(columnType)
            
            column.dataType = columnType
            ' カラムがNULL可であるかを取得
            columnNull = .Cells(currentRow, ConstantsTable.R_COLUMN_NULL_COL)
            
            If columnNull <> "" Then
            
                column.isNull = True
            Else
            
                column.isNull = False
            End If
            
            ' カラムのデフォルト値を取得
            columnDef = .Cells(currentRow, ConstantsTable.R_COLUMN_DEF_COL)
            column.default = columnDef
            
            ' カラムがPKかどうかを取得
            columnPk = .Cells(currentRow, ConstantsTable.R_COLUMN_PK_COL)
            
            If columnPk <> "" Then
            
                tableConstPk.addColumn column
            End If
            
            ' カラムがUKかどうかを取得
            columnUk = .Cells(currentRow, ConstantsTable.R_COLUMN_UK_COL)
            
            If columnUk <> "" Then
            
                columnUk = replace(columnUk, vbCr, "")
                columnUkArray = Split(columnUk, newline)
                
                For i = LBound(columnUkArray) To UBound(columnUkArray)
                
                    ' テーブル制約情報がリストに格納されているかを確認する
                    Set tableConst = tableConstList.getItem(columnUkArray(i))
                    ' テーブル制約情報が存在しない場合
                    If tableConst Is Nothing Then
                    
                        ' 新しく制約情報を生成する
                        Set tableConst = New ValDbDefineTableConstraints
                        tableConst.schemaName = table.schemaName
                        tableConst.tableName = table.tableName
                        tableConst.constraintType = tableConstUk
                        tableConst.constraintName = columnUkArray(i)
                    End If
                    
                    tableConst.addColumn column
                Next
                
            End If
            
            ' カラムがFKかどうかを取得
            columnFk = .Cells(currentRow, ConstantsTable.R_COLUMN_REFER_COL)
            If columnFk <> "" Then
            
                columnFk = replace(columnFk, vbCr, "")
                columnFkArray = Split(columnFk, newline)
                
                For i = LBound(columnFkArray) To UBound(columnFkArray)
                
                    ' テーブル制約情報がリストに格納されているかを確認する
                    Set tableConst = tableConstList.getItem(columnFkArray(i))
                    ' テーブル制約情報が存在しない場合
                    If tableConst Is Nothing Then
                    
                        ' 新しく制約情報を生成する
                        Set tableConst = New ValDbDefineTableConstraints
                        tableConst.schemaName = table.schemaName
                        tableConst.tableName = table.tableName
                        tableConst.constraintType = tableConstFk
                        tableConst.constraintName = columnFkArray(i)
                    End If
                    
                    tableConst.addColumn column
                Next
                
            End If
            
            table.addColumn column
            
            ' カラム生成完了イベント
            RaiseEvent readColumn(column)
        
            ' 行を一つ進める
            currentRow = currentRow + 1
            
            If sheetRowSize < currentRow Then
            
                Exit Do
            End If
            
            ' ループを繰り返し行うかどうかを判定するための文字列を取得する
            continue = .Cells(currentRow, ConstantsTable.R_COLUMN_NAME_COL)
        
        Loop
        ' ---------------------------------------------------------------------
    
        ' PK制約となるカラムが１件以上ある場合、テーブル制約情報として追加する
        If tableConstPk.columnList.count > 0 Then
        
            tableConstPk.schemaName = table.schemaName
            tableConstPk.tableName = table.tableName
            tableConstPk.constraintName = "PrimaryKey"
            tableConstPk.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk
            
            tableConstList.setItem tableConstPk
        End If
        
        ' カラムリスト生成完了イベント
        RaiseEvent readColumnList(columnList)
    
    End With ' sheet

End Function

' =========================================================
' ▽シート上のレコード情報を全て読み込む
'
' 概要　　　：シート上のレコード情報を全て読み込む
'
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：レコードリスト
'
' =========================================================
Public Function readRecordListAll(ByRef tableSheet As ValTableWorksheet) As ValRecordList

    ' 戻り値
    Dim ret As ValRecordList
    
    ' レコードサイズ
    Dim recordSize As Long
    
    ' レコードサイズを取得する
    recordSize = getRecordSize(tableSheet)
    
    ' サイズが 0 以下の場合、空の戻り値を返し終了する
    If recordSize <= 0 Then
    
        ' 戻り値を設定する
        Set readRecordListAll = New ValRecordList
        Exit Function
    End If
            
    ' 戻り値を設定する
    Set ret = readRecordList(tableSheet, 1, recordSize)
    
    Set readRecordListAll = ret

End Function

' =========================================================
' ▽シート上のレコード情報を任意の範囲で読み込む
'
' 概要　　　：シート上のレコード情報を任意の範囲で読み込む
'
' 　　　　　　○例
' 　　　　　　・1件読み込む場合
' 　　　　　　　row     = 1
' 　　　　　　　rowSize = 1
'
' 　　　　　　・5～10件読み込む場合
' 　　　　　　　row     = 5
' 　　　　　　　rowSize = 10
'
' 　　　　　　・1000件読み込む場合
' 　　　　　　　row     = 1
' 　　　　　　　rowSize = 1000
'
'
' 引数　　　：tableSheet テーブルシート
' 　　　　　　row        読み込み開始行
' 　　　　　　rowSize    読み込みサイズ
' 戻り値　　：レコードリスト
'
' =========================================================
Public Function readRecordList(ByRef tableSheet As ValTableWorksheet _
                             , ByVal row As Long _
                             , ByVal rowSize As Long) As ValRecordList

    ' テーブル定義
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    
    ' 戻り値
    Dim recordList As ValRecordList
    ' 戻り値を生成する
    Set recordList = New ValRecordList
        
    ' レコード
    Dim records As Variant
    
    ' カラムサイズ
    Dim columnSize As Long
    
    ' カラムサイズを取得する
    columnSize = table.columnList.count
    ' サイズが 0 以下の場合、空の戻り値を返し終了する
    If columnSize <= 0 Then
    
        ' 戻り値を設定する
        Set readRecordList = recordList
        Exit Function
    End If
    
    ' ◇レコードリスト生成処理が開始されることをイベントとして発信する
    RaiseEvent readRecordListBegin(rowSize)
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        row = row + ConstantsTable.U_RECORD_OFFSET_ROW - 1
    
        ' セルからレコードを取得する
        records = ExcelUtil.copyCellsToArray( _
                                 sheet _
                               , row _
                               , rowSize _
                               , ConstantsTable.U_COLUMN_OFFSET_COL _
                               , columnSize)
        
    
        ' レコード配列取得もとのオフセットRangeオブジェクトを設定する
        Set recordList.valuesCellOffset = sheet.Cells(row, ConstantsTable.U_COLUMN_OFFSET_COL)
    
        ' レコード配列を設定する
        recordList.values = records
    
    Else
    
        row = row + ConstantsTable.R_RECORD_OFFSET_COL - 1
        
        ' セルからレコードを取得する
        records = ExcelUtil.copyCellsToArray( _
                                 sheet _
                               , ConstantsTable.R_COLUMN_OFFSET_ROW _
                               , columnSize _
                               , row _
                               , rowSize)
                               
        ' X・Y方向を逆に入れ替える
        Dim swapArray As Variant
        
        swapArray = VBUtil.transposeDim(records)
        
        ' 一度 records 配列を削除する
        Erase records
        
        ' records 配列に入れ替えた配列を再設定する
        records = swapArray
    
        ' レコード配列取得もとのオフセットRangeオブジェクトを設定する
        Set recordList.valuesCellOffset = sheet.Cells(ConstantsTable.R_COLUMN_OFFSET_ROW, row)
    
        ' レコード配列を設定する
        recordList.values = records
            
    End If
    
    ' 戻り値を設定する
    Set readRecordList = recordList
    
    ' ◇レコードリスト生成処理が完了したことをイベントとして発信する
    RaiseEvent readRecordList(recordList)

End Function

' =========================================================
' ▽レコード情報（範囲選択）の読込
'
' 概要　　　：レコード情報をセルから取得してクラスに格納する
'
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：レコード情報クラス
'
' =========================================================
Public Function readRecordListBySelectionArea(ByRef tableSheet As ValTableWorksheet) As ValCollection

    ' テーブル定義
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    
    ' 戻り値（ValRecordListのコレクション）
    Dim recordLists As ValCollection
    ' 戻り値を初期化する
    Set recordLists = New ValCollection

    ' 選択領域の一意性チェック用のコレクションオブジェクト
    ' 同じ領域が選択されていないかをチェックする
    ' Selectin.Areasは、同じセルが複数回選択された場合でも同じセル情報が複数格納される。
    ' 例えば、A1セルを2回 Ctrl+マウス左クリック で選択されると、A1セルが2つ情報として格納される。そのためにチェックを行う必要がある。
    Dim uniqueCheck    As New ValCollection
    ' 行番号または列番号をキーとして格納する
    Dim uniqueCheckKey As String

    ' レコードリスト
    Dim recordList As ValRecordList
        
    ' レコード
    Dim records As Variant
    
    ' 選択領域
    Dim selectionArea  As Areas
    ' レンジオブジェクト
    Dim selectionRange As Range
    
    ' 選択領域の総数
    Dim cnt As Long
    
    Dim rowsCount As Long
    Dim colsCount As Long
    
    Dim i As Long
            

    
    ' カラムサイズ
    Dim columnSize As Long: columnSize = table.columnList.count
        
    Set selectionArea = ActiveWindow.RangeSelection.Areas
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ' 選択領域に関するチェックを行い、レコードの総数をカウントする
        For Each selectionRange In selectionArea
    
            cnt = cnt + selectionRange.Rows.count
            
            ' 選択範囲が入力領域の範囲内かをチェック
            If selectionRange.row < ConstantsTable.U_RECORD_OFFSET_ROW Then
            
                ' 選択範囲がレコード開始オフセットより小さい場合
                ' エラーを発行する
                err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SELECTION _
                        , _
                        , ConstantsError.ERR_DESC_OUT_OF_RANGE_SELECTION
            End If

            ' 選択領域からセルを１件ずつ取り出し番号としてチェック用コレクションに格納
            ' チェック用コレクションに既に格納済みであればエラーとする
            For i = 1 To selectionRange.Rows.count
                
                ' キーを取得する
                uniqueCheckKey = CStr(selectionRange.Offset.row + i - 1)
                
                ' 重複している場合
                If uniqueCheck.exist(uniqueCheckKey) = True Then
                
                    ' エラーを発行する
                    err.Raise ERR_NUMBER_DUPLICATE_SELECTION_CELL _
                            , _
                            , ERR_DESC_DUPLICATE_SELECTION_CELL
                ' 重複していない場合
                Else
                
                    ' 格納
                    uniqueCheck.setItem True, uniqueCheckKey
                End If
                
            Next
        Next
        
        ' ◇レコードリスト生成処理が開始されることをイベントとして発信する
        RaiseEvent readRecordListBySelectionBegin(cnt)
    
        For Each selectionRange In selectionArea
        
            records = ExcelUtil.copyCellsToArray( _
                                 ActiveWindow.ActiveSheet _
                               , selectionRange.row _
                               , selectionRange.Rows.count _
                               , ConstantsTable.U_COLUMN_OFFSET_COL _
                               , columnSize)
    
            ' レコードリストを初期化する
            Set recordList = New ValRecordList
            
            ' レコードリストに配列を設定する
            recordList.values = records
            ' レコードリストにRangeオブジェクトを設定する
            Set recordList.valuesCellOffset = selectionRange
            
            ' 戻り値にレコードリストを設定する
            recordLists.setItem recordList
            
            ' ◇レコードリストが1つ生成されたことをイベントとして発信する
            RaiseEvent readRecordListBySelectionOne(recordList)
        Next
    
    Else
    
        ' 選択領域に関するチェックを行い、レコードの総数をカウントする
        For Each selectionRange In selectionArea
    
            cnt = cnt + selectionRange.Columns.count
            
            ' 選択範囲が入力領域の範囲内かをチェック
            If selectionRange.column < ConstantsTable.R_RECORD_OFFSET_COL Then
            
                ' 選択範囲がレコード開始オフセットより小さい場合
                ' エラーを発行する
                err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SELECTION _
                        , _
                        , ConstantsError.ERR_DESC_OUT_OF_RANGE_SELECTION
            End If

            ' 選択領域からセルを１件ずつ取り出し番号としてチェック用コレクションに格納
            ' チェック用コレクションに既に格納済みであればエラーとする
            For i = 1 To selectionRange.Columns.count
                
                ' キーを取得する
                uniqueCheckKey = CStr(selectionRange.Offset.column + i - 1)
                
                ' 重複している場合
                If uniqueCheck.exist(uniqueCheckKey) = True Then
                
                    ' エラーを発行する
                    err.Raise ERR_NUMBER_DUPLICATE_SELECTION_CELL _
                            , _
                            , ERR_DESC_DUPLICATE_SELECTION_CELL
                ' 重複していない場合
                Else
                
                    ' 格納
                    uniqueCheck.setItem True, uniqueCheckKey
                End If
                
            Next

        Next
        
        ' ◇レコードリスト生成処理が開始されることをイベントとして発信する
        RaiseEvent readRecordListBySelectionBegin(cnt)
    
        For Each selectionRange In selectionArea
        
            records = ExcelUtil.copyCellsToArray( _
                                 ActiveWindow.ActiveSheet _
                               , ConstantsTable.R_COLUMN_OFFSET_ROW _
                               , columnSize _
                               , selectionRange.column _
                               , selectionRange.Columns.count)
    
            ' X・Y方向を逆に入れ替える
            Dim swapArray As Variant
            
            swapArray = VBUtil.transposeDim(records)
            
            ' 一度 records 配列を削除する
            Erase records
            
            ' records 配列に入れ替えた配列を再設定する
            records = swapArray
    
            ' レコードリストを初期化する
            Set recordList = New ValRecordList
            
            ' レコードリストに配列を設定する
            recordList.values = records
            ' レコードリストにRangeオブジェクトを設定する
            Set recordList.valuesCellOffset = selectionRange
            
            ' 戻り値にレコードリストを設定する
            recordLists.setItem recordList
            
            ' ◇レコードリストが1つ生成されたことをイベントとして発信する
            RaiseEvent readRecordListBySelectionOne(recordList)
        Next
    
    End If

    ' ◇レコードリスト生成処理が完了したことをイベントとして発信する
    RaiseEvent readRecordListBySelection(recordLists)

    ' 戻り値を設定する
    Set readRecordListBySelectionArea = recordLists
End Function

' =========================================================
' ▽テーブルシートのカラムサイズ取得
'
' 概要　　　：テーブルシートからカラムサイズの取得を行う
'
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：カラムサイズ
'
' =========================================================
Public Function getColumnSize(ByRef tableSheet As ValTableWorksheet) As Long

    ' 戻り値変数
    Dim ret As Long
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ret = ExcelUtil.getCellEndOfLastInputtedCol(sheet, ConstantsTable.U_COLUMN_NAME_ROW) - ConstantsTable.U_COLUMN_OFFSET_COL + 1
    
    Else
    
        ret = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsTable.R_COLUMN_NAME_COL) - ConstantsTable.R_COLUMN_OFFSET_ROW + 1
    End If
    
    ' 戻り値を設定する
    getColumnSize = ret
    
End Function

' =========================================================
' ▽テーブルシートの行サイズ取得
'
' 概要　　　：テーブルシートから行サイズの取得を行う
'
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：行サイズ
'
' =========================================================
Public Function getRecordSize(ByRef tableSheet As ValTableWorksheet) As Long


    ' 戻り値変数
    Dim ret As Long
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ret = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsTable.U_RECORD_NUM_COL)
        ret = ret - ConstantsTable.U_RECORD_OFFSET_ROW + 1
    
    Else
    
        ret = ExcelUtil.getCellEndOfLastInputtedCol(sheet, ConstantsTable.R_RECORD_NUM_ROW)
        ret = ret - ConstantsTable.R_RECORD_OFFSET_COL + 1
    End If
    
    ' 戻り値を設定する
    getRecordSize = ret

End Function

Public Function convertColumnType(ByVal columnType As String) As String

    ' 正規表現オブジェクトを生成する
    Static reg As Object
    
    If reg Is Nothing Then
    
        ' IE5.0以上必須
        Set reg = CreateObject("VBScript.RegExp")
        
        If reg Is Nothing Then
        
            err.Raise ConstantsError.ERR_NUMBER_REG_EXP_NOT_CREATED _
                    , "" _
                    , ConstantsError.ERR_DESC_REG_EXP_NOT_CREATED
        
        End If
    
        With reg
            ' パターン文字列
            .Pattern = "\(.*?\)"
            ' 大文字と小文字を区別しない
            .IgnoreCase = True
            ' 文字列全体を検索
            .Global = True
            
        End With
    
    End If
    
    convertColumnType = reg.replace(columnType, "")
    
End Function
