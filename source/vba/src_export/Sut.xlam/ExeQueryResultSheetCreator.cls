VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeQueryResultSheetCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリ結果シートを生成するクラス
'
' 作成者　：Ison
' 履歴　　：2013/06/23 新規作成
'
' 特記事項：
' *********************************************************

' =========================================================
' プロパティ
' =========================================================
' アプリケーション設定情報
Private m_applicationSetting As ValApplicationSetting
' 処理対象ワークブック
Private m_book               As Workbook
' 処理対象ワークブック
Private m_sheet              As Worksheet

Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
End Property

Public Property Let sheet(ByVal vNewValue As Variant)

    Set m_sheet = vNewValue
    
End Property

Public Property Let book(ByVal vNewValue As Variant)

    Set m_book = vNewValue
    
End Property

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
' ▽クエリ結果シート判定
'
' 概要　　　：対象シートが
' 　　　　　　クエリ結果シートであるかを判定する。
'
' 引数　　　：sheet 対象シート
' 戻り値　　：True テーブルシートの場合
' =========================================================
Public Function isQueryResultSheet(ByRef sheet As Worksheet) As Boolean

    ' テーブルシート判定用のオートシェイプ
    Dim s As Shape

    ' 戻り値を初期化する
    isQueryResultSheet = False
    
    ' ------------------------------------------------
    ' エラーが発生しても継続する
    On Error Resume Next

    Set s = sheet.Shapes(ConstantsTable.QUERY_RESULT_SHEET_CHECK_IMAGE)
    
    ' 取得できた場合
    If Not s Is Nothing Then
    
        isQueryResultSheet = True
        Exit Function
    End If
    
    On Error GoTo 0
    ' ------------------------------------------------

End Function

' =========================================================
' ▽シートを生成する
'
' 概要　　　：シートを生成する
' 引数　　　：
'
' =========================================================
Public Function createSheet() As Worksheet

    Dim sheetName As String
    sheetName = ExcelUtil.convertUniqSheetName(m_book, ConstantsTable.QUERY_RESULT_SHEET_DEFAULT_NAME)
    
    ' テーブル名
    Dim tableName     As String
    ' シートオブジェクト
    Dim sheet         As Worksheet
    
    ' シートの削除（存在していたら）と挿入を実行する
    deleteAndInsertSheet sheetName
    
    ' 新たにコピーされたシートオブジェクトを取得する
    Set sheet = m_book.Worksheets(sheetName)
    
    ' ▽新シートにテーブル定義情報を埋め込む
    With sheet

        ' シートを可視にする
        .visible = xlSheetVisible

    End With

    Set m_sheet = sheet
    Set m_book = sheet.parent
    Set createSheet = sheet
    
End Function

' =========================================================
' ▽シートの削除・挿入
'
' 概要　　　：シートが既に存在する場合は削除し、挿入を行う。
' 引数　　　：tableSheet      テーブル情報
' 　　　　　　targetSheetName 対象シート名
'
' =========================================================
Private Sub deleteAndInsertSheet(ByRef targetSheetName As String)

    ' コピーシート
    Dim copySheetName As String
    copySheetName = ConstantsTable.QUERY_RESULT_SHEET_NAME_TEMPLATE

    ' 削除されたシートの位置に再びシートを作成したいので
    ' 削除対象シートの位置を調べ隣合わせになっているシートオブジェクトを取得する。
    ' （シートのコピー関数が
    ' 　あるシートを基点としてコピーするといった方法にしか対応していないためこのような処理が必要になる）
    On Error Resume Next

    ' 削除対象シート
    Dim targetSheet     As Worksheet
    ' 削除対象シートの左隣にあるシート
    Dim targetSheetPrev As Worksheet
    ' 削除対象シートの右隣にあるシート
    Dim targetsheetNext As Worksheet

    ' 削除対象シートを取得する
    Set targetSheet = m_book.Worksheets(targetSheetName)

    ' 削除対象シートが取得できた場合（元々存在していない場合もある）
    If Not targetSheet Is Nothing Then

        ' ブック内にシートの数が1つだけの場合（削除対象シート含めて）
        If m_book.Worksheets.count = 1 Then
        
            ' 何もしない
        ' 削除対象シートが左端の場合
        ElseIf targetSheet.index = 1 Then

            Set targetsheetNext = m_book.Worksheets(2)

        ' 削除対象シートが右端の場合
        ElseIf targetSheet.index = m_book.Worksheets.count Then

            Set targetSheetPrev = m_book.Worksheets(targetSheet.index - 1)

        ' 削除対象シートが中間（端にはない）
        Else
        
            Set targetSheetPrev = m_book.Worksheets(targetSheet.index - 1)
        End If


    End If

    ' Excelではシートの数が0になるのを許さないので、この状態でシートを削除すると
    ' ダミーのシートを勝手に挿入してしまう。
    ' したがって、対象のブックのシート数が1つの場合は、前もってシートを挿入しておき、処理終了後に削除するようにする
    Dim dummySheetName As String
    
    If m_book.Worksheets.count = 1 Then
    
        m_book.Worksheets.Add
        dummySheetName = m_book.ActiveSheet.name
    End If
    
    ' シート削除実行
    ExcelUtil.deleteSheet m_book, targetSheetName
    
    ' 左隣のシートが取得できている場合
    If Not targetSheetPrev Is Nothing Then
    
        ExcelUtil.copySheet SutWorkbook _
                          , copySheetName _
                          , m_book _
                          , targetSheetName _
                          , targetSheetPrev.name _
                          , "after"
    
    ' 右隣のシートが取得できている場合
    ElseIf Not targetsheetNext Is Nothing Then
    
        ExcelUtil.copySheet SutWorkbook _
                          , copySheetName _
                          , m_book _
                          , targetSheetName _
                          , targetsheetNext.name _
                          , "before"
    
    ' 削除対象シートが存在していない場合
    Else
    
        ExcelUtil.copySheetAppend SutWorkbook _
                                , copySheetName _
                                , m_book _
                                , targetSheetName
        
    
    End If
    
    ' ダミーシートが挿入されている場合、ダミーシートを削除する
    If dummySheetName <> "" Then
    
        m_book.Worksheets(dummySheetName).delete
    End If
    
    On Error GoTo 0

End Sub

' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：recordOffset レコード反映時のオフセット位置
' 　　　　　　columnOffset レコード反映時のオフセット位置
'
' 戻り値　　：
' =========================================================
Public Sub insertEmptyCellForResultCount(ByVal recordOffset As Long, ByVal columnOffset As Long)

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_RESULT_ROW _
                          , ConstantsTable.QUERY_RESULT_COL)
                  

    ' コピー先のRangeオブジェクトを取得する
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address)

    ' 入力領域をコピーする
    srcRange.copy Destination:=desRange
    
    ' 高さをコピーする
    ExcelUtil.changeRowHeight desRange, srcRange.RowHeight
    
End Sub
' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：recordOffset レコード反映時のオフセット位置
' 　　　　　　columnOffset レコード反映時のオフセット位置
'
' 戻り値　　：
' =========================================================
Public Sub insertEmptyCellForError(ByVal recordOffset As Long, ByVal columnOffset As Long)

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_ERROR_ROW _
                          , ConstantsTable.QUERY_ERROR_COL)

    ' コピー先のRangeオブジェクトを取得する
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address)

    ' 入力領域をコピーする
    srcRange.copy Destination:=desRange
    
    ' 高さをコピーする
    ExcelUtil.changeRowHeight desRange, srcRange.RowHeight

End Sub

' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：recordOffset レコード反映時のオフセット位置
' 　　　　　　columnOffset レコード反映時のオフセット位置
'
' 戻り値　　：
' =========================================================
Public Sub insertEmptyCellForTitle(ByVal recordOffset As Long, ByVal columnOffset As Long)

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_TITLE_ROW _
                          , ConstantsTable.QUERY_TITLE_COL)
                  

    ' コピー先のRangeオブジェクトを取得する
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address)

    ' 入力領域をコピーする
    srcRange.copy Destination:=desRange
    
    ' 高さをコピーする
    ExcelUtil.changeRowHeight desRange, srcRange.RowHeight
    
End Sub

' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：recordOffset レコード反映時のオフセット位置
' 　　　　　　columnOffset レコード反映時のオフセット位置
'
' 戻り値　　：
' =========================================================
Public Sub insertEmptyCellForResultSet(ByVal recordOffset As Long, ByVal columnOffset As Long)

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_RESULTSET_ROW _
                          , ConstantsTable.QUERY_RESULTSET_COL)
                  

    ' コピー先のRangeオブジェクトを取得する
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address)

    ' 入力領域をコピーする
    srcRange.copy Destination:=desRange
    
    ' 高さをコピーする
    ExcelUtil.changeRowHeight desRange, srcRange.RowHeight
    
End Sub

' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：recordOffset レコード反映時のオフセット位置
' 　　　　　　recordSize   レコードサイズ
' 　　　　　　columnOffset カラム反映時のオフセット位置
' 　　　　　　columnSize   カラムサイズ
' 　　　　　　isGroup      グループ化の有無
'
' 戻り値　　：
' =========================================================
Public Sub insertEmptyCellForRecord(ByVal recordOffset As Long, ByVal recordSize As Long, ByVal columnOffset As Long, ByVal columnSize As Long _
                                    , Optional ByVal isGroup As Boolean = True)

    ' カラムの数
    Dim columnCount As Long
    ' シートの最大列数
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(m_sheet)
    
    ' カラムの数がシートの最大列数を超えていないかをチェックする
    If columnSize > sheetColSize - columnOffset + 1 Then
    
        ' 最大列数を超えている場合、列数を調整する
        columnCount = sheetColSize - columnOffset + 1
        
    Else
    
        ' 最大列数を超えていない場合、そのまま設定
        columnCount = columnSize
    End If

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_RECORD_ROW _
                          , ConstantsTable.QUERY_RECORD_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset + recordSize - 1 _
                                    , columnOffset + columnCount - 1).Address)

    ' 入力領域をコピーする
    srcRange.copy Destination:=desRange

    ' ------------------------------------------------------------
    ' セル書式の設定
    ' ------------------------------------------------------------
    ' フォントを設定
    desRange.Font.name = m_applicationSetting.cellFontName
    ' フォントサイズを設定
    desRange.Font.size = m_applicationSetting.cellFontSize
    ' 折り返しを設定
    desRange.WrapText = m_applicationSetting.cellWordwrap
    ' 列幅を設定
    ExcelUtil.changeColWidth desRange, m_applicationSetting.cellWidth
    ' 行高を設定
    If m_applicationSetting.lineHeightAutoAdjust Then
        ExcelUtil.changeRowHeight desRange, -1
    Else
        ExcelUtil.changeRowHeight desRange, m_applicationSetting.cellHeight
    End If
    

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_ROWNUMBER_ROW _
                          , ConstantsTable.QUERY_ROWNUMBER_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                     , ConstantsTable.QUERY_COLUMN_TITLE_COL).Address & ":" & _
                        m_sheet.Cells(recordOffset + recordSize - 1 _
                                    , ConstantsTable.QUERY_COLUMN_TITLE_COL).Address)

    ' レコード番号をコピーする
    srcRange.copy Destination:=desRange

    desRange.Formula = "=ROW() - " & (recordOffset - 1)
    
    If isGroup Then
        ' グループ化する
        desRange.Rows.Group
    End If

End Sub


' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：recordOffset レコード反映時のオフセット位置
' 　　　　　　recordSize   レコードサイズ
' 　　　　　　columnOffset カラム反映時のオフセット位置
' 　　　　　　columnSize   カラムサイズ
'
' 戻り値　　：
' =========================================================
Public Sub insertEmptyCellForColumn(ByVal recordOffset As Long, ByVal columnOffset As Long, ByVal columnSize As Long)

    ' カラムの数
    Dim columnCount As Long
    ' シートの最大列数
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(m_sheet)
    
    ' カラムの数がシートの最大列数を超えていないかをチェックする
    If columnSize > sheetColSize - columnOffset + 1 Then
    
        ' 最大列数を超えている場合、列数を調整する
        columnCount = sheetColSize - columnOffset + 1
        
    Else
    
        ' 最大列数を超えていない場合、そのまま設定
        columnCount = columnSize
    End If

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_HEADER_ROW _
                          , ConstantsTable.QUERY_HEADER_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset + columnCount - 1).Address)

    ' 入力領域をコピーする
    srcRange.copy Destination:=desRange

    ' ------------------------------------------------------------
    ' セル書式の設定
    ' ------------------------------------------------------------
    ' フォントを設定
    desRange.Font.name = m_applicationSetting.cellFontName
    ' フォントサイズを設定
    desRange.Font.size = m_applicationSetting.cellFontSize
    ' 折り返しを設定
    desRange.WrapText = m_applicationSetting.cellWordwrap
    ' 列幅を設定
    ExcelUtil.changeColWidth desRange, m_applicationSetting.cellWidth
    ' 行高を設定
    If m_applicationSetting.lineHeightAutoAdjust Then
        ExcelUtil.changeRowHeight desRange, -1
    Else
        ExcelUtil.changeRowHeight desRange, m_applicationSetting.cellHeight
    End If

End Sub

' =========================================================
' ▽レコードリストのシート反映（レコード用）
'
' 概要　　　：レコードリストをシートに反映する
' 引数　　　：recordList   レコードリスト
' 　　　　　　recordOffset レコード反映時のオフセット位置
'
' 戻り値　　：True  正常
' 　　　　　　False シートへの反映失敗
'
' =========================================================
Public Function reflectSheetForRecord(ByRef recordList As Variant _
                                    , ByVal recordOffset As Long) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = True

    ' レコードサイズ
    Dim rowSize As Long
    ' カラムサイズ
    Dim colSize As Long

    ' 配列からレコードサイズを取得する
    rowSize = VBUtil.arraySize(recordList)
    ' 配列からカラムサイズを取得する
    colSize = VBUtil.arraySize(recordList, 2)

    ' ------------------------------------------------------------
    ' シートの範囲チェック
    ' ------------------------------------------------------------
    If ExcelUtil.checkOverMaxRow(m_sheet, recordOffset, rowSize) = False Then
    
        ' サイズをオーバーしているので調整する
        rowSize = ExcelUtil.getSizeOfSheetRow(m_sheet) - recordOffset + 1
    
        ' 全レコード反映できないので戻り値をFalseに設定
        ret = False
    End If
        
    ' ------------------------------------------------------------
    ' 配列の内容をシートに反映する
    ' ------------------------------------------------------------
    ExcelUtil.copyArrayToCells recordList, m_sheet, recordOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL, rowSize, colSize

    reflectSheetForRecord = ret
    
    ' 戻り値が False の場合、エラーを発行する
    If ret = False Then
    
        err.Raise ERR_NUMBER_OUT_OF_RANGE_SHEET _
                , _
                , ERR_DESC_OUT_OF_RANGE_SHEET
    End If

End Function


' =========================================================
' ▽レコードリストのシート反映（カラム用）
'
' 概要　　　：レコードリストをシートに反映する
' 引数　　　：headerList   レコードリスト
' 　　　　　　headerOffset レコード反映時のオフセット位置
'
' 戻り値　　：True  正常
' 　　　　　　False シートへの反映失敗
'
' =========================================================
Public Function reflectSheetForColumn(ByRef headerList As Variant _
                                    , ByVal headerOffset As Long) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = True

    ' レコードサイズ
    Dim rowSize As Long: rowSize = 1
    ' カラムサイズ
    Dim colSize As Long

    ' 配列からカラムサイズを取得する
    colSize = VBUtil.arraySize(headerList)

    ' ------------------------------------------------------------
    ' シートの範囲チェック
    ' ------------------------------------------------------------
    If ExcelUtil.checkOverMaxCol(m_sheet, headerOffset, colSize) = False Then
    
        ' サイズをオーバーしているので調整する
        colSize = ExcelUtil.getSizeOfSheetCol(m_sheet) - headerOffset + 1
    
        ' 全レコード反映できないので戻り値をFalseに設定
        ret = False
    End If

    ' ------------------------------------------------------------
    ' 配列の内容をシートに反映する
    ' ------------------------------------------------------------
    ExcelUtil.copyArrayToCellsForColumns headerList, m_sheet, headerOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL, colSize

    reflectSheetForColumn = ret

    
    ' 戻り値が False の場合、エラーを発行する
    If ret = False Then
    
        err.Raise ERR_NUMBER_OUT_OF_RANGE_SHEET _
                , _
                , ERR_DESC_OUT_OF_RANGE_SHEET
    End If

End Function

' =========================================================
' ▽レコードのオフセット位置取得
'
' 概要　　　：取得したレコードを反映するオフセット位置を取得する。
'
' 引数　　　：appendSheet 最後尾への追加有無
' 戻り値　　：レコードのオフセット位置
'
' =========================================================
Public Function getRecordOffset(ByVal appendSheet As Boolean) As Long

    ' 戻り値
    Dim recordOffset As Long
    
    ' 既存の行を削除せずに、最後の位置からレコードを追加する
    If appendSheet = True Then
        
        ' 最終入力行を取得し＋１する
        recordOffset = ExcelUtil.getCellEndOfLastInputtedRow(m_sheet, ConstantsTable.QUERY_COLUMN_TITLE_COL) + 1
        
        ' 上記で取得した値がシートの範囲を超えているかをチェックする
        If ExcelUtil.checkOverMaxRow(m_sheet, recordOffset) = False Then
        
            ' エラーを発行する
            err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SHEET _
                    , _
                    , ConstantsError.ERR_DESC_OUT_OF_RANGE_SHEET
        End If
                
    ' 既存の行を削除して、最初の位置からレコードを追加する
    Else
    
        ' 定数からレコードオフセットを取得する
        recordOffset = ConstantsTable.QUERY_COLUMN_OFFSET_ROW
    
    End If

    ' 戻り値を設定する
    getRecordOffset = recordOffset

End Function

