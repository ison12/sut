VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSnapDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スナップショットの比較結果を出力するクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/01/06　新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' イベント
' ________________________________________________________

' =========================================================
' ▽シート作成イベント（開始）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event createSnapshotDiffSheetBegin()

' =========================================================
' ▽シート作成イベント
'
' 概要　　　：
' 引数　　　：sheet シート
'
' =========================================================
Public Event createSnapshotDiffSheet(ByRef sheet As Worksheet)

' =========================================================
' ▽レコードセットの比較イベント（開始）
'
' 概要　　　：
' 引数　　　：count 件数
'
' =========================================================
Public Event diffRecordSetBegin(ByVal count As Long)

' =========================================================
' ▽レコードセットの比較イベント（1件完了）
'
' 概要　　　：
' 引数　　　：count 件数
'
' =========================================================
Public Event diffRecordSetOne()

' =========================================================
' ▽レコードセットの比較イベント（完了）
'
' 概要　　　：
' 引数　　　：count 件数
'
' =========================================================
Public Event diffRecordSet()

' ---------------------------------------------------------
' □プロパティ
' ---------------------------------------------------------
' キャンセルフラグ
Public canceled      As Boolean

' 処理対象ワークブック
Private m_book       As Workbook
' アプリケーション設定情報
Private m_applicationSetting As ValApplicationSetting

' =========================================================
' ▽プロパティ ブック
' =========================================================
Public Property Let book(ByVal vNewValue As Variant)

    Set m_book = vNewValue
End Property

' =========================================================
' ▽アプリケーション設定情報
' =========================================================
Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
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
' ▽実行
'
' 概要　　　：対象となるクエリの実行結果をシートに反映する
' 引数　　　：diff1  クエリリスト
' 　　  　　：diff2  クエリリスト
' 戻り値　　：レコード取得件数
'
' =========================================================
Public Function execute(ByRef diff1 As ValSnapRecordsSet, ByRef diff2 As ValSnapRecordsSet) As Long

    On Error GoTo err
    
    Dim diffSheet As Worksheet
    
    If diff1.list.count <> diff2.list.count Then
        ' 結果セットの件数が違うのはエラー（発生しないはず）
        err.Raise _
            ConstantsError.ERR_NUMBER_SNAP_DIFF__EXEC_ERROR, _
            err.Source, _
            ConstantsError.ERR_DESC_SNAP_DIFF__EXEC_ERROR & vbNewLine & _
            "結果セットの件数が異なるためエラーになりました。", _
            err.HelpFile, _
            err.HelpContext
    End If
    
    ' シートを生成する
    Set diffSheet = createSheet
    
    ' キャンセルフラグをチェックする
    checkCancel
    
    ' 比較実行
    diffRecordsSet diffSheet, diff1, diff2

    Exit Function

err:

    ' エラーハンドラで別の関数を呼び出すとエラー情報が消えてしまうことがあるので
    ' 構造体にエラー情報を保存しておく
    Dim errT As errInfo: errT = VBUtil.swapErr

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ▽レコードセット比較
'
' 概要　　　：レコードセットの比較を実行する
' 引数　　　：sheet  シート
'           ：diff1  比較対象1
' 　　  　　：diff2  比較対象2
' 戻り値　　：レコード取得件数
'
' =========================================================
Private Function diffRecordsSet(ByRef sheet As Worksheet, ByRef diff1 As ValSnapRecordsSet, ByRef diff2 As ValSnapRecordsSet)

    Dim recordSetIndex As Long
    
    Dim diff1Rec As ValSnapRecords
    Dim diff2Rec As ValSnapRecords
    
    Dim offsetRow As Long
    offsetRow = ConstantsSnapDiff.RESULT_START_ROW
    
    ' ◇レコードセットの比較イベント開始
    RaiseEvent diffRecordSetBegin(diff1.list.count)
    
    ' 名前をセルに書き込む
    sheet.Cells(ConstantsSnapDiff.NAME_ROW, ConstantsSnapDiff.NAME_COL) = diff1.name
    ' 全体の変更有無をセルに書き込む
    setModifyCell sheet, ConstantsSnapDiff.MODIFY_ALL_ROW, ConstantsSnapDiff.MODIFY_ALL_COL, False ' 全体

    recordSetIndex = 0
    For Each diff1Rec In diff1.list
    
        ' キャンセルフラグをチェックする
        checkCancel
    
        Set diff2Rec = diff2.list.item(recordSetIndex + 1)
        
        offsetRow = diffRecords(recordSetIndex, sheet, offsetRow, diff1Rec, diff2Rec)

        recordSetIndex = recordSetIndex + 1
        
        ' ◇レコードセットの比較イベント
        RaiseEvent diffRecordSetOne

    Next
        
    ' ◇レコードセットの比較イベント完了
    RaiseEvent diffRecordSet

End Function

' =========================================================
' ▽レコードリスト比較
'
' 概要　　　：レコードセットの比較を実行する（SQLの実行単位）
' 引数　　　：recordSetIndex  レコードセット開始行
'           ：sheet           シート
' 　　  　　：offsetRow       開始行
' 　　  　　：diff1Rec        比較対象1
' 　　  　　：diff2Rec        比較対象2
' 戻り値　　：レコード取得件数
'
' =========================================================
Private Function diffRecords(ByVal recordSetIndex As Long, ByRef sheet As Worksheet, ByVal offsetRow As Long, ByRef diff1Rec As ValSnapRecords, ByRef diff2Rec As ValSnapRecords) As Long

    Dim totalCount    As Long
    Dim noChangeCount As Long
    Dim insertCount   As Long
    Dim updateCount   As Long
    Dim deleteCount   As Long

    Dim offsetRowHeader As Long: offsetRowHeader = offsetRow + ConstantsSnapDiff.RESULT_START_OFFSET_HEADER
    Dim offsetRowReocrd As Long: offsetRowReocrd = offsetRow + ConstantsSnapDiff.RESULT_START_OFFSET_RECORD

    Dim endOffsetRow As Long

    Dim header1Index As Long
    Dim header2Index As Long
    
    Dim recordIndex As Long
    Dim recordNum As Long
    
    Dim pkStr As String

    Dim diff1RecMap As ValCollection
    Dim diff2RecMap As ValCollection
    
    Dim diffDeletedRecord As Variant
    
    Dim diff2RecIndex As Long
    Dim headerIndex As Long
    
    Dim isUpdate As Boolean
    Dim isOutRecord As Boolean

    diffRecords = offsetRow

    If UBound(diff1Rec.headerList) - LBound(diff1Rec.headerList) <> _
       UBound(diff2Rec.headerList) - LBound(diff2Rec.headerList) Then
        ' ヘッダの件数が違うのはエラー（通常は発生しないはず）
        err.Raise _
            ConstantsError.ERR_NUMBER_SNAP_DIFF__EXEC_ERROR, _
            err.Source, _
            ConstantsError.ERR_DESC_SNAP_DIFF__EXEC_ERROR & vbNewLine & _
            "結果セットのヘッダの件数が異なるためエラーになりました。", _
            err.HelpFile, _
            err.HelpContext
    End If
    
    For header1Index = LBound(diff1Rec.headerList) To UBound(diff1Rec.headerList)
    
        If diff1Rec.headerList(header1Index) <> diff2Rec.headerList(header1Index) Then
            ' ヘッダ名が違うのはエラー（通常は発生しないはず）
            err.Raise _
                ConstantsError.ERR_NUMBER_SNAP_DIFF__EXEC_ERROR, _
                err.Source, _
                ConstantsError.ERR_DESC_SNAP_DIFF__EXEC_ERROR & vbNewLine & _
                "結果セットのヘッダ名が異なるためエラーになりました。", _
                err.HelpFile, _
                err.HelpContext
        End If
        
    Next
    
    ' --------------------------------------------
    ' 比較範囲をセルに出力する
    ' --------------------------------------------
    insertResultRange sheet, offsetRow
    
    ' 比較有無
    setModifyCell sheet, offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.MODIFY_COL, False ' 個別
    
    ' --------------------------------------------
    ' ヘッダをセルに出力する
    ' --------------------------------------------
    ' レコードリストの空欄の領域を出力
    insertHeaderListCell sheet, offsetRowHeader, VBUtil.arraySize(diff1Rec.headerList)
    ' 内容を出力
    ExcelUtil.copyArrayToCellsForColumns diff1Rec.headerList, sheet, offsetRowHeader, ConstantsSnapDiff.HEADER_COL
    
    Set diff1RecMap = diff1Rec.convertToMapKeyIsPkValIsReocrdIndex
    Set diff2RecMap = diff2Rec.convertToMapKeyIsPkValIsReocrdIndex
    
    ' diff1を基準に比較する
    If Not IsEmpty(diff1Rec.recordList) Then
    
        ' --------------------------------------------
        ' レコードリストをセルに出力する
        ' --------------------------------------------
        ' レコードリストの空欄の領域を出力
        insertRecordListCell sheet, offsetRowReocrd, VBUtil.arraySize(diff1Rec.recordList), VBUtil.arraySize(diff1Rec.recordList, 2)
        ' 内容を出力
        ExcelUtil.copyArrayToCells diff1Rec.recordList, sheet, offsetRowReocrd, ConstantsSnapDiff.RECORD_COL
    
        totalCount = totalCount + VBUtil.arraySize(diff1Rec.recordList)
    
        recordIndex = 0
        For recordIndex = LBound(diff1Rec.recordList) To UBound(diff1Rec.recordList)
            
            pkStr = diff1Rec.getPkStr(recordIndex)
            
            If diff2RecMap.exist(pkStr) = False Then
                ' PKに該当するレコードが存在しない場合
                ' INSERT
                sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_INSERT

                ' 背景色を変更する
                sheet.Range( _
                    Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL).Address, _
                    Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL + VBUtil.arraySize(diff1Rec.headerList) - 1).Address _
                ).Interior.Color = _
                    RGB( _
                        ConstantsSnapDiff.INSERT_COLOR_R, _
                        ConstantsSnapDiff.INSERT_COLOR_G, _
                        ConstantsSnapDiff.INSERT_COLOR_B)
                        
                ' 変更有無をセルに書き込む
                setModifyCell sheet, ConstantsSnapDiff.MODIFY_ALL_ROW, ConstantsSnapDiff.MODIFY_ALL_COL, True ' 全体
                setModifyCell sheet, offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.MODIFY_COL, True ' 個別
                
                insertCount = insertCount + 1
            Else
                ' PKに該当するレコードが存在する場合
                
                diff2RecIndex = diff2RecMap.getItem(pkStr, vbLong)
                
                isUpdate = False
                For headerIndex = LBound(diff1Rec.headerList) To UBound(diff1Rec.headerList)
                    
                    If diff1Rec.recordList(recordIndex, headerIndex) <> diff2Rec.recordList(diff2RecIndex, headerIndex) Then
                    
                        ' 値が異なる場合、変更の前の値をセルにコメントとして埋め込む
                        ExcelUtil.addCommentForWorkSheet sheet, _
                            offsetRowReocrd, _
                            ConstantsSnapDiff.RECORD_COL + headerIndex, _
                            diff2Rec.recordList(diff2RecIndex, headerIndex)
                            
                        ' 背景色を変更する
                        sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL + headerIndex).Interior.Color = _
                            RGB( _
                                ConstantsSnapDiff.UPDATE_COLOR_R, _
                                ConstantsSnapDiff.UPDATE_COLOR_G, _
                                ConstantsSnapDiff.UPDATE_COLOR_B)
                    
                        isUpdate = True
                    End If
                    
                Next
                
                If isUpdate Then
                    ' UPDATE
                    sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_UPDATE
                    ' 変更有無をセルに書き込む
                    setModifyCell sheet, ConstantsSnapDiff.MODIFY_ALL_ROW, ConstantsSnapDiff.MODIFY_ALL_COL, True ' 全体
                    setModifyCell sheet, offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.MODIFY_COL, True ' 個別
                
                    updateCount = updateCount + 1
                Else
                    ' 変更がない場合
                    sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_NOCHANGE
                    
                    noChangeCount = noChangeCount + 1
                End If
                
            End If
            
            ' 連番
            recordNum = recordNum + 1
            sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_NUM_COL) = recordNum
            
            offsetRowReocrd = offsetRowReocrd + 1
            
        Next
        
        isOutRecord = True
    
    End If
    
    ' diff2を基準に比較する
    If Not IsEmpty(diff2Rec.recordList) Then
        
        recordIndex = 0
        For recordIndex = LBound(diff2Rec.recordList) To UBound(diff2Rec.recordList)
            
            pkStr = diff2Rec.getPkStr(recordIndex)
            
            If diff1RecMap.exist(pkStr) = False Then
                ' PKに該当するレコードが存在しない場合
                ' DELETE
                
                diffDeletedRecord = convert2to1Array(diff2Rec.recordList, recordIndex)
                
                ' --------------------------------------------
                ' レコードリストをセルに出力する
                ' 1次元配列を出力するので、ExcelUtil.copyArrayToCellsForColumnsを使用する
                ' --------------------------------------------
                ' レコードリストの空欄の領域を出力
                insertRecordListCell sheet, offsetRowReocrd, 1, VBUtil.arraySize(diffDeletedRecord)
                ' 内容を出力
                ExcelUtil.copyArrayToCellsForColumns diffDeletedRecord, sheet, offsetRowReocrd, ConstantsSnapDiff.RECORD_COL
                sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_DELETE
    
                deleteCount = deleteCount + 1
                totalCount = totalCount + 1
            
                ' 背景色を変更する
                sheet.Range( _
                    Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL).Address, _
                    Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL + VBUtil.arraySize(diffDeletedRecord) - 1).Address _
                ).Interior.Color = _
                    RGB( _
                        ConstantsSnapDiff.DELETE_COLOR_R, _
                        ConstantsSnapDiff.DELETE_COLOR_G, _
                        ConstantsSnapDiff.DELETE_COLOR_B)
            
                ' 変更有無をセルに書き込む
                setModifyCell sheet, ConstantsSnapDiff.MODIFY_ALL_ROW, ConstantsSnapDiff.MODIFY_ALL_COL, True ' 全体
                setModifyCell sheet, offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.MODIFY_COL, True ' 個別
                
                ' 連番
                recordNum = recordNum + 1
                sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_NUM_COL) = recordNum
            
                isOutRecord = True
                offsetRowReocrd = offsetRowReocrd + 1

            End If
            
        Next
        
    End If
    
    If isOutRecord Then
        ' レコード出力がある
        endOffsetRow = offsetRowReocrd + 1
    
    Else
        ' レコード出力がない（offsetRowRecord変数が一度もインクリメントされていないので、+1余分に加算）
        endOffsetRow = offsetRowReocrd + 2
        
        sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_NORECORD
        
        ' 最初の列の罫線を解除する
        sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL + 1).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
        sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL + 1).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone

    End If
    
    Dim pkListStr As String
    
    ' SQL番号
    sheet.Cells(offsetRow + RESULT_START_OFFSET_SQL, ConstantsSnapDiff.SQL_NUM_COL) = recordSetIndex + 1
    ' SQL
    sheet.Cells(offsetRow + RESULT_START_OFFSET_SQL, ConstantsSnapDiff.SQL_COL) = diff1Rec.sql
    ' PKリスト
    sheet.Cells(offsetRow + RESULT_START_OFFSET_SQL, ConstantsSnapDiff.PKEY_COL) = Join(diff1Rec.convertPkListToArray, vbNewLine)
    ' 備考
    sheet.Cells(offsetRow + RESULT_START_OFFSET_SQL, ConstantsSnapDiff.MEMO_COL) = diff1Rec.memo
    
    ' 件数などをセルに書き込む
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.TOTAL_COUNT_COL) = totalCount
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.NOCHANGE_COUNT_COL) = noChangeCount
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.INSERT_COUNT_COL) = insertCount
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.UPDATE_COUNT_COL) = updateCount
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.DELETE_COUNT_COL) = deleteCount
    
    ' ヘッダとレコード部の行をグループ化する
    sheet.Range( _
        Cells(offsetRow + RESULT_START_OFFSET_HEADER, 1).Address, _
        Cells(endOffsetRow - 2, 1).Address).Rows.Group
    
    diffRecords = endOffsetRow
    
End Function

' =========================================================
' ▽シートを生成する
'
' 概要　　　：指定されたテーブルのシートを生成する
' 引数　　　：
' 戻り値    ：シート
'
' =========================================================
Private Function createSheet() As Worksheet

    ' シート名
    Dim sheetName     As String

    ' ◇スナップショット比較シート生成イベント
    RaiseEvent createSnapshotDiffSheetBegin

    sheetName = "比較結果_" & Format(Now, "yyyymmdd_hhmmss")

    ' コピーシート
    Dim copySheetName As String
    copySheetName = ConstantsSnapDiff.SHEET_NAME_TEMPLATE

    ExcelUtil.copySheetAppend SutWorkbook _
                            , copySheetName _
                            , m_book _
                            , sheetName
                            
    ' 新たにコピーされたシートオブジェクトを取得する
    Set createSheet = m_book.Worksheets(sheetName)
                            
    ' ▽新シートにテーブル定義情報を埋め込む
    With createSheet

        ' シートを可視にする
        .visible = xlSheetVisible

    End With
    
    ' ◇スナップショット比較シート生成完了イベント
    RaiseEvent createSnapshotDiffSheet(createSheet)

End Function

' =========================================================
' ▽比較結果セル範囲を挿入
'
' 概要　　　：
' 引数　　　：sheet        ワークシート
' 　　　　　　recordOffset レコード反映時のオフセット位置
'
' 戻り値　　：
' =========================================================
Private Sub insertResultRange(ByRef sheet As Worksheet, ByVal recordOffset As Long)

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                    .Range( _
                        Cells(ConstantsSnapDiff.RESULT_RANGE_START_ROW _
                            , ConstantsSnapDiff.RESULT_RANGE_START_COL).Address & ":" & _
                        Cells(ConstantsSnapDiff.RESULT_RANGE_END_ROW _
                            , ConstantsSnapDiff.RESULT_RANGE_END_COL).Address)
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.RESULT_RANGE_START_COL).Address)

    ' 入力領域をコピーする
    srcRange.Copy Destination:=desRange

End Sub

' =========================================================
' ▽ヘッダセルリストを挿入
'
' 概要　　　：
' 引数　　　：sheet        ワークシート
' 　　　　　　recordOffset レコード反映時のオフセット位置
' 　　　　　　columnSize   カラムサイズ
'
' 戻り値　　：
' =========================================================
Private Sub insertHeaderListCell(ByRef sheet As Worksheet, ByVal recordOffset As Long, ByVal columnSize As Long)

    ' カラムの数
    Dim columnCount As Long
    
    ' シートの最大列数
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(sheet)
    
    ' カラムの数がシートの最大列数を超えていないかをチェックする
    If columnSize > sheetColSize - ConstantsSnapDiff.RECORD_COL + 1 Then
    
        ' 最大列数を超えている場合、列数を調整する
        columnCount = sheetColSize - ConstantsSnapDiff.RECORD_COL + 1
        
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
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_HEADER _
                          , ConstantsSnapDiff.HEADER_COL)
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.HEADER_COL).Address & ":" & _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.HEADER_COL + columnCount - 1).Address)

    ' 入力領域をコピーする
    srcRange.Copy Destination:=desRange

    ' カラム幅の調整
    ExcelUtil.changeColWidth desRange, srcRange.EntireColumn.ColumnWidth

    ' ------------------------------------------------------------
    ' レコード番号
    ' ------------------------------------------------------------
    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_HEADER _
                          , ConstantsSnapDiff.RECORD_NUM_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.RECORD_NUM_COL).Address)

    srcRange.Copy Destination:=desRange

    ' ------------------------------------------------------------
    ' レコード操作内容
    ' ------------------------------------------------------------
    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_HEADER _
                          , ConstantsSnapDiff.RECORD_MODIFY_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.RECORD_MODIFY_COL).Address)
                            
    srcRange.Copy Destination:=desRange

End Sub

' =========================================================
' ▽レコードセルリストを挿入
'
' 概要　　　：
' 引数　　　：sheet        ワークシート
' 　　　　　　recordOffset レコード反映時のオフセット位置
' 　　　　　　recordSize   レコードサイズ
' 　　　　　　columnSize   カラムサイズ
'
' 戻り値　　：
' =========================================================
Private Sub insertRecordListCell(ByRef sheet As Worksheet, ByVal recordOffset As Long, ByVal recordSize As Long, ByVal columnSize As Long)

    ' カラムの数
    Dim columnCount As Long
    
    ' シートの最大列数
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(sheet)
    
    ' カラムの数がシートの最大列数を超えていないかをチェックする
    If columnSize > sheetColSize - ConstantsSnapDiff.RECORD_COL + 1 Then
    
        ' 最大列数を超えている場合、列数を調整する
        columnCount = sheetColSize - ConstantsSnapDiff.RECORD_COL + 1
        
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
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_RECORD _
                          , ConstantsSnapDiff.RECORD_COL)
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        Cells(recordOffset + recordSize - 1 _
                            , ConstantsSnapDiff.RECORD_COL + columnCount - 1).Address)

    ' 入力領域をコピーする
    srcRange.Copy Destination:=desRange

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

    ' ------------------------------------------------------------
    ' レコード番号
    ' ------------------------------------------------------------
    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_RECORD _
                          , ConstantsSnapDiff.RECORD_NUM_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        Cells(recordOffset + recordSize - 1 _
                            , ConstantsSnapDiff.RECORD_NUM_COL).Address)

    srcRange.Copy Destination:=desRange

    ' ------------------------------------------------------------
    ' レコード操作内容
    ' ------------------------------------------------------------
    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_RECORD _
                          , ConstantsSnapDiff.RECORD_MODIFY_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        Cells(recordOffset + recordSize - 1 _
                            , ConstantsSnapDiff.RECORD_MODIFY_COL).Address)
                            
    srcRange.Copy Destination:=desRange

End Sub

' =========================================================
' ▽変更有無のセル情報を設定する
'
' 概要　　　：
' 引数　　　：sheet    ワークシート
' 　　　　　　row      行
' 　　　　　　col      列
' 　　　　　　modify   変更有無
'
' 戻り値　　：
' =========================================================
Private Sub setModifyCell(ByRef sheet As Worksheet, ByVal row As Long, ByVal col As Long, ByVal modify As Boolean)

    Dim targetRange As Range

    ' コピー先のRangeオブジェクトを取得する
    Set targetRange = sheet _
                     .Range(Cells(row, col).Address)

    If modify Then
        targetRange.value = ConstantsSnapDiff.MODIFY_ON
        targetRange.Interior.Color = _
            RGB( _
                ConstantsSnapDiff.MODIFY_CHANGE_COLOR_R, _
                ConstantsSnapDiff.MODIFY_CHANGE_COLOR_G, _
                ConstantsSnapDiff.MODIFY_CHANGE_COLOR_B)
    Else
        targetRange.value = ConstantsSnapDiff.MODIFY_OFF
        targetRange.Interior.Color = _
            RGB( _
                ConstantsSnapDiff.MODIFY_NOCHANGE_COLOR_R, _
                ConstantsSnapDiff.MODIFY_NOCHANGE_COLOR_G, _
                ConstantsSnapDiff.MODIFY_NOCHANGE_COLOR_B)
    End If

End Sub

' =========================================================
' ▽キャンセルフラグチェック
'
' 概要　　　：キャンセルフラグをチェックしキャンセルされていた場合に
' 　　　　　　エラーを発行する。
'
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub checkCancel()

    If canceled = True Then

        err.Raise ConstantsError.ERR_NUMBER_PROC_CANCEL _
                , err.Source _
                , ConstantsError.ERR_DESC_PROC_CANCEL
    End If

End Sub


