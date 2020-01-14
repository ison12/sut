VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' メニュー操作時の処理 - Data
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/03/14　新規作成
'
' 特記事項：
' *********************************************************

' SELECTの再実行に関するレジストリのキー
Private Const REG_SUB_KEY_SELECT_LATEST                    As String = "select_latest"
    
' SELECTの再実行のSQL本体
Private Const REG_SUB_KEY_SELECT_LATEST_SQL                As String = "sql"
' SELECTの再実行の条件指定の場合の追加有無フラグ
Private Const REG_SUB_KEY_SELECT_LATEST_TYPE               As String = "type"
' SELECTの再実行の種類（全て）
Private Const REG_SUB_KEY_SELECT_LATEST_TYPE_ALL           As String = "1"
' SELECTの再実行の種類（条件指定）
Private Const REG_SUB_KEY_SELECT_LATEST_TYPE_CONDITION     As String = "2"
' SELECTの再実行の条件指定の場合の追加有無フラグ
Private Const REG_SUB_KEY_SELECT_LATEST_APPEND             As String = "append"
' SELECTの再実行の条件指定の場合の追加有無フラグ
Private Const REG_SUB_KEY_SELECT_LATEST_APPEND_TRUE        As String = "True"

' セルカラーインデックス　エラー時
Private Const CELL_COLOR_INDEX_ERROR        As Long = 40
' セルカラーインデックス　成功時
Private Const CELL_COLOR_INDEX_SUCCESS      As Long = 15
' セルカラーインデックス　0件更新
Private Const CELL_COLOR_INDEX_ZERO_UPDATED As Long = 36


' プログレスフォーム（イベント受信用に定義）
Private WithEvents frmProgress                  As sutredlib.Progress
Attribute frmProgress.VB_VarHelpID = -1
' SELECT SQL生成フォーム
Private WithEvents frmSelectConditionCreatorVar As frmSelectConditionCreator
Attribute frmSelectConditionCreatorVar.VB_VarHelpID = -1
' レコードの追加・削除フォーム
Private WithEvents frmRecAppender               As frmRecordAppender
Attribute frmRecAppender.VB_VarHelpID = -1
' レコード情報生成
Private WithEvents tableSheetReader             As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1
' 更新系クエリー生成
Private WithEvents update                       As ExeUpdate
Attribute update.VB_VarHelpID = -1
' 更新系クエリー生成（選択）
Private WithEvents updateSelection_             As ExeUpdate
Attribute updateSelection_.VB_VarHelpID = -1
' 参照系クエリー生成
Private WithEvents select_                      As ExeSelect
Attribute select_.VB_VarHelpID = -1
' クエリー実行生成
Private WithEvents queryExecution               As ExeQueryExecution
Attribute queryExecution.VB_VarHelpID = -1
' クエリエディタ
Private WithEvents frmQueryEditor               As sutredlib.queryEditor
Attribute frmQueryEditor.VB_VarHelpID = -1
' 一括実行フォーム
Private WithEvents frmQueryBatch                As sutredlib.DbQueryBatch
Attribute frmQueryBatch.VB_VarHelpID = -1
' クエリ結果フォーム
Private WithEvents frmQueryResult               As sutredlib.queryResult
Attribute frmQueryResult.VB_VarHelpID = -1

' クエリ結果情報
Private QueryResultInfo As sutredlib.QueryResultInfo
' クエリ結果件数
' 処理
Private queryResultProcessCnt As Long
' 成功
Private queryResultSuccessCnt As Long
' エラー
Private queryResultErrorCnt As Long

' アプリケーション設定
Private applicationSetting As ValApplicationSetting

' 削除レコードリスト
Private deleteRecordList As ValCollection

' ________________________________________________________
' イベント定義
' ________________________________________________________
Private Sub frmProgress_OnCancel()

    ' キャンセルフラグをONにする
    If Not tableSheetReader Is Nothing Then
        tableSheetReader.canceled = True
    End If
    
    If Not update Is Nothing Then
        update.canceled = True
    End If
    
    If Not updateSelection_ Is Nothing Then
        updateSelection_.canceled = True
    End If
    
    If Not select_ Is Nothing Then
        select_.canceled = True
    End If
    
    If Not queryExecution Is Nothing Then
        queryExecution.canceled = True
    End If

End Sub

' =========================================================
' ▼クエリエディタ共通の実行処理
' =========================================================
Private Sub onExecuteQueryCommon(ByVal queryList As sutredlib.IGenericCollection, ByVal isAppend As Boolean, ByVal isBeginTran As Boolean)

    On Error GoTo err
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgress.Show ExcelUtil.getApplicationHWnd, True
    Else
        frmProgress.Show ExcelUtil.getApplicationHWnd, False
    End If
    frmProgress.title = "クエリ実行" & createProgressTitleOptionalMessage
    
    ' エディタが開いている間に接続状態が変更されている可能性があるので、コネクションを設定する
    Dim connForQueryEditor As Object
    Set connForQueryEditor = Main.getDBConnection
    
    Set queryExecution.conn = connForQueryEditor
    queryExecution.sheet = ActiveSheet
    queryExecution.applicationSetting = Main.getApplicationSetting ' エディタが開いている間に接続状態が変更されている可能性があるので、アプリケーション設定情報を取得する
    queryExecution.canceled = False
    queryExecution.isBeginTran = isBeginTran
    queryExecution.appendSheet = isAppend
    queryExecution.execute queryList
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing

    frmProgress.Close
    
    frmQueryEditor.enable
    
    Exit Sub
err:
    
    frmProgress.Close
    
    frmQueryEditor.enable
    
    ' エラー発生
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▼クエリエディタイベント - クエリ全て実行
' =========================================================
Private Sub frmQueryEditor_OnExecuteQueryAll(ByVal queryList As sutredlib.IGenericCollection, Optional ByVal isAppend As Boolean, Optional ByVal isBeginTran As Boolean)

    onExecuteQueryCommon queryList, isAppend, isBeginTran

End Sub

Private Sub frmQueryEditor_OnExecuteQuerySelection(ByVal queryList As sutredlib.IGenericCollection, Optional ByVal isAppend As Boolean, Optional ByVal isBeginTran As Boolean)

    onExecuteQueryCommon queryList, isAppend, isBeginTran

End Sub

' =========================================================
' ▼クエリエディタイベント - トランザクションの開始
' =========================================================
Private Sub frmQueryEditor_OnBeginTrans(ByRef isError As Boolean)
    
    ' エディタが開いている間に接続状態が変更されている可能性があるので、コネクションを設定する
    Dim connForQueryEditor As Object
    Set connForQueryEditor = Main.getDBConnection
    
    On Error Resume Next
    
    ' 一旦トランザクションをロールバックする
    connForQueryEditor.RollbackTrans
    connForQueryEditor.begintrans
    
    On Error GoTo 0
    
    isError = False
    
End Sub

' =========================================================
' ▼クエリエディタイベント - コミットの開始
' =========================================================
Private Sub frmQueryEditor_OnCommitTrans(ByRef isError As Boolean)
    
    On Error GoTo err

    ' エディタが開いている間に接続状態が変更されている可能性があるので、コネクションを設定する
    Dim connForQueryEditor As Object
    Set connForQueryEditor = Main.getDBConnection
    
    connForQueryEditor.CommitTrans

    isError = False
    
    Exit Sub
err:
    ' トランザクションが開始されていない可能性あり
    VBUtil.showMessageBoxForWarning "トランザクションが開始されていません。" _
                                    , ConstantsCommon.APPLICATION_NAME _
                                    , err

    isError = True
    

End Sub

' =========================================================
' ▼クエリエディタイベント - ロールバックの開始
' =========================================================
Private Sub frmQueryEditor_OnRollbackTrans(ByRef isError As Boolean)
    
    On Error GoTo err

    ' エディタが開いている間に接続状態が変更されている可能性があるので、コネクションを設定する
    Dim connForQueryEditor As Object
    Set connForQueryEditor = Main.getDBConnection
    
    connForQueryEditor.RollbackTrans

    isError = False
    
    Exit Sub
err:
    
    ' トランザクションが開始されていない可能性あり
    VBUtil.showMessageBoxForWarning "トランザクションが開始されていません。" _
                                    , ConstantsCommon.APPLICATION_NAME _
                                    , err


    isError = True

End Sub

' =========================================================
' ▼クエリ一括フォームイベント - 実行
' =========================================================
Private Sub frmQueryBatch_OnRun(ByVal tableWorksheets As sutredlib.IGenericCollection)

    Dim var As sutredlib.DbQueryBatchTableWorksheet
    Dim tableSheet As ValTableWorksheet
    
    For Each var In tableWorksheets
    
        ' テーブルシートを取得する
        Set tableSheet = var.tableWorksheet
        ' 対象シートを設定する
        Set update.sheet = tableSheet.sheet
        ' 対象シートを設定する
        Set select_.sheet = tableSheet.sheet
        
        ' 対象シートをアクティブにする
        update.sheet.activate
            
        If var.BatchQueryType = BATCH_QUERY_TYPE.insertUpdate Then
        
            ' 登録・更新
            Call insertUpdateAll
        
        ElseIf var.BatchQueryType = BATCH_QUERY_TYPE.insert Then
        
            ' 登録
            Call insertAll
        
        ElseIf var.BatchQueryType = BATCH_QUERY_TYPE.update Then
        
            ' 更新
            Call updateAll
        
        ElseIf var.BatchQueryType = BATCH_QUERY_TYPE.deleteAll Then
        
            ' 削除
            Call deleteAllOfTable
        
        ElseIf var.BatchQueryType = BATCH_QUERY_TYPE.deleteOnSheet Then
        
            ' 削除
            Call deleteAll
        
        ElseIf var.BatchQueryType = BATCH_QUERY_TYPE.selectAll Then
        
            ' 参照
            Call selectAll
        
        ElseIf var.BatchQueryType = BATCH_QUERY_TYPE.selectCondition Then
        
            ' 対象シートを設定する
            Set tableSheetReader.sheet = tableSheet.sheet
            ' 参照（条件あり）
            Call selectCondition
        
        ElseIf var.BatchQueryType = BATCH_QUERY_TYPE.selectReExec Then
        
            ' 参照（再実行）
            Call selectReExecute
        
        End If
        
    Next
    
    showQueryResult
    
End Sub

' =========================================================
' ▼クエリ結果フォームイベント - 詳細情報選択時
' =========================================================
Private Sub frmQueryResult_OnSelectQueryResultDetail(ByVal info As sutredlib.IQueryResultInfo, ByVal detail As sutredlib.IQueryResultDetailInfo)

    Dim sheet As ValTableWorksheet
    Set sheet = info.tableWorksheet
    
    Dim wsheet As Worksheet
    
    If sheet Is Nothing Then
    
        Exit Sub
    End If
    
    Set wsheet = sheet.sheet
    wsheet.activate
    
    On Error Resume Next
    
    If detail.cell <> "" Then
    
        wsheet.Range(detail.cell).activate
    
    End If
    
    On Error GoTo 0

End Sub

' =========================================================
' ▼参照イベント - テーブル情報読み込み開始
' =========================================================
Private Sub select__readTableInfoBegin()

    #If (DEBUG_MODE = 1) Then
    
    #End If

End Sub

' =========================================================
' ▼参照イベント - テーブル情報読み込み完了
' =========================================================
Private Sub select__readTableInfo(tableSheet As ValTableWorksheet)

    #If (DEBUG_MODE = 1) Then
        
        Debug.Print tableSheet.table.schemaTableName
    #End If

    ' クエリ結果情報を生成する
    If Not QueryResultInfo Is Nothing Then
    
        QueryResultInfo.tableWorksheet = tableSheet
        QueryResultInfo.processCount = 0
        QueryResultInfo.successCount = 0
        QueryResultInfo.errorCount = 0
    End If

End Sub

' =========================================================
' ▼参照イベント - カラムリスト読み込み開始
' =========================================================
Private Sub select__readColumnListBegin(ByVal cnt As Long)

    frmProgress.InitPri cnt, "カラム定義読み込み中"
    
    DoEvents
End Sub

' =========================================================
' ▼参照イベント - カラムリスト読み込み
' =========================================================
Private Sub select__readColumn(column As ValDbDefineColumn)

    #If (DEBUG_MODE = 1) Then
        
        Debug.Print column.columnName & " - " & column.dataTypeFormal
    #End If

    frmProgress.InclimentPri
    
    DoEvents
End Sub

' =========================================================
' ▼参照イベント - カラムリスト読み込み完了
' =========================================================
Private Sub select__readColumnList(columnList As sutredlib.GenericCollection)

    frmProgress.PriMessage = "カラム定義読み込み完了"
    
    DoEvents
End Sub

' =========================================================
' ▼参照イベント - SQL実行中
' =========================================================
Private Sub select__ExecuteSql(sql As String)

    frmProgress.InitPri -1, "SQL実行中"
    frmProgress.PriCount = 0
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    DoEvents

    SutWhite.showHourglassWindowOnCenterPt
End Sub

' =========================================================
' ▼参照イベント - SQL実行完了
' =========================================================
Private Sub select__ExecuteSqlSuccessful(sql As String)

    frmProgress.PriMessage = "SQL実行完了"
    
    DoEvents

    SutWhite.closeHourglassWindow
    
End Sub

' =========================================================
' ▼参照イベント - SQL実行エラー
' =========================================================
Private Sub select__ExecuteSqlError(sql As String, sqlState As String, message As String)

    frmProgress.PriMessage = "SQL実行エラー"
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    ' クエリ結果情報に詳細を追加する
    
    If Not QueryResultInfo Is Nothing Then
    
        Dim detail As New sutredlib.QueryResultDetailInfo
        detail.cell = "-"
        detail.sqlState = sqlState
        detail.message = message
        detail.query = sql
        
        QueryResultInfo.addDetail detail
    End If

    DoEvents

    SutWhite.closeHourglassWindow
    
End Sub

' =========================================================
' ▼参照イベント - レコード取得中
' =========================================================
Private Sub select__ReflectionRecordGetBegin(ByVal recordListSize As Long)

    frmProgress.PriMessage = "レコード取得中"

    DoEvents

    SutWhite.showHourglassWindowOnCenterPt

End Sub

' =========================================================
' ▼参照イベント - レコード取得
' =========================================================
Private Sub select__ReflectionRecordGet(ByVal recordListSizeActual As Long, ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriCount = frmProgress.PriCount + recordListSize
    frmProgress.PriMessage = "レコード取得完了"
    
    DoEvents

    SutWhite.closeHourglassWindow

End Sub

' =========================================================
' ▼参照イベント - レコード変換中
' =========================================================
Private Sub select__ReflectionRecordConvertBegin(ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriMessage = "レコード情報変換中"
    
    DoEvents

End Sub

' =========================================================
' ▼参照イベント - レコード変換完了
' =========================================================
Private Sub select__ReflectionRecordConvert(ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriMessage = "レコード情報変換完了"
    
    DoEvents

End Sub

' =========================================================
' ▼参照イベント - レコード反映中
' =========================================================
Private Sub select__ReflectionRecordBegin(ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriMessage = "レコード シート出力中"
    
    DoEvents

End Sub

' =========================================================
' ▼参照イベント - レコード反映完了
' =========================================================
Private Sub select__ReflectionRecord(ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriMessage = "レコード シート出力完了"
    
    DoEvents

End Sub

' =========================================================
' ▼更新イベント - テーブル情報読み込み
' =========================================================
Private Sub update_readTableInfo(table As ValTableWorksheet)

    ' クエリ結果情報を生成する
    Set QueryResultInfo = New sutredlib.QueryResultInfo
    QueryResultInfo.tableWorksheet = table
    QueryResultInfo.processCount = 0
    QueryResultInfo.successCount = 0
    QueryResultInfo.errorCount = 0

End Sub

' =========================================================
' ▼更新イベント - カラムリスト読み込み開始
' =========================================================
Private Sub update_readColumnListBegin(ByVal cnt As Long)

    frmProgress.InitPri cnt, "カラム定義読み込み中"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント - カラムリスト読み込み
' =========================================================
Private Sub update_readColumn(column As ValDbDefineColumn)

    frmProgress.InclimentPri
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント - カラムリスト読み込み完了
' =========================================================
Private Sub update_readColumnList(columnList As sutredlib.GenericCollection)

    frmProgress.PriMessage = "カラム定義読み込み完了"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント - レコードリストサイズ読み込み
' =========================================================
Private Sub update_readRecordListSize(ByVal cnt As Long)

    frmProgress.InitPri cnt, "レコードリストサイズ読み込み完了"
End Sub

' =========================================================
' ▼更新イベント - レコードリスト読み込み開始
' =========================================================
Private Sub update_readRecordListBegin(ByVal cnt As Long)

    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        ' 分割して処理を実行していることが分かるように
        ' 処理レコードの範囲をメッセージ（その他）に埋め込む
        frmProgress.PriMessage = "・・・"
        frmProgress.InitSec applicationSetting.recProcessCountCustom, "レコードリスト読み込み中"
        
    Else
    
        frmProgress.PriMessage = "レコードリスト読み込み中"
        
    End If
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント - レコードリスト読み込み
' =========================================================
Private Sub update_readRecordList(recordList As ValRecordList)

    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        ' 分割して処理を実行していることが分かるように
        ' 処理レコードの範囲をメッセージ（その他）に埋め込む
        
        frmProgress.SecMessage = "レコードリスト読み込み完了"
        frmProgress.SecCount = VBUtil.arraySize(recordList.values)
        
    Else
    
        frmProgress.PriMessage = "レコードリスト読み込み完了"
        frmProgress.PriCount = VBUtil.arraySize(recordList.values)
    End If
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント - SQL実行開始
' =========================================================
Private Sub update_ExecuteSqlBegin(ByVal count As Long)

    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        ' 分割して処理を実行していることが分かるように
        ' 処理レコードの範囲をメッセージ（その他）に埋め込む
        
        frmProgress.InitSec count, "SQL実行中"
    Else
    
        frmProgress.InitPri count, "SQL実行中"
        
    End If
    
    DoEvents
    
End Sub

' =========================================================
' ▼更新イベント - SQL実行
' =========================================================
Private Sub update_ExecuteSql(ByRef tableSheet As ValTableWorksheet _
                            , ByRef recordList As ValRecordList _
                            , ByVal recordListIndex As Long _
                            , ByRef sql As String _
                            , ByVal count As Long)

    On Error Resume Next

    QueryResultInfo.processCount = QueryResultInfo.processCount + 1
    If count > 0 Then
        QueryResultInfo.successCount = QueryResultInfo.successCount + 1
    End If

    ' レンジオブジェクト
    Dim r  As Range
    
    ' 結果件数が1件以上の場合のみインクリメントする
    If count > 0 Then
        
        ' 処理単位が設定されている場合
        If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        
            frmProgress.InclimentSec
        Else
        
            frmProgress.InclimentPri
        End If
        
        ' クエリー種類がDELETEの場合
        If update.queryType = update.QUERY_TYPE_DELETE Then
        
            ' レコードリストからセルのオフセット位置を取得（レンジオブジェクト）
            Set r = recordList.valuesCellOffset
            
            If tableSheet.recFormat = sutredlib.recFormatToUnder Then
                ' 削除レコードリストに行番号を格納
                deleteRecordList.setItem (r.row + recordListIndex - 1)
            
            Else
                ' 削除レコードリストに列番号を格納
                deleteRecordList.setItem (r.column + recordListIndex - 1)
            End If
            
        End If
    End If
    
    DoEvents

    ' レコードリストが設定されている場合
    If Not recordList Is Nothing Then
        
        ' コメントを埋め込むレンジオブジェクトを算出する
        Set r = getCellOfInsertedComment(tableSheet, recordList, recordListIndex)
        
        ' 更新結果が0件の場合
        If count = 0 Then
    
            Dim queryResultDetail As New sutredlib.QueryResultDetailInfo
            
            ' コメントを追加する
            ExcelUtil.addComment tableSheet.sheetName _
                               , r.row _
                               , r.column _
                               , "更新結果が0件でした。"
                               
            ' セルの背景を塗りつぶす
            ExcelUtil.fillBgColor tableSheet.sheetName _
                                , r.row _
                                , r.column _
                                , r.row _
                                , r.column _
                                , CELL_COLOR_INDEX_ZERO_UPDATED
    
            ' クエリ結果のセルを設定
            queryResultDetail.cell = ExcelUtil.convertExcelNumberToAlpha(r.row, r.column)
            ' クエリ結果のメッセージを設定
            queryResultDetail.message = "更新結果が0件でした。"
            ' クエリ結果のSQLを設定
            queryResultDetail.query = sql
        
            ' クエリ結果情報に詳細を設定する
            QueryResultInfo.addDetail queryResultDetail
            
        ' 正常に更新できた場合
        Else
        
            ' コメントを削除する
            ExcelUtil.deleteComment tableSheet.sheetName _
                               , r.row _
                               , r.column
                               
            ' セルの背景を元に戻す
            ExcelUtil.fillBgColor tableSheet.sheetName _
                                , r.row _
                                , r.column _
                                , r.row _
                                , r.column _
                                , CELL_COLOR_INDEX_SUCCESS

        End If
    End If
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    On Error GoTo 0

End Sub

' =========================================================
' ▼更新イベント - SQL実行 成功
' =========================================================
Private Sub update_ExecuteSqlSuccessful(ByVal count As Long)

    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        frmProgress.SecMessage = "SQL実行完了"
        frmProgress.PriCount = frmProgress.PriCount + count
        
    Else
    
        frmProgress.PriMessage = "SQL実行完了"
    End If
    
    DoEvents

End Sub

' =========================================================
' ▼更新イベント - SQL実行 エラー
' =========================================================
Private Sub update_ExecuteSqlError(ByRef tableSheet As ValTableWorksheet _
                                 , ByRef recordList As ValRecordList _
                                 , ByVal recordListIndex As Long _
                                 , ByRef sql As String _
                                 , ByRef sqlState As String _
                                 , ByRef errorDesc As String)

    On Error Resume Next

    QueryResultInfo.processCount = QueryResultInfo.processCount + 1
    QueryResultInfo.errorCount = QueryResultInfo.errorCount + 1
    
    Dim queryResultDetail As New sutredlib.QueryResultDetailInfo

    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        frmProgress.SecMessage = "SQL実行エラー"
    Else
    
        frmProgress.PriMessage = "SQL実行エラー"
    End If
    
    DoEvents
    
    ' レコードリストが設定されている場合
    If Not recordList Is Nothing Then
        ' レンジオブジェクト
        Dim r As Range
        ' コメントを埋め込むレンジオブジェクトを算出する
        Set r = getCellOfInsertedComment(tableSheet, recordList, recordListIndex)
        
        ' コメントを追加する
        ExcelUtil.addComment tableSheet.sheetName _
                           , r.row _
                           , r.column _
                           , errorDesc & vbNewLine & vbNewLine & sql
                           
        ' セルの背景を塗りつぶす
        ExcelUtil.fillBgColor tableSheet.sheetName _
                            , r.row _
                            , r.column _
                            , r.row _
                            , r.column _
                            , CELL_COLOR_INDEX_ERROR
    
        ' クエリ結果のセルを設定
        queryResultDetail.cell = ExcelUtil.convertExcelNumberToAlpha(r.row, r.column)
        ' クエリ結果のSQLStateを設定
        queryResultDetail.sqlState = sqlState
        ' クエリ結果のメッセージを設定
        queryResultDetail.message = errorDesc
        ' クエリ結果のSQLを設定
        queryResultDetail.query = sql
    
    Else
    
        ' クエリ結果のSQLStateを設定
        queryResultDetail.sqlState = sqlState
        ' クエリ結果のメッセージを設定
        queryResultDetail.message = errorDesc
        ' クエリ結果のSQLを設定
        queryResultDetail.query = sql
        
    End If
    
    ' クエリ結果情報に詳細を設定する
    QueryResultInfo.addDetail queryResultDetail
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    On Error GoTo 0
    
End Sub

' =========================================================
' ▼更新イベント - SQL実行 キャンセル
' =========================================================
Private Sub update_ExecuteSqlCancel(ByVal count As Long)

    Dim current As Long
    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        frmProgress.PriCount = frmProgress.PriCount - count
        frmProgress.SecCount = 0
        frmProgress.SecMessage = "SQL実行キャンセル"
    Else
    
        frmProgress.PriCount = frmProgress.PriCount - count
        frmProgress.PriMessage = "SQL実行キャンセル"
    End If

    ' クエリー種類がDELETEの場合
    ' キャンセルされた数だけ、削除対象レコードリストの要素を削除する
    If update.queryType = update.QUERY_TYPE_DELETE Then
    
        ' 削除レコードリストから要素を全て削除
        ' （新たにオブジェクトを生成することで削除を行う）
        Set deleteRecordList = New ValCollection
        
    End If

    DoEvents

End Sub

' =========================================================
' ▼更新イベント - テーブル情報読み込み
' =========================================================
Private Sub updateSelection__readTableInfo(table As ValTableWorksheet)

    ' クエリ結果情報を生成する
    Set QueryResultInfo = New sutredlib.QueryResultInfo
    QueryResultInfo.tableWorksheet = table
    QueryResultInfo.processCount = 0
    QueryResultInfo.successCount = 0
    QueryResultInfo.errorCount = 0

End Sub

' =========================================================
' ▼更新イベント（選択） - カラムリスト読み込み開始
' =========================================================
Private Sub updateSelection__readColumnListBegin(ByVal cnt As Long)

    frmProgress.InitPri cnt, "カラム定義読み込み中"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - カラムリスト読み込み
' =========================================================
Private Sub updateSelection__readColumn(column As ValDbDefineColumn)

    frmProgress.InclimentPri
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - カラムリスト読み込み完了
' =========================================================
Private Sub updateSelection__readColumnList(columnList As sutredlib.GenericCollection)

    frmProgress.PriMessage = "カラム定義読み込み完了"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - セル選択時 レコードリスト読み込み開始
' =========================================================
Private Sub updateSelection__readRecordListBySelectionBegin(ByVal cnt As Long)

    frmProgress.InitPri cnt, "レコードリスト読み込み中"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - セル選択時 レコードリスト読み込み
' =========================================================
Private Sub updateSelection__readRecordListBySelectionOne(recordList As ValRecordList)

    Dim count As Long
    
    count = VBUtil.arraySize(recordList.values)
    
    frmProgress.PriCount = frmProgress.PriCount + count
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - セル選択時 レコードリスト読み込み
' =========================================================
Private Sub updateSelection__readRecordListBySelection(recordLists As ValCollection)

    frmProgress.PriMessage = "レコードリスト読み込み完了"
    
    ' カウント数を一旦0クリアする（カウント総数はそのまま）
    ' ExecuteSqlBeginイベントはセル選択時の場合は複数回呼び出されるので
    ' カウントをリセットするような処理はせずに、その手前で呼び出される本イベントでリセットを実施する
    frmProgress.PriCount = 0
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - SQL実行開始
' =========================================================
Private Sub updateSelection__ExecuteSqlBegin(ByVal count As Long)

    frmProgress.PriMessage = "SQL実行中"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - SQL実行
' =========================================================
Private Sub updateSelection__ExecuteSql(ByRef tableSheet As ValTableWorksheet _
                            , ByRef recordList As ValRecordList _
                            , ByVal recordListIndex As Long _
                            , ByRef sql As String _
                            , ByVal count As Long)

    On Error Resume Next

    QueryResultInfo.processCount = QueryResultInfo.processCount + 1
    If count > 0 Then
        QueryResultInfo.successCount = QueryResultInfo.successCount + 1
    End If

    ' レンジオブジェクト
    Dim r  As Range
    
    ' 結果件数が1件以上の場合のみインクリメントする
    If count > 0 Then
        
        frmProgress.InclimentPri
        
        ' クエリー種類がDELETEの場合
        If updateSelection_.queryType = updateSelection_.QUERY_TYPE_DELETE Then
        
            ' レコードリストからセルのオフセット位置を取得（レンジオブジェクト）
            Set r = recordList.valuesCellOffset
            
            If tableSheet.recFormat = sutredlib.recFormatToUnder Then
                ' 削除レコードリストに行番号を格納
                deleteRecordList.setItem (r.row + recordListIndex - 1)
            
            Else
                ' 削除レコードリストに列番号を格納
                deleteRecordList.setItem (r.column + recordListIndex - 1)
            End If
            
        End If
    End If
    
    DoEvents

    ' レコードリストが設定されている場合
    If Not recordList Is Nothing Then
        
        ' コメントを埋め込むレンジオブジェクトを算出する
        Set r = getCellOfInsertedComment(tableSheet, recordList, recordListIndex)
        
        ' 更新結果が0件の場合
        If count = 0 Then
            
            Dim queryResultDetail As New sutredlib.QueryResultDetailInfo
            
            ' コメントを追加する
            ExcelUtil.addComment tableSheet.sheetName _
                               , r.row _
                               , r.column _
                               , "更新結果が0件でした。"
                               
            ' セルの背景を塗りつぶす
            ExcelUtil.fillBgColor tableSheet.sheetName _
                                , r.row _
                                , r.column _
                                , r.row _
                                , r.column _
                                , CELL_COLOR_INDEX_ZERO_UPDATED
        
    
            ' クエリ結果のセルを設定
            queryResultDetail.cell = ExcelUtil.convertExcelNumberToAlpha(r.row, r.column)
            ' クエリ結果のメッセージを設定
            queryResultDetail.message = "更新結果が0件でした。"
            ' クエリ結果のSQLを設定
            queryResultDetail.query = sql
            
            ' クエリ結果情報に詳細を追加する
            QueryResultInfo.addDetail queryResultDetail
            
        ' 正常に更新できた場合
        Else
        
            ' コメントを削除する
            ExcelUtil.deleteComment tableSheet.sheetName _
                               , r.row _
                               , r.column
                               
            ' セルの背景を塗りつぶす
            ExcelUtil.fillBgColor tableSheet.sheetName _
                                , r.row _
                                , r.column _
                                , r.row _
                                , r.column _
                                , CELL_COLOR_INDEX_SUCCESS

        End If
    End If
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    On Error GoTo 0

End Sub

' =========================================================
' ▼更新イベント（選択） - SQL実行 成功
' =========================================================
Private Sub updateSelection__ExecuteSqlSuccessful(ByVal count As Long)

    frmProgress.PriMessage = "SQL実行完了"

    DoEvents

End Sub

' =========================================================
' ▼更新イベント（選択） - SQL実行 エラー
' =========================================================
Private Sub updateSelection__ExecuteSqlError(ByRef tableSheet As ValTableWorksheet _
                                 , ByRef recordList As ValRecordList _
                                 , ByVal recordListIndex As Long _
                                 , ByRef sql As String _
                                 , ByRef sqlState As String _
                                 , ByRef errorDesc As String)

    On Error Resume Next
    
    Dim queryResultDetail As New sutredlib.QueryResultDetailInfo
    
    QueryResultInfo.processCount = QueryResultInfo.processCount + 1
    QueryResultInfo.errorCount = QueryResultInfo.errorCount + 1

    frmProgress.PriMessage = "SQL実行エラー"
    
    DoEvents
    
    ' レコードリストが設定されている場合
    If Not recordList Is Nothing Then
        ' レンジオブジェクト
        Dim r As Range
        ' コメントを埋め込むレンジオブジェクトを算出する
        Set r = getCellOfInsertedComment(tableSheet, recordList, recordListIndex)
        
        ' コメントを追加する
        ExcelUtil.addComment tableSheet.sheetName _
                           , r.row _
                           , r.column _
                           , errorDesc & vbNewLine & vbNewLine & sql
                           
        ' セルの背景を塗りつぶす
        ExcelUtil.fillBgColor tableSheet.sheetName _
                            , r.row _
                            , r.column _
                            , r.row _
                            , r.column _
                            , CELL_COLOR_INDEX_ERROR
    
    
        ' クエリ結果のセルを設定
        queryResultDetail.cell = ExcelUtil.convertExcelNumberToAlpha(r.row, r.column)
        ' クエリ結果のSQLステータスを設定
        queryResultDetail.sqlState = sqlState
        ' クエリ結果のメッセージを設定
        queryResultDetail.message = errorDesc
        ' クエリ結果のSQLを設定
        queryResultDetail.query = sql
        
    Else
    
        ' クエリ結果のSQLステータスを設定
        queryResultDetail.sqlState = sqlState
        ' クエリ結果のメッセージを設定
        queryResultDetail.message = errorDesc
        ' クエリ結果のSQLを設定
        queryResultDetail.query = sql
    
    End If
    
    ' クエリ結果情報に詳細を追加する
    QueryResultInfo.addDetail queryResultDetail
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    On Error GoTo 0
    
End Sub

' =========================================================
' ▼更新イベント（選択） - SQL実行 キャンセル
' =========================================================
Private Sub updateSelection__ExecuteSqlCancel(ByVal count As Long)

    Dim current As Long
    
    current = frmProgress.PriCount
    current = current - count
    
    frmProgress.PriCount = current
    frmProgress.PriMessage = "SQL実行キャンセル"

    ' クエリー種類がDELETEの場合
    ' キャンセルされた数だけ、削除対象レコードリストの要素を削除する
    If updateSelection_.queryType = updateSelection_.QUERY_TYPE_DELETE Then
    
        Dim i As Long
        For i = 1 To count
        
            ' 最後尾の要素を削除
            deleteRecordList.removeByIndex deleteRecordList.count
        Next
        
    End If

    DoEvents

End Sub

' =========================================================
' ▼クエリー実行イベント - SQL実行中
' =========================================================
Private Sub queryExecution_ExecuteSql(sql As String)

    frmProgress.InitPri -1, "SQL実行中"
    frmProgress.PriCount = 0
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    DoEvents

    SutWhite.showHourglassWindowOnCenterPt
End Sub

' =========================================================
' ▼クエリー実行イベント - SQLキャンセル
' =========================================================
Private Sub queryExecution_ExecuteSqlCancel(ByVal count As Long)

    frmProgress.PriMessage = "SQL実行キャンセル"

    DoEvents

    SutWhite.closeHourglassWindow
    
End Sub

' =========================================================
' ▼クエリー実行イベント - SQL実行完了
' =========================================================
Private Sub queryExecution_ExecuteSqlSuccessful(sql As String)

    frmProgress.PriMessage = "SQL実行完了"
    
    DoEvents

    SutWhite.closeHourglassWindow
    
End Sub

' =========================================================
' ▼クエリー実行イベント - SQL実行エラー
' =========================================================
Private Sub queryExecution_ExecuteSqlError(sql As String, sqlState As String, message As String)

    frmProgress.PriMessage = "SQL実行エラー"
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    ' クエリ結果情報に詳細を追加する
    
    If Not QueryResultInfo Is Nothing Then
    
        Dim detail As New sutredlib.QueryResultDetailInfo
        detail.cell = "-"
        detail.sqlState = sqlState
        detail.message = message
        detail.query = sql
        
        QueryResultInfo.addDetail detail
    End If

    DoEvents

    SutWhite.closeHourglassWindow
    
End Sub

' =========================================================
' ▼クエリー実行イベント - レコード取得中
' =========================================================
Private Sub queryExecution_ReflectionRecordGetBegin(ByVal recordListSize As Long)

    frmProgress.PriMessage = "レコード取得中"

    DoEvents

    SutWhite.showHourglassWindowOnCenterPt

End Sub

' =========================================================
' ▼クエリー実行イベント - レコード取得
' =========================================================
Private Sub queryExecution_ReflectionRecordGet(ByVal recordListSizeActual As Long, ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriCount = frmProgress.PriCount + recordListSize
    frmProgress.PriMessage = "レコード取得完了"
    
    DoEvents

    SutWhite.closeHourglassWindow

End Sub

' =========================================================
' ▼クエリー実行イベント - レコード変換中
' =========================================================
Private Sub queryExecution_ReflectionRecordConvertBegin(ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriMessage = "レコード情報変換中"
    
    DoEvents

End Sub

' =========================================================
' ▼クエリー実行イベント - レコード変換完了
' =========================================================
Private Sub queryExecution_ReflectionRecordConvert(ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriMessage = "レコード情報変換完了"
    
    DoEvents

End Sub

' =========================================================
' ▼クエリー実行イベント - レコード反映中
' =========================================================
Private Sub queryExecution_ReflectionRecordBegin(ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriMessage = "レコード シート出力中"
    
    DoEvents

End Sub

' =========================================================
' ▼クエリー実行イベント - レコード反映完了
' =========================================================
Private Sub queryExecution_ReflectionRecord(ByVal recordListSize As Long, recordList As Variant)

    frmProgress.PriMessage = "レコード シート出力完了"
    
    DoEvents

End Sub

' =========================================================
' ▽コメント挿入位置のセル取得
'
' 概要　　　：コメントを挿入する位置のセルを取得する。
' 引数　　　：tableSheet      テーブルシート
' 　　　　　　recordList      レコードリスト
' 　　　　　　recordListIndex レコードリストインデックス
' 戻り値　　：コメントを挿入する位置のセル（Rangeオブジェクト）
'
' =========================================================
Private Function getCellOfInsertedComment(ByRef tableSheet As ValTableWorksheet _
                                        , ByRef recordList As ValRecordList _
                                        , ByVal recordListIndex As Long) As Range

    ' レンジオブジェクト
    Dim r As Range
    ' コメントを埋め込むセルの行位置
    Dim row As Long
    ' コメントを埋め込むセルの列位置
    Dim col As Long
    
    ' ベースとなるレンジオブジェクトを取得する
    Set r = recordList.valuesCellOffset
        
    If tableSheet.recFormat = sutredlib.recFormatToUnder Then
    
        ' 行をレコード配列の現在位置から算出する
        row = r.row + recordListIndex - LBound(recordList.values)
        ' 列はレコード番号を設定
        col = ConstantsTable.U_RECORD_NUM_COL
        
        ' 戻り値を設定する
        Set getCellOfInsertedComment = tableSheet.sheet.Cells(row, col)
    Else
    
        ' 行はレコード番号を設定
        row = ConstantsTable.R_RECORD_NUM_ROW
        ' 列をレコード配列の現在位置から算出する
        col = r.column + recordListIndex - LBound(recordList.values)
    
        ' 戻り値を設定する
        Set getCellOfInsertedComment = tableSheet.sheet.Cells(row, col)
    End If
    
End Function

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    If Not frmProgress Is Nothing Then
        frmProgress.Close
    End If
    If Not frmQueryEditor Is Nothing Then
        frmQueryEditor.Close
    End If
    If Not frmQueryResult Is Nothing Then
        frmQueryResult.Close
    End If

    ' プログレスフォームを破棄する
    Set frmProgress = Nothing
    ' SELECT条件生成フォームを破棄する
    Set frmSelectConditionCreatorVar = Nothing
    ' レコード追加・削除フォームを破棄する
    Set frmRecAppender = Nothing
    ' クエリエディタを破棄する
    Set frmQueryEditor = Nothing
    ' 一括クエリフォームを破棄する
    Set frmQueryBatch = Nothing
    ' クエリ結果フォームを破棄する
    Set frmQueryResult = Nothing
    
    ' レコード情報生成オブジェクトを破棄する
    Set tableSheetReader = Nothing
    
    ' 更新クエリー実行オブジェクトを破棄する
    Set update = Nothing
    Set updateSelection_ = Nothing
    
    ' 参照クエリー実行オブジェクトを破棄する
    Set select_ = Nothing
    
    ' クエリー実行オブジェクトを破棄する
    Set queryExecution = Nothing
    
    ' アプリケーション設定情報を破棄する
    Set applicationSetting = Nothing
    
    ' 削除レコードリストを破棄する
    Set deleteRecordList = Nothing

End Sub

' =========================================================
' ▽初期化
'
' 概要　　　：初期化処理を実行する。
' 引数　　　：appSetting アプリケーション設定情報
' 　　　　　　appSettingColFmt アプリケーション設定情報（カラム書式）
' 　　　　　　conn       DBコネクション
'             clearQueryResults クエリ結果を消去する
'
' =========================================================
Public Sub init(ByRef appSetting As ValApplicationSetting _
              , ByRef appSettingColFmt As ValApplicationSettingColFormat _
              , ByRef conn As Object _
              , Optional ByVal clearQueryResults As Boolean = True)

    ' フォームを設定する
    Set frmProgress = New sutredlib.Progress
    
    ' フォームをロードする
    Load frmSelectConditionCreator
    ' フォームを設定する
    Set frmSelectConditionCreatorVar = frmSelectConditionCreator

    ' フォームをロードする
    Load frmRecordAppender
    ' フォームを設定する
    Set frmRecAppender = frmRecordAppender
    
    Set QueryResultInfo = Nothing
    
    ' クエリエディタを初期化する
    If frmQueryEditor Is Nothing Then
    
        Set frmQueryEditor = New sutredlib.queryEditor
        frmQueryEditor.ExcelApplication = Excel.Application
    End If
    
    ' クエリ結果を初期化する
    If frmQueryResult Is Nothing Then
    
        ' フォームを設定する
        Set frmQueryResult = New sutredlib.queryResult
        frmQueryResult.ExcelApplication = Excel.Application
    End If
    
    If clearQueryResults Then
    
        ' クエリ結果をクリアする
        frmQueryResult.ClearInfos
    End If

    ' レコード情報生成オブジェクトを生成する
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = ActiveSheet

    ' 更新クエリー実行オブジェクトを生成する
    Set update = New ExeUpdate
    Set update.conn = conn
    Set update.sheet = ActiveSheet
    Set update.applicationSetting = appSetting
    Set update.applicationSettingColFmt = appSettingColFmt
    
    ' 更新クエリー実行オブジェクトを生成する
    Set updateSelection_ = New ExeUpdate
    Set updateSelection_.conn = conn
    Set updateSelection_.sheet = ActiveSheet
    Set updateSelection_.applicationSetting = appSetting
    Set updateSelection_.applicationSettingColFmt = appSettingColFmt
    
    ' 参照クエリー実行オブジェクトを生成する
    Set select_ = New ExeSelect
    Set select_.conn = conn
    Set select_.sheet = ActiveSheet
    Set select_.applicationSetting = appSetting
    Set select_.applicationSettingColFmt = appSettingColFmt
    
    ' 参照クエリー実行オブジェクトを生成する
    Set queryExecution = New ExeQueryExecution
    Set queryExecution.conn = conn
    queryExecution.sheet = ActiveSheet
    queryExecution.applicationSetting = appSetting
    
    ' アプリケーション設定情報を取得する
    Set applicationSetting = appSetting

    ' 削除レコードリストを初期化する
    Set deleteRecordList = New ValCollection
    
End Sub

' =========================================================
' ▽INSERT and UPDATE実行（全レコード）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub insertUpdateAll()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet

    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgress.Show ExcelUtil.getApplicationHWnd, True
    Else
        frmProgress.Show ExcelUtil.getApplicationHWnd, False
    End If
    frmProgress.title = "INSERT and UPDATE - 全て" & createProgressTitleOptionalMessage

    ' クエリを実行する
    update.queryType = update.QUERY_TYPE_INSERT_UPDATE
    resultCount = update.execute
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
        QueryResultInfo.queryType = insertUpdate
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = insertUpdate
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If

End Sub

' =========================================================
' ▽INSERT and UPDATE実行（選択範囲）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub insertUpdateSelection()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    frmProgress.Show ExcelUtil.getApplicationHWnd, False
    frmProgress.title = "INSERT and UPDATE - 範囲選択"

    ' クエリを実行する
    updateSelection_.queryType = updateSelection_.QUERY_TYPE_INSERT_UPDATE
    resultCount = updateSelection_.executeSelection
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
        QueryResultInfo.queryType = insertUpdate
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = insertUpdate
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If

End Sub
' =========================================================
' ▽INSERT実行（全レコード）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub insertAll()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet

    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgress.Show ExcelUtil.getApplicationHWnd, True
    Else
        frmProgress.Show ExcelUtil.getApplicationHWnd, False
    End If
    frmProgress.title = "INSERT - 全て" & createProgressTitleOptionalMessage

    ' クエリを実行する
    update.queryType = update.QUERY_TYPE_INSERT
    resultCount = update.execute
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
        QueryResultInfo.queryType = insert
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = insert
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If

End Sub

' =========================================================
' ▽INSERT実行（選択範囲）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub insertSelection()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    frmProgress.Show ExcelUtil.getApplicationHWnd, False
    frmProgress.title = "INSERT - 範囲選択"

    ' クエリを実行する
    updateSelection_.queryType = updateSelection_.QUERY_TYPE_INSERT
    resultCount = updateSelection_.executeSelection
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
        QueryResultInfo.queryType = insert
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = insert
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If

End Sub

' =========================================================
' ▽UPDATE実行（全レコード）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub updateAll()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgress.Show ExcelUtil.getApplicationHWnd, True
    Else
        frmProgress.Show ExcelUtil.getApplicationHWnd, False
    End If
    frmProgress.title = "UPDATE - 全て" & createProgressTitleOptionalMessage
    
    ' クエリを実行する
    update.queryType = update.QUERY_TYPE_UPDATE
    resultCount = update.execute
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
        QueryResultInfo.queryType = sutredlib.update
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = sutredlib.update
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If
    
End Sub

' =========================================================
' ▽UPDATE実行（選択範囲）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub updateSelection()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    frmProgress.Show ExcelUtil.getApplicationHWnd, False
    frmProgress.title = "UPDATE - 範囲選択"
    
    ' クエリを実行する
    updateSelection_.queryType = updateSelection_.QUERY_TYPE_UPDATE
    resultCount = updateSelection_.executeSelection
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
        QueryResultInfo.queryType = sutredlib.update
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = sutredlib.update
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If

End Sub

' =========================================================
' ▽DELETE実行（全レコード）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub deleteAll()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
        
    ' 削除確認
    If confirmDelete = False Then
    
        Exit Sub
    End If
    
    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgress.Show ExcelUtil.getApplicationHWnd, True
    Else
        frmProgress.Show ExcelUtil.getApplicationHWnd, False
    End If
    frmProgress.title = "DELETE - 全て" & createProgressTitleOptionalMessage
    
    ' クエリを実行する
    update.queryType = update.QUERY_TYPE_DELETE
    resultCount = update.execute
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
        QueryResultInfo.queryType = deleteOnSheet
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    If resultCount > 0 Then
        ' 削除に成功したレコードをセル上から消す
        deleteCellsByDeleteRecordList
    
    End If
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = deleteOnSheet
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If
    
End Sub

' =========================================================
' ▽DELETE実行（選択範囲）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub deleteSelection()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' 削除確認
    If confirmDelete = False Then
    
        Exit Sub
    End If

    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    frmProgress.Show ExcelUtil.getApplicationHWnd, False
    frmProgress.title = "DELETE - 範囲選択"
    
    ' クエリを実行する
    updateSelection_.queryType = updateSelection_.QUERY_TYPE_DELETE
    resultCount = updateSelection_.executeSelection
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = deleteOnSheet
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    If resultCount > 0 Then
        ' 削除に成功したレコードをセル上から消す
        deleteCellsByDeleteRecordList
    End If

    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgress.Close
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            QueryResultInfo.queryType = deleteOnSheet
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If

End Sub

' =========================================================
' ▽DELETE実行（テーブル上の全レコード）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub deleteAllOfTable()

    On Error GoTo err

    Set QueryResultInfo = New sutredlib.QueryResultInfo
    QueryResultInfo.queryType = sutredlib.deleteAll
    QueryResultInfo.processCount = 0
    QueryResultInfo.successCount = 0
    QueryResultInfo.errorCount = 0

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' 削除確認
    If confirmDelete = False Then
    
        Exit Sub
    End If
    
    ' テーブル上の全レコード削除ではプログレスウィンドウは表示しない
    
    ' 実行件数
    Dim resultCount As Long
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' クエリを実行する
    update.queryType = update.QUERY_TYPE_DELETE_ALL
    resultCount = update.executeQueryDeleteOfAllTable
    
    ' 確認ダイアログを表示する
    If VBUtil.showMessageBoxForYesNo("シート上の全レコードを削除します。" & vbNewLine & "よろしいですか？" _
            , ConstantsCommon.APPLICATION_NAME) = WinAPI_User.IDYES Then
    
        ' セルの削除・挿入用に利用するオブジェクト
        Dim tableSheetCreator As New ExeTableSheetCreator
    
        ' アプリケーション設定情報を設定
        tableSheetCreator.applicationSetting = applicationSetting
        ' 入力領域を全て削除する
        tableSheetCreator.deleteCellOfRecord update.readTableSheet
        ' 入力領域を新たに生成する
        tableSheetCreator.insertEmptyCell update.readTableSheet
    End If
    
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
        QueryResultInfo.tableWorksheet = update.readTableSheet
        QueryResultInfo.processCount = resultCount
        QueryResultInfo.successCount = resultCount
        QueryResultInfo.errorCount = 0
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' クエリエラーが発生した場合
    If err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        ' クエリ結果情報を設定する
        If Not QueryResultInfo Is Nothing Then
            Dim detail As New sutredlib.QueryResultDetailInfo
            detail.cell = "-"
            detail.message = err.Description
            detail.query = update.createdSql
            
            If Not dbConn Is Nothing Then
            
                If dbConn.Errors.count > 0 Then
                
                    ' エラーオブジェクトを取得する
                    Dim error As Object
                    Set error = dbConn.Errors.item(0)
                    detail.sqlState = error.sqlState
                End If
            End If
            
            QueryResultInfo.addDetail detail
            QueryResultInfo.tableWorksheet = update.readTableSheet
            QueryResultInfo.processCount = 1
            QueryResultInfo.successCount = 0
            QueryResultInfo.errorCount = 1
        
            frmQueryResult.AddInfo QueryResultInfo
        End If
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If

End Sub

' =========================================================
' ▽セル削除
'
' 概要　　　：削除レコードリストを基にセルを削除する。
' 引数　　　：
'
' =========================================================
Private Sub deleteCellsByDeleteRecordList()

    ' 確認ダイアログを表示する
    If VBUtil.showMessageBoxForYesNo("削除に成功したレコードをシート上から削除します。" & vbNewLine & "よろしいですか？" _
            , ConstantsCommon.APPLICATION_NAME) = WinAPI_User.IDNO Then
    
        Exit Sub
    End If
    
    ' レコード配列
    Dim recordArray As Variant
    ' インデックス
    Dim i As Long
    
    With update.sheet
        
        ' 削除レコードリストを配列に変換
        recordArray = deleteRecordList.toArray(vbVariant)
        
        ' 配列かをチェック
        If IsArray(recordArray) = True Then
            
            ' クイックソートを実行する
            VBUtil.quickSort recordArray
            
            ' 配列を1件ずつ処理する
            ' セルを削除すると、行の場合は下から上へ、列の場合は右から左へ全体がシフトするため
            ' 大きい番号から削除を実行する必要がある
            i = UBound(recordArray)
            Do While i >= LBound(recordArray)
            
                If tableSheetReader.getRowFormat = sutredlib.recFormatToUnder Then
                
                    ' 削除する（行単位で削除）
                    .Cells(recordArray(i), 1).EntireRow.delete
                Else
                
                    ' 削除する（列単位で削除）
                    .Cells(1, recordArray(i)).EntireColumn.delete
                End If
            
                i = i - 1
            Loop
        End If
        
        ' 全ての入力領域が削除されているかを表すフラグ
        Dim isAllDeleted As Boolean
        
        ' 番号セルに値が設定されているかを調べることでレコードセルが削除されているかどうかを調べる
        ' 行フォーマットごとに調べる箇所を変えてやる
        If tableSheetReader.getRowFormat = sutredlib.recFormatToUnder Then
        
            If .Cells(ConstantsTable.U_RECORD_OFFSET_ROW, U_RECORD_NUM_COL) = "" Then
            
                isAllDeleted = True
            End If
            
        Else
        
            If .Cells(ConstantsTable.R_RECORD_NUM_ROW, R_RECORD_OFFSET_COL) = "" Then
            
                isAllDeleted = True
            End If
            
        End If
        
        ' 全ての入力領域が削除されている場合
        If isAllDeleted = True Then
            
            ' シートを読み込む
            Dim tableSheet As ValTableWorksheet
            Set tableSheet = tableSheetReader.readTableInfo
            
            ' セルの削除・挿入用に利用するオブジェクト
            Dim tableSheetCreator As New ExeTableSheetCreator
            tableSheetCreator.applicationSetting = applicationSetting
            ' 入力領域を新たに生成する
            tableSheetCreator.insertEmptyCell tableSheet

        End If
        
        
    End With
    
End Sub

' =========================================================
' ▽SELECT実行（条件なし）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub selectAll()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' クエリ結果を初期化しておく
    Set QueryResultInfo = New sutredlib.QueryResultInfo
    QueryResultInfo.queryType = sutredlib.selectAll
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' レコード取得件数
    Dim resultCount

    frmProgress.Show ExcelUtil.getApplicationHWnd, False
    frmProgress.title = "SELECT - 全て" & createProgressTitleOptionalMessage

    select_.appendSheet = False
    resultCount = select_.executeSelect
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing

    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
    
        QueryResultInfo.processCount = resultCount
        QueryResultInfo.successCount = resultCount
        QueryResultInfo.errorCount = 0
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    ' --------------------------------------------------------------
    ' レジストリに実行したSQLを前回実行したSQLの情報として設定する
    ' --------------------------------------------------------------
    ' テーブル定義
    Dim tableSheet As ValTableWorksheet
    ' テーブル定義を取得する
    Set tableSheet = tableSheetReader.readTableInfo(True)

    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SELECT_LATEST & "\" & tableSheet.table.schemaTableName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    registry.setValue REG_SUB_KEY_SELECT_LATEST_TYPE, REG_SUB_KEY_SELECT_LATEST_TYPE_ALL
    registry.setValue REG_SUB_KEY_SELECT_LATEST_APPEND, ""
    registry.setValue REG_SUB_KEY_SELECT_LATEST_SQL, ""

    Set registry = Nothing

    SutWhite.closeHourglassWindow
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    SutWhite.closeHourglassWindow
    
    frmProgress.Close

    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing And err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        QueryResultInfo.processCount = 0
        QueryResultInfo.successCount = 0
        QueryResultInfo.errorCount = 1
        frmQueryResult.AddInfo QueryResultInfo
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    
    End If

End Sub

' =========================================================
' ▽SELECT実行（条件あり）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub selectCondition()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' SELECT SQL生成フォームを表示する
    frmSelectConditionCreatorVar.ShowExt vbModal _
                                    , Main.getApplicationSetting _
                                    , Main.getApplicationSettingColFormat _
                                    , Main.getDBConnection
    
    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽行の追加
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub recordAdd()

    On Error GoTo err

    frmRecAppender.ShowExt vbModal, ActiveSheet, applicationSetting
    
    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽クエリエディタ
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub queryEditor()

    On Error GoTo err

    ' クエリエディタを表示する
    frmQueryEditor.Show ExcelUtil.getApplicationHWnd
    
    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽一括クエリ
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub queryBatch()

    On Error GoTo err

    ' テーブルリスト
    Dim tableList As New sutredlib.GenericCollection
    
    ' テーブルシート読込オブジェクト
    Dim tableSheetReader As ExeTableSheetReader
        
    ' ブック
    Dim book  As Workbook
    ' シート
    Dim sheet As Worksheet
    
    Dim v1 As ValTableWorksheet
    Dim v2 As sutredlib.DbQueryBatchTableWorksheet
    
    ' アクティブブックをbook変数に格納する
    Set book = ActiveWorkbook
    
    ' ブックに含まれているシートを1件ずつ処理する
    For Each sheet In book.Worksheets
    
        Set tableSheetReader = New ExeTableSheetReader
        Set tableSheetReader.sheet = sheet
        
        ' 対象シートがテーブルシートの場合
        If tableSheetReader.isTableSheet = True Then
        
            Set v1 = tableSheetReader.readTableInfo(True)
            
            Set v2 = New sutredlib.DbQueryBatchTableWorksheet
            v2.tableWorksheet = v1
            v2.BatchQueryType = none
            
            ' テーブルシートを読み込んでリストに設定する（テーブル情報のみ取得する）
            tableList.Add v2
        
        End If
    
    Next
    
    ' フォームを設定する
    Set frmQueryBatch = New sutredlib.DbQueryBatch
    
    frmQueryBatch.tableWorksheets = tableList
    frmQueryBatch.Show
    
    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽クエリ結果にエラーがあるかどうかを調査する。
'
' 概要　　　：
' 戻り値　　：true エラー発生
'
' =========================================================
Public Function isErrorQueryResult() As Boolean

    ' クエリ結果を取得する
    Dim i As Long
    Dim queryResultCount As Long
    Dim queryResultObj As QueryResultInfo
    
    queryResultCount = frmQueryResult.GetInfoCount
    
    For i = 1 To queryResultCount
    
        Set queryResultObj = frmQueryResult.GetInfo(i)
        
        If Not queryResultObj Is Nothing Then
        
            If queryResultObj.errorCount > 0 Then
            
                isErrorQueryResult = True
                Exit Function
            End If
            
        End If
    Next

    isErrorQueryResult = False
    
End Function

' =========================================================
' ▽クエリ結果表示
'
' 概要　　　：
'
' =========================================================
Public Sub showQueryResult()

    On Error GoTo err
    
    ' クエリ結果フォームを表示する
    frmQueryResult.Show ExcelUtil.getApplicationHWnd

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽クエリ結果表示
'
' 概要　　　：
'
' =========================================================
Public Sub showQueryResultWhenSettingResult()

    On Error GoTo err
    
    
    If applicationSetting.queryResultShowWhenNormal = False Then
    
        If isErrorQueryResult Then
        
            ' クエリ結果が設定済みの場合、表示する
            If frmQueryResult.GetInfoCount > 0 Then
                ' クエリ結果フォームを表示する
                frmQueryResult.Show ExcelUtil.getApplicationHWnd
            End If
        End If
        
    Else
    
        ' クエリ結果が設定済みの場合、表示する
        If frmQueryResult.GetInfoCount > 0 Then
            ' クエリ結果フォームを表示する
            frmQueryResult.Show ExcelUtil.getApplicationHWnd
        End If
    End If

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽SELECT条件生成フォーム呼び出し後のイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Private Sub frmSelectConditionCreatorVar_ok(ByVal sql As String, ByVal append As Boolean)

    On Error GoTo err
    
    ' クエリ結果を初期化しておく
    Set QueryResultInfo = New sutredlib.QueryResultInfo
    QueryResultInfo.queryType = sutredlib.selectCondition

    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' レコード取得件数
    Dim resultCount
    
    frmProgress.Show ExcelUtil.getApplicationHWnd, False
    frmProgress.title = "SELECT - 条件指定" & createProgressTitleOptionalMessage
    
    ' テーブル定義
    Dim tableSheet As ValTableWorksheet
    
    ' テーブル定義を取得する
    Set tableSheet = tableSheetReader.readTableInfo
    
    select_.appendSheet = append
    resultCount = select_.executeSelectByQuery(sql)
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
    
        QueryResultInfo.processCount = resultCount
        QueryResultInfo.successCount = resultCount
        QueryResultInfo.errorCount = 0
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    ' --------------------------------------------------------------
    ' レジストリに実行したSQLを前回実行したSQLの情報として設定する
    ' --------------------------------------------------------------
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SELECT_LATEST & "\" & tableSheet.table.schemaTableName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    registry.setValue REG_SUB_KEY_SELECT_LATEST_TYPE, REG_SUB_KEY_SELECT_LATEST_TYPE_CONDITION
    registry.setValue REG_SUB_KEY_SELECT_LATEST_APPEND, CStr(append)
    registry.setValue REG_SUB_KEY_SELECT_LATEST_SQL, sql

    Set registry = Nothing
    
    SutWhite.closeHourglassWindow
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    SutWhite.closeHourglassWindow
    
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing And err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        QueryResultInfo.queryType = sutredlib.selectCondition
        QueryResultInfo.processCount = 0
        QueryResultInfo.successCount = 0
        QueryResultInfo.errorCount = 1
        frmQueryResult.AddInfo QueryResultInfo
        
    Else

        ' エラー発生
        Main.ShowErrorMessage
    
    End If

End Sub

' =========================================================
' ▽SELECT実行（再実行）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub selectReExecute()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' クエリ結果を初期化しておく
    Set QueryResultInfo = New sutredlib.QueryResultInfo
    QueryResultInfo.queryType = sutredlib.selectReExec
    
    ' テーブル定義
    Dim tableSheet As ValTableWorksheet
    ' テーブル定義を取得する
    Set tableSheet = tableSheetReader.readTableInfo
    
    ' ------------------------------------------------
    ' 前回実行されたSELECT文の情報をレジストリから取得する
    ' ------------------------------------------------
    Dim latestType   As String      ' SQLの種類
    Dim latestSql    As String      ' SQL
    Dim latestAppend As String      ' 追加有無フラグ
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SELECT_LATEST & "\" & tableSheet.table.schemaTableName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Dim value    As String
    Dim valueKey As String
    
    ' SQL種類を取得する
    ' 前回指定されたSELECT文は無いと見なして、条件無しのSELECT文を発行する
    valueKey = REG_SUB_KEY_SELECT_LATEST_TYPE
    If registry.getValue(valueKey, value) = False Then
    
        ' レジストリに情報自体存在しない場合は、”全て”をデフォルト値として実行する
        latestType = REG_SUB_KEY_SELECT_LATEST_TYPE_ALL
        
    Else
        latestType = value
    End If
    
    ' SQLを取得する
    valueKey = REG_SUB_KEY_SELECT_LATEST_SQL
    registry.getValue valueKey, value
    latestSql = value
    
    ' 追加有無を取得する
    valueKey = REG_SUB_KEY_SELECT_LATEST_APPEND
    registry.getValue valueKey, value
    latestAppend = value
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' レコード取得件数
    Dim resultCount
    
    frmProgress.Show ExcelUtil.getApplicationHWnd, False
    frmProgress.title = "SELECT - 再実行" & createProgressTitleOptionalMessage
    
    If latestType = REG_SUB_KEY_SELECT_LATEST_TYPE_CONDITION Then
        
        If latestAppend = REG_SUB_KEY_SELECT_LATEST_APPEND_TRUE Then
        
            select_.appendSheet = True
        Else
        
            select_.appendSheet = False
        End If
        
        resultCount = select_.executeSelectByQuery(latestSql)
    
    Else
    
        resultCount = select_.executeSelect
    
    End If
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing Then
    
        QueryResultInfo.processCount = resultCount
        QueryResultInfo.successCount = resultCount
        QueryResultInfo.errorCount = 0
        frmQueryResult.AddInfo QueryResultInfo
    End If
    
    SutWhite.closeHourglassWindow
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    SutWhite.closeHourglassWindow
    
    frmProgress.Close
    
    ' クエリ結果情報を設定する
    If Not QueryResultInfo Is Nothing And err.Number = ERR_NUMBER_SQL_EXECUTE_FAILED Then
    
        QueryResultInfo.processCount = 0
        QueryResultInfo.successCount = 0
        QueryResultInfo.errorCount = 1
        frmQueryResult.AddInfo QueryResultInfo
    Else
    
        ' エラー発生
        Main.ShowErrorMessage
    End If
    
End Sub

' =========================================================
' ▽削除実行確認
'
' 概要　　　：DELETEを実行してもいいかどうかを確認する。
' 引数　　　：
' 戻り値　　：True DELETEを実行する
'
' =========================================================
Private Function confirmDelete() As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False
    
    ' メッセージボックスの戻り値
    Dim msgResult As Long
    
    ' メッセージを表示する
    msgResult = VBUtil.showMessageBoxForYesNo("本当に削除してもよろしいですか？" _
                     , ConstantsCommon.APPLICATION_NAME)
    
    ' YESが押下された場合
    If msgResult = WinAPI_User.IDYES Then
    
        ret = True
    End If
    
    confirmDelete = ret

End Function

' =========================================================
' ▽プログレスバーのタイトルに表示するオプションメッセージを生成する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Function createProgressTitleOptionalMessage() As String

    ' レコード処理件数が指定されている場合、タイトルにレコード処理件数を付加する
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        createProgressTitleOptionalMessage = " ・・・" & applicationSetting.recProcessCountCustom & "件単位で実行"
    End If

End Function
