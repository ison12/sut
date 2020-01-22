VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' メニュー操作時の処理 - File
'
' 作成者　：Ison
' 履歴　　：2009/03/14　新規作成
'
' 特記事項：
' *********************************************************

' プログレスフォーム
Private WithEvents frmProgressVar               As frmProgress
Attribute frmProgressVar.VB_VarHelpID = -1
' ファイル出力フォーム
Private WithEvents frmFileOutputVar             As frmFileOutput
Attribute frmFileOutputVar.VB_VarHelpID = -1
' DBクエリ一括実行フォーム
Private WithEvents frmDbQueryBatchVar           As frmDBQueryBatch
Attribute frmDbQueryBatchVar.VB_VarHelpID = -1
' レコード情報生成
Private WithEvents tableSheetReader             As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1

' 更新系クエリー生成
Private WithEvents update                       As ExeUpdate
Attribute update.VB_VarHelpID = -1
' 更新系クエリー生成（選択）
Private WithEvents updateSelection_             As ExeUpdate
Attribute updateSelection_.VB_VarHelpID = -1

' ファイル出力は "全て" と "範囲選択" の2種類に分かれる。
' 更新系クエリーオブジェクト(ExeUpdate)も分けられるので、どちらが実行されるかを判定するオブジェクトが必要になる。以下がその判定用オブジェクト。
' ファイル出力フォームの呼び出し前に、以下に"update"・"updateSelection_"の何れかを格納し
' ファイル出力フォームのOKイベントで、このオブジェクトに対して操作を実行する。
Private updateCurrent As ExeUpdate

' アプリケーション設定
Private applicationSetting As ValApplicationSetting

' ________________________________________________________
' イベント定義
' ________________________________________________________
Private Sub frmProgressVar_cancel()

    ' キャンセルフラグをONにする
    tableSheetReader.canceled = True
    update.canceled = True
    updateSelection_.canceled = True
    
End Sub

' =========================================================
' ▽ファイル出力
'
' 概要　　　：ファイル出力を実行する。
' 引数　　　：table      テーブルオブジェクト１
' 　　　　　　queryType  クエリ種類
' 　　　　　　filePath   ファイルパス
' 　　　　　　encode     エンコード
' 　　　　　　newline    改行コード
'
' =========================================================
Private Function execFileOutput(table As ValTableWorksheet _
                            , queryType As DB_QUERY_BATCH_TYPE _
                            , filePath As String _
                            , encode As String _
                            , newline As String) As Boolean

    On Error GoTo err

    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' 実行件数
    Dim resultCount As Long
    
    ' 更新オブジェクトのファイル出力に必要なオプションプロパティを設定する
    ' ファイル出力フラグ
    update.isFileOutput = True
    ' ファイルパス
    update.filePath = VBUtil.concatFilePath(filePath, table.sheetName & ".sql")
    ' 文字コード
    update.charSet = encode
    ' 改行コード
    update.newline = newline
    
    ' クエリ種類の設定
    If queryType = DB_QUERY_BATCH_TYPE.insert Then

        update.queryType = update.QUERY_TYPE_INSERT

    ElseIf queryType = DB_QUERY_BATCH_TYPE.update Then

        update.queryType = update.QUERY_TYPE_UPDATE

    ElseIf queryType = DB_QUERY_BATCH_TYPE.deleteOnSheet Then

        update.queryType = update.QUERY_TYPE_DELETE
        
    ElseIf queryType = DB_QUERY_BATCH_TYPE.deleteAll Then

        update.queryType = update.QUERY_TYPE_DELETE_ALL

    End If
    
    ' プログレスウィンドウを表示する
    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgressVar.ShowExt vbModeless, True
    Else
        frmProgressVar.ShowExt vbModeless, False
    End If
    frmProgressVar.title = "ファイル出力 - " & frmFileOutputVar.lblHeader.Caption
    
    resultCount = update.execute
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgressVar.HideExt
    
    execFileOutput = True
    
    Exit Function
    
err:
    
    execFileOutput = False
    
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgressVar.HideExt

End Function

' =========================================================
' ▼クエリ一括フォームイベント - 実行
' =========================================================
Private Sub frmDbQueryBatchVar_ok(ByVal dbQueryBatchMode As DB_QUERY_BATCH_MODE _
                                , ByVal filePath As String _
                                , ByVal characterCode As String _
                                , ByVal newline As String _
                                , ByVal tableWorksheets As ValCollection)

    Dim var As ValDbQueryBatchTableWorksheet
    Dim tableSheet As ValTableWorksheet
        
    Dim ret As Boolean
        
    For Each var In tableWorksheets.col

        ' テーブルシートを取得する
        Set tableSheet = var.tableWorksheet
        ' 対象シートを設定する
        Set update.sheet = tableSheet.sheet

        ' 対象シートをアクティブにする
        update.sheet.activate
        
        ' ファイル出力を実行する
        ret = execFileOutput(var.tableWorksheet, var.dbQueryBatchType, filePath, characterCode, newline)
        
        If ret = False Then
        
            err.Number = ERR_NUMBER_FILE_OUTPUT_FAILED
            err.Description = ERR_DESC_FILE_OUTPUT_FAILED
            Main.ShowErrorMessage
            
            Exit Sub
        End If
        
    Next
    
    VBUtil.showMessageBoxForInformation "ファイルが出力されました。", "Sut"

End Sub

' =========================================================
' ▼更新イベント - カラムリスト読み込み開始
' =========================================================
Private Sub update_readColumnListBegin(ByVal cnt As Long)

    frmProgressVar.initPri cnt, "カラム定義読み込み中"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント - カラムリスト読み込み
' =========================================================
Private Sub update_readColumn(column As ValDbDefineColumn)

    frmProgressVar.inclimentPri
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント - カラムリスト読み込み完了
' =========================================================
Private Sub update_readColumnList(columnList As ValCollection)

    frmProgressVar.priMessage = "カラム定義読み込み完了"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント - レコードリストサイズ読み込み
' =========================================================
Private Sub update_readRecordListSize(ByVal cnt As Long)

    frmProgressVar.initPri cnt, "レコードリストサイズ読み込み完了"
End Sub

' =========================================================
' ▼更新イベント - レコードリスト読み込み開始
' =========================================================
Private Sub update_readRecordListBegin(ByVal cnt As Long)

    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        ' 分割して処理を実行していることが分かるように
        ' 処理レコードの範囲をメッセージ（その他）に埋め込む
        
        frmProgressVar.priMessage = "・・・"
        frmProgressVar.initSec applicationSetting.recProcessCountCustom, "レコードリスト読み込み中"
        
    Else
    
        frmProgressVar.priMessage = "レコードリスト読み込み中"
        
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
        
        frmProgressVar.secMessage = "レコードリスト読み込み完了"
        frmProgressVar.secCount = VBUtil.arraySize(recordList.values)
        
    Else
    
        frmProgressVar.priMessage = "レコードリスト読み込み完了"
        frmProgressVar.priCount = VBUtil.arraySize(recordList.values)
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
        
        frmProgressVar.initSec count, "SQLファイル出力中"
    Else
    
        frmProgressVar.initPri count, "SQLファイル出力中"
        
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

    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
    
        frmProgressVar.inclimentSec
    Else
    
        frmProgressVar.inclimentPri
    End If
    
    DoEvents

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
    
        frmProgressVar.secMessage = "SQLファイル出力完了"
        frmProgressVar.priCount = frmProgressVar.priCount + count
        
    Else
    
        frmProgressVar.priMessage = "SQLファイル出力完了"
    End If
    
    DoEvents

End Sub

' =========================================================
' ▼更新イベント（選択） - カラムリスト読み込み開始
' =========================================================
Private Sub updateSelection__readColumnListBegin(ByVal cnt As Long)

    frmProgressVar.initPri cnt, "カラム定義読み込み中"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - カラムリスト読み込み
' =========================================================
Private Sub updateSelection__readColumn(column As ValDbDefineColumn)

    frmProgressVar.inclimentPri
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - カラムリスト読み込み完了
' =========================================================
Private Sub updateSelection__readColumnList(columnList As ValCollection)

    frmProgressVar.priMessage = "カラム定義読み込み完了"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - セル選択時 レコードリスト読み込み開始
' =========================================================
Private Sub updateSelection__readRecordListBySelectionBegin(ByVal cnt As Long)

    frmProgressVar.initPri cnt, "レコードリスト読み込み中"

    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - セル選択時 レコードリスト読み込み
' =========================================================
Private Sub updateSelection__readRecordListBySelectionOne(recordList As ValRecordList)

    Dim count As Long
    
    count = VBUtil.arraySize(recordList.values)
    
    frmProgressVar.priCount = frmProgressVar.priCount + count
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - セル選択時 レコードリスト読み込み
' =========================================================
Private Sub updateSelection__readRecordListBySelection(recordLists As ValCollection)

    frmProgressVar.priMessage = "レコードリスト読み込み完了"
    
    ' カウント数を一旦0クリアする（カウント総数はそのまま）
    ' ExecuteSqlBeginイベントはセル選択時の場合は複数回呼び出されるので
    ' カウントをリセットするような処理はせずに、その手前で呼び出される本イベントでリセットを実施する
    frmProgressVar.priCount = 0
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - SQL実行開始
' =========================================================
Private Sub updateSelection__ExecuteSqlBegin(ByVal count As Long)

    frmProgressVar.priMessage = "SQLファイル出力中"
    
    DoEvents
End Sub

' =========================================================
' ▼更新イベント（選択） - SQL実行
' =========================================================
Private Sub updateSelection__ExecuteSql(ByRef table As ValTableWorksheet _
                            , ByRef recordList As ValRecordList _
                            , ByVal recordListIndex As Long _
                            , ByRef sql As String _
                            , ByVal count As Long)

    frmProgressVar.inclimentPri

    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If

    DoEvents
    
End Sub

' =========================================================
' ▼更新イベント（選択） - SQL実行 成功
' =========================================================
Private Sub updateSelection__ExecuteSqlSuccessful(ByVal count As Long)

    frmProgressVar.priMessage = "SQL実行完了"
    
    DoEvents

End Sub

' =========================================================
' ▼ファイル出力フォーム OKボタン押下時のイベントプロシージャ
' =========================================================
Private Sub frmFileOutputVar_ok(ByVal filePath As String, ByVal characterCode As String, ByVal newline As String)

    On Error GoTo err
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' 実行件数
    Dim resultCount As Long
    
    ' 更新オブジェクトのファイル出力に必要なオプションプロパティを設定する
    ' ファイル出力フラグ
    updateCurrent.isFileOutput = True
    ' ファイルパス
    updateCurrent.filePath = filePath
    ' 文字コード
    updateCurrent.charSet = characterCode
    ' 改行コード
    updateCurrent.newline = newline
    
    ' プログレスウィンドウを表示する
    ' 処理単位が設定されている場合
    If applicationSetting.recProcessCount = applicationSetting.REC_PROCESS_COUNT_COSTOM Then
        frmProgressVar.ShowExt vbModeless, True
    Else
        frmProgressVar.ShowExt vbModeless, False
    End If
    frmProgressVar.title = "ファイル出力 - " & frmFileOutputVar.lblHeader.Caption
    
    ' 更新を実行する
    If updateCurrent Is update Then
    
        resultCount = updateCurrent.execute
    Else
    
        resultCount = updateCurrent.executeSelection
    End If
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgressVar.HideExt
    
    ' メッセージボックスに実行件数を表示する
    VBUtil.showMessageBoxForInformation resultCount & "件処理しました。", ConstantsCommon.APPLICATION_NAME
    
    Exit Sub
    
err:
    
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgressVar.HideExt

    ' エラー発生
    Main.ShowErrorMessage
    
End Sub


' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' プログレスフォームを破棄する
    Set frmProgress = Nothing
    ' ファイル出力フォームを破棄する
    Set frmFileOutputVar = Nothing
    ' 一括クエリフォームを破棄する
    Set frmDbQueryBatchVar = Nothing
    
    ' レコード情報生成オブジェクトを破棄する
    Set tableSheetReader = Nothing
    
    ' 更新クエリー実行オブジェクトを破棄する
    Set update = Nothing
    Set updateSelection_ = Nothing
    Set updateCurrent = Nothing
    
    ' アプリケーション設定情報を破棄する
    Set applicationSetting = Nothing
    
End Sub

' =========================================================
' ▽初期化
'
' 概要　　　：初期化処理を実行する。
' 引数　　　：appSetting  アプリケーション設定情報
' 　　　　　　appSettingColFmt アプリケーション設定情報（カラム書式）
' 　　　　　　conn        DBコネクション
'
' =========================================================
Public Sub init(ByRef appSetting As ValApplicationSetting _
              , ByRef appSettingColFmt As ValApplicationSettingColFormat _
              , ByRef conn As Object)

    ' フォームを設定する
    Load frmProgress
    ' フォームを設定する
    Set frmProgressVar = frmProgress
    
    ' フォームをロードする
    Load frmFileOutput
    ' フォームを設定する
    Set frmFileOutputVar = frmFileOutput
    
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
    
    ' アプリケーション設定情報を取得する
    Set applicationSetting = appSetting

End Sub

' =========================================================
' ▽INSERT + UPDATE 実行（全レコード）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub insertUpdateAll()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
        
    ' テーブルシートオブジェクト
    Dim tableSheet As ValTableWorksheet
    ' テーブルシートオブジェクトを取得する
    Set tableSheet = tableSheetReader.readTableInfo

    ' 対象となる更新系クエリーを保存しておく
    Set updateCurrent = update
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_INSERT_UPDATE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "INSERT + UPDATE SQL" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽INSERT + UPDATE 実行（選択範囲）
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub insertUpdateSelection()

    On Error GoTo err

    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
        
    ' テーブルオブジェクト
    Dim tableSheet As ValTableWorksheet
    ' テーブルオブジェクトを取得する
    Set tableSheet = tableSheetReader.readTableInfo

    ' 対象となる更新系クエリーを保存しておく
    Set updateCurrent = updateSelection_
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_INSERT_UPDATE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "INSERT + UPDATE SQL（範囲選択）" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

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
        
    ' テーブルシートオブジェクト
    Dim tableSheet As ValTableWorksheet
    ' テーブルシートオブジェクトを取得する
    Set tableSheet = tableSheetReader.readTableInfo

    ' 対象となる更新系クエリーを保存しておく
    Set updateCurrent = update
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_INSERT
    
    frmFileOutputVar.ShowExt vbModal _
                            , "INSERT SQL" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

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
        
    ' テーブルオブジェクト
    Dim tableSheet As ValTableWorksheet
    ' テーブルオブジェクトを取得する
    Set tableSheet = tableSheetReader.readTableInfo

    ' 対象となる更新系クエリーを保存しておく
    Set updateCurrent = updateSelection_
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_INSERT
    
    frmFileOutputVar.ShowExt vbModal _
                            , "INSERT SQL（範囲選択）" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

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
        
    ' テーブルオブジェクト
    Dim tableSheet As ValTableWorksheet
    ' テーブルオブジェクトを取得する
    Set tableSheet = tableSheetReader.readTableInfo

    ' 対象となる更新系クエリーを保存しておく
    Set updateCurrent = update
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_UPDATE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "UPDATE SQL" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

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
        
    ' テーブルオブジェクト
    Dim tableSheet As ValTableWorksheet
    ' テーブルオブジェクトを取得する
    Set tableSheet = tableSheetReader.readTableInfo

    ' 対象となる更新系クエリーを保存しておく
    Set updateCurrent = updateSelection_
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_UPDATE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "UPDATE SQL（範囲選択）" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

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
        
    ' テーブルシートオブジェクト
    Dim tableSheet As ValTableWorksheet
    ' テーブルオブジェクトを取得する
    Set tableSheet = tableSheetReader.readTableInfo

    ' 対象となる更新系クエリーを保存しておく
    Set updateCurrent = update
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_DELETE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "DELETE SQL" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

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
        
    ' テーブルシートオブジェクト
    Dim tableSheet As ValTableWorksheet
    ' テーブルシートオブジェクトを取得する
    Set tableSheet = tableSheetReader.readTableInfo

    ' 対象となる更新系クエリーを保存しておく
    Set updateCurrent = updateSelection_
    updateCurrent.queryType = updateCurrent.QUERY_TYPE_DELETE
    
    frmFileOutputVar.ShowExt vbModal _
                            , "DELETE SQL（範囲選択）" _
                            , tableSheet.table.schemaTableName & ".sql"

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽一括ファイル出力
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub batchFile()

    On Error GoTo err
        
    ' ブック
    Dim book  As Workbook
    ' アクティブブックをbook変数に格納する
    Set book = ActiveWorkbook
    
    ' フォームを設定する
    Load frmDBQueryBatch
    Set frmDbQueryBatchVar = frmDBQueryBatch
    
    frmDbQueryBatchVar.ShowExt vbModal, FileOutput, insertUpdate, book
    
    Set frmDbQueryBatchVar = Nothing

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Sub

