VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' メニュー操作時の処理 - Diff
'
' 作成者　：Ison
' 履歴　　：2018/12/30　新規作成
'
' 特記事項：
' *********************************************************

' プログレスフォーム
Private WithEvents frmProgressVar               As frmProgress
Attribute frmProgressVar.VB_VarHelpID = -1
' DB接続フォーム（イベント受信用に定義）
Private WithEvents frmDBConnectVar              As frmDBConnect
Attribute frmDBConnectVar.VB_VarHelpID = -1
' スナップショット取得フォーム
Private WithEvents frmSnapShotVar               As frmSnapshot
Attribute frmSnapShotVar.VB_VarHelpID = -1
' スナップショット取得フォーム（比較）
Private WithEvents frmSnapshotDiffVar           As frmSnapshotDiff
Attribute frmSnapshotDiffVar.VB_VarHelpID = -1

' スナップショットSQL定義シート生成
Private WithEvents snapSqlDefineSheetCreator    As ExeSnapSqlDefineSheetCreator
Attribute snapSqlDefineSheetCreator.VB_VarHelpID = -1
' スナップショットSQL定義シート読み取り
Private WithEvents snapSqlDefineSheetReader     As ExeSnapSqlDefineSheetReader
Attribute snapSqlDefineSheetReader.VB_VarHelpID = -1
' スナップショットSQL実行
Private WithEvents snapQueryExecution           As ExeSnapQueryExecution
Attribute snapQueryExecution.VB_VarHelpID = -1
' スナップショット比較
Private WithEvents snapDiff                     As ExeSnapDiff
Attribute snapDiff.VB_VarHelpID = -1

' アプリケーション設定
Private applicationSetting As ValApplicationSetting

' DBコネクションオブジェクト
Private dbConn As Object
' SQL定義情報
Private sqlDefineWorkSheet As ValSnapSqlDefineWorkSheet
' スナップショット
Private snapshotSet As ValCollection

' ________________________________________________________
' イベント定義
' ________________________________________________________
Private Sub frmProgressVar_cancel()

    ' キャンセルフラグをONにする
    If Not snapSqlDefineSheetCreator Is Nothing Then
        snapSqlDefineSheetCreator.canceled = True
    End If
    
    If Not snapSqlDefineSheetReader Is Nothing Then
        snapSqlDefineSheetReader.canceled = True
    End If
    
    If Not snapQueryExecution Is Nothing Then
        snapQueryExecution.canceled = True
    End If
    
    If Not snapDiff Is Nothing Then
        snapDiff.canceled = True
    End If

End Sub

' =========================================================
' ▽DB接続フォームでOKボタンが押下された場合に発生するイベント
'
' 概要　　　：
'
' =========================================================
Private Sub frmDBConnectVar_ok(ByVal connStr As String, ByVal connSimpleStr As String, ByVal connectInfo As ValDBConnectInfo)

    On Error GoTo err

    ' DB接続されている場合
    If Not Main.dbConn Is Nothing Then
    
        ' 切断する
        Main.SutUpdateDbConn Nothing, Empty, Empty
    End If
    
    ' DB接続し変数に設定する
    Dim conn As Object
    Set conn = ADOUtil.connectDb(connStr)
    
    ' DB接続文字列を設定する
    Main.SutUpdateDbConn conn, connStr, connSimpleStr
    
    Exit Sub
    
err:

    ' エラー発生
    Main.SutUpdateDbConn Nothing, Empty, Empty

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽DB接続フォームで処理がキャンセルされた場合に発生するイベント
'
' 概要　　　：
'
' =========================================================
Private Sub frmDBConnectVar_cancel()

    On Error GoTo err

    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▼スナップショット取得フォーム - スナップショット取得
' =========================================================
Private Sub frmSnapShotVar_execSnapshot(sheet As Worksheet)
    
    On Error GoTo err
    
    Dim ssKey  As String
    Dim ssList As ValCollection
    Dim ss     As ValSnapRecordsSet

    Set snapSqlDefineSheetReader = New ExeSnapSqlDefineSheetReader
    Set snapSqlDefineSheetReader.sheet = sheet
    Set snapSqlDefineSheetReader.conn = dbConn
    
    ' SQL定義情報を読み取る
    Set sqlDefineWorkSheet = snapSqlDefineSheetReader.readSqlDefineInfo
    
    ' SQLを実行しスナップショットを取得する
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    frmProgressVar.ShowExt vbModeless, False
    frmProgressVar.title = "クエリ実行"
    
    ' エディタが開いている間に接続状態が変更されている可能性があるので、コネクションを設定する
    Dim conn As Object
    Set conn = Main.getDBConnection
    
    Set snapQueryExecution = New ExeSnapQueryExecution
    Set snapQueryExecution.conn = conn
    snapQueryExecution.applicationSetting = Main.getApplicationSetting ' エディタが開いている間に接続状態が変更されている可能性があるので、アプリケーション設定情報を取得する
    snapQueryExecution.canceled = False
    
    ' -----------------------------------------------------------
    ' スナップショットデータの格納
    Set ss = snapQueryExecution.execute(sqlDefineWorkSheet)
    
    ssKey = sheet.parent.name & " " & sheet.name
    
    Set ssList = snapshotSet.getItem(ssKey)
    If ssList Is Nothing Then
        Set ssList = New ValCollection
        snapshotSet.setItem ssList, sheet.parent.name & " " & sheet.name
    End If
    
    ssList.setItem ss
    frmSnapShotVar_changeSql sheet
    ' -----------------------------------------------------------
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing

    frmProgressVar.HideExt
    frmSnapShotVar.Show vbModeless ' アクティブにするために表示メソッドを呼び出す
    
    Exit Sub
err:

    ' エラー情報を退避する
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    frmProgressVar.HideExt
    frmSnapShotVar.Show vbModeless ' アクティブにするために表示メソッドを呼び出す

    ' 退避したエラー情報を設定しなおす
    VBUtil.setErr errT

    ' エラー発生
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▼スナップショット取得フォーム - スナップショット取得キャンセル
' =========================================================
Private Sub frmSnapShotVar_cancel()

End Sub

' =========================================================
' ▼スナップショット取得フォーム - DB変更
' =========================================================
Private Sub frmSnapShotVar_changeDb()

    ' DB接続フォームを表示する
    frmDBConnect.ShowExt vbModal

End Sub

' =========================================================
' ▼スナップショット取得フォーム - SQL定義変更
' =========================================================
Private Sub frmSnapShotVar_changeSql(sheet As Worksheet)

    Dim ssKey  As String
    Dim ssList As ValCollection
    Dim ss     As ValSnapRecordsSet

    ssKey = sheet.parent.name & " " & sheet.name
    
    Set ssList = snapshotSet.getItem(ssKey)
    If Not ssList Is Nothing Then
    
        frmSnapShotVar.clearSnapshot
        
        For Each ss In ssList.col
            frmSnapShotVar.addSnapshot Format(ss.getDate, "yyyy/mm/dd hh:nn:ss") & " - " & ss.recordCount & "件", Empty
        Next
    
    Else
    
        frmSnapShotVar.clearSnapshot
        
    End If

End Sub

' =========================================================
' ▼スナップショット取得フォーム - スナップショットリストクリア
' =========================================================
Private Sub frmSnapShotVar_clearSnapshot(sheet As Worksheet)

    Dim resultMsg As Long
    
    resultMsg = VBUtil.showMessageBoxForYesNo("スナップショットをクリアしてもよろしいですか？" _
                   & vbNewLine _
                   & vbNewLine & "・[はい] → クリア" _
                   & vbNewLine & "・[いいえ] → 何もしない" _
                , ConstantsCommon.APPLICATION_NAME)

    ' いいえ押下
    If resultMsg = WinAPI_User.IDNO Then
    
        Exit Sub
    End If

    Dim ssKey  As String
    Dim ssList As ValCollection
    Dim ss     As ValSnapRecordsSet

    ssKey = sheet.parent.name & " " & sheet.name
    
    Set ssList = snapshotSet.getItem(ssKey)
    If Not ssList Is Nothing Then
        snapshotSet.remove ssKey
    End If
    
    frmSnapShotVar.clearSnapshot

End Sub

' =========================================================
' ▼スナップショット取得フォーム - スナップショットリストクリア
' =========================================================
Private Sub frmSnapShotVar_showSnapshotDiff(ByRef sheet As Worksheet)

    Dim ssKey  As String
    Dim ssList As ValCollection
    Dim ss     As ValSnapRecordsSet

    ssKey = sheet.parent.name & " " & sheet.name
    
    Set ssList = snapshotSet.getItem(ssKey)
    If Not ssList Is Nothing Then
        ' モーダル表示するのでロードを常に実行する
        If VBUtil.unloadFormIfChangeActiveBook(frmSnapshotDiff) Then Unload frmSnapshotDiff
        Load frmSnapshotDiff
        Set frmSnapshotDiffVar = frmSnapshotDiff
        frmSnapshotDiffVar.ShowExt vbModal, ssList
        
    Else
    
        VBUtil.showMessageBoxForWarning "スナップショットがありません。" _
                , ConstantsCommon.APPLICATION_NAME _
                , Nothing

    End If

End Sub

' =========================================================
' ▼スナップショット取得比較フォーム - キャンセル
' =========================================================
Private Sub frmSnapshotDiffVar_Cancel()

End Sub

' =========================================================
' ▼スナップショット取得比較フォーム - 比較実行
' =========================================================
Private Sub frmSnapshotDiffVar_execDiff(ByRef snapShotList As ValCollection, ByVal srcIndex As Long, ByVal desIndex As Long)

    On Error GoTo err
    
    ' モーダルフォームを非表示にする
    frmSnapshotDiffVar.Hide

    ' プログレスウィンドウを表示する
    frmProgressVar.ShowExt vbModeless, False
    frmProgressVar.title = "スナップショット比較"

    Set snapDiff = New ExeSnapDiff
    snapDiff.book = ActiveWorkbook
    snapDiff.applicationSetting = applicationSetting
    
    Dim valSnapRecordsSet1 As ValSnapRecordsSet
    Set valSnapRecordsSet1 = snapShotList.getItemByIndex(srcIndex + 1) ' リストの要素は、1始まりなので、+1する
    
    Dim valSnapRecordsSet2 As ValSnapRecordsSet
    Set valSnapRecordsSet2 = snapShotList.getItemByIndex(desIndex + 1) ' リストの要素は、1始まりなので、+1する
    
    snapDiff.execute valSnapRecordsSet1, valSnapRecordsSet2
    
    frmProgressVar.HideExt
    
    ' モーダルフォームを元に戻す
    frmSnapshotDiffVar.Show vbModal

    Exit Sub
err:
    
    ' エラー情報を退避する
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    frmProgressVar.HideExt
    
    ' 退避したエラー情報を設定しなおす
    VBUtil.setErr errT
    
    ' エラー発生
    Main.ShowErrorMessage
    
    ' モーダルフォームを元に戻す
    frmSnapshotDiffVar.Show

End Sub

' =========================================================
' ▼スナップショットSQL定義シート作成 - シート作成開始
' =========================================================
Private Sub snapSqlDefineSheetCreator_createSnapshotSqlDefineSheetBegin()

    frmProgressVar.initPri 1, "スナップショットSQL定義シート作成中"
    
    DoEvents

End Sub

' =========================================================
' ▼スナップショットSQL定義シート作成 - シート作成完了
' =========================================================
Private Sub snapSqlDefineSheetCreator_createSnapshotSqlDefineSheet()

    frmProgressVar.inclimentPri
    
    DoEvents

End Sub

' =========================================================
' ▼スナップショットSQL実行 - SQLリスト実行開始
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlListBegin(ByVal count As Long)

    frmProgressVar.initPri count, "SQL実行中"

End Sub

' =========================================================
' ▼スナップショットSQL実行 - SQLリスト実行終了
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlListComplete()

    frmProgressVar.priMessage = "SQL実行完了"
    
    DoEvents

    ' 長時間の処理のカーソル解除処理
    Dim cursorWait As New ExcelCursorWait: cursorWait.forceRestore

End Sub

' =========================================================
' ▼スナップショットSQL実行 - SQL実行開始
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlBegin(sql As String)

End Sub

' =========================================================
' ▼スナップショットSQL実行 - SQL実行成功
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlSuccessful(sql As String)

    frmProgressVar.inclimentPri
    
    DoEvents
    
End Sub

' =========================================================
' ▼スナップショットSQL実行 - SQL実行エラー
' =========================================================
Private Sub snapQueryExecution_ExecuteSqlError(ByRef sheet As Worksheet, ByVal recordSetIndex As Long, sql As String, sqlState As String, message As String)

    frmProgressVar.priMessage = "SQL実行エラー"
    
    ' デバッグモードのみの実行
    #If (DEBUG_MODE = 1) Then
    
        Debug.Print sql
    #End If
    
    DoEvents
    
    sheet.activate
    
    ' メッセージボックスに実行件数を表示する
    VBUtil.showMessageBoxForError _
        recordSetIndex & "行目でエラーが発生しました。" & vbNewLine & _
        "SQLSTATE=" & sqlState & vbNewLine & _
        " MESSAGE=" & message & vbNewLine & _
        "     SQL=" & sql _
        , ConstantsCommon.APPLICATION_NAME _
        , err
    
End Sub

' =========================================================
' ▼スナップショットの比較実行 - シート生成開始
' =========================================================
Private Sub snapDiff_createSnapshotDiffSheetBegin()

    frmProgressVar.priMessage = "比較結果シート生成中"

    DoEvents

    ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
    Dim cursorWait As New ExcelCursorWait: cursorWait.init True

End Sub

' =========================================================
' ▼スナップショットの比較実行 - シート生成
' =========================================================
Private Sub snapDiff_createSnapshotDiffSheet(sheet As Worksheet)

    frmProgressVar.priMessage = "比較結果シート生成完了"

    DoEvents

    ' 長時間の処理のカーソル解除処理
    Dim cursorWait As New ExcelCursorWait: cursorWait.forceRestore

End Sub

' =========================================================
' ▼スナップショットの比較実行 - レコードセット比較開始
' =========================================================
Private Sub snapDiff_diffRecordSetBegin(ByVal count As Long)

    frmProgressVar.initPri count, "レコードセット比較中"

    DoEvents

    ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
    Dim cursorWait As New ExcelCursorWait: cursorWait.init True

End Sub

' =========================================================
' ▼スナップショットの比較実行 - レコードセット比較開始
' =========================================================
Private Sub snapDiff_diffRecordSetOne()

    frmProgressVar.inclimentPri

    DoEvents

End Sub

' =========================================================
' ▼スナップショットの比較実行 - レコードセット比較
' =========================================================
Private Sub snapDiff_diffRecordSet()

    frmProgressVar.priMessage = "レコードセット比較完了"

    DoEvents

    ' 長時間の処理のカーソル解除処理
    Dim cursorWait As New ExcelCursorWait: cursorWait.forceRestore

End Sub

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set snapshotSet = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' プログレスフォームを破棄する
    Set frmProgress = Nothing
    ' イベント受信用のDB接続フォームを破棄する
    Set frmDBConnectVar = Nothing
    ' スナップショットフォームを破棄する
    Set frmSnapShotVar = Nothing
    Set frmSnapshotDiffVar = Nothing
    
    ' アプリケーション設定情報を破棄する
    Set applicationSetting = Nothing
    ' DBコネクションを破棄する
    Set dbConn = Nothing
    ' SQL定義情報を破棄する
    Set sqlDefineWorkSheet = Nothing
    
    Set snapSqlDefineSheetCreator = Nothing
    Set snapSqlDefineSheetReader = Nothing
    Set snapQueryExecution = Nothing
    Set snapDiff = Nothing
    
    Set snapshotSet = Nothing

End Sub

' =========================================================
' ▽初期化
'
' 概要　　　：初期化処理を実行する。
' 引数　　　：appSetting アプリケーション設定情報
' 　　　　　　appSettingColFmt アプリケーション設定情報（カラム書式）
' 　　　　　　conn       DBコネクション
'
' =========================================================
Public Sub init(ByRef appSetting As ValApplicationSetting _
              , ByRef appSettingColFmt As ValApplicationSettingColFormat _
              , ByRef conn As Object)

    ' DB接続フォームをロードする
    If VBUtil.unloadFormIfChangeActiveBook(frmDBConnect) Then Unload frmDBConnect
    Load frmDBConnect
    ' イベント受信用のDB接続フォームを初期化する
    Set frmDBConnectVar = frmDBConnect
    
    ' スナップショットフォームをロードする
    If VBUtil.unloadFormIfChangeActiveBook(frmSnapshot) Then Unload frmSnapshot
    Load frmSnapshot
    ' スナップショットフォームを設定する
    Set frmSnapShotVar = frmSnapshot

    ' スナップショット比較フォームをロードする
    If VBUtil.unloadFormIfChangeActiveBook(frmSnapshotDiff) Then Unload frmSnapshotDiff
    Load frmSnapshotDiff
    ' スナップショット比較フォームを設定する
    Set frmSnapshotDiffVar = frmSnapshotDiff

    ' フォームを設定する
    If VBUtil.unloadFormIfChangeActiveBook(frmProgress) Then Unload frmProgress
    Load frmProgress
    ' フォームを設定する
    Set frmProgressVar = frmProgress
    
    ' アプリケーション設定情報を取得する
    Set applicationSetting = appSetting
    
    ' DBコネクションを取得する
    Set dbConn = conn
    
    ' SQL定義情報を初期化する
    Set sqlDefineWorkSheet = Nothing

End Sub

' =========================================================
' ▽スナップショットSQLシート追加
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub createNewSheetSnapSqlDefine()

    On Error GoTo err

    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' プログレスウィンドウを表示する
    frmProgressVar.ShowExt vbModeless, False
    frmProgressVar.title = "スナップショットSQLシート追加"

    DoEvents

    ' ブックオブジェクト
    Dim book As Workbook
    
    ' ブックオブジェクトを取得する
    Set book = ActiveWorkbook
    
    ' ブックオブジェクトのチェック
    If book Is Nothing Then
    
        Set book = Application.Workbooks.Add
        book.activate
    End If

    Set snapSqlDefineSheetCreator = New ExeSnapSqlDefineSheetCreator
    snapSqlDefineSheetCreator.applicationSetting = applicationSetting
    snapSqlDefineSheetCreator.book = ActiveWorkbook
    snapSqlDefineSheetCreator.createSheet

    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgressVar.HideExt
    
    Exit Sub
    
err:

    Set longTimeProcessing = Nothing
    
    ' プログレスウィンドウを非表示にする
    frmProgressVar.HideExt
    
    ' エラー発生
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽スナップショットフォームの表示
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Sub showSnapshot()

    On Error GoTo err

    frmSnapShotVar.ShowExt vbModeless, applicationSetting, dbConn, dbConnStr
    
    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽DBコネクションの更新
'
' 概要　　　：conn    DBコネクション
'             connStr DB接続文字列
'
' =========================================================
Public Sub updateDbConn(ByRef conn As Object, ByVal connStr As String)

    On Error GoTo err
    
    Set dbConn = conn
    dbConnStr = connStr
    
    If Not frmSnapShotVar Is Nothing Then
        frmSnapShotVar.updateDbConn dbConn, dbConnStr
    End If

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

