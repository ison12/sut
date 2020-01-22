VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' メニュー操作時の処理 - テーブル
'
' 作成者　：Ison
' 履歴　　：2009/01/24　新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' メンバ変数
' ________________________________________________________
' DBエクスプローラ
Private WithEvents frmDBExplorerVar        As frmDBExplorer
Attribute frmDBExplorerVar.VB_VarHelpID = -1
' テーブルシート作成フォーム
Private WithEvents frmTableSheetCreatorVar As frmTableSheetCreator
Attribute frmTableSheetCreatorVar.VB_VarHelpID = -1
' テーブルシートリストフォーム
Private WithEvents frmTableSheetListVar    As frmTableSheetList
Attribute frmTableSheetListVar.VB_VarHelpID = -1
' テーブルシートリストフォーム
Private WithEvents frmTableSheetUpdateVar  As frmTableSheetUpdate
Attribute frmTableSheetUpdateVar.VB_VarHelpID = -1
' テーブルシート生成
Private WithEvents tableSheetCreator       As ExeTableSheetCreator
Attribute tableSheetCreator.VB_VarHelpID = -1
' プログレスフォーム
Private WithEvents frmProgressVar          As frmProgress
Attribute frmProgressVar.VB_VarHelpID = -1
' レコード情報生成
Private WithEvents tableSheetReader        As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1

' DBコネクション
Private dbConn As Object
' アプリケーション設定情報
Private applicationSetting As ValApplicationSetting

' ________________________________________________________
' イベント
' ________________________________________________________
' =========================================================
' ▼DBエクスプローラ Excelシートに出力
' =========================================================
Private Sub frmDBExplorerVar_export(ByVal tableList As ValCollection, ByVal recFormat As REC_FORMAT)

    createEmptyBook
    
    Dim tableSheetList As ValCollection
    Set tableSheetList = createDataOfTableSheetList(tableList, recFormat)
    
    frmTableSheetCreatorVar_complete tableSheetList
    
End Sub

' =========================================================
' ▼DBエクスプローラ 閉じる
' =========================================================
Private Sub frmDBExplorerVar_closed()

End Sub

' =========================================================
' ▼テーブルシート作成 - テーブルシート作成開始（複数）
' =========================================================
Private Sub tableSheetCreator_createTableSheetsBegin(tableSheetList As ValCollection)

    frmProgressVar.initPri tableSheetList.count, "テーブルシート作成中"
    
    DoEvents
End Sub

' =========================================================
' ▼テーブルシート作成 - テーブルシート作成完了（複数）
' =========================================================
Private Sub tableSheetCreator_createTableSheets(tableSheetList As ValCollection)

    frmProgressVar.priMessage = "テーブルシート作成完了"
    frmProgressVar.priCount = tableSheetList.count
    
    DoEvents
End Sub

' =========================================================
' ▼テーブルシート作成 - テーブルシート作成開始
' =========================================================
Private Sub tableSheetCreator_createTableSheetBegin(tableSheet As ValTableWorksheet)

    frmProgressVar.priMessage = "テーブルシート作成中・・・" & tableSheet.table.schemaTableName
    
    DoEvents
End Sub

' =========================================================
' ▼テーブルシート作成 - テーブルシート作成完了
' =========================================================
Private Sub tableSheetCreator_createTableSheet(tableSheet As ValTableWorksheet)

    frmProgressVar.inclimentPri
    
    DoEvents
End Sub

' =========================================================
' ▼処理がキャンセルされた場合に呼び出されるイベント
' =========================================================
Private Sub frmProgressVar_cancel()

    If Not tableSheetCreator Is Nothing Then
        tableSheetCreator.canceled = True
    End If
    
    If Not tableSheetReader Is Nothing Then
        tableSheetReader.canceled = True
    End If

End Sub

' =========================================================
' ▼テーブルシート作成フォーム - キャンセル時に呼び出されるイベント
' =========================================================
Private Sub frmTableSheetCreatorVar_cancel()

End Sub

' =========================================================
' ▼テーブルシート作成フォーム - 完了時に呼び出されるイベント
' =========================================================
Private Sub frmTableSheetCreatorVar_complete(createTargetTable As ValCollection)

    On Error GoTo err

    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init

    ' ブックオブジェクト
    Dim book As Workbook
    
    ' ブックオブジェクトを取得する
    Set book = ActiveWorkbook
    
    ' ブックオブジェクトのチェック
    If book Is Nothing Then
    
        Set book = Application.Workbooks.Add
        book.activate
    End If

    frmProgressVar.ShowExt vbModeless, False
    frmProgressVar.title = "テーブルシート作成"
    
    DoEvents

    Set tableSheetCreator = New ExeTableSheetCreator
    tableSheetCreator.applicationSetting = applicationSetting
    tableSheetCreator.book = ActiveWorkbook
    tableSheetCreator.createSheets createTargetTable

    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    frmProgressVar.HideExt
    
    Exit Sub
err:

    Set longTimeProcessing = Nothing
    
    frmProgressVar.HideExt
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▼テーブルシート更新フォーム - OKボタン押下時に呼び出されるイベント
' =========================================================
Private Sub frmTableSheetUpdateVar_ok(ByVal recFormat As REC_FORMAT)

    On Error GoTo err

    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' ----------------------------------------------
    ' テーブルシートから一度テーブル情報を読み込む
    Dim srcTable As ValTableWorksheet
    
    Dim tableSheetReader As ExeTableSheetReader
    
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = dbConn
    Set tableSheetReader.sheet = ActiveSheet
    
    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    Set srcTable = tableSheetReader.readTableInfo
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' DBから最新のテーブル情報を読み込む
    ' DBメタ情報取得用オブジェクト生成クラス
    Dim dbObjFactory As New DbObjectFactory
    ' DBメタ情報取得用クラス
    Dim DbMetaInfoGetter  As IDbMetaInfoGetter
    
    Set DbMetaInfoGetter = dbObjFactory.createMetaInfoGetterObject(dbConn)
    
    ' テーブルシート
    Dim desTableSheet As New ValTableWorksheet
    ' テーブル情報
    Dim desTable      As ValDbDefineTable
    ' テーブル情報を取得する
    Set desTable = DbMetaInfoGetter.getTable(srcTable.table.schemaName, srcTable.table.tableName)
    
    If desTable Is Nothing Then
    
        err.Raise ConstantsError.ERR_NUMBER_NOT_EXIST_TABLE_INFO _
                , _
                , ConstantsError.ERR_DESC_NOT_EXIST_TABLE_INFO

        Exit Sub
    End If
    
    ' 最新のテーブル情報の行フォーマットをフォームで指定したものに設定する
    Set desTableSheet.table = desTable
    Set desTableSheet.sheet = srcTable.sheet
    desTableSheet.recFormat = recFormat
    desTableSheet.omitsSchema = srcTable.omitsSchema
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' テーブルシートを生成する
    
    ' テーブルシート生成用のオブジェクトを生成する
    ' イベントを受信しないようにするために、メンバ変数（tableSheetCreator）ではなくローカル変数を利用する。
    ' ここでの処理は、プログレスバーに進捗状況を表示しない。
    Dim tsCreator As ExeTableSheetCreator
    
    Set tsCreator = New ExeTableSheetCreator
    tsCreator.applicationSetting = applicationSetting
    tsCreator.book = ActiveWorkbook
    tsCreator.createSheet desTableSheet
    ' ----------------------------------------------
    
    ' ExcelLongTimeProcessingにて、Application.interactiveを無効にした状態で
    ' ダイアログ等を開いたり閉じたりすると、ウィンドウのZオーダーが乱れてしまう。（処理終了後に、Excelがアクティブにならない）
    ' そのため、ExcelLongTimeProcessingオブジェクトを明示的に破棄する
    Set longTimeProcessing = Nothing
    
    Exit Sub

err:

    Set longTimeProcessing = Nothing
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▼テーブルシート一覧フォーム - シート選択時に呼び出されるイベント
' =========================================================
Private Sub frmTableSheetListVar_selected(table As ValTableWorksheet)

    On Error GoTo err

    table.sheet.activate

    Exit Sub

err:

        

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
    
    Set frmDBExplorerVar = Nothing
    Set frmTableSheetCreatorVar = Nothing
    Set frmTableSheetListVar = Nothing
    Set frmTableSheetUpdateVar = Nothing
    Set frmProgressVar = Nothing
    Set tableSheetCreator = Nothing
    Set tableSheetReader = Nothing

    Set applicationSetting = Nothing
    Set dbConn = Nothing
End Sub

' =========================================================
' ▽初期化
'
' 概要　　　：初期化処理を実行する。
' 引数　　　：appSetting        アプリケーション設定情報
' 　　　　　　conn               DBコネクション
'
' =========================================================
Public Sub init(ByRef appSetting As ValApplicationSetting _
              , ByRef conn As Object)

    ' フォームを設定する
    Load frmProgress
    ' フォームを設定する
    Set frmProgressVar = frmProgress

    ' DBエクスプローラフォームをロードする
    Load frmDBExplorer
    ' イベント受信用の変数にDBエクスプローラフォームを設定する
    Set frmDBExplorerVar = frmDBExplorer

    ' テーブルシート作成フォームをロードする
    Load frmTableSheetCreator
    ' イベント受信用の変数にテーブルシート作成フォームを設定する
    Set frmTableSheetCreatorVar = frmTableSheetCreator
    
    ' テーブルシート一覧フォームをロードする
    Load frmTableSheetList
    ' イベント受信用の変数にテーブルシート一覧フォームを設定する
    Set frmTableSheetListVar = frmTableSheetList
    
    ' テーブルシート更新フォームをロードする
    Load frmTableSheetUpdate
    ' イベント受信用の変数にテーブルシート更新フォームを設定する
    Set frmTableSheetUpdateVar = frmTableSheetUpdate
    
    ' レコード情報生成オブジェクトを生成する
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = ActiveSheet
    
    ' アプリケーション設定情報を設定する
    Set applicationSetting = appSetting
    ' DBコネクションを設定する
    Set dbConn = conn
    
End Sub

' =========================================================
' ▽DBコネクションの更新
'
' 概要　　　：
'
' =========================================================
Public Sub updateDbConn(ByRef dbConn As Object)

    On Error GoTo err
    
    ' DBエクスプローラのコネクションを再設定する
    If Not frmDBExplorer Is Nothing Then
    
        Load frmDBExplorer
        Set frmDBExplorerVar = frmDBExplorer
        
        frmDBExplorerVar.DbConnection = dbConn
    End If

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽DBエクスプローラ表示
'
' 概要　　　：
'
' =========================================================
Public Sub showDbExplorer()

    On Error GoTo err
    
    ' DBエクスプローラフォームを表示する
    frmDBExplorerVar.ShowExt vbModeless, dbConn

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽テーブルシート一覧表示
'
' 概要　　　：
'
' =========================================================
Public Sub showTableSheetList()

    On Error GoTo err
    
    ' テーブルシート一覧フォームを表示する
    frmTableSheetListVar.ShowExt vbModeless

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽テーブルシート作成
'
' 概要　　　：
'
' =========================================================
Public Sub createTableSheet()

    On Error GoTo err

    createEmptyBook
    ' テーブル作成フォームを表示する
    frmTableSheetCreatorVar.ShowExt vbModal, applicationSetting, dbConn

    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽テーブルシート更新
'
' 概要　　　：
'
' =========================================================
Public Sub updateTableSheet()

    On Error GoTo err
    
    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    ' テーブルシート更新フォームを表示する
    frmTableSheetUpdateVar.ShowExt vbModal
    
    Exit Sub
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽テーブルシートのデータオブジェクトを生成する。
'
' 概要　　　：
'
' =========================================================
Private Function createDataOfTableSheet(ByRef table As ValDbDefineTable _
                                      , ByVal recFormat As REC_FORMAT) As ValTableWorksheet

    ' テーブルシートを生成する
    Dim tableSheet As ValTableWorksheet
    Set tableSheet = New ValTableWorksheet
    
    ' スキーマの省略有無を設定する
    If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
        tableSheet.omitsSchema = True
    Else
        tableSheet.omitsSchema = False
    End If
    ' 行フォーマットの生成
    tableSheet.recFormat = recFormat
    ' テーブル定義の設定
    Set tableSheet.table = table

    Set createDataOfTableSheet = tableSheet

End Function

' =========================================================
' ▽テーブル定義リストが格納されたリストオブジェクトを基にテーブルシートのリストオブジェクトを生成する。
'
' 概要　　　：
'
' =========================================================
Private Function createDataOfTableSheetList(ByRef tableList As ValCollection _
                                          , ByVal recFormat As REC_FORMAT) As ValCollection
                                          
    Dim tableSheetList As ValCollection
    Set tableSheetList = New ValCollection
    
    Dim table      As ValDbDefineTable
    Dim tableSheet As ValTableWorksheet
    
    ' テーブルリストからテーブル定義を取り出しテーブルシート情報に変換して
    ' リストに設定する
    For Each table In tableList.col
    
        Set tableSheet = createDataOfTableSheet(table, recFormat)
        tableSheetList.setItem tableSheet
    Next
    
    Set createDataOfTableSheetList = tableSheetList

End Function

' =========================================================
' ▽ブックが表示されていない場合、空のブックを表示する。
'
' 概要　　　：
'
' =========================================================
Private Function createEmptyBook()

    ' ブックオブジェクト
    Dim book As Workbook
    
    ' ブックオブジェクトを取得する
    Set book = ActiveWorkbook
    
    ' ブックオブジェクトのチェック
    If book Is Nothing Then
    
        Set book = Application.Workbooks.Add
        book.activate
    End If

End Function
