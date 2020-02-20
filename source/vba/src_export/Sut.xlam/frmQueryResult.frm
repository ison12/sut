VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryResult 
   Caption         =   "クエリ結果"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15900
   OleObjectBlob   =   "frmQueryResult.frx":0000
End
Attribute VB_Name = "frmQueryResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' クエリ結果フォーム
'
' 作成者　：Ison
' 履歴　　：2020/01/18　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽テーブルを選択した場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：tableSheet テーブルシート
'           : row        行番号
'
' =========================================================
Public Event selectedDetail(ByRef tableSheet As ValTableWorksheet, ByVal cell As String)

' =========================================================
' ▽テーブルを選択した場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：tableSheet テーブルシート
'
' =========================================================
Public Event selected(ByRef tableSheet As ValTableWorksheet)

' =========================================================
' ▽閉じるボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event closed()

' クエリ結果詳細情報の一件毎の表示（子画面）
Private WithEvents frmQueryResultDetailVar As frmQueryResultDetail
Attribute frmQueryResultDetailVar.VB_VarHelpID = -1

' テーブルリストでの選択項目インデックス
Private tableSheetSelectedIndex As Long
' テーブルリストでの選択項目オブジェクト
Private tableSheetSelectedItem As ValDbQueryBatchTableWorksheet

' クエリ結果情報リスト
Private queryResultSetInfoParam As ValQueryResultSetInfo
' テーブルリスト
Private tableSheetList  As CntListBox

' 対象ブック
Private targetBook As Workbook
' 対象ブックを取得する
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal              モーダルまたはモードレス表示指定
'             queryResultSetInfo クエリ結果セット情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal queryResultSetInfo As ValQueryResultSetInfo)

    ' パラメータ設定
    Set queryResultSetInfoParam = queryResultSetInfo

    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal

End Sub

' =========================================================
' ▽フォーム非表示
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub HideExt()

    deactivate
    
    Main.storeFormPosition Me.name, Me
    Me.Hide
    
End Sub

' =========================================================
' ▽フォーム初期化時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
    ' ロード時点のアクティブブックを保持しておく
    Set targetBook = ExcelUtil.getActiveWorkbook
    ' 初期化処理を実行する
    initial

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォーム破棄時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Terminate()

    On Error GoTo err
    
    ' 破棄処理を実行する
    unInitial
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォームアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Activate()

End Sub

' =========================================================
' ▽フォームの閉じる時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' 本処理では処理自体をキャンセルする
        Cancel = True
        ' 以下のイベント経由で閉じる
        cmdClose_Click
    End If
    
End Sub

' =========================================================
' ▽テーブルシートリスト　選択肢変更時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstTableSheet_Change()

    selectedTable
End Sub

' =========================================================
' ▽詳細ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDetail_Click()


    Dim selectedList As ValCollection
    Set selectedList = tableSheetList.getSelectedList
    
    If selectedList.count <= 0 Then
    
        ' 終了する
        Exit Sub
    End If

    Dim queryResultInfo As ValQueryResultInfo
    Set queryResultInfo = selectedList.getItemByIndex(1)

    If VBUtil.unloadFormIfChangeActiveBook(frmQueryResultDetail) Then Unload frmQueryResultDetail
    Load frmQueryResultDetail
    Set frmQueryResultDetailVar = frmQueryResultDetail
    frmQueryResultDetail.ShowExt vbModal, queryResultInfo
                            
    Set frmQueryResultDetailVar = Nothing
    
End Sub

' =========================================================
' ▽閉じるボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdClose_Click()

    On Error GoTo err
    
    ' フォームを閉じる
    HideExt
    
    ' 閉じるイベントを送信する
    RaiseEvent closed

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽クエリ結果詳細の選択時のイベント処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmQueryResultDetailVar_selected(tableSheet As ValTableWorksheet, ByVal cell As String)

    RaiseEvent selectedDetail(tableSheet, cell)
End Sub

' =========================================================
' ▽クエリ結果詳細の閉じる処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmQueryResultDetailVar_closed()

End Sub

' =========================================================
' ▽初期化処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub initial()

    ' 閉じるボタンを非表示にする
    cmdClose.Width = 0
    
End Sub

' =========================================================
' ▽後始末処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub unInitial()

End Sub

' =========================================================
' ▽アクティブ時の処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()
    
    lblErrorMessage.visible = False
    
    Dim queryResultInfo As ValQueryResultInfo
    
    ' テーブルシートリストに表示情報を反映する
    Set tableSheetList = New CntListBox: tableSheetList.init lstTableSheet
    addTableSheetList queryResultSetInfoParam.queryResultInfoList

    Dim i As Long: i = 0
    Dim selectedIndex As Long: selectedIndex = -1
    
    For Each queryResultInfo In tableSheetList.collection.col
    
        If queryResultInfo.sheetName = ActiveSheet.name Then
            selectedIndex = i
        End If
    
        i = i + 1
    Next
    
    If selectedIndex <> -1 Then
        ' アクティブシートを選択状態にする
        tableSheetList.setSelectedIndex selectedIndex
    End If

    ' エラーがある場合に、エラーメッセージを表示する
    Dim erroredResultInfoCount As Long
    
    erroredResultInfoCount = 0
    For Each queryResultInfo In tableSheetList.collection.col
    
        If queryResultInfo.errorCount > 0 Then
        
            erroredResultInfoCount = erroredResultInfoCount + 1
        End If
    
    Next
    
    If erroredResultInfoCount > 0 Then
    
        lblErrorMessage.visible = True
        lblErrorMessage.Caption = "処理結果にエラーがあります。対象のシートを選択してエラー内容を確認してください。"
    End If

End Sub

' =========================================================
' ▽ノンアクティブ時の処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub deactivate()

End Sub

' =========================================================
' ▽テーブル選択時の処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub selectedTable()

    Dim selectedList As ValCollection
    
    Dim queryResultInfo As ValQueryResultInfo
    Dim tableSheet      As ValTableWorksheet

    Set selectedList = tableSheetList.getSelectedList

    If selectedList.count >= 1 Then
    
        Set queryResultInfo = selectedList.getItemByIndex(1)
        
        If Not queryResultInfo.tableWorksheet Is Nothing Then
            Set tableSheet = queryResultInfo.tableWorksheet
            RaiseEvent selected(tableSheet)
        End If
        
    End If

End Sub

' =========================================================
' ▽テーブルシートリストを追加
'
' 概要　　　：
' 引数　　　：valTableSheetList テーブルシートリスト
'     　　　  isAppend              追加有無フラグ
' 戻り値　　：
'
' =========================================================
Private Sub addTableSheetList(ByVal valTableSheetList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    tableSheetList.addAll valTableSheetList _
                       , "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName", "processErrorCount" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ▽テーブルシートを追加
'
' 概要　　　：
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：
'
' =========================================================
Private Sub addTableSheet(ByVal tableSheet As ValDbQueryBatchTableWorksheet)
    
    tableSheetList.addItemByProp tableSheet, "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName", "processErrorCount"
    
End Sub

' =========================================================
' ▽テーブルシートを変更
'
' 概要　　　：
' 引数　　　：index インデックス
'     　　　  rec   テーブルシート
' 戻り値　　：
'
' =========================================================
Private Sub setTableSheet(ByVal index As Long, ByVal rec As ValDbQueryBatchTableWorksheet)
    
    tableSheetList.setItem index, rec, "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName", "processErrorCount"
    
End Sub


