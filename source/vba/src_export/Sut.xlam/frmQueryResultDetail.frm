VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryResultDetail 
   Caption         =   "クエリ結果詳細"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15900
   OleObjectBlob   =   "frmQueryResultDetail.frx":0000
End
Attribute VB_Name = "frmQueryResultDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリ結果詳細フォーム
'
' 作成者　：Ison
' 履歴　　：2020/02/19　新規作成
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
Public Event selected(ByRef tableSheet As ValTableWorksheet, ByVal cell As String)

' =========================================================
' ▽閉じるボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event closed()

' 詳細情報リストでの選択項目インデックス
Private detailInfoSelectedIndex As Long
' 詳細情報リストでの選択項目オブジェクト
Private detailInfoSelectedItem As ValQueryResultDetailInfo

' クエリ結果情報
Private queryResultInfoParam As ValQueryResultInfo
' 詳細情報リスト
Private detailInfoList  As CntListBox

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
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal queryResultInfo As ValQueryResultInfo)

    ' パラメータ設定
    Set queryResultInfoParam = queryResultInfo

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
' ▽詳細情報リスト　選択肢変更時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstDetailInfo_Change()

    selectedTable
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
' ▽詳細情報のコピークリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDetailInfoCopy_Click()

    Dim selectedIndex As Long
    Dim selectedItem As ValQueryResultDetailInfo
    
    ' 現在選択されているインデックスを取得
    selectedIndex = detailInfoList.getSelectedIndex

    ' 未選択の場合
    If selectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    Set selectedItem = detailInfoList.getSelectedItem
    
    WinAPI_Clipboard.SetClipboard selectedItem.tabbedInfoHeader & vbNewLine & getDetailInfoForClipboardFormat(selectedItem)
    
End Sub

' =========================================================
' ▽詳細情報の全てコピーボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdAllDetailInfoCopy_Click()

    Dim data As New StringBuilder
    Dim var As Variant
    
    Dim i As Long
    
    For Each var In detailInfoList.collection.col
        If i <= 0 Then
            data.append var.tabbedInfoHeader & vbNewLine
        End If
        data.append getDetailInfoForClipboardFormat(var)
        i = i + 1
    Next
    
    WinAPI_Clipboard.SetClipboard data.str

End Sub

' =========================================================
' ▽詳細情報のクリップボードフォーマット形式文字列取得
'
' 概要　　　：詳細情報のクリップボードフォーマット形式文字列を取得する。
' 引数　　　：var 詳細情報
' 戻り値　　：詳細情報のクリップボードフォーマット形式文字列取得
'
' =========================================================
Private Function getDetailInfoForClipboardFormat(ByVal var As ValQueryResultDetailInfo) As String

    getDetailInfoForClipboardFormat = var.tabbedInfo & vbNewLine

End Function

' =========================================================
' ▽初期化処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub initial()

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
    
    Dim queryResultDetailInfo As ValQueryResultDetailInfo
    
    ' 詳細情報リストに表示情報を反映する
    Set detailInfoList = New CntListBox: detailInfoList.init lstDetailInfo
    addDetailInfoList queryResultInfoParam.detailList

    detailInfoList.setSelectedIndex 0
    
    txtSheetName.value = queryResultInfoParam.sheetNameOrSheetTableName

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

    Dim selected As ValQueryResultDetailInfo
    Set selected = detailInfoList.getSelectedItem

    If Not selected Is Nothing Then
        RaiseEvent selected(queryResultInfoParam.tableWorksheet, selected.cell)
    End If

End Sub

' =========================================================
' ▽詳細情報リストを追加
'
' 概要　　　：
' 引数　　　：valDetailInfoList     詳細情報リスト
'     　　　  isAppend              追加有無フラグ
' 戻り値　　：
'
' =========================================================
Private Sub addDetailInfoList(ByVal valDetailInfoList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    detailInfoList.addAll valDetailInfoList _
                       , "cell", "messageWithSqlState", "queryWithoutNewLine" _
                       , isAppend:=isAppend
    
End Sub


