VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetList 
   Caption         =   "テーブルシート一覧"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5940
   OleObjectBlob   =   "frmTableSheetList.frx":0000
End
Attribute VB_Name = "frmTableSheetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' テーブルシート一覧フォーム
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/04/03　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
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

' テーブルリスト
Private tableSheetList  As CntListBox

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal モーダルまたはモードレス表示指定
' 　　　　　　conn  DBコネクション
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants)

    activate

    ' デフォルトフォーカスコントロールを設定する
    lstTableSheet.SetFocus

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

    ' 最前面表示にする
    ExcelUtil.setUserFormTopMost Me

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

    ' テーブルシートリストを初期化する
    Set tableSheetList = New CntListBox: tableSheetList.control = lstTableSheet
    
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

    ' テーブルシートリストを破棄する
    Set tableSheetList = Nothing
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

    ' テーブルリスト
    Dim tableList As ValCollection
    
    ' テーブルシート読込オブジェクト
    Dim tableSheetReader As ExeTableSheetReader
    Set tableSheetReader = New ExeTableSheetReader
        
    ' ブック
    Dim book  As Workbook
    ' シート
    Dim sheet As Worksheet
    
    ' アクティブブックをbook変数に格納する
    Set book = ActiveWorkbook
    
    ' テーブルリストを初期化する
    Set tableList = New ValCollection
    
    ' ブックに含まれているシートを1件ずつ処理する
    For Each sheet In book.Worksheets
    
        Set tableSheetReader.sheet = sheet
        
        ' 対象シートがテーブルシートの場合
        If tableSheetReader.isTableSheet = True Then
        
            ' テーブルシートを読み込んでリストに設定する（テーブル情報のみ取得する）
            tableList.setItem tableSheetReader.readTableInfo(True)
        
        End If
    
    Next
    
    ' リストコントロールにテーブルシート情報を追加する
    tableSheetList.addNestedProperty tableList, "Table", "SchemaTableName", "TableComment"
    
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
' ▽テーブルシートリスト　ダブルクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstTableSheet_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    selectedTable
End Sub

' =========================================================
' ▽テーブルシートリスト　キー押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstTableSheet_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
    
        selectedTable
    End If
    
End Sub

' =========================================================
' ▽閉じるボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnClose_Click()

    RaiseEvent closed

    Me.HideExt
    
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
    
    Dim tableSheet As ValTableWorksheet

    Set selectedList = tableSheetList.selectedList

    If selectedList.count >= 1 Then
    
        Set tableSheet = selectedList.getItemByIndex(1)
        
        RaiseEvent selected(tableSheet)
    End If

End Sub
