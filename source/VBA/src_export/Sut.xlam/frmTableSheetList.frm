VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetList 
   Caption         =   "テーブルシート一覧"
   ClientHeight    =   8415.001
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10905
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
' 作成者　：Ison
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

' フィルタなし状態のテーブルリスト
Private tableSheetWithoutFilterList As ValCollection
' テーブルリスト
Private tableSheetList  As CntListBox

Private inFilterProcess As Boolean

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
    Set tableSheetList = New CntListBox: tableSheetList.init lstTableSheet
    
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
    Dim tableDistinctList As ValCollection
    Dim tableList As ValCollection
    Dim tableWorksheet As ValTableWorksheet
    
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
    Set tableSheetWithoutFilterList = New ValCollection
    
    Dim i As Long: i = 0
    Dim selectedIndex As Long: selectedIndex = -1
    
    ' ブックに含まれているシートを1件ずつ処理する
    For Each sheet In book.Worksheets
    
        Set tableSheetReader.sheet = sheet
        
        ' 対象シートがテーブルシートの場合
        If tableSheetReader.isTableSheet = True Then
        
            ' テーブルシートを読み込んでリストに設定する（テーブル情報のみ取得する）
            Set tableWorksheet = tableSheetReader.readTableInfo(True)
            
            tableList.setItem tableWorksheet
            tableSheetWithoutFilterList.setItem tableWorksheet
            
            If tableWorksheet.sheetName = ActiveSheet.name Then
                selectedIndex = i
            End If
        
            i = i + 1
        End If
    
    Next
    
    ' リストコントロールにテーブルシート情報を追加する
    addTableSheetList tableList, False
    
    If selectedIndex <> -1 Then
        ' アクティブシートを選択状態にする
        tableSheetList.setSelectedIndex selectedIndex
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
' ▽テーブルシートリスト更新ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdTableSheetListUpdate_Click()

    activate
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
' ▽フィルタコンボボックス変更時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboFilter_Change()

    On Error GoTo err

    Dim currentFilterText As String

    ' 本イベントプロシージャ内部で、同コントロールを変更することによる変更イベントが
    ' 再帰的に発生しても良いように
    ' フラグを参照して再実行されないようにする判定を実施
    If inFilterProcess = False Then

        inFilterProcess = True
    
        currentFilterText = cboFilter.text
        
        filterTableSheetList "*" & currentFilterText & "*"
        
        clearFilterCondition True
    
        inFilterProcess = False

    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ▽フィルタトグル全般の変更時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub tglFilterA_Click()
    filterToggle tglFilterA, "A"
End Sub
Private Sub tglFilterB_Click()
    filterToggle tglFilterB, "B"
End Sub
Private Sub tglFilterC_Click()
    filterToggle tglFilterC, "C"
End Sub
Private Sub tglFilterD_Click()
    filterToggle tglFilterD, "D"
End Sub
Private Sub tglFilterE_Click()
    filterToggle tglFilterE, "E"
End Sub
Private Sub tglFilterF_Click()
    filterToggle tglFilterF, "F"
End Sub
Private Sub tglFilterG_Click()
    filterToggle tglFilterG, "G"
End Sub
Private Sub tglFilterH_Click()
    filterToggle tglFilterH, "H"
End Sub
Private Sub tglFilterI_Click()
    filterToggle tglFilterI, "I"
End Sub
Private Sub tglFilterJ_Click()
    filterToggle tglFilterJ, "J"
End Sub
Private Sub tglFilterK_Click()
    filterToggle tglFilterK, "K"
End Sub
Private Sub tglFilterL_Click()
    filterToggle tglFilterL, "L"
End Sub
Private Sub tglFilterM_Click()
    filterToggle tglFilterM, "M"
End Sub
Private Sub tglFilterN_Click()
    filterToggle tglFilterN, "N"
End Sub
Private Sub tglFilterO_Click()
    filterToggle tglFilterO, "O"
End Sub
Private Sub tglFilterP_Click()
    filterToggle tglFilterP, "P"
End Sub
Private Sub tglFilterQ_Click()
    filterToggle tglFilterQ, "Q"
End Sub
Private Sub tglFilterR_Click()
    filterToggle tglFilterR, "R"
End Sub
Private Sub tglFilterS_Click()
    filterToggle tglFilterS, "S"
End Sub
Private Sub tglFilterT_Click()
    filterToggle tglFilterT, "T"
End Sub
Private Sub tglFilterU_Click()
    filterToggle tglFilterU, "U"
End Sub
Private Sub tglFilterV_Click()
    filterToggle tglFilterV, "V"
End Sub
Private Sub tglFilterW_Click()
    filterToggle tglFilterW, "W"
End Sub
Private Sub tglFilterX_Click()
    filterToggle tglFilterX, "X"
End Sub
Private Sub tglFilterY_Click()
    filterToggle tglFilterY, "Y"
End Sub
Private Sub tglFilterZ_Click()
    filterToggle tglFilterZ, "Z"
End Sub
Private Sub tglFilterOther_Click()
    
    ' Otherの処理だけ「～以外」という検索なので別の処理として定義
    
    On Error GoTo err

    ' 本イベントプロシージャ内部で、同コントロールを変更することによる変更イベントが
    ' 再帰的に発生しても良いように
    ' フラグを参照して再実行されないようにする判定を実施
    If inFilterProcess = False Then

        inFilterProcess = True
        
        If tglFilterOther.value = True Then
            ' アルファベット以外の文字で始まる情報で検索
            filterTableSheetListForRegExp "[^a-zA-Z]*"
            
            clearFilterCondition
            tglFilterOther.value = True
        Else
            filterTableSheetListForRegExp ""
        End If
        
        inFilterProcess = False
        
    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ▽トグル系フィルタ条件の共通処理
'
' 概要　　　：
' 引数　　　：state   トグルボタン
'     　　　  keyword キーワード
' 戻り値　　：
'
' =========================================================
Private Sub filterToggle(ByVal state As ToggleButton, ByVal keyword As String)

    On Error GoTo err

    If inFilterProcess = False Then

        inFilterProcess = True
        
        If state.value = True Then
            filterTableSheetList keyword & "*"
            
            clearFilterCondition
            state.value = True
        Else
            filterTableSheetList ""
        End If
        
        inFilterProcess = False
        
    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ▽フィルタ条件のクリア処理
'
' 概要　　　：
' 引数　　　：isNotClearComboFilter コンボボックスのフィルタをクリアするかどうかのフラグ
' 戻り値　　：
'
' =========================================================
Private Sub clearFilterCondition(Optional ByVal isNotClearComboFilter As Boolean = False)

    tglFilterA.value = False
    tglFilterB.value = False
    tglFilterC.value = False
    tglFilterD.value = False
    tglFilterE.value = False
    tglFilterF.value = False
    tglFilterG.value = False
    tglFilterH.value = False
    tglFilterI.value = False
    tglFilterJ.value = False
    tglFilterK.value = False
    tglFilterL.value = False
    tglFilterM.value = False
    tglFilterN.value = False
    tglFilterO.value = False
    tglFilterP.value = False
    tglFilterQ.value = False
    tglFilterR.value = False
    tglFilterS.value = False
    tglFilterT.value = False
    tglFilterU.value = False
    tglFilterV.value = False
    tglFilterW.value = False
    tglFilterX.value = False
    tglFilterY.value = False
    tglFilterZ.value = False
    tglFilterOther.value = False
    
    If isNotClearComboFilter = False Then
        cboFilter.text = ""
    End If
    
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

' =========================================================
' ▽テーブルシートリストをフィルタする処理
'
' 概要　　　：テーブルシートリストをフィルタする処理
' 引数　　　：filterKeyword         フィルタキーワード
' 戻り値　　：
'
' =========================================================
Private Sub filterTableSheetList(ByVal filterKeyword As String)

    If filterKeyword = "" Then
        ' フィルタ文字がない場合は、全ての情報を表示する
        tableSheetList.addAll tableSheetWithoutFilterList, "sheetNameOrSheetTableName", "TableComment"
        Exit Sub
    End If

    Dim filterTableSheetList As ValCollection
    Set filterTableSheetList = VBUtil.filterWildcard(tableSheetWithoutFilterList, "table.tableName", filterKeyword)
    
    addTableSheetList filterTableSheetList, False

End Sub


' =========================================================
' ▽テーブルシートリストをフィルタする処理（正規表現版）
'
' 概要　　　：テーブルシートリストをフィルタする処理
' 引数　　　：filterKeyword         フィルタキーワード
' 戻り値　　：
'
' =========================================================
Private Sub filterTableSheetListForRegExp(ByVal filterKeyword As String)

    Dim filterTableSheetList As ValCollection
    Set filterTableSheetList = VBUtil.filterRegExp(tableSheetWithoutFilterList, "table.tableName", filterKeyword)
    
    addTableSheetList filterTableSheetList, False

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
Private Sub addTableSheetList(ByVal valTableSheetList As ValCollection, Optional ByVal isAppend As Boolean = True)
    
    tableSheetList.addAll valTableSheetList _
                       , "sheetNameOrSheetTableName", "tableComment" _
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
    
    tableSheetList.addItemByProp tableSheet, "sheetNameOrSheetTableName", "tableComment"
    
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
    
    tableSheetList.setItem index, rec, "sheetNameOrSheetTableName", "tableComment"
    
End Sub
