VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBExplorer 
   Caption         =   "DBエクスプローラ"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7950
   OleObjectBlob   =   "frmDBExplorer.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDBExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBエクスプローラ
'
' 作成者　：Ison
' 履歴　　：2020/01/18　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽OKボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：tableList  テーブルリスト
'             recFormat  レコードフォーマット
' =========================================================
Public Event export(ByVal tableList As ValCollection _
                  , ByVal recFormat As REC_FORMAT)

' =========================================================
' ▽閉じるボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event closed()

Private Const REG_SUB_KEY_DB_EXPLORER_OPTION As String = "db_explorer"

' DBコネクションオブジェクト
Private dbConn As Object
' スキーマリスト
Private schemaInfoList  As CntListBox
' テーブルリスト
Private tableInfoList   As CntListBox
' テーブルリストのフィルタ条件なしのリスト
Private tableWithoutFilterList As ValCollection

Private inFilterProcess As Boolean

' =========================================================
' ▽DBコネクション設定
'
' 概要　　　：
' 引数　　　：vNewValue DBコネクション
' 戻り値　　：
'
' =========================================================
Public Property Let DbConnection(ByVal vNewValue As Variant)

    Set dbConn = vNewValue
    
    ' スキーマシートを読み込む
    readSchemaInfo
    ' テーブルシートを読み込む
    readTableInfo
    
End Property

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal  モーダルまたはモードレス表示指定
'             conn   DBコネクション
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef conn As Object)

    ' DBコネクションを設定する
    Set dbConn = conn
    
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
' ▽スキーマコンボボックス変更時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboSchema_Change()

    On Error GoTo err

    inFilterProcess = True
    
    clearFilterCondition False
    readTableInfo
    
    inFilterProcess = False
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
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
        
        filterTableInfoList "*" & currentFilterText & "*"
        
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
            filterTableInfoListForRegExp "[^a-zA-Z]*"
            
            clearFilterCondition
            tglFilterOther.value = True
        Else
            filterTableInfoListForRegExp ""
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
            filterTableInfoList keyword & "*"
            
            clearFilterCondition
            state.value = True
        Else
            filterTableInfoList ""
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
' ▽フィルタ条件の適用処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub applyFilterCondition()

    If cboFilter.text <> "" Then
        cboFilter_Change
        Exit Sub
    End If
    
    If tglFilterA.value = True Then
        tglFilterA_Click
    ElseIf tglFilterB.value = True Then
        tglFilterB_Click
    ElseIf tglFilterC.value = True Then
        tglFilterC_Click
    ElseIf tglFilterD.value = True Then
        tglFilterD_Click
    ElseIf tglFilterE.value = True Then
        tglFilterE_Click
    ElseIf tglFilterF.value = True Then
        tglFilterF_Click
    ElseIf tglFilterG.value = True Then
        tglFilterG_Click
    ElseIf tglFilterH.value = True Then
        tglFilterH_Click
    ElseIf tglFilterI.value = True Then
        tglFilterI_Click
    ElseIf tglFilterJ.value = True Then
        tglFilterJ_Click
    ElseIf tglFilterK.value = True Then
        tglFilterK_Click
    ElseIf tglFilterL.value = True Then
        tglFilterL_Click
    ElseIf tglFilterM.value = True Then
        tglFilterM_Click
    ElseIf tglFilterN.value = True Then
        tglFilterN_Click
    ElseIf tglFilterO.value = True Then
        tglFilterO_Click
    ElseIf tglFilterP.value = True Then
        tglFilterP_Click
    ElseIf tglFilterQ.value = True Then
        tglFilterQ_Click
    ElseIf tglFilterR.value = True Then
        tglFilterR_Click
    ElseIf tglFilterS.value = True Then
        tglFilterS_Click
    ElseIf tglFilterT.value = True Then
        tglFilterT_Click
    ElseIf tglFilterU.value = True Then
        tglFilterU_Click
    ElseIf tglFilterV.value = True Then
        tglFilterV_Click
    ElseIf tglFilterW.value = True Then
        tglFilterW_Click
    ElseIf tglFilterX.value = True Then
        tglFilterX_Click
    ElseIf tglFilterY.value = True Then
        tglFilterY_Click
    ElseIf tglFilterZ.value = True Then
        tglFilterZ_Click
    ElseIf tglFilterOther.value = True Then
        tglFilterOther_Click
    End If
    
End Sub

' =========================================================
' ▽全ての選択肢を選択済みにするボタンのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdSelectedAll_Click()

    tableInfoList.setSelectedAll True

End Sub

' =========================================================
' ▽全ての選択肢を選択解除にするボタンのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdUnselectedAll_Click()

    tableInfoList.setSelectedAll False

End Sub

' =========================================================
' ▽エクスポート（↓）ボタンのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdExportToUnder_Click()
    
    exportProcess recFormatToUnder
End Sub

' =========================================================
' ▽エクスポート（→）ボタンのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdExportToRight_Click()

    exportProcess recFormatToRight
End Sub

' =========================================================
' ▽エクスポート処理
'
' 概要　　　：
' 引数　　　：recFormat 行フォーマット
' 戻り値　　：
'
' =========================================================
Private Sub exportProcess(ByVal recFormat As REC_FORMAT)

    On Error GoTo err
    
    Dim exportTargets As ValCollection
    Set exportTargets = tableInfoList.selectedList
    
    If exportTargets.count <= 0 Then
        err.Raise ERR_NUMBER_NOT_SELECTED_TABLE _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_TABLE _
                , err.HelpFile _
                , err.HelpContext
        Exit Sub
    End If
    
    RaiseEvent export(exportTargets, recFormat)

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
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
' ▽初期化処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub initial()
    
    ' リスト系コントロールの初期化
    Set schemaInfoList = New CntListBox: schemaInfoList.init cboSchema
    Set tableInfoList = New CntListBox: tableInfoList.init lstTable

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

    ' DBエクスプローラオプションを読み込む
    restoreDBExplorerOption

    ' スキーマシートを読み込む
    readSchemaInfo
    ' テーブルシートを読み込む
    readTableInfo
    
    ' フィルタ条件を適用する
    applyFilterCondition
    
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
' ▽DBエクスプローラオプションを保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeDBExplorerOption()

    On Error GoTo err
    
    Dim j As Long
    
    Dim dbExplorerOption(0 To 29 _
                       , 0 To 1) As Variant
    
    dbExplorerOption(j, 0) = cboSchema.name
    dbExplorerOption(j, 1) = cboSchema.value: j = j + 1
    
    dbExplorerOption(j, 0) = cboFilter.name
    dbExplorerOption(j, 1) = cboFilter.value: j = j + 1

    dbExplorerOption(j, 0) = tglFilterA.name
    dbExplorerOption(j, 1) = tglFilterA.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterB.name
    dbExplorerOption(j, 1) = tglFilterB.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterC.name
    dbExplorerOption(j, 1) = tglFilterC.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterD.name
    dbExplorerOption(j, 1) = tglFilterD.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterE.name
    dbExplorerOption(j, 1) = tglFilterE.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterF.name
    dbExplorerOption(j, 1) = tglFilterF.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterG.name
    dbExplorerOption(j, 1) = tglFilterG.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterH.name
    dbExplorerOption(j, 1) = tglFilterH.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterI.name
    dbExplorerOption(j, 1) = tglFilterI.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterJ.name
    dbExplorerOption(j, 1) = tglFilterJ.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterK.name
    dbExplorerOption(j, 1) = tglFilterK.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterL.name
    dbExplorerOption(j, 1) = tglFilterL.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterM.name
    dbExplorerOption(j, 1) = tglFilterM.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterN.name
    dbExplorerOption(j, 1) = tglFilterN.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterO.name
    dbExplorerOption(j, 1) = tglFilterO.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterP.name
    dbExplorerOption(j, 1) = tglFilterP.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterQ.name
    dbExplorerOption(j, 1) = tglFilterQ.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterR.name
    dbExplorerOption(j, 1) = tglFilterR.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterS.name
    dbExplorerOption(j, 1) = tglFilterS.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterT.name
    dbExplorerOption(j, 1) = tglFilterT.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterU.name
    dbExplorerOption(j, 1) = tglFilterU.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterV.name
    dbExplorerOption(j, 1) = tglFilterV.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterW.name
    dbExplorerOption(j, 1) = tglFilterW.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterX.name
    dbExplorerOption(j, 1) = tglFilterX.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterY.name
    dbExplorerOption(j, 1) = tglFilterY.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterZ.name
    dbExplorerOption(j, 1) = tglFilterZ.value: j = j + 1
    dbExplorerOption(j, 0) = tglFilterOther.name
    dbExplorerOption(j, 1) = tglFilterOther.value: j = j + 1
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_EXPLORER_OPTION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' レジストリに情報を設定する
    registry.setValues dbExplorerOption
    
    Set registry = Nothing
        
    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet
    
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, cboSchema.name, cboSchema.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, cboFilter.name, cboFilter.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterA.name, tglFilterA.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterB.name, tglFilterB.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterC.name, tglFilterC.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterD.name, tglFilterD.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterE.name, tglFilterE.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterF.name, tglFilterF.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterG.name, tglFilterG.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterH.name, tglFilterH.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterI.name, tglFilterI.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterJ.name, tglFilterJ.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterK.name, tglFilterK.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterL.name, tglFilterL.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterM.name, tglFilterM.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterN.name, tglFilterN.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterO.name, tglFilterO.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterP.name, tglFilterP.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterQ.name, tglFilterQ.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterR.name, tglFilterR.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterS.name, tglFilterS.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterT.name, tglFilterT.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterU.name, tglFilterU.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterV.name, tglFilterV.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterW.name, tglFilterW.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterX.name, tglFilterX.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterY.name, tglFilterY.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterZ.name, tglFilterZ.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG, tglFilterOther.name, tglFilterOther.value

    ' ----------------------------------------------

    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽DBエクスプローラオプションを読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreDBExplorerOption()

    On Error GoTo err
        
    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' 設定情報シートが存在する
        
        Set bookPropVal = bookProp.getValues(ConstantsBookProperties.TABLE_DB_EXPLORER_DIALOG)
        If bookPropVal.count > 0 Then
            ' 設定情報が存在するので、フォームに反映する
            
            inFilterProcess = True
            
            cboSchema.value = bookPropVal.getItem(cboSchema.name, vbString)
            cboFilter.value = bookPropVal.getItem(cboFilter.name, vbString)
            tglFilterA.value = bookPropVal.getItem(tglFilterA.name, vbString)
            tglFilterB.value = bookPropVal.getItem(tglFilterB.name, vbString)
            tglFilterC.value = bookPropVal.getItem(tglFilterC.name, vbString)
            tglFilterD.value = bookPropVal.getItem(tglFilterD.name, vbString)
            tglFilterE.value = bookPropVal.getItem(tglFilterE.name, vbString)
            tglFilterF.value = bookPropVal.getItem(tglFilterF.name, vbString)
            tglFilterG.value = bookPropVal.getItem(tglFilterG.name, vbString)
            tglFilterH.value = bookPropVal.getItem(tglFilterH.name, vbString)
            tglFilterI.value = bookPropVal.getItem(tglFilterI.name, vbString)
            tglFilterJ.value = bookPropVal.getItem(tglFilterJ.name, vbString)
            tglFilterK.value = bookPropVal.getItem(tglFilterK.name, vbString)
            tglFilterL.value = bookPropVal.getItem(tglFilterL.name, vbString)
            tglFilterM.value = bookPropVal.getItem(tglFilterM.name, vbString)
            tglFilterN.value = bookPropVal.getItem(tglFilterN.name, vbString)
            tglFilterO.value = bookPropVal.getItem(tglFilterO.name, vbString)
            tglFilterP.value = bookPropVal.getItem(tglFilterP.name, vbString)
            tglFilterQ.value = bookPropVal.getItem(tglFilterQ.name, vbString)
            tglFilterR.value = bookPropVal.getItem(tglFilterR.name, vbString)
            tglFilterS.value = bookPropVal.getItem(tglFilterS.name, vbString)
            tglFilterT.value = bookPropVal.getItem(tglFilterT.name, vbString)
            tglFilterU.value = bookPropVal.getItem(tglFilterU.name, vbString)
            tglFilterV.value = bookPropVal.getItem(tglFilterV.name, vbString)
            tglFilterW.value = bookPropVal.getItem(tglFilterW.name, vbString)
            tglFilterX.value = bookPropVal.getItem(tglFilterX.name, vbString)
            tglFilterY.value = bookPropVal.getItem(tglFilterY.name, vbString)
            tglFilterZ.value = bookPropVal.getItem(tglFilterZ.name, vbString)
            tglFilterOther.value = bookPropVal.getItem(tglFilterOther.name, vbString)

            inFilterProcess = False
            
            applyFilterCondition
            
            Exit Sub
        End If
    End If
    ' ----------------------------------------------

    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_EXPLORER_OPTION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Dim retStr As String
            
    inFilterProcess = True
    
    registry.getValue cboSchema.name, retStr: cboSchema.value = retStr
    registry.getValue cboFilter.name, retStr: cboFilter.value = retStr
    registry.getValue tglFilterA.name, retStr: tglFilterA.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterB.name, retStr: tglFilterB.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterC.name, retStr: tglFilterC.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterD.name, retStr: tglFilterD.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterE.name, retStr: tglFilterE.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterF.name, retStr: tglFilterF.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterG.name, retStr: tglFilterG.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterH.name, retStr: tglFilterH.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterI.name, retStr: tglFilterI.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterJ.name, retStr: tglFilterJ.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterK.name, retStr: tglFilterK.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterL.name, retStr: tglFilterL.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterM.name, retStr: tglFilterM.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterN.name, retStr: tglFilterN.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterO.name, retStr: tglFilterO.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterP.name, retStr: tglFilterP.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterQ.name, retStr: tglFilterQ.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterR.name, retStr: tglFilterR.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterS.name, retStr: tglFilterS.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterT.name, retStr: tglFilterT.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterU.name, retStr: tglFilterU.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterV.name, retStr: tglFilterV.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterW.name, retStr: tglFilterW.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterX.name, retStr: tglFilterX.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterY.name, retStr: tglFilterY.value = VBUtil.convertBoolStrToBool(retStr)
    registry.getValue tglFilterZ.name, retStr: tglFilterZ.value = VBUtil.convertBoolStrToBool(retStr)

    inFilterProcess = False
    
    Set registry = Nothing
    
    Exit Sub
    
err:

    inFilterProcess = False
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽スキーマ情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub readSchemaInfo()

    On Error GoTo err
    
    Dim var As ValCollection
    
    If ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        ' 切断状態
        
        Set var = New ValCollection
        addSchemaInfoList var
        
    Else
        ' 接続状態
    
        ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
        Dim cursorWait As New ExcelCursorWait: cursorWait.init
    
        ' スキーマ定義を取得する
        Dim dbObjFactory As New DbObjectFactory
        
        Dim dbInfo As IDbMetaInfoGetter
        Set dbInfo = dbObjFactory.createMetaInfoGetterObject(dbConn)
           
        Set var = dbInfo.getSchemaList
        
        ' スキーマリストボックスにリストを追加する
        addSchemaInfoList var
        
        ' 長時間の処理が終了したのでマウスカーソルを元に戻す
        cursorWait.destroy
        
    End If

    Exit Sub
    
err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ▽テーブル情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub readTableInfo()

    On Error GoTo err

    Dim var  As ValCollection

    If ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        ' 切断状態
        
        Set var = New ValCollection
        addTableInfoList var
        
        Set tableWithoutFilterList = var.copy
        
    Else
        ' 接続状態

        ' 選択済みのスキーマ情報を取得
        If schemaInfoList.count > 0 Then
        
            ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
            Dim cursorWait As New ExcelCursorWait: cursorWait.init
        
            If schemaInfoList.getSelectedIndex = -1 Then
                ' 選択がない場合は、先頭を選択状態にする
                schemaInfoList.setSelectedIndex 0
            End If
            
            Dim selectedSchemaList As New ValCollection
            Dim selectedSchema As ValDbDefineSchema
            Set selectedSchema = schemaInfoList.getSelectedItem(vbObject)
            selectedSchemaList.setItem selectedSchema
            
            ' テーブル定義を取得する
            Dim dbObjFactory As New DbObjectFactory
            
            Dim dbInfo As IDbMetaInfoGetter
            Set dbInfo = dbObjFactory.createMetaInfoGetterObject(dbConn)
            
            Set var = dbInfo.getTableList(selectedSchemaList)
            
            ' テーブルリストボックスにリストを追加する
            addTableInfoList var
            
            Set tableWithoutFilterList = var.copy
            
            ' 長時間の処理が終了したのでマウスカーソルを元に戻す
            cursorWait.destroy
            
        Else
            ' スキーマが存在しない場合
            Set var = New ValCollection
            addTableInfoList var
        
            Set tableWithoutFilterList = var.copy
        End If
    End If

    Exit Sub
    
err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ▽テーブルリストをフィルタする処理
'
' 概要　　　：テーブルリストをフィルタする処理
' 引数　　　：filterKeyword         フィルタキーワード
' 戻り値　　：
'
' =========================================================
Private Sub filterTableInfoList(ByVal filterKeyword As String)

    Dim filterTableInfoList As ValCollection
    Set filterTableInfoList = VBUtil.filterWildcard(tableWithoutFilterList, "tableName", filterKeyword)
    
    addTableInfoList filterTableInfoList, False

End Sub

' =========================================================
' ▽テーブルリストをフィルタする処理（正規表現版）
'
' 概要　　　：テーブルリストをフィルタする処理
' 引数　　　：filterKeyword         フィルタキーワード
' 戻り値　　：
'
' =========================================================
Private Sub filterTableInfoListForRegExp(ByVal filterKeyword As String)

    Dim filterTableInfoList As ValCollection
    Set filterTableInfoList = VBUtil.filterRegExp(tableWithoutFilterList, "tableName", filterKeyword)
    
    addTableInfoList filterTableInfoList, False

End Sub

' =========================================================
' ▽スキーマリストを追加
'
' 概要　　　：
' 引数　　　：valSchemaInfoList スキーマリスト
'     　　　  isAppend          追加有無フラグ
' 戻り値　　：
'
' =========================================================
Private Sub addSchemaInfoList(ByVal valSchemaInfoList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    schemaInfoList.addAll valSchemaInfoList _
                       , "schemaName" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ▽テーブルリストを追加
'
' 概要　　　：
' 引数　　　：valtableInfoList テーブルリスト
'     　　　  isAppend     追加有無フラグ
' 戻り値　　：
'
' =========================================================
Private Sub addTableInfoList(ByVal valTableInfoList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    tableInfoList.addAll valTableInfoList _
                       , "tableName", "tableComment" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ▽テーブルを追加
'
' 概要　　　：
' 引数　　　：table テーブル
' 戻り値　　：
'
' =========================================================
Private Sub addTable(ByVal table As ValDbDefineTable)
    
    tableInfoList.addItemByProp table, "tableName", "tableComment"
    
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
Private Sub setTable(ByVal index As Long, ByVal rec As ValDbDefineTable)
    
    tableInfoList.setItem index, rec, "tableName", "tableComment"
    
End Sub
