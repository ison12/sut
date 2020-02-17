VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectSelector 
   Caption         =   "接続情報の選択"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12675
   OleObjectBlob   =   "frmDBConnectSelector.frx":0000
End
Attribute VB_Name = "frmDBConnectSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' DB接続選択フォーム
'
' 作成者　：Ison
' 履歴　　：2020/01/14　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽決定した際に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：connectInfo DB接続情報
'
' =========================================================
Public Event ok(ByVal connectInfo As ValDBConnectInfo)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' フォームモード
Private formMode As DB_CONNECT_INFO_TYPE

' DB接続情報リスト コントロール
Private dbConnectList As CntListBox
' DB接続情報リスト（フィルタ条件適用なし）
Private dbConnectWithoutFilterList As ValCollection

' DB接続情報リストでの選択項目インデックス
Private dbConnectSelectedIndex As Long
' DB接続情報リストでの選択項目オブジェクト
Private dbConnectSelectedItem As ValDBConnectInfo

Private inFilterProcess As Boolean

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
' 引数　　　：modal モーダルまたはモードレス表示指定
'     　　　  fm    フォームモード
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal fm As DB_CONNECT_INFO_TYPE)

    ' フォームモード
    formMode = fm

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
' ▽フォームアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()

    If formMode = DB_CONNECT_INFO_TYPE.favorite Then
        lblFormModeName.Caption = "設定情報"
    Else
        lblFormModeName.Caption = "履歴情報"
    End If

    restoreDbConnectInfo formMode
    
    ' フィルタ条件を適用する
    cboFilter.text = ""
    applyFilterCondition

End Sub

' =========================================================
' ▽フォームディアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub deactivate()

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
        
        'filterConnectList currentFilterText ' 完全一致
        filterConnectList "*" & currentFilterText & "*" ' 中間一致
    
        inFilterProcess = False

    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
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
    
End Sub

' =========================================================
' ▽OKボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err

    ' 現在選択されているインデックスを取得
    dbConnectSelectedIndex = dbConnectList.getSelectedIndex

    ' 未選択の場合
    If dbConnectSelectedIndex = -1 Then
        err.Raise ERR_NUMBER_NOT_SELECTED_DB_CONNECT _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_DB_CONNECT _
                , err.HelpFile _
                , err.HelpContext
        ' 終了する
        Exit Sub
    End If
    
    ' フォームを閉じる
    HideExt

    ' 現在選択されている項目を取得
    Set dbConnectSelectedItem = dbConnectList.getSelectedItem
    
    ' OKイベントを送信する
    RaiseEvent ok(dbConnectSelectedItem)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽キャンセルボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdCancel_Click()

    On Error GoTo err
    
    ' フォームを閉じる
    HideExt
    
    ' キャンセルイベントを送信する
    RaiseEvent cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽DB接続リストのダブルクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstDbConnectList_DblClick(ByVal cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

' =========================================================
' ▽DB接続リスト キー押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstDbConnectList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        cmdOk_Click
    End If
    
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
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' 本処理では処理自体をキャンセルする
        cancel = True
        ' 以下のイベント経由で閉じる
        cmdCancel_Click
    End If
    
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
' ▽設定情報の生成
' =========================================================
Private Function createApplicationProperties(ByVal formMode As DB_CONNECT_INFO_TYPE) As ApplicationProperties
    
    ' フォーム名を取得する
    Dim subName As String
    
    If formMode = DB_CONNECT_INFO_TYPE.favorite Then
        subName = "frmDBConnectFavorite"
    Else
        subName = Me.name & "History"
    End If
    
    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & subName & ".ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ▽DB接続情報を保存する
'
' 概要　　　：
' 引数　　　：formMode フォームモード
' 戻り値　　：
'
' =========================================================
Private Sub storeDBConnectInfo(ByVal formMode As DB_CONNECT_INFO_TYPE)

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties(formMode)
    
    
    ' 書き込みデータ
    Dim val As New ValDBConnectInfo
    Dim values As New ValCollection
    
    Dim i As Long: i = 1
    For Each val In dbConnectList.collection.col
        
        values.setItem Array(i & "_" & "no", i)
        values.setItem Array(i & "_" & "name", val.name)
        values.setItem Array(i & "_" & "type", val.type_)
        values.setItem Array(i & "_" & "dsn", val.dsn)
        values.setItem Array(i & "_" & "host", val.host)
        values.setItem Array(i & "_" & "port", val.port)
        values.setItem Array(i & "_" & "db", val.db)
        values.setItem Array(i & "_" & "user", val.user)
        values.setItem Array(i & "_" & "password", val.password)
        values.setItem Array(i & "_" & "option", val.option_)
        
        i = i + 1
    Next
        
    ' データを書き込む
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
        
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽DB接続の履歴情報情報を読み込む
'
' 概要　　　：
' 引数　　　：formMode フォームモード
' 戻り値　　：
'
' =========================================================
Private Sub restoreDbConnectInfo(ByVal formMode As DB_CONNECT_INFO_TYPE)

    On Error GoTo err
        
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties(formMode)
    
    ' 接続情報
    Dim connectInfoList As ValCollection
    Set connectInfoList = New ValCollection
    Dim connectInfo As ValDBConnectInfo
    

    ' データを読み込む
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    Dim i As Long: i = 1
    Do While True
    
        val = values.getItem(i & "_" & "no", vbVariant)
        If Not IsArray(val) Then
            Exit Do
        End If
        
        Set connectInfo = New ValDBConnectInfo
                    
        val = values.getItem(i & "_" & "name", vbVariant): If IsArray(val) Then connectInfo.name = val(2)
        val = values.getItem(i & "_" & "type", vbVariant): If IsArray(val) Then connectInfo.type_ = val(2)
        val = values.getItem(i & "_" & "dsn", vbVariant): If IsArray(val) Then connectInfo.dsn = val(2)
        val = values.getItem(i & "_" & "host", vbVariant): If IsArray(val) Then connectInfo.host = val(2)
        val = values.getItem(i & "_" & "port", vbVariant): If IsArray(val) Then connectInfo.port = val(2)
        val = values.getItem(i & "_" & "db", vbVariant): If IsArray(val) Then connectInfo.db = val(2)
        val = values.getItem(i & "_" & "user", vbVariant): If IsArray(val) Then connectInfo.user = val(2)
        val = values.getItem(i & "_" & "password", vbVariant): If IsArray(val) Then connectInfo.password = val(2)
        val = values.getItem(i & "_" & "option", vbVariant): If IsArray(val) Then connectInfo.option_ = val(2)
        
        connectInfoList.setItem connectInfo
    
        i = i + 1
    Loop
    
    Set dbConnectList = New CntListBox: dbConnectList.init lstDbConnectList
    addDbConnectList connectInfoList
    Set dbConnectWithoutFilterList = dbConnectList.collection.copy
        
    ' 先頭を選択する
    dbConnectList.setSelectedIndex 0

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽DB接続の履歴情報情報を追加
'
' 概要　　　：
' 引数　　　：connectInfo DB接続情報
'     　      formMode    フォームモード
' 戻り値　　：
'
' =========================================================
Public Function registDbConnectInfo(ByVal connectInfo As ValDBConnectInfo, ByVal formMode As DB_CONNECT_INFO_TYPE)

    On Error GoTo err
    
    ' -------------------------------------------------------
    ' DB接続履歴を再ロードして最新にする
    ' -------------------------------------------------------
    restoreDbConnectInfo formMode
    
    ' -------------------------------------------------------
    ' 重複を取り除いた履歴情報を生成する
    ' -------------------------------------------------------
    Dim dbConnectDistinctList As New ValCollection
    Dim dbConnect As ValDBConnectInfo
    
    For Each dbConnect In dbConnectList.collection.col
        
        If dbConnect.displayName <> connectInfo.displayName Then
            ' 降順で表示するので、追加する要素は先頭に追加していく
            dbConnectDistinctList.setItem dbConnect
        End If
        
    Next
    
    ' -------------------------------------------------------
    ' 履歴情報を先頭に追加する
    ' -------------------------------------------------------
    dbConnectDistinctList.setItemByIndexBefore connectInfo, 1
    
    ' -------------------------------------------------------
    ' DB接続履歴に重複を取り除いたリストで入れ替える
    ' -------------------------------------------------------
    dbConnectList.removeAll
    addDbConnectList dbConnectDistinctList
    
    ' -------------------------------------------------------
    ' DB接続履歴を保存する
    ' -------------------------------------------------------
    storeDBConnectInfo formMode

    Exit Function
    
err:

    Main.ShowErrorMessage

End Function

' =========================================================
' ▽DB接続情報を追加
'
' 概要　　　：
' 引数　　　：connectInfoList DB接続情報リスト
' 戻り値　　：
'
' =========================================================
Private Sub addDbConnectList(ByVal connectInfoList As ValCollection)
    
    dbConnectList.addAll connectInfoList, "displayName"
    
End Sub

' =========================================================
' ▽DB接続情報を追加
'
' 概要　　　：
' 引数　　　：connectInfo DB接続情報
' 戻り値　　：
'
' =========================================================
Private Sub addDbConnect(ByVal connectInfo As ValDBConnectInfo)
    
    dbConnectList.addItemByProp connectInfo, "displayName"
    
End Sub

' =========================================================
' ▽テーブルシートリストをフィルタする処理
'
' 概要　　　：テーブルシートリストをフィルタする処理
' 引数　　　：filterKeyword         フィルタキーワード
' 戻り値　　：
'
' =========================================================
Private Sub filterConnectList(ByVal filterKeyword As String)

    Dim filterConnectList As ValCollection
    Set filterConnectList = VBUtil.filterWildcard(dbConnectWithoutFilterList, "displayName", filterKeyword)
    
    addDbConnectList filterConnectList

End Sub


