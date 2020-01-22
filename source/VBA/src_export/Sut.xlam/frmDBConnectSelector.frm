VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectSelector 
   Caption         =   "接続情報の選択"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12630
   OleObjectBlob   =   "frmDBConnectSelector.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

' レジストリキー
Private Const REG_SUB_KEY_DB_CONNECT_FAVORITE As String = "db_favorite"
' レジストリキー
Private Const REG_SUB_KEY_DB_CONNECT_HISTORY  As String = "db_history"

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
        
        filterConnectList "*" & currentFilterText & "*"
    
        inFilterProcess = False

    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
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
' ▽DB接続情報を保存する
'
' 概要　　　：
' 引数　　　：formMode フォームモード
' 戻り値　　：
'
' =========================================================
Private Sub storeDBConnectInfo(ByVal formMode As DB_CONNECT_INFO_TYPE)

    On Error GoTo err
    
    ' レジストリのサブキーを決定する
    Dim regSubKey As String
    
    If formMode = DB_CONNECT_INFO_TYPE.favorite Then
        regSubKey = REG_SUB_KEY_DB_CONNECT_FAVORITE
    Else
        regSubKey = REG_SUB_KEY_DB_CONNECT_HISTORY
    End If
    
    Dim i, j As Long
    ' レジストリ操作クラス
    Dim registry As RegistryManipulator
    
    ' -------------------------------------------------------
    ' 全ての情報をレジストリから一旦削除する
    ' -------------------------------------------------------
    ' レジストリ操作クラスを初期化する
    Set registry = New RegistryManipulator
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, regSubKey) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    Dim key     As Variant
    Dim keyList As ValCollection
    
    Set keyList = registry.getKeyList
    
    For Each key In keyList.col
        registry.delete key
    Next
    
    ' -------------------------------------------------------
    ' 全ての情報をレジストリに保存する
    ' -------------------------------------------------------
    Dim dbConnectInfo As ValDBConnectInfo
    Dim dbConnectArray(0 To 9 _
                             , 0 To 1) As Variant
    
    i = 0
     For Each dbConnectInfo In dbConnectList.collection.col
        
        ' レジストリ操作クラスを初期化する
        Set registry = New RegistryManipulator
        registry.init RegKeyConstants.HKEY_CURRENT_USER _
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, regSubKey & "\" & i) _
                    , RegAccessConstants.KEY_ALL_ACCESS _
                    , True

        j = 0
        dbConnectArray(j, 0) = "no"
        dbConnectArray(j, 1) = j: j = j + 1
        dbConnectArray(j, 0) = "name"
        dbConnectArray(j, 1) = dbConnectInfo.name: j = j + 1
        dbConnectArray(j, 0) = "type"
        dbConnectArray(j, 1) = dbConnectInfo.type_: j = j + 1
        dbConnectArray(j, 0) = "dsn"
        dbConnectArray(j, 1) = dbConnectInfo.dsn: j = j + 1
        dbConnectArray(j, 0) = "host"
        dbConnectArray(j, 1) = dbConnectInfo.host: j = j + 1
        dbConnectArray(j, 0) = "port"
        dbConnectArray(j, 1) = dbConnectInfo.port: j = j + 1
        dbConnectArray(j, 0) = "db"
        dbConnectArray(j, 1) = dbConnectInfo.db: j = j + 1
        dbConnectArray(j, 0) = "user"
        dbConnectArray(j, 1) = dbConnectInfo.user: j = j + 1
        dbConnectArray(j, 0) = "password"
        dbConnectArray(j, 1) = dbConnectInfo.password: j = j + 1
        dbConnectArray(j, 0) = "option"
        dbConnectArray(j, 1) = dbConnectInfo.option_: j = j + 1
        
        ' レジストリに情報を設定する
        registry.setValues dbConnectArray
    
        Set registry = Nothing

        
        i = i + 1
    Next

        
    Exit Sub
    
err:
    
    Set registry = Nothing

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
    
    ' レジストリのサブキーを決定する
    Dim regSubKey As String
    
    If formMode = DB_CONNECT_INFO_TYPE.favorite Then
        regSubKey = REG_SUB_KEY_DB_CONNECT_FAVORITE
    Else
        regSubKey = REG_SUB_KEY_DB_CONNECT_HISTORY
    End If
    
    ' お気に入りの接続情報
    Dim connectInfoList As ValCollection
    Set connectInfoList = New ValCollection
    Dim connectInfo As ValDBConnectInfo
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
                
    Dim key     As Variant
    Dim keyList As ValCollection

    ' -------------------------------------------------------
    ' 全ての情報をレジストリから取得する（インデックス番号リストの取得）
    ' -------------------------------------------------------
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, regSubKey) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Set keyList = registry.getKeyList

    Set registry = Nothing
    
    ' -------------------------------------------------------
    ' 全ての詳細情報をレジストリから取得する
    ' -------------------------------------------------------
    Dim valueNameList As ValCollection
    Dim valueList As ValCollection
    
    For Each key In keyList.col
    
        ' レジストリ操作クラスを初期化する
        Set registry = New RegistryManipulator
        registry.init RegKeyConstants.HKEY_CURRENT_USER _
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, regSubKey & "\" & key) _
                    , RegAccessConstants.KEY_ALL_ACCESS _
                    , True
                    
        registry.getValueList valueNameList, valueList

        Set connectInfo = New ValDBConnectInfo
        connectInfo.name = valueList.getItem("name", vbVariant)
        connectInfo.type_ = valueList.getItem("type", vbVariant)
        connectInfo.dsn = valueList.getItem("dsn", vbVariant)
        connectInfo.host = valueList.getItem("host", vbVariant)
        connectInfo.port = valueList.getItem("port", vbVariant)
        connectInfo.db = valueList.getItem("db", vbVariant)
        connectInfo.user = valueList.getItem("user", vbVariant)
        connectInfo.password = valueList.getItem("password", vbVariant)
        connectInfo.option_ = valueList.getItem("option", vbVariant)
        
        connectInfoList.setItem connectInfo
                    
        Set registry = Nothing
    Next
    
    Set dbConnectList = New CntListBox: dbConnectList.init lstDbConnectList
    addDbConnectList connectInfoList
    Set dbConnectWithoutFilterList = dbConnectList.collection.copy
        
    ' 先頭を選択する
    dbConnectList.setSelectedIndex 0

    Exit Sub
    
err:
    
    Set registry = Nothing
    
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


