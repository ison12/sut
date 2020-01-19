VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectHistory 
   Caption         =   "DB接続の履歴"
   ClientHeight    =   9120.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   OleObjectBlob   =   "frmDBConnectHistory.frx":0000
End
Attribute VB_Name = "frmDBConnectHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB接続履歴フォーム
'
' 作成者　：Hideki Isobe
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
Private Const REG_SUB_KEY_DB_CONNECT_HISTORY As String = "db_history"

' DB接続の履歴情報リスト コントロール
Private dbConnectHistoryList As CntListBox
'DB接続のお気に入り情報リスト（フィルタ条件適用なし）
Private dbConnectHistoryWithoutFilterList As ValCollection

' DB接続の履歴情報リストでの選択項目インデックス
Private dbConnectHistorySelectedIndex As Long
' DB接続の履歴情報リストでの選択項目オブジェクト
Private dbConnectHistorySelectedItem As ValDBConnectInfo

Private inFilterProcess As Boolean


' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal モーダルまたはモードレス表示指定
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants)

    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show vbModal
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

    restoreDbConnectHistory
    
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
    dbConnectHistorySelectedIndex = dbConnectHistoryList.getSelectedIndex

    ' 未選択の場合
    If dbConnectHistorySelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If
    
    ' フォームを閉じる
    HideExt

    ' 現在選択されている項目を取得
    Set dbConnectHistorySelectedItem = dbConnectHistoryList.getSelectedItem
    
    ' OKイベントを送信する
    RaiseEvent ok(dbConnectHistorySelectedItem)
    
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
' ▽DB接続履歴リストのダブルクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstDbConnectHistoryList_DblClick(ByVal cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

' =========================================================
' ▽DB接続履歴リスト キー押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstDbConnectHistoryList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
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
' ▽DB接続のお気に入り情報情報を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeDBConnectHistory()

    On Error GoTo err
    
    Dim i, j As Long
    ' レジストリ操作クラス
    Dim registry As RegistryManipulator
    
    ' -------------------------------------------------------
    ' 全ての情報をレジストリから一旦削除する
    ' -------------------------------------------------------
    ' レジストリ操作クラスを初期化する
    Set registry = New RegistryManipulator
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_HISTORY) _
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
    Dim dbConnectFavoriteArray(0 To 9 _
                             , 0 To 1) As Variant
    
    i = 0
     For Each dbConnectInfo In dbConnectHistoryList.collection.col
        
        ' レジストリ操作クラスを初期化する
        Set registry = New RegistryManipulator
        registry.init RegKeyConstants.HKEY_CURRENT_USER _
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_HISTORY & "\" & i) _
                    , RegAccessConstants.KEY_ALL_ACCESS _
                    , True

        j = 0
        dbConnectFavoriteArray(j, 0) = "no"
        dbConnectFavoriteArray(j, 1) = j: j = j + 1
        dbConnectFavoriteArray(j, 0) = "name"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.name: j = j + 1
        dbConnectFavoriteArray(j, 0) = "type"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.type_: j = j + 1
        dbConnectFavoriteArray(j, 0) = "dsn"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.dsn: j = j + 1
        dbConnectFavoriteArray(j, 0) = "host"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.host: j = j + 1
        dbConnectFavoriteArray(j, 0) = "port"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.port: j = j + 1
        dbConnectFavoriteArray(j, 0) = "db"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.db: j = j + 1
        dbConnectFavoriteArray(j, 0) = "user"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.user: j = j + 1
        dbConnectFavoriteArray(j, 0) = "password"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.password: j = j + 1
        dbConnectFavoriteArray(j, 0) = "option"
        dbConnectFavoriteArray(j, 1) = dbConnectInfo.option_: j = j + 1
        
        ' レジストリに情報を設定する
        registry.setValues dbConnectFavoriteArray
    
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
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreDbConnectHistory()

    On Error GoTo err
    
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
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_HISTORY) _
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
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_HISTORY & "\" & key) _
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
    
    Set dbConnectHistoryList = New CntListBox: dbConnectHistoryList.init lstDbConnectHistoryList
    addDbConnectHistoryList connectInfoList
    Set dbConnectHistoryWithoutFilterList = dbConnectHistoryList.collection.copy
        
    ' 先頭を選択する
    dbConnectHistoryList.setSelectedIndex 0

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
' 戻り値　　：
'
' =========================================================
Public Function registDbConnectInfo(ByVal connectInfo As ValDBConnectInfo)

    On Error GoTo err
    
    ' -------------------------------------------------------
    ' DB接続履歴を再ロードして最新にする
    ' -------------------------------------------------------
    restoreDbConnectHistory
    
    ' -------------------------------------------------------
    ' 重複を取り除いた履歴情報を生成する
    ' -------------------------------------------------------
    Dim dbConnectHistoryDistinctList As New ValCollection
    Dim dbConnectHistory As ValDBConnectInfo
    
    For Each dbConnectHistory In dbConnectHistoryList.collection.col
        
        If dbConnectHistory.displayName <> connectInfo.displayName Then
            ' 降順で表示するので、追加する要素は先頭に追加していく
            dbConnectHistoryDistinctList.setItem dbConnectHistory
        End If
        
    Next
    
    ' -------------------------------------------------------
    ' 履歴情報を先頭に追加する
    ' -------------------------------------------------------
    dbConnectHistoryDistinctList.setItemByIndexBefore connectInfo, 1
    
    ' -------------------------------------------------------
    ' DB接続履歴に重複を取り除いたリストで入れ替える
    ' -------------------------------------------------------
    dbConnectHistoryList.removeAll
    addDbConnectHistoryList dbConnectHistoryDistinctList
    
    ' -------------------------------------------------------
    ' DB接続履歴を保存する
    ' -------------------------------------------------------
    storeDBConnectHistory

    Exit Function
    
err:

    Main.ShowErrorMessage

End Function

' =========================================================
' ▽DB接続の履歴情報を追加
'
' 概要　　　：
' 引数　　　：connectInfoList DB接続情報リスト
' 戻り値　　：
'
' =========================================================
Private Sub addDbConnectHistoryList(ByVal connectInfoList As ValCollection)
    
    dbConnectHistoryList.addAll connectInfoList, "displayName"
    
End Sub

' =========================================================
' ▽DB接続の履歴情報を追加
'
' 概要　　　：
' 引数　　　：connectInfo DB接続情報
' 戻り値　　：
'
' =========================================================
Private Sub addDbConnectHistory(ByVal connectInfo As ValDBConnectInfo)
    
    dbConnectHistoryList.addItemByProp connectInfo, "displayName"
    
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
    Set filterConnectList = VBUtil.filterWildcard(dbConnectHistoryWithoutFilterList, "displayName", filterKeyword)
    
    addDbConnectHistoryList filterConnectList

End Sub
