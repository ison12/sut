VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectFavorite 
   Caption         =   "DB接続の管理"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   OleObjectBlob   =   "frmDBConnectFavorite.frx":0000
End
Attribute VB_Name = "frmDBConnectFavorite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB接続お気に入りフォーム
'
' 作成者　：Ison
' 履歴　　：2020/01/14　新規作成
'
' 特記事項：
' *********************************************************

Implements IDbConnectListener

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

' B接続のお気に入り情報の新規作成最大数
Private Const DB_CONNECT_FAVORITE_NEW_CREATED_OVER_SIZE As String = "DB接続のお気に入り情報は最大${count}まで登録可能です。"

' DB接続フォーム
Private WithEvents frmDBConnectVar As frmDBConnect
Attribute frmDBConnectVar.VB_VarHelpID = -1

' DB接続のお気に入り情報リスト コントロール
Private dbConnectFavoriteList As CntListBox
'DB接続のお気に入り情報リスト（フィルタ条件適用なし）
Private dbConnectFavoriteWithoutFilterList As ValCollection

' DB接続のお気に入り情報リストでの選択項目インデックス
Private dbConnectFavoriteSelectedIndex As Long
' DB接続のお気に入り情報リストでの選択項目オブジェクト
Private dbConnectFavoriteSelectedItem As ValDBConnectInfo

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

    restoredbConnectFavorite
    
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

    ' Nothingを設定することでイベントを受信しないようにする
    Set frmDBConnectVar = Nothing
    
    ' フィルタを解除する
    cboFilter.text = ""
    
    ' 情報を記録する
    storeDBConnectFavorite

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
        
        If currentFilterText <> "" Then
            changeEnabledListManipulationControl False
        Else
            changeEnabledListManipulationControl True
        End If
        
        filterConnectList "*" & currentFilterText & "*"
    
        inFilterProcess = False

    End If
    
    Exit Sub
    
err:
    
    inFilterProcess = False
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

' =========================================================
' ▽リスト操作関連のコントロール類のEnabledフラグを制御する処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub changeEnabledListManipulationControl(ByVal enabled As Boolean)

    cmdAdd.enabled = enabled
    cmdDelete.enabled = enabled
    cmdUp.enabled = enabled
    cmdDown.enabled = enabled
    cmdDbConnectFavoritePaste.enabled = enabled
    
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
    dbConnectFavoriteSelectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' 未選択の場合
    If dbConnectFavoriteSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If
    
    ' フォームを閉じる
    HideExt

    ' 現在選択されている項目を取得
    Set dbConnectFavoriteSelectedItem = dbConnectFavoriteList.getSelectedItem
    
    ' OKイベントを送信する
    RaiseEvent ok(dbConnectFavoriteSelectedItem)
    
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
' ▽DB接続お気に入りリストのダブルクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstDbConnectFavoriteList_DblClick(ByVal cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

' =========================================================
' ▽DB接続お気に入りリスト キー押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstDbConnectFavoriteList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        cmdOk_Click
    End If
    
End Sub

' =========================================================
' ▽新規ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdAdd_Click()
    
    ' リストボックスのサイズ
    Dim cnt As Long
    ' リストボックスのサイズを取得する
    cnt = dbConnectFavoriteList.collection.count
    
    ' ポップアップの数が最大登録数を超えているかチェックする
    If cnt >= ConstantsCommon.DB_CONNECT_FAVORITE_NEW_CREATED_MAX_SIZE Then
    
        ' メッセージを表示する
        Dim mess As String
        mess = replace(DB_CONNECT_FAVORITE_NEW_CREATED_OVER_SIZE, "${count}", ConstantsCommon.DB_CONNECT_FAVORITE_NEW_CREATED_MAX_SIZE)
        
        VBUtil.showMessageBoxForInformation mess _
                                          , ConstantsCommon.APPLICATION_NAME
        Exit Sub
    End If
    
    ' ポップアップメニューオブジェクトをリストに追加する
    Dim dbConnectFavorite As ValDBConnectInfo
    Set dbConnectFavorite = New ValDBConnectInfo
    
    dbConnectFavorite.name = ConstantsCommon.DB_CONNECT_FAVORITE_DEFAULT_NAME & " " & (cnt + 1)
    
    Dim list As New ValCollection
    list.setItem dbConnectFavorite
    
    addDbConnectFavorite dbConnectFavorite
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
    
    dbConnectFavoriteList.setSelectedIndex cnt
    dbConnectFavoriteList.control.SetFocus
    
End Sub

' =========================================================
' ▽編集ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdEdit_Click()

    editFavorite
End Sub

' =========================================================
' ▽名称の編集ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdEditName_Click()

    editFavoriteName
End Sub

Private Sub editFavorite()

    ' 現在選択されているインデックスを取得
    dbConnectFavoriteSelectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' 未選択の場合
    If dbConnectFavoriteSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' 現在選択されている項目を取得
    Set dbConnectFavoriteSelectedItem = dbConnectFavoriteList.getSelectedItem
    
    Set frmDBConnectVar = New frmDBConnect
    frmDBConnectVar.ShowExt vbModal, dbConnectFavoriteSelectedItem, Me
                            
    Set frmDBConnectVar = Nothing

End Sub

Private Sub editFavoriteName()

    ' 現在選択されているインデックスを取得
    dbConnectFavoriteSelectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' 未選択の場合
    If dbConnectFavoriteSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' 現在選択されている項目を取得
    Set dbConnectFavoriteSelectedItem = dbConnectFavoriteList.getSelectedItem
    
    ' DbConnectInfo.Nameプロパティの入力を行うプロンプトを表示する
    Dim inputedName As String
    inputedName = InputBox("DB接続情報の名前を編集します。名前を入力してください。", "DB接続の名称編集", dbConnectFavoriteSelectedItem.name)
    If StrPtr(inputedName) = 0 Then
        ' キャンセルボタンが押下された場合
        Exit Sub
    End If
    
    dbConnectFavoriteSelectedItem.name = inputedName
    
    setDbConnectFavorite dbConnectFavoriteSelectedIndex, dbConnectFavoriteSelectedItem
    dbConnectFavoriteList.control.SetFocus
    
End Sub

' =========================================================
' ▽DB接続のお気に入り情報設定フォームのOKボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub IDbConnectListener_connect(connectInfo As ValDBConnectInfo)

    Dim v As ValDBConnectInfo
    Set v = dbConnectFavoriteList.getItem(dbConnectFavoriteSelectedIndex)
    
    v.dsn = connectInfo.dsn
    v.type_ = connectInfo.type_
    v.host = connectInfo.host
    v.port = connectInfo.port
    v.db = connectInfo.db
    v.user = connectInfo.user
    v.password = connectInfo.password
    v.option_ = connectInfo.option_

    setDbConnectFavorite dbConnectFavoriteSelectedIndex, v
    
    dbConnectFavoriteList.control.SetFocus

End Sub

' =========================================================
' ▽DB接続のお気に入り情報設定フォームのキャンセルボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub IDbConnectListener_cancel()

    dbConnectFavoriteList.control.SetFocus
End Sub

' =========================================================
' ▽削除ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDelete_Click()

    Dim selectedIndex As Long
    
    ' 現在選択されているインデックスを取得
    selectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' 未選択の場合
    If selectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    dbConnectFavoriteList.removeItem selectedIndex
    dbConnectFavoriteList.control.SetFocus
    
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy

End Sub

' =========================================================
' ▽上へボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdUp_Click()

    On Error GoTo err
    
    ' 選択済みインデックス
    Dim selectedIndex As Long
    
    ' 現在リストで選択されているインデックスを取得する
    selectedIndex = dbConnectFavoriteList.getSelectedIndex
    
    ' 未選択の場合
    If selectedIndex = -1 Then
        ' 終了する
        Exit Sub
    End If

    If selectedIndex > 0 Then
    
        dbConnectFavoriteList.swapItem _
                          selectedIndex _
                        , selectedIndex - 1 _
                        , vbObject _
                        , 1
                              
        dbConnectFavoriteList.setSelectedIndex selectedIndex - 1
            
    End If
    
    dbConnectFavoriteList.control.SetFocus
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽下へボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDown_Click()

    On Error GoTo err
    
    ' 選択済みインデックス
    Dim selectedIndex As Long
    
    ' 現在リストで選択されているインデックスを取得する
    selectedIndex = dbConnectFavoriteList.getSelectedIndex
    
        ' 未選択の場合
    If selectedIndex = -1 Then
        ' 終了する
        Exit Sub
    End If

    If selectedIndex < dbConnectFavoriteList.count - 1 Then
    
        dbConnectFavoriteList.swapItem _
                          selectedIndex _
                        , selectedIndex + 1 _
                        , vbObject _
                        , 1
                              
        dbConnectFavoriteList.setSelectedIndex selectedIndex + 1
            
    End If
    
    dbConnectFavoriteList.control.SetFocus
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽パラメータコピー時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDbConnectFavoriteCopy_Click()

    Dim selectedIndex As Long
    Dim selectedItem As ValDBConnectInfo
    
    ' 現在選択されているインデックスを取得
    selectedIndex = dbConnectFavoriteList.getSelectedIndex

    ' 未選択の場合
    If selectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    Set selectedItem = dbConnectFavoriteList.getSelectedItem
    
    WinAPI_Clipboard.SetClipboard _
        selectedItem.tabbedInfoHeader & vbNewLine & _
        getDbConnectFavoriteForClipboardFormat(selectedItem)
    
End Sub

' =========================================================
' ▽全パラメータコピー時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdAllDbConnectFavoriteCopy_Click()

    Dim data As New StringBuilder
    Dim var As Variant
    
    Dim i As Long
    
    For Each var In dbConnectFavoriteList.collection.col
        If i <= 0 Then
            data.append var.tabbedInfoHeader & vbNewLine
        End If
        data.append getDbConnectFavoriteForClipboardFormat(var)
        i = i + 1
    Next
    
    WinAPI_Clipboard.SetClipboard data.str

End Sub

' =========================================================
' ▽DB接続のお気に入り情報のクリップボードフォーマット形式文字列取得
'
' 概要　　　：DB接続のお気に入り情報のクリップボードフォーマット形式文字列を取得する。
' 引数　　　：var DB接続のお気に入り情報
' 戻り値　　：DB接続のお気に入り情報のクリップボードフォーマット形式文字列取得
'
' =========================================================
Private Function getDbConnectFavoriteForClipboardFormat(ByVal var As ValDBConnectInfo) As String

    getDbConnectFavoriteForClipboardFormat = var.tabbedInfo & vbNewLine

End Function

' =========================================================
' ▽DB接続のお気に入り情報をクリップボードから貼付け
'
' 概要　　　：DB接続のお気に入り情報をクリップボードから貼付けする。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmddbConnectFavoritePaste_Click()

    Dim var As Variant
    Dim dbConnectFavoriteRawList As ValCollection
    
    Dim dbConnectFavoriteObj As ValDBConnectInfo
    Dim dbConnectFavoriteObjList As New ValCollection

    Dim clipBoard As String
    clipBoard = WinAPI_Clipboard.GetClipboard
    
    Dim CsvParser As New CsvParser: CsvParser.init vbTab
    Set dbConnectFavoriteRawList = CsvParser.parse(clipBoard)
    
    For Each var In dbConnectFavoriteRawList.col
        
        Set dbConnectFavoriteObj = New ValDBConnectInfo
    
        If var.count >= 9 Then
            dbConnectFavoriteObj.name = var.getItemByIndex(1, vbVariant)
            dbConnectFavoriteObj.type_ = var.getItemByIndex(2, vbVariant)
            dbConnectFavoriteObj.dsn = var.getItemByIndex(3, vbVariant)
            dbConnectFavoriteObj.host = var.getItemByIndex(4, vbVariant)
            dbConnectFavoriteObj.port = var.getItemByIndex(5, vbVariant)
            dbConnectFavoriteObj.db = var.getItemByIndex(6, vbVariant)
            dbConnectFavoriteObj.user = var.getItemByIndex(7, vbVariant)
            dbConnectFavoriteObj.password = var.getItemByIndex(8, vbVariant)
            dbConnectFavoriteObj.option_ = var.getItemByIndex(9, vbVariant)
            
            If dbConnectFavoriteObj.tabbedInfo <> dbConnectFavoriteObj.tabbedInfoHeader Then
                dbConnectFavoriteObjList.setItem dbConnectFavoriteObj
            End If
            
        End If
    
    Next
    
    addDbConnectFavoriteList dbConnectFavoriteObjList, isAppend:=True
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy

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
' ▽フォームディアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Deactivate()

End Sub

' =========================================================
' ▽フォーム閉じる時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)

    deactivate

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
Private Sub storeDBConnectFavorite()

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
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_FAVORITE) _
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
     For Each dbConnectInfo In dbConnectFavoriteList.collection.col
        
        ' レジストリ操作クラスを初期化する
        Set registry = New RegistryManipulator
        registry.init RegKeyConstants.HKEY_CURRENT_USER _
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_FAVORITE & "\" & i) _
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
' ▽DB接続のお気に入り情報情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoredbConnectFavorite()

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
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_FAVORITE) _
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
                    , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT_FAVORITE & "\" & key) _
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
    
    cboFilter.text = ""
    Set dbConnectFavoriteList = New CntListBox: dbConnectFavoriteList.init lstDbConnectFavoriteList
    addDbConnectFavoriteList connectInfoList
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
    
    ' 先頭を選択する
    dbConnectFavoriteList.setSelectedIndex 0
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽DB接続のお気に入り情報を追加
'
' 概要　　　：
' 引数　　　：connectInfo DB接続情報
' 戻り値　　：
'
' =========================================================
Public Function registDbConnectInfo(ByVal connectInfo As ValDBConnectInfo)

    On Error GoTo err
    
    ' -------------------------------------------------------
    ' DB接続お気に入り情報を再ロードして最新にする
    ' -------------------------------------------------------
    restoredbConnectFavorite
    
    ' -------------------------------------------------------
    ' DB接続お気に入り情報の末尾に情報を追加する
    ' -------------------------------------------------------
    cboFilter.text = ""
    addDbConnectFavorite connectInfo
    Set dbConnectFavoriteWithoutFilterList = dbConnectFavoriteList.collection.copy
    
    ' -------------------------------------------------------
    ' DB接続お気に入り情報を保存する
    ' -------------------------------------------------------
    storeDBConnectFavorite

    Exit Function
    
err:

    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽DB接続のお気に入り情報を追加
'
' 概要　　　：
' 引数　　　：connectInfoList DB接続情報リスト
'     　　　  isAppend        追加有無フラグ
' 戻り値　　：
'
' =========================================================
Private Sub addDbConnectFavoriteList(ByVal connectInfoList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    dbConnectFavoriteList.addAll connectInfoList, "displayName", isAppend:=isAppend
    
End Sub

' =========================================================
' ▽DB接続のお気に入り情報を追加
'
' 概要　　　：
' 引数　　　：connectInfo DB接続情報
' 戻り値　　：
'
' =========================================================
Private Sub addDbConnectFavorite(ByVal connectInfo As ValDBConnectInfo)
    
    dbConnectFavoriteList.addItemByProp connectInfo, "displayName"
    
End Sub

' =========================================================
' ▽DBカラム書式設定情報を変更
'
' 概要　　　：
' 引数　　　：index インデックス
'     　　　  rec   DB接続情報
' 戻り値　　：
'
' =========================================================
Private Sub setDbConnectFavorite(ByVal index As Long, ByVal rec As ValDBConnectInfo)
    
    dbConnectFavoriteList.setItem index, rec, "displayName"
    
End Sub

' =========================================================
' 接続情報リストをフィルタする処理
'
' 概要　　　：接続情報リストをフィルタする処理
' 引数　　　：filterKeyword         フィルタキーワード
' 戻り値　　：
'
' =========================================================
Private Sub filterConnectList(ByVal filterKeyword As String)

    Dim filterConnectList As ValCollection
    Set filterConnectList = VBUtil.filterWildcard(dbConnectFavoriteWithoutFilterList, "displayName", filterKeyword)
    
    addDbConnectFavoriteList filterConnectList

End Sub
