VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnectFavorite 
   Caption         =   "DB接続の管理"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12675
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
' 引数　　　：
'
' =========================================================
Public Event ok()

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' B接続のお気に入り情報の新規作成最大数
Private Const DB_CONNECT_FAVORITE_NEW_CREATED_OVER_SIZE As String = "DB接続のお気に入り情報は最大${count}まで登録可能です。"

' DB接続フォーム
Private WithEvents frmDBConnectVar As frmDBConnect
Attribute frmDBConnectVar.VB_VarHelpID = -1

' DB接続のお気に入り情報リスト コントロール
Private dbConnectFavoriteList As CntListBox

' DB接続のお気に入り情報リストでの選択項目インデックス
Private dbConnectFavoriteSelectedIndex As Long
' DB接続のお気に入り情報リストでの選択項目オブジェクト
Private dbConnectFavoriteSelectedItem As ValDBConnectInfo

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
' 引数  　　：modal                    モーダルまたはモードレス表示指定
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
    
    ' 情報を記録する
    storeDBConnectFavorite
    
    ' フォームを閉じる
    HideExt
    
    ' OKイベントを送信する
    RaiseEvent ok
    
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
    editFavorite
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
    
    If VBUtil.unloadFormIfChangeActiveBook(frmDBConnect) Then Unload frmDBConnect
    Load frmDBConnect
    Set frmDBConnectVar = frmDBConnect
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

    Dim i As Long
    
    Dim var As ValCollection
    Dim dbConnectFavoriteRawList As ValCollection
    
    Dim dbConnectFavoriteObj As ValDBConnectInfo
    Dim dbConnectFavoriteObjList As New ValCollection

    Dim clipBoard As String
    clipBoard = WinAPI_Clipboard.GetClipboard
    
    Dim CsvParser As New CsvParser: CsvParser.init vbTab
    Set dbConnectFavoriteRawList = CsvParser.parse(clipBoard)
    
    For Each var In dbConnectFavoriteRawList.col
        
        Set dbConnectFavoriteObj = New ValDBConnectInfo
    
        ' 不足分を補完する（最終列が未入力の場合など、一列不足することがあるため）
        For i = 1 To 9 - var.count
            var.setItem ""
        Next
    
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
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & Me.name & ".ini"

    Set createApplicationProperties = appProp
    
End Function

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
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' 書き込みデータ
    Dim val As ValDBConnectInfo
    Dim values As New ValCollection
    
    Dim i As Long: i = 1
    For Each val In dbConnectFavoriteList.collection.col
        
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
' ▽DB接続のお気に入り情報情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoredbConnectFavorite()

    On Error GoTo err
            
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
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
    
    Set dbConnectFavoriteList = New CntListBox: dbConnectFavoriteList.init lstDbConnectFavoriteList
    addDbConnectFavoriteList connectInfoList
        
    ' 先頭を選択する
    dbConnectFavoriteList.setSelectedIndex 0

    Exit Sub
    
err:
    
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
    addDbConnectFavorite connectInfo
    
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
