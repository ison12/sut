VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryParameter 
   Caption         =   "クエリパラメータ設定"
   ClientHeight    =   8595.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8055
   OleObjectBlob   =   "frmQueryParameter.frx":0000
End
Attribute VB_Name = "frmQueryParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' クエリパラメータ定義フォーム
'
' 作成者　：Ison
' 履歴　　：2019/12/04　新規作成
'
' 特記事項：
' *********************************************************

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
Public Event Cancel()

' クエリパラメータの新規作成最大数
Private Const QUERY_PARAMETER_NEW_CREATED_OVER_SIZE As String = "クエリパラメータは最大${count}まで登録可能です。"

' クエリパラメータ設定情報の一件毎の編集（子画面）
Private WithEvents frmQueryParameterSettingVar As frmQueryParameterSetting
Attribute frmQueryParameterSettingVar.VB_VarHelpID = -1

' クエリパラメータリスト コントロール
Private queryParameterList As CntListBox

' クエリパラメータリストでの選択項目インデックス
Private queryParameterSelectedIndex As Long
' クエリパラメータリストでの選択項目オブジェクト
Private queryParameterSelectedItem As ValQueryParameter

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

    restoreQueryParameter
    
    lblDescription.Caption = replace(replace(lblDescription.Caption, "$es", ConstantsTable.QUERY_PARAMETER_ENCLOSE_START), "$ee", ConstantsTable.QUERY_PARAMETER_ENCLOSE_END)
    
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
    Set frmQueryParameterSettingVar = Nothing

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
    storeQueryParameter
    
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
    RaiseEvent Cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽クエリパラメータリストのダブルクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstQueryParameterList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    editQueryParameter
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
    cnt = queryParameterList.collection.count
    
    ' ポップアップの数が最大登録数を超えているかチェックする
    If cnt >= ConstantsCommon.QUERY_PARAMETER_NEW_CREATED_MAX_SIZE Then
    
        ' メッセージを表示する
        Dim mess As String
        mess = replace(QUERY_PARAMETER_NEW_CREATED_OVER_SIZE, "${count}", ConstantsCommon.QUERY_PARAMETER_NEW_CREATED_MAX_SIZE)
        
        VBUtil.showMessageBoxForInformation mess _
                                          , ConstantsCommon.APPLICATION_NAME
        Exit Sub
    End If
    
    ' ポップアップメニューオブジェクトをリストに追加する
    Dim queryParameter As ValQueryParameter
    Set queryParameter = New ValQueryParameter
    
    queryParameter.name = ConstantsCommon.QUERY_PARAMETER_DEFAULT_NAME & "_" & (cnt + 1)
    
    Dim list As New ValCollection
    list.setItem queryParameter
    
    addQueryParameter queryParameter
    
    queryParameterList.setSelectedIndex cnt
    queryParameterList.control.SetFocus
    
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

    editQueryParameter
End Sub

Private Sub editQueryParameter()

    ' 現在選択されているインデックスを取得
    queryParameterSelectedIndex = queryParameterList.getSelectedIndex

    ' 未選択の場合
    If queryParameterSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' 現在選択されている項目を取得
    Set queryParameterSelectedItem = queryParameterList.getSelectedItem
    
    If VBUtil.unloadFormIfChangeActiveBook(frmQueryParameterSetting) Then Unload frmQueryParameterSetting
    Load frmQueryParameterSetting
    Set frmQueryParameterSettingVar = frmQueryParameterSetting
    frmQueryParameterSetting.ShowExt vbModal, queryParameterSelectedItem
                            
    Set frmQueryParameterSettingVar = Nothing

End Sub

' =========================================================
' ▽クエリパラメータ設定フォームのOKボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：queryParameter クエリパラメータ情報
' 戻り値　　：
'
' =========================================================
Private Sub frmQueryParameterSettingVar_ok(ByVal queryParameter As ValQueryParameter)

    Dim v As ValQueryParameter
    Set v = queryParameterList.getItem(queryParameterSelectedIndex)
    
    v.name = queryParameter.name
    v.value = queryParameter.value

    setQueryParameter queryParameterSelectedIndex, v
    
    queryParameterList.control.SetFocus

End Sub

' =========================================================
' ▽クエリパラメータ設定フォームのキャンセルボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmQueryParameterSettingVar_cancel()

    queryParameterList.control.SetFocus
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
    selectedIndex = queryParameterList.getSelectedIndex

    ' 未選択の場合
    If selectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    queryParameterList.removeItem selectedIndex
    queryParameterList.control.SetFocus

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
    selectedIndex = queryParameterList.getSelectedIndex
    
    ' 未選択の場合
    If selectedIndex = -1 Then
        ' 終了する
        Exit Sub
    End If

    If selectedIndex > 0 Then
    
        queryParameterList.swapItem _
                          selectedIndex _
                        , selectedIndex - 1 _
                        , vbObject _
                        , 2
                              
        queryParameterList.setSelectedIndex selectedIndex - 1
            
    End If
    
    queryParameterList.control.SetFocus
        
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
    selectedIndex = queryParameterList.getSelectedIndex
    
        ' 未選択の場合
    If selectedIndex = -1 Then
        ' 終了する
        Exit Sub
    End If

    If selectedIndex < queryParameterList.count - 1 Then
    
        queryParameterList.swapItem _
                          selectedIndex _
                        , selectedIndex + 1 _
                        , vbObject _
                        , 2
                              
        queryParameterList.setSelectedIndex selectedIndex + 1
            
    End If
    
    queryParameterList.control.SetFocus
        
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
Private Sub cmdQueryParameterCopy_Click()

    Dim selectedIndex As Long
    Dim selectedItem As ValQueryParameter
    
    ' 現在選択されているインデックスを取得
    selectedIndex = queryParameterList.getSelectedIndex

    ' 未選択の場合
    If selectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    Set selectedItem = queryParameterList.getSelectedItem
    
    WinAPI_Clipboard.SetClipboard selectedItem.tabbedInfoHeader & vbNewLine & getQueryParameterForClipboardFormat(selectedItem)
    
End Sub

' =========================================================
' ▽全パラメータコピー時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdAllQueryParameterCopy_Click()

    Dim data As New StringBuilder
    Dim var As Variant
    
    Dim i As Long
    
    For Each var In queryParameterList.collection.col
        If i <= 0 Then
            data.append var.tabbedInfoHeader & vbNewLine
        End If
        data.append getQueryParameterForClipboardFormat(var)
        i = i + 1
    Next
    
    WinAPI_Clipboard.SetClipboard data.str

End Sub

' =========================================================
' ▽クエリパラメータのクリップボードフォーマット形式文字列取得
'
' 概要　　　：クエリパラメータのクリップボードフォーマット形式文字列を取得する。
' 引数　　　：var クエリパラメータ
' 戻り値　　：クエリパラメータのクリップボードフォーマット形式文字列取得
'
' =========================================================
Private Function getQueryParameterForClipboardFormat(ByVal var As ValQueryParameter) As String

    getQueryParameterForClipboardFormat = var.tabbedInfo & vbNewLine

End Function

' =========================================================
' ▽クエリパラメータをクリップボードから貼付け
'
' 概要　　　：クエリパラメータをクリップボードから貼付けする。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdQueryParameterPaste_Click()

    Dim var As Variant
    Dim queryParameterRawList As ValCollection
    
    Dim queryParameterObj As ValQueryParameter
    Dim queryParameterObjList As New ValCollection

    Dim clipBoard As String
    clipBoard = WinAPI_Clipboard.GetClipboard
    
    Dim CsvParser As New CsvParser: CsvParser.init vbTab
    Set queryParameterRawList = CsvParser.parse(clipBoard)
    
    For Each var In queryParameterRawList.col
    
        Set queryParameterObj = New ValQueryParameter
    
        If var.count >= 1 Then
            queryParameterObj.name = var.getItemByIndex(1, vbVariant)
        End If
    
        If var.count >= 2 Then
            queryParameterObj.value = var.getItemByIndex(2, vbVariant)
        End If
        
        If queryParameterObj.tabbedInfoHeader <> queryParameterObj.tabbedInfo Then
            queryParameterObjList.setItem queryParameterObj
        End If
    
    Next
    
    addQueryParameterList queryParameterObjList, isAppend:=True

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
    appProp.initWorksheet targetBook, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & Me.name & ".ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ▽クエリパラメータ情報を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeQueryParameter()

    On Error GoTo err
    
    Dim queryParameterList_ As New ValQueryParameterList
    queryParameterList_.init targetBook
    queryParameterList_.list = queryParameterList.collection
    queryParameterList_.writeForData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽クエリパラメータ情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreQueryParameter()

    On Error GoTo err
    
    Dim queryParameterList_ As New ValQueryParameterList
    queryParameterList_.init targetBook
    queryParameterList_.readForData

    Set queryParameterList = New CntListBox: queryParameterList.init lstQueryParameterList
    
    addQueryParameterList queryParameterList_.list
    
    ' 先頭を選択する
    queryParameterList.setSelectedIndex 0
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽クエリパラメータリストを追加
'
' 概要　　　：
' 引数　　　：valQueryParameterList クエリパラメータリスト
'     　　　  isAppend              追加有無フラグ
' 戻り値　　：
'
' =========================================================
Private Sub addQueryParameterList(ByVal ValQueryParameterList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    queryParameterList.addAll ValQueryParameterList _
                       , "name" _
                       , "value" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ▽クエリパラメータを追加
'
' 概要　　　：
' 引数　　　：queryParameter クエリパラメータ
' 戻り値　　：
'
' =========================================================
Private Sub addQueryParameter(ByVal queryParameter As ValQueryParameter)
    
    queryParameterList.addItemByProp queryParameter, "name", "value"
    
End Sub

' =========================================================
' ▽クエリパラメータを変更
'
' 概要　　　：
' 引数　　　：index インデックス
'     　　　  rec   クエリパラメータ
' 戻り値　　：
'
' =========================================================
Private Sub setQueryParameter(ByVal index As Long, ByVal rec As ValQueryParameter)
    
    queryParameterList.setItem index, rec, "name", "value"
    
End Sub
