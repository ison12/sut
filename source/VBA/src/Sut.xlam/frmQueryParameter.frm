VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryParameter 
   Caption         =   "クエリパラメータの設定"
   ClientHeight    =   8355.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8250.001
   OleObjectBlob   =   "frmQueryParameter.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
' 作成者　：Hideki Isobe
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

' クエリパラメータ設定情報
Private WithEvents frmQueryParameterSettingVar As frmQueryParameterSetting
Attribute frmQueryParameterSettingVar.VB_VarHelpID = -1

' クエリパラメータリスト コントロール
Private queryParameterList As CntListBox

' クエリパラメータリストでの選択項目インデックス
Private queryParameterSelectedIndex As Long
' クエリパラメータリストでの選択項目オブジェクト
Private queryParameterSelectedItem As ValQueryParameter

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
    
    Main.storeFormPosition Me.name, Me

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
    Set queryParameter = New ValQueryParameter: queryParameter.name = ConstantsCommon.QUERY_PARAMETER_DEFAULT_NAME
    
    queryParameter.name = QUERY_PARAMETER_DEFAULT_NAME & " " & (cnt + 1)
    
    Dim list As New ValCollection
    list.setItem queryParameter
    
    queryParameterList.addItemByProp queryParameter, "name", "value"
    
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
    
    Load frmQueryParameterSetting
    Set frmQueryParameterSettingVar = frmQueryParameterSetting
    frmQueryParameterSetting.ShowExt vbModal, queryParameterSelectedItem
                            
    Set frmQueryParameterSettingVar = Nothing

End Sub

' =========================================================
' ▽クエリパラメータ設定フォームのOKボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmQueryParameterSettingVar_ok(ByVal ValQueryParameter As ValQueryParameter)

    Dim v As ValQueryParameter
    Set v = queryParameterList.getItem(queryParameterSelectedIndex)
    
    v.name = ValQueryParameter.name
    v.value = ValQueryParameter.value

    queryParameterList.setItem queryParameterSelectedIndex, v, "name", "value"
    
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
Private Sub frmQueryParameterSettingVar_Cancel()

    queryParameterList.control.SetFocus
End Sub

' =========================================================
' ▽メニュー設定フォームのリセットボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmMenuSettingVar_reset(appSettingShortcut As ValApplicationSettingShortcut _
                                  , ByRef Cancel As Boolean)

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
    
    WinAPI_Clipboard.SetClipboard getQueryParameterForClipboardFormat(selectedItem)
    
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
    
    For Each var In queryParameterList.collection.col
        data.append getQueryParameterForClipboardFormat(var)
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

    getQueryParameterForClipboardFormat = """" & replace(var.name, """", """""") & """" & vbTab & """" & replace(var.value, """", """""") & """" & vbNewLine

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
        
        queryParameterObjList.setItem queryParameterObj
    
    Next
    
    queryParameterList.addAll queryParameterObjList, "name", "value", True
    

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
' ▽クエリパラメータ情報を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeQueryParameter()

    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet
    
    Dim var As Variant
    Dim i As Long
    
    bookProp.removeAllValue ConstantsBookProperties.TABLE_QUERY_PARAMETER_DIALOG
    
    i = 0
    For Each var In queryParameterList.collection.col
    
        bookProp.setValue ConstantsBookProperties.TABLE_QUERY_PARAMETER_DIALOG, "name_" & i, var.name
        bookProp.setValue ConstantsBookProperties.TABLE_QUERY_PARAMETER_DIALOG, "value_" & i, var.value
    
        i = i + 1
    Next
    ' ----------------------------------------------


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

    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' 設定情報シートが存在する
        Set bookPropVal = bookProp.getValuesOfElementArray(ConstantsBookProperties.TABLE_QUERY_PARAMETER_DIALOG)
    Else
        Set bookPropVal = New ValCollection
    End If
    ' ----------------------------------------------

    Dim ValQueryParameterList As New ValQueryParameterList
    ValQueryParameterList.setListFromFlatRecords bookPropVal

    Set queryParameterList = New CntListBox: queryParameterList.init lstQueryParameterList
    
    queryParameterList.addAll ValQueryParameterList.list _
                       , "name" _
                       , "value"
    
End Sub



