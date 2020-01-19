VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBColumnFormat 
   Caption         =   "DBカラム書式設定"
   ClientHeight    =   8550.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15105
   OleObjectBlob   =   "frmDBColumnFormat.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDBColumnFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBカラム書式設定フォーム
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
Public Event ok(ByVal dbColumnFormatInfo As ValDbColumnFormatInfo)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' DBカラム書式編集フォーム
Private WithEvents frmDBColumnFormatSettingVar As frmDBColumnFormatSetting
Attribute frmDBColumnFormatSettingVar.VB_VarHelpID = -1

' DBカラム書式設定情報リスト（フォーム表示時点での情報）
Private dbColumnFormatInfoParam As ValDbColumnFormatInfo

' DBカラム書式設定情報リスト コントロール
Private dbColumnFormatList As CntListBox

' DBカラム書式設定情報リストでの選択項目インデックス
Private dbColumnFormatSelectedIndex As Long
' DBカラム書式設定情報リストでの選択項目オブジェクト
Private dbColumnFormatSelectedItem As ValDbColumnTypeColInfo

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal モーダルまたはモードレス表示指定
'     　　　：info  DBカラム書式情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal info As ValDbColumnFormatInfo)

    Set dbColumnFormatInfoParam = info

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

    Set dbColumnFormatList = New CntListBox: dbColumnFormatList.init lstDbColumnFormatList
    addDbColumnFormatList dbColumnFormatInfoParam.columnList
    
    ' 先頭を選択する
    dbColumnFormatList.setSelectedIndex 0
    
    lblDbName.Caption = DBUtil.getDbmsTypeName(dbColumnFormatInfoParam.dbName)

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
    Set frmDBColumnFormatSettingVar = Nothing
    
End Sub

' =========================================================
' ▽デフォルトボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDefault_Click()

    ' DBオブジェクト生成クラス
    Dim dbObjFactory As New DbObjectFactory
    ' カラム書式情報取得オブジェクト
    Dim dbColumnType As IDbColumnType
    ' カラム書式情報のデフォルト値情報を取得する
    Set dbColumnType = dbObjFactory.createColumnType(dbColumnFormatInfoParam.dbName)
    
    ' 反映する
    addDbColumnFormatList dbColumnType.getDefaultColumnFormat

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
    
    ' フォームを閉じる
    HideExt

    ' OKイベント送信時に設定する情報を生成する
    Dim var As New ValDbColumnFormatInfo
    var.dbName = dbColumnFormatInfoParam.dbName ' DB名
    Set var.columnList = dbColumnFormatList.collection ' リストボックス内の情報を取得して設定
    
    ' OKイベントを送信する
    RaiseEvent ok(var)
    
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
    cnt = dbColumnFormatList.collection.count
    
    ' ポップアップメニューオブジェクトをリストに追加する
    Dim dbColumnFormat As ValDbColumnTypeColInfo
    Set dbColumnFormat = New ValDbColumnTypeColInfo
    
    dbColumnFormat.columnName = ConstantsCommon.DB_COLUMN_FORMAT_DEFAULT_NAME & " " & (cnt + 1)
    
    addDbColumnFormat dbColumnFormat
    
    dbColumnFormatList.setSelectedIndex cnt
    dbColumnFormatList.control.SetFocus
    
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

    editDbColumnFormat
End Sub

Private Sub editDbColumnFormat()

    ' 現在選択されているインデックスを取得
    dbColumnFormatSelectedIndex = dbColumnFormatList.getSelectedIndex

    ' 未選択の場合
    If dbColumnFormatSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' 現在選択されている項目を取得
    Set dbColumnFormatSelectedItem = dbColumnFormatList.getSelectedItem
    
    Load frmDBColumnFormatSetting
    Set frmDBColumnFormatSettingVar = frmDBColumnFormatSetting
    frmDBColumnFormatSettingVar.ShowExt vbModal, dbColumnFormatSelectedItem
                            
    Set frmDBColumnFormatSettingVar = Nothing

End Sub

' =========================================================
' ▽DBカラム書式設定（子フォーム）のOKボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmDBColumnFormatSettingVar_ok(ByVal dbColumnTypeColInfo As ValDbColumnTypeColInfo)

    Dim v As ValDbColumnTypeColInfo
    Set v = dbColumnFormatList.getItem(dbColumnFormatSelectedIndex)
    
    v.columnName = dbColumnTypeColInfo.columnName
    v.formatUpdate = dbColumnTypeColInfo.formatUpdate
    v.formatSelect = dbColumnTypeColInfo.formatSelect

    setDbColumnFormat dbColumnFormatSelectedIndex, v
    
    dbColumnFormatList.control.SetFocus
End Sub

' =========================================================
' ▽DBカラム書式設定（子フォーム）のキャンセルボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmDBColumnFormatSettingVar_cancel()

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
    selectedIndex = dbColumnFormatList.getSelectedIndex

    ' 未選択の場合
    If selectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    dbColumnFormatList.removeItem selectedIndex
    dbColumnFormatList.control.SetFocus
    
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
    selectedIndex = dbColumnFormatList.getSelectedIndex
    
    ' 未選択の場合
    If selectedIndex = -1 Then
        ' 終了する
        Exit Sub
    End If

    If selectedIndex > 0 Then
    
        dbColumnFormatList.swapItem _
                          selectedIndex _
                        , selectedIndex - 1 _
                        , vbObject _
                        , 1
                              
        dbColumnFormatList.setSelectedIndex selectedIndex - 1
            
    End If
    
    dbColumnFormatList.control.SetFocus
        
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
    selectedIndex = dbColumnFormatList.getSelectedIndex
    
        ' 未選択の場合
    If selectedIndex = -1 Then
        ' 終了する
        Exit Sub
    End If

    If selectedIndex < dbColumnFormatList.count - 1 Then
    
        dbColumnFormatList.swapItem _
                          selectedIndex _
                        , selectedIndex + 1 _
                        , vbObject _
                        , 1
                              
        dbColumnFormatList.setSelectedIndex selectedIndex + 1
            
    End If
    
    dbColumnFormatList.control.SetFocus
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
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
' ▽DBカラム書式設定情報を追加
'
' 概要　　　：
' 引数　　　：list DBカラム書式設定情報リスト
' 戻り値　　：
'
' =========================================================
Private Sub addDbColumnFormatList(ByVal list As ValCollection)
    
    dbColumnFormatList.addAll list, "columnName", "formatUpdate", "formatSelect"
    
End Sub

' =========================================================
' ▽DBカラム書式設定情報を追加
'
' 概要　　　：
' 引数　　　：rec DBカラム書式設定情報
' 戻り値　　：
'
' =========================================================
Private Sub addDbColumnFormat(ByVal rec As ValDbColumnTypeColInfo)
    
    dbColumnFormatList.addItemByProp rec, "columnName", "formatUpdate", "formatSelect"
    
End Sub

' =========================================================
' ▽DBカラム書式設定情報を変更
'
' 概要　　　：
' 引数　　　：index インデックス
'     　　　  rec   DBカラム書式設定情報
' 戻り値　　：
'
' =========================================================
Private Sub setDbColumnFormat(ByVal index As Long, ByVal rec As ValDbColumnTypeColInfo)
    
    dbColumnFormatList.setItem index, rec, "columnName", "formatUpdate", "formatSelect"
    
End Sub
