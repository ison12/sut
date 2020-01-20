VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBQueryBatchTypeSetting 
   Caption         =   "クエリ一括実行のクエリ種類変更"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7065
   OleObjectBlob   =   "frmDBQueryBatchTypeSetting.frx":0000
End
Attribute VB_Name = "frmDBQueryBatchTypeSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBクエリバッチのクエリ種類の一件毎の編集（子画面）
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/08　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽決定した際に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：dbQueryBatchType DBクエリバッチ種類
'
' =========================================================
Public Event ok(ByVal dbQueryBatchType As DB_QUERY_BATCH_TYPE)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' シート名（フォーム表示時点での情報）
Private sheetNameParam As String
' DBクエリバッチ種類（フォーム表示時点での情報）
Private dbQueryBatchTypeParam As DB_QUERY_BATCH_TYPE
' DBクエリバッチ種類の選択肢リスト
Private dbQueryBatchTypeSelectList As ValCollection
' DBクエリバッチ種類コンボボックスリスト
Private dbQueryBatchTypeComboList As CntListBox

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal モーダルまたはモードレス表示指定
' 　　　　　　sheetName                     シート名
' 　　　　　　dbQueryBatchType              DBクエリバッチ種類の初期値
' 　　　　　　valDbQueryBatchTypeSelectList DBクエリバッチ種類の選択肢リスト
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                    , ByVal sheetName As String _
                    , ByVal dbQueryBatchType As DB_QUERY_BATCH_TYPE _
                    , ByVal valDbQueryBatchTypeSelectList As ValCollection)

    ' パラメータを設定
    sheetNameParam = sheetName
    dbQueryBatchTypeParam = dbQueryBatchType
    Set dbQueryBatchTypeSelectList = valDbQueryBatchTypeSelectList

    activate
    
    ' デフォルトフォーカスコントロールを設定する
    cboDbQueryBatchType.SetFocus
    
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

    ' シート名を設定する
    txtSheetName.text = sheetNameParam
    
    ' DBクエリバッチ種類のコンボボックスを初期化する
    Set dbQueryBatchTypeComboList = New CntListBox: dbQueryBatchTypeComboList.init cboDbQueryBatchType
    dbQueryBatchTypeComboList.addAll dbQueryBatchTypeSelectList, "dbQueryBatchTypeName"
    
    ' DBクエリバッチ種類コンボボックスのアクティブな選択項目を設定する
    Dim v As ValDbQueryBatchType
    Dim i As Long
    
    i = 0
    For Each v In dbQueryBatchTypeComboList.collection.col
        
        If v.dbQueryBatchType = dbQueryBatchTypeParam Then
            dbQueryBatchTypeComboList.setSelectedIndex i
            Exit For
        End If
        
        i = i + 1
    Next
    
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
' ▽OKボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err
    
    ' 未選択の場合
    If dbQueryBatchTypeComboList.getSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If
    
    ' 選択肢が何もなしの場合
    If dbQueryBatchTypeComboList.getSelectedItem.dbQueryBatchType = DB_QUERY_BATCH_TYPE.none Then
    
        ' 終了する
        Exit Sub
    End If
    
    ' フォームを閉じる
    HideExt

    ' OKイベントを送信する
    RaiseEvent ok(dbQueryBatchTypeComboList.getSelectedItem.dbQueryBatchType)
    
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



