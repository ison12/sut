VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBColumnFormatSetting 
   Caption         =   "DBカラム書式設定の編集"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   OleObjectBlob   =   "frmDBColumnFormatSetting.frx":0000
End
Attribute VB_Name = "frmDBColumnFormatSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBカラム書式設定の一件毎の編集（子画面）
'
' 作成者　：Ison
' 履歴　　：2019/12/08　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽決定した際に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：dbColumnTypeColInfo DBカラム書式設定
'
' =========================================================
Public Event ok(ByVal dbColumnTypeColInfo As ValDbColumnTypeColInfo)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' DBカラム書式設定情報（フォーム表示時点での情報）
Private dbColumnTypeColInfoParam As ValDbColumnTypeColInfo

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
' 引数　　　：modal               モーダルまたはモードレス表示指定
' 　　　　　　dbColumnTypeColInfo DBカラム書式設定情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal dbColumnTypeColInfo As ValDbColumnTypeColInfo)

    Set dbColumnTypeColInfoParam = dbColumnTypeColInfo
    
    activate
    
    ' デフォルトフォーカスコントロールを設定する
    txtColumnName.SetFocus
    
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

    txtColumnName.value = dbColumnTypeColInfoParam.columnName
    txtFormatUpdate.value = dbColumnTypeColInfoParam.formatUpdate
    txtFormatSelect.value = dbColumnTypeColInfoParam.formatSelect
    
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
    
    ' 入力情報を、対応する変数に格納する
    Dim dbColumnTypeColInfo As New ValDbColumnTypeColInfo
    dbColumnTypeColInfo.columnName = txtColumnName.value
    dbColumnTypeColInfo.formatUpdate = txtFormatUpdate.value
    dbColumnTypeColInfo.formatSelect = txtFormatSelect.value

    ' OKイベントを送信する
    RaiseEvent ok(dbColumnTypeColInfo)
    
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

Private Sub onPasteValue()

    Me.ActiveControl.text = "$value"

End Sub


