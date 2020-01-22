VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetUpdate 
   Caption         =   "テーブルシート更新"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5940
   OleObjectBlob   =   "frmTableSheetUpdate.frx":0000
End
Attribute VB_Name = "frmTableSheetUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' テーブルシート更新フォーム
'
' 作成者　：Ison
' 履歴　　：2009/04/03　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽処理が完了した場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：recFormat 行フォーマット
'
' =========================================================
Public Event ok(ByVal recFormat As REC_FORMAT)

' =========================================================
' ▽処理がキャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

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
' ▽アクティブ時の処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()

    ' 行フォーマット
    Dim recFormat As REC_FORMAT
    
    ' アクティブなテーブルシートの行フォーマットを取得し
    ' オプションボタンに反映する
        
    ' テーブルシート読込オブジェクト
    Dim tableSheetReader As ExeTableSheetReader
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.sheet = ActiveSheet
        
    recFormat = tableSheetReader.getRowFormat
    
    If recFormat = REC_FORMAT.recFormatToUnder Then
    
        optRowFormatToUnder.value = True
    
    Else
    
        optRowFormatToRight.value = True
    End If
    
        
End Sub

' =========================================================
' ▽ノンアクティブ時の処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub deactivate()

End Sub

' =========================================================
' ▽OKボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdOk_Click()

    ' 行フォーマット定数を利用するために、ValTableを生成する
    Dim table As ValDbDefineTable
    ' 行フォーマット
    Dim recFormat As REC_FORMAT
    
    ' オプションボタンで選択されている値を
    ' Long型の行フォーマット定数に変換する。
    If optRowFormatToUnder.value = True Then
    
        ' 行フォーマットXのラジオボタンが選択されている場合
        recFormat = REC_FORMAT.recFormatToUnder
    
    Else
    
        ' 行フォーマットYのラジオボタンが選択されている場合
        recFormat = REC_FORMAT.recFormatToRight
        
    End If
    
    RaiseEvent ok(recFormat)
    HideExt
End Sub

' =========================================================
' ▽キャンセルボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdCancel_Click()

    RaiseEvent cancel
    HideExt
End Sub
