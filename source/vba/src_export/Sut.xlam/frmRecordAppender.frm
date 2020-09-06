VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecordAppender 
   Caption         =   "行の追加・削除"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4620
   OleObjectBlob   =   "frmRecordAppender.frx":0000
End
Attribute VB_Name = "frmRecordAppender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' ワークシートの行数を変更するフォーム
'
' 作成者　：Ison
' 履歴　　：2009/11/15　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽OKボタン押下イベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event ok(ByVal recCount As Long)

' =========================================================
' ▽キャンセルボタン押下イベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' 処理対象ワークシート
Public sheet As Worksheet
' アプリケーション設定情報
Private applicationSetting As ValApplicationSetting

' 処理対象テーブルオブジェクト
Public tableSheet As ValTableWorksheet
' 既存の行数
Private recCountOrign As Long

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
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef sheet As Worksheet, ByRef aps As ValApplicationSetting)

    Set Me.sheet = sheet
    Set applicationSetting = aps
    
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
    
    ' テーブルシート生成オブジェクト
    Dim tableSheetSheetCreator As New ExeTableSheetCreator
    tableSheetSheetCreator.book = sheet.parent
    tableSheetSheetCreator.applicationSetting = applicationSetting
    
    ' 変更後の行数
    Dim recCount    As Long
    ' 既存の行数と変更後の行数の差
    Dim recCountDiff As Long
    
    ' テキストボックスから変更後の行数を取得する
    recCount = txtRecCount.value
    ' 既存の行数と変更後の行数の差分を取得する
    recCountDiff = recCount - recCountOrign
    
    Dim recStart As Long
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        recStart = ConstantsTable.U_RECORD_OFFSET_ROW
    Else
    
        recStart = ConstantsTable.R_RECORD_OFFSET_COL
    End If
    
    
    ' 行の削除
    If recCountDiff < 0 Then
    
        tableSheetSheetCreator.deleteCellOfRecord tableSheet, recStart + recCount
    
    ' 行の追加
    ElseIf recCountDiff > 0 Then
    
        tableSheetSheetCreator.insertEmptyCell tableSheet, recStart + recCountOrign, recCountDiff
            
    ' 何もしない
    Else
    
    
    End If
    
    ' OKイベントを送信する
    RaiseEvent ok(recCount)
    
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

' =========================================================
' ▽アクティブ時の処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()

    ' ----------------------------------------------
    ' テーブルシートから一度テーブル情報を読み込む
    Dim srctableSheet As ValTableWorksheet
    
    Dim tableSheetReader As ExeTableSheetReader
    
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.sheet = ActiveSheet
    
    ' テーブルシートかどうかを確認する。（失敗した場合、エラーが発行される）
    tableSheetReader.validTableSheet
    
    Set srctableSheet = tableSheetReader.readTableInfo
    ' ----------------------------------------------
    
    ' レコード件数
    Dim recCount As Long
    ' レコード件数の取得
    recCount = tableSheetReader.getRecordSize(srctableSheet)
    
    ' テキストボックスにレコード件数を設定する
    txtRecCount.value = recCount
    
    ' テキストボックスにフォーカスを与え全選択状態にする
    txtRecCount.SetFocus
    txtRecCount.SelStart = 0
    txtRecCount.SelLength = Len(txtRecCount)
    
    ' 処理対象テーブルオブジェクトを設定
    Set tableSheet = srctableSheet
    ' 既存のレコード件数
    recCountOrign = recCount
    
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
' ▽行数テキストボックスのチェック
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtRecCount_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' 空の場合、エラー
    If txtRecCount.text = "" Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        ' アラートを表示する
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        changeControlPropertyByValidTrue txtRecCount

    ' テキストボックスの値が整数かをチェックする
    ElseIf validInteger(txtRecCount.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        ' アラートを表示する
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
        
        changeControlPropertyByValidFalse txtRecCount
    
    ' 数値範囲チェック
    ElseIf CDec(txtRecCount.text) < 1 Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse txtRecCount
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecCount
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽テキストボックスチェック成功時のコントロール変更処理
'
' 概要　　　：
' 引数　　　：cnt コントロール
' 戻り値　　：
'
' =========================================================
Public Sub changeControlPropertyByValidTrue(ByRef cnt As MSForms.control)

    With cnt
        .BackColor = &H80000005
        .ForeColor = &H80000012
    
    End With

End Sub

' =========================================================
' ▽テキストボックスチェック失敗時のコントロール変更処理
'
' 概要　　　：
' 引数　　　：cnt コントロール
' 戻り値　　：
'
' =========================================================
Public Sub changeControlPropertyByValidFalse(ByRef cnt As MSForms.control)

    With cnt
        ' テキスト全体を選択する
        .SelStart = 0
        .SelLength = Len(.text)
        
        .BackColor = RGB(&HFF, &HFF, &HCC)
        .ForeColor = reverseRGB(&HFF, &HFF, &HCC)
        
    End With

End Sub

