VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryParameterSetting 
   Caption         =   "クエリパラメータの編集"
   ClientHeight    =   8445.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7095
   OleObjectBlob   =   "frmQueryParameterSetting.frx":0000
End
Attribute VB_Name = "frmQueryParameterSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' クエリパラメータの一件毎の編集（子画面）
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
' 引数　　　：queryParameter クエリパラメータ情報
'
' =========================================================
Public Event ok(ByVal queryParameter As ValQueryParameter)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' クエリパラメータ情報（フォーム表示時点での情報）
Private queryParameterParam As ValQueryParameter

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
' 　　　　　　var   アプリケーション設定情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal queryParameter As ValQueryParameter)

    Set queryParameterParam = queryParameter
    
    activate
    
    ' デフォルトフォーカスコントロールを設定する
    txtParameter.SetFocus
    
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

    lblErrorMessage.Caption = ""
    
    txtParameter.value = queryParameterParam.name
    txtValue.value = queryParameterParam.value
    
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
    
    ' キーコードを分解し、対応する変数に格納する
    Dim queryParameter As New ValQueryParameter
    queryParameter.name = txtParameter.value
    queryParameter.value = txtValue.value

    ' OKイベントを送信する
    RaiseEvent ok(queryParameter)
    
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
' ▽DB接続変更ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDBConnectedChange_Click()

    On Error GoTo err
    
    Main.disconnectDB
    
    ' DBコネクションを取得する
    Dim dbConn As Object
    Set dbConn = Main.getDBConnection
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽テストボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdTest_Click()

    On Error GoTo err
    
    Const MSG_TITLE As String = "SELECTテスト結果"
    
    ' DBコネクションを取得する
    Dim dbConn As Object
    Set dbConn = Main.getDBConnection

    ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
    Dim cursorWait As New ExcelCursorWait: cursorWait.init
    
    Dim resultSet   As Object
    Dim resultRecs  As Variant
    Dim resultVal   As Variant

    Set resultSet = ADOUtil.querySelect(dbConn, txtValue.text, 0)
    
    ' ------------------------------------------------
    ' 戻り値

    ' レコードセットがEOFではない場合
    If Not resultSet.EOF Then
        ' レコードセットから全レコードを取得する
        resultRecs = resultSet.getRows(1)
        resultVal = resultRecs(0, 0)
    Else
        ' 空を返す
        resultVal = Empty
    End If
    ' ------------------------------------------------
    
    ADOUtil.closeRecordSet resultSet

    ' 長時間の処理が終了したのでマウスカーソルを元に戻す
    cursorWait.destroy
    
    If isNull(resultVal) Then
        VBUtil.showMessageBoxForInformation "取得データ：" & "NULL", MSG_TITLE
    ElseIf VBUtil.arraySize(resultRecs) <= 0 Then
        VBUtil.showMessageBoxForInformation "取得データ：" & "NULL (取得レコードが0件)", MSG_TITLE
    Else
        VBUtil.showMessageBoxForInformation "取得データ：" & CStr(resultVal), MSG_TITLE
    End If
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽パラメータ名入力時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtParameter_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:


    ' 正規表現オブジェクトを生成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        ' 検索対象文字列
        .Pattern = "^" & "([a-z0-9_-]|[^\u0000-\u007F])+" & "$"
        ' 大文字小文字無視フラグ
        .IgnoreCase = True
        ' 文字列全体を繰り返し検索するフラグ
        .Global = False
    End With


    ' 未入力時
    If txtParameter.text = "" Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtParameter
    
    ' 不正な文字入力
    ElseIf reg.test(txtParameter.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NOT_ALPHA_NUM_MARK_FULL
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtParameter
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtParameter
    End If
    
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


