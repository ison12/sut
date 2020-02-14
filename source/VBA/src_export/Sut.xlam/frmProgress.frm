VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "処理状況"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6690
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
   Tag             =   "168"
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' プログレスフォーム
'
' 作成者　：Ison
' 履歴　　：2020/01/21　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' セカンダリプログレスの表示有無
Private enableSecProgressParam As Boolean

' 対象ブック
Private targetBook As Workbook
' 対象ブックを取得する
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ▽タイトルプロパティ
' =========================================================
Public Property Get title() As String
    title = lblTitle.Caption
End Property

Public Property Let title(ByVal vNewValue As String)
    lblTitle.Caption = vNewValue
End Property

' =========================================================
' ▽プライマリ件数
' =========================================================
Public Property Get priCount() As Long
    priCount = CLng(lblPriCount.Caption)
End Property

Public Property Let priCount(ByVal vNewValue As Long)
    lblPriCount.Caption = CStr(vNewValue)
    updatePrimaryProgressBar
End Property

' =========================================================
' ▽プライマリ合計件数
' =========================================================
Public Property Get priCountOfAll() As Long
    priCountOfAll = CLng(lblPriCountOfAll.Caption)
End Property

Public Property Let priCountOfAll(ByVal vNewValue As Long)
    lblPriCountOfAll.Caption = CStr(vNewValue)
    updatePrimaryProgressBar
End Property

' =========================================================
' ▽プライマリメッセージ
' =========================================================
Public Property Get priMessage() As String
    priMessage = lblPriMessage.Caption
End Property

Public Property Let priMessage(ByVal vNewValue As String)
    lblPriMessage.Caption = vNewValue
End Property

' =========================================================
' ▽セカンダリ件数
' =========================================================
Public Property Get secCount() As Long
    secCount = CLng(lblSecCount.Caption)
End Property

Public Property Let secCount(ByVal vNewValue As Long)
    lblSecCount.Caption = CStr(vNewValue)
    updateSecondaryProgressBar
End Property

' =========================================================
' ▽セカンダリ合計件数
' =========================================================
Public Property Get secCountOfAll() As Long
    secCountOfAll = CLng(lblSecCountOfAll.Caption)
End Property

Public Property Let secCountOfAll(ByVal vNewValue As Long)
    lblSecCountOfAll.Caption = CStr(vNewValue)
    updateSecondaryProgressBar
End Property

' =========================================================
' ▽セカンダリメッセージ
' =========================================================
Public Property Get secMessage() As String
    secMessage = lblSecMessage.Caption
End Property

Public Property Let secMessage(ByVal vNewValue As String)
    lblSecMessage.Caption = vNewValue
End Property

' =========================================================
' ▽プライマリ情報の初期化
'
' 概要　　　：
' 引数　　　：all     合計件数
' 　　　　　　message メッセージ
' 戻り値　　：
'
' =========================================================
Public Function initPri(ByVal all As Long, ByVal message As String)

    priCount = 0
    priCountOfAll = all
    lblPriMessage.Caption = message
    
End Function

' =========================================================
' ▽セカンダリ情報の初期化
'
' 概要　　　：
' 引数　　　：all     合計件数
' 　　　　　　message メッセージ
' 戻り値　　：
'
' =========================================================
Public Function initSec(ByVal all As Long, ByVal message As String)
    
    secCount = 0
    secCountOfAll = all
    lblSecMessage.Caption = message

End Function

' =========================================================
' ▽プライマリ情報の件数更新（現在値に+1カウントする）
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Function inclimentPri()
    
    priCount = priCount + 1

End Function

' =========================================================
' ▽セカンダリ情報の件数更新（現在値に+1カウントする）
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Function inclimentSec()
    
    secCount = secCount + 1

End Function

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal               モーダルまたはモードレス表示指定
' 　　　　　　enableSecProgress   セカンダリプログレスの表示有無
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByVal enableSecProgress As Boolean)

    ' パラメータの設定
    enableSecProgressParam = enableSecProgress
    
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
' ▽キャンセル確認メッセージの確認ダイアログの表示
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：ダイアログで押下されたボタン
'
' =========================================================
Private Function showCancelConfDialog() As Long

    showCancelConfDialog = VBUtil.showMessageBoxForYesNo("キャンセルしてもよろしいですか？", ConstantsCommon.APPLICATION_NAME)

End Function

' =========================================================
' ▽フォームアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()

    ' 調整位置
    Dim frmAdjustValue    As Double: frmAdjustValue = 50        ' セカンダリプログレスを表示しない場合の高さの調整

    If enableSecProgressParam = True Then
        ' セカンダリプログレスを表示する場合
        lblSecMessage.visible = True
        lblSecCount.visible = True
        lblSecCountSeparator.visible = True
        lblSecCountOfAll.visible = True
        lblSecProgressBg.visible = True
        lblSecProgressFg.visible = True
        
        ' フォームの位置調整
        frmProgress.Height = CLng(frmProgress.Tag)
    Else
        ' セカンダリプログレスを表示する場合しない場合
        lblSecMessage.visible = False
        lblSecCount.visible = False
        lblSecCountSeparator.visible = False
        lblSecCountOfAll.visible = False
        lblSecProgressBg.visible = False
        lblSecProgressFg.visible = False
        
        ' フォームの位置調整
        frmProgress.Height = CLng(frmProgress.Tag) - frmAdjustValue
    End If
    
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
' ▽プログレスバーの更新
'
' 概要　　　：
' 引数　　　：cntCurrent 現在値を表示するコントロール
'     　　　  cntAll     最大値を表示するコントロール
'     　　　  valCurrent 現在値
'     　　　  valAll     最大値
' 戻り値　　：
'
' =========================================================
Private Sub updateProgressBar(ByRef cntCurrent As MSForms.label _
                            , ByRef cntAll As MSForms.label _
                            , ByVal valCurrent As Long _
                            , ByVal valAll As Long)
                         
    If valAll <= 0 Then
        ' 0除算しないようにチェックする
        cntCurrent.Width = 0
        Exit Sub
    End If
                         
    cntCurrent.Width = CDbl(cntAll.Width) * (CDbl(valCurrent) / CDbl(valAll))
    
End Sub

' =========================================================
' ▽プライマリプログレスバーの更新
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub updatePrimaryProgressBar()

    updateProgressBar lblPriProgressFg, lblPriProgressBg, CLng(lblPriCount.Caption), CLng(lblPriCountOfAll.Caption)
    
End Sub

' =========================================================
' ▽セカンダリプログレスバーの更新
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub updateSecondaryProgressBar()

    updateProgressBar lblSecProgressFg, lblSecProgressBg, CLng(lblSecCount.Caption), CLng(lblSecCountOfAll.Caption)
    
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
    
        If showCancelConfDialog = 6 Then
        
            ' キャンセルイベントを送信する
            RaiseEvent Cancel
            
        End If
        
        Cancel = True
        
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

