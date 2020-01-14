VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' メニュー操作時の処理 - ヘルプ
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/05/18　新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' メンバ変数
' ________________________________________________________
' ■オプションフォーム（イベント受信用に定義）
Private WithEvents frmLicenceInfoVar As frmLicenceInfo
Attribute frmLicenceInfoVar.VB_VarHelpID = -1

' ライセンス認証有無フラグ
Private m_authenticatedLicence As Boolean
' ライセンス情報
Private m_licenceInfo As ValLicenceInfo

' ________________________________________________________
' イベント
' ________________________________________________________
' =========================================================
' ▽ライセンス情報表示フォーム OKボタン押下時のイベント
'
' 概要　　　：
'
' =========================================================
Private Sub frmLicenceInfoVar_ok()

End Sub

' =========================================================
' ▽ライセンス情報表示フォーム キャンセルボタン押下時のイベント
'
' 概要　　　：
'
' =========================================================
Private Sub frmLicenceInfoVar_cancel()

End Sub

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    
    ' イベント受信用のオプションフォームを破棄する
    Set frmLicenceInfoVar = Nothing
End Sub

' =========================================================
' ▽初期化
'
' 概要　　　：初期化処理を実行する。
' 引数　　　：authenticatedLicence ライセンス認証フラグ
' 　　　　　　licenceInfo ライセンス情報
'
' =========================================================
Public Sub init(ByVal authenticatedLicence As Boolean _
              , ByRef licenceInfo As ValLicenceInfo)

    ' イベント受信用のオプションフォームを初期化する
    Load frmLicenceInfo
    Set frmLicenceInfoVar = frmLicenceInfo
    
    ' ライセンス認証フラグを設定
    m_authenticatedLicence = authenticatedLicence
    ' ライセンス情報を設定
    Set m_licenceInfo = licenceInfo

End Sub

' =========================================================
' ▽ライセンス情報表示
'
' 概要　　　：
'
' =========================================================
Public Sub showLicenceInfo()

    On Error GoTo err
    
    ' ライセンス情報フォームを表示する
    frmLicenceInfoVar.ShowExt vbModal, m_authenticatedLicence, m_licenceInfo
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub
