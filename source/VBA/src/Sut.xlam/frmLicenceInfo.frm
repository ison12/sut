VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLicenceInfo 
   Caption         =   "ライセンスについて"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6585
   OleObjectBlob   =   "frmLicenceInfo.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmLicenceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ライセンス情報を表示（登録）するフォーム
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/05/18　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽OKボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event ok()

' =========================================================
' ▽閉じるボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event closed()

' エラーメッセージ
Private Const ERR_MSG_AUTH_FAILED               As String = "認証に失敗しました。"
' インフォメッセージ
Private Const INFO_MSG_AUTH_SUCCESS             As String = "認証されました。"
' インフォメッセージ（試用期間日付）
Private Const INFO_MSG_PROBATION_DAY            As String = "後、${date}日使用できます。購入を希望される場合は、ライセンス登録を行ってください。詳しくはマニュアルを参照してください。"
' インフォメッセージ（試用期間オーバー）
Private Const INFO_MSG_OVER_PROBATION_DAY       As String = "試用期間が過ぎています。継続して利用する場合は、ライセンス登録をお願いします。詳しくはマニュアルを参照してください。"
' インフォメッセージ（認証完了）
Private Const INFO_MSG_AUTH_LICENCE_COMPLETED   As String = "認証が完了しています。"
' インフォメッセージ（ライセンス入力欄のメッセージ）
Private Const INFO_MSG_AUTH_LICENCE_AREA        As String = "発行したユーザIDとライセンスキーを入力して､認証ボタンを押下してください｡"
' インフォメッセージ（ライセンス入力欄のメッセージ2）
Private Const INFO_MSG_AUTH_LICENCE_AREA2       As String = "本ソフトウェアは次の方にライセンスされています。"

' ライセンス認証有無フラグ
Private m_authenticatedLicence As Boolean
' ライセンス認証情報
Private m_licenceInfo As ValLicenceInfo

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal  モーダルまたはモードレス表示指定
' 　　　　　　authenticatedLicence ライセンスが認証有無フラグ
' 　　　　　　licenceInfo ライセンス情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByVal authenticatedLicence As Boolean _
                 , ByRef licenceInfo As ValLicenceInfo)

    
    ' メンバを設定
    m_authenticatedLicence = authenticatedLicence
    Set m_licenceInfo = licenceInfo
    
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

    ' 最前面表示にする
    ExcelUtil.setUserFormTopMost Me

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
    
    ' コントロールのチェック
    If validAllControl = False Then
    
        Exit Sub
    End If
    
    ' ----------------------------------------------------
    ' ライセンス認証処理
    ' ----------------------------------------------------
    ' ライセンス情報（一時的に生成）
    Dim tmpLicenceInfo As New ValLicenceInfo
    ' ライセンス情報の設定
    tmpLicenceInfo.userId = txtUserId.value
    tmpLicenceInfo.password = txtPassword.value
    
    ' 認証オブジェクト
    Dim author As New ExeAuthenticateLicence
    ' 認証オブジェクトにライセンス情報を設定する
    author.init tmpLicenceInfo
    
    ' 認証を実施する
    If author.executeAuthor = False Then
    
        ' 失敗した場合
        lblErrorMessage.Caption = ERR_MSG_AUTH_FAILED
        Exit Sub
        
    End If
    ' ----------------------------------------------------
    
    ' ファイル出力オプションを書き込む
    storeLicenceInfo

    ' 成功メッセージを表示
    VBUtil.showMessageBoxForInformation INFO_MSG_AUTH_SUCCESS, ConstantsCommon.APPLICATION_NAME

    ' フォームを閉じる
    HideExt
    
    ' OKイベントを送信する
    RaiseEvent ok
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub


' =========================================================
' ▽閉じるボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdClose_Click()

    On Error GoTo err
    
    ' フォームを閉じる
    HideExt
    
    ' 閉じるイベントを送信する
    RaiseEvent closed

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

    ' 認証情報オブジェクトを破棄する
    Set m_licenceInfo = Nothing
    
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
    
    ' 認証済み
    If m_authenticatedLicence = True Then
    
        ' 認証済みメッセージを設定する
        lblMessage.Caption = INFO_MSG_AUTH_LICENCE_COMPLETED
        ' 認証済みメッセージをライセンス入力欄に設定する
        lblMessageAuthLicence.Caption = INFO_MSG_AUTH_LICENCE_AREA2
        
        ' 認証項目を開く
        openItemOfLicence
        ' 認証項目を認証済み用に変更する
        changeItemOfLicenceAuthenticatedLicence
        
    ' 未認証
    Else
    
        ' 未認証メッセージをライセンス入力欄に設定する
        lblMessageAuthLicence.Caption = INFO_MSG_AUTH_LICENCE_AREA
        
        ' 認証オブジェクト
        Dim author As New ExeAuthenticateLicence
        
        ' ソフトの使用開始日
        Dim fromDate As Date
        ' 使用開始日を取得する
        fromDate = author.getProbationDate
    
        ' 試用期間の範囲内
        If author.isRangeProbation(fromDate) = True Then
        
            lblMessage.Caption = replace(INFO_MSG_PROBATION_DAY _
                                            , "${date}" _
                                            , author.getRemainderProbationDay(fromDate))
        
        
        ' 試用期間の範囲外
        Else
        
            lblMessage.Caption = INFO_MSG_OVER_PROBATION_DAY
        End If
        
        ' 認証項目を未認証用に変更する
        changeItemOfLicenceNonAuthenticatedLicence
        
        ' ライセンス登録ボタンを擬似的に押下する
        tglRegistLicence_Click
    
    End If
    
    ' エラーメッセージを消去する
    lblErrorMessage.Caption = ""
    
    ' ライセンス情報を読み込む
    restoreLicenceInfo
    
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
' ▽ライセンス認証トグルボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub tglRegistLicence_Click()

    If tglRegistLicence.value = True Then
    
        openItemOfLicence
        
    ElseIf tglRegistLicence.value = False Then
    
        closeItemOfLicence
    End If
    
End Sub

Private Sub openItemOfLicence()

    ' 関連するコントロールを表示する
    fraLicenceInfo.visible = True
    cmdOk.visible = True
    
    cmdClose.Top = 183
    
    frmLicenceInfo.Height = 228.75
    
End Sub

Private Sub closeItemOfLicence()

    ' 関連するコントロールを非表示にする
    fraLicenceInfo.visible = False
    cmdOk.visible = False
    
    cmdClose.Top = fraLicenceInfo.Top
    
    frmLicenceInfo.Height = 136.5
    
End Sub

Private Sub changeItemOfLicenceAuthenticatedLicence()

    tglRegistLicence.Enabled = False
    txtUserId.Locked = True
    txtPassword.Locked = True
    cmdOk.Enabled = False
    
End Sub

Private Sub changeItemOfLicenceNonAuthenticatedLicence()

    tglRegistLicence.Enabled = True
    txtUserId.Locked = False
    txtPassword.Locked = False
    cmdOk.Enabled = True

End Sub

' =========================================================
' ▽全コントロールのチェック
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：True チェックOK
'
' =========================================================
Private Function validAllControl() As Boolean

    ' 全てのコントロールのチェックを実施する
    If validUserId = False Then
    
        validAllControl = False
        
    ElseIf validPassword = False Then
    
        validAllControl = False
        
    Else
    
        validAllControl = True
    
    End If
    
End Function

' =========================================================
' ▽ユーザIDのチェック
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：True チェックOK
'
' =========================================================
Private Function validUserId() As Boolean

    ' 必須チェック
    If txtUserId.value = "" Then
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse txtUserId
        
        validUserId = False
        
        txtUserId.SetFocus
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidTrue txtUserId
        
        validUserId = True
    
    End If

End Function

' =========================================================
' ▽パスワードのチェック
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：True チェックOK
'
' =========================================================
Private Function validPassword() As Boolean

    ' 必須チェック
    If txtPassword.value = "" Then
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse txtPassword
        
        txtPassword.SetFocus
        validPassword = False
    
    ' 16進数チェック
    ElseIf VBUtil.validHex(txtPassword.value) = False Then
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INVALID
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse txtPassword
        
        txtPassword.SetFocus
        validPassword = False
    
    ' サイズチェック（サイズが2の倍数ではない場合）
    ElseIf Len(txtPassword.value) Mod 2 <> 0 Then
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INVALID_SIZE
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse txtPassword

        txtPassword.SetFocus
        validPassword = False
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidTrue txtPassword
        
        validPassword = True
    
    End If
    

End Function

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

' =========================================================
' ▽ライセンス情報を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeLicenceInfo()

    On Error GoTo err
    
    m_licenceInfo.userId = txtUserId.value
    m_licenceInfo.password = txtPassword.value
    
    ' レジストリに情報を保存する
    m_licenceInfo.writeForRegistry
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽ライセンス情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreLicenceInfo()

    On Error GoTo err
    
    ' レジストリから情報は取得しない
    ' （レジストリからの情報読み込みは前段階で行っているため、ここでは行わない）
    txtUserId.value = m_licenceInfo.userId
    txtPassword.value = m_licenceInfo.password
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub


