VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeAuthenticateLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ライセンス認証を実施するクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/05/18　新規作成
'
' 特記事項：
' *********************************************************

' ライセンス情報
Private m_licenceInfo As ValLicenceInfo

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ▽初期化処理
' =========================================================
Public Sub init(ByRef licenceInfo As ValLicenceInfo)

    Set m_licenceInfo = licenceInfo
End Sub

' =========================================================
' ▽ユーザ認証
'
' 概要　　　：ユーザ認証を実施する。
' 引数　　　：
' 戻り値　　：True 認証OK
'
' =========================================================
Public Function executeAuthor() As Boolean

    On Error GoTo err

    ' パスワードのバイト配列
    Dim passwordByte() As Byte
    ' パスワードをバイト配列に変換する
    passwordByte = SutGray.ConvertHexToBinaryDataWrapper(m_licenceInfo.password)
    
    ' 複合化情報のバイト配列
    Dim decryptBuffer()  As Byte
    ' 複合化情報の文字列
    Dim decryptBufferStr As String
    
    ' 入力されたパスワードを復元する
    decryptBuffer = SutGray.DecryptWrapper(SESSION_KEY_PASSWORD, passwordByte)
    ' 複合されたバッファをANSI→UNICODE変換し文字列として格納する
    decryptBufferStr = StrConv(decryptBuffer, vbUnicode)
    
    ' 複合された文字列とユーザIDを比較する
    If decryptBufferStr = m_licenceInfo.userId Then
    
        ' 一致している場合、認証OK
        executeAuthor = True
        
    Else
    
        ' 一致しなかった場合、認証NG
        executeAuthor = False
    End If
    
    Exit Function

err:

    ' DLLの呼び出しに失敗した場合（警告レベル）
    If ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING = err.Number Then
    
        executeAuthor = False
        
    ' その他のエラーの場合
    Else
    
        ' 上位モジュールにエラーを発行する
        err.Raise err.Number _
                , err.souce _
                , err.Description _
                , err.HelpFile _
                , err.HelpContext
    End If


End Function

' =========================================================
' ▽試用期間に関する日付の取得
'
' 概要　　　：試用期間に関する日付を取得する。
' 　　　　　　（ソフトの使用開始日付）
' 引数　　　：
' 戻り値　　：ソフトの使用開始日付
'
' =========================================================
Public Function getProbationDate() As Date

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME_ALIAS _
                                                  , "" _
                                                  , ConstantsCommon.PROBATION_REG_DIR) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    ' 日付取得 成功有無フラグ
    Dim dateGetSuccess As Boolean
    ' 日付設定 成功有無フラグ
    Dim dateSetSuccess As Boolean
    
    ' 日付
    Dim date_             As Date
    ' 日付（文字列型）
    Dim dateStr           As String
    ' 日付（バイト型）
    Dim dateByte()        As Byte
    ' 暗号化済み日付
    Dim dateEncryptByte() As Byte
    
    ' レジストリから日付を取得する（暗号化済み）
    dateGetSuccess = registry.getValueByByte(ConstantsCommon.PROBATION_REG_KEY, dateEncryptByte)
    
    ' レジストリから日付を取得できた場合
    If dateGetSuccess = True Then
    
        ' ---------------------------------------------------------
        ' 暗号化データの復元
        ' ---------------------------------------------------------
        ' エラーの発生に備える
        On Error Resume Next
        
        ' 複合化情報のバイト配列
        Dim decryptBuffer()  As Byte
        ' 複合化情報の文字列
        Dim decryptBufferStr As String
        
        ' データを復元する
        decryptBuffer = SutGray.DecryptWrapper(ConstantsCommon.SESSION_KEY_PASSWORD, dateEncryptByte)
        
        ' データの復元に失敗
        If err.Number = ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING Then
        
            decryptBufferStr = "2000-01-01"
            
        ' データの復元に成功
        ElseIf err.Number = 0 Then
        
            ' 複合されたバッファを文字列に変換する
            decryptBufferStr = CStr(decryptBuffer)
            
        ' その他のエラー
        Else
        
            GoTo err
        End If
        
        ' ---------------------------------------------------------
        
        ' エラーが発生した場合、errブロックに移動する
        On Error GoTo err
        
        ' 複合化情報を日付に変換する
        dateStr = Format(decryptBufferStr, "yyyy-mm-dd")
        
        ' 日付として認識可能な場合
        If IsDate(dateStr) = True Then
        
            ' そのまま変換する
            date_ = CDate(dateStr)
            
        ' 日付として認識できない場合（データが改竄された可能性あり）
        Else
        
            ' 現在の日付から -N 日した結果を取得する
            date_ = DateAdd("d", -1 * ConstantsCommon.PROBATION_DAY, Now)
            
        End If

    ' レジストリから日付を取得できなかった場合
    Else
    
        ' 現在日付を取得
        date_ = Now
        ' 現在日付を文字列に変換
        dateStr = Format(date_, "yyyy-mm-dd")
        ' 文字列をバイト配列に変換
        dateByte = dateStr
        
        ' 日付を暗号化する
        dateEncryptByte = SutGray.EncryptWrapper(ConstantsCommon.SESSION_KEY_PASSWORD, dateByte)
        
        ' レジストリに日付を設定する
        dateSetSuccess = registry.setValueByByte(ConstantsCommon.PROBATION_REG_KEY, dateEncryptByte)
        
        ' 日付の設定に失敗した場合
        If dateSetSuccess = False Then
        
            err.Raise ConstantsError.ERR_NUMBER_REGISTRY_ACCESS_FAILED _
                    , _
                    , ConstantsError.ERR_DESC_REGISTRY_ACCESS_FAILED
        
        End If
    
    End If
    
    ' 戻り値に日付を設定する
    getProbationDate = date_
    
    Exit Function
    
err:
    
    ' 上位モジュールにエラーを発行する
    err.Raise err.Number _
            , err.Source _
            , err.Description _
            , err.HelpFile _
            , err.HelpContext

End Function

' =========================================================
' ▽試用期間の判定
'
' 概要　　　：試用期間内で使用されているかを判定する。
' 引数　　　：fromDate ソフトの使用開始日付
' 戻り値　　：True 試用期間内
'
' =========================================================
Public Function isRangeProbation(ByVal fromDate As Date) As Boolean

    ' 試用期間の最終日
    Dim toDate As Date
    
    ' 試用期間の最終日を計算する
    toDate = DateAdd("d", ConstantsCommon.PROBATION_DAY - 1, fromDate)
    
    ' 試用期間の最終日の範囲内であるか
    If Format(toDate, "yyyy-mm-dd") >= Format(Now, "yyyy-mm-dd") Then
    
        ' 範囲内
        isRangeProbation = True
    Else
    
        ' 範囲外
        isRangeProbation = False
    End If

End Function

' =========================================================
' ▽試用期間の残り日数の取得
'
' 概要　　　：試用期間の残り日数の取得
' 　　　　　　戻り値 <= 0 が成立する場合は、ソフトの使用はできない。
' 引数　　　：fromDate ソフトの使用開始日
' 戻り値　　：残り日数
'
' =========================================================
Public Function getRemainderProbationDay(ByVal fromDate As Date) As Long

    getRemainderProbationDay = ConstantsCommon.PROBATION_DAY - DateDiff("d", fromDate, Now)

End Function
