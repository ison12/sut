VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValLicenceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' *********************************************************
' ライセンス情報
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/05/14　新規作成
'
' 特記事項：
' *********************************************************

' =========================================================
' ▽メンバ
' =========================================================
' ユーザID
Public userId As String
' パスワード
Public password As String

' =========================================================
' ▽定数プロパティ
' 　Classモジュールでは、Public Constが定義できないのでプロパティで代用する
' =========================================================
' キー値 レジストリキー　ライセンス情報
Public Property Get REG_SUB_KEY_LICENCE_INFO() As String

    REG_SUB_KEY_LICENCE_INFO = "licence_info"
End Property

' キー値 ユーザID
Public Property Get KEY_USER_ID() As String

    KEY_USER_ID = "userId"
End Property

' キー値 パスワード
Public Property Get KEY_PASSWORD() As String

    KEY_PASSWORD = "password"
End Property

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
' ▽レジストリから情報を読み込む
' =========================================================
Public Sub readForRegistry()

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_LICENCE_INFO) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Dim retUserId   As String
    Dim retPassword As String

    registry.getValue KEY_USER_ID, retUserId
    registry.getValue KEY_PASSWORD, retPassword

    userId = retUserId
    password = retPassword
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽レジストリに情報を書き込む
' =========================================================
Public Sub writeForRegistry()

    On Error GoTo err
    
    Dim j As Long
    
    Dim licenceInfo(0 To 1 _
                  , 0 To 1) As Variant
    
    
    licenceInfo(j, 0) = KEY_USER_ID
    licenceInfo(j, 1) = userId: j = j + 1

    licenceInfo(j, 0) = KEY_PASSWORD
    licenceInfo(j, 1) = password: j = j + 1

    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_LICENCE_INFO) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' レジストリに情報を設定する
    registry.setValues licenceInfo
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub



