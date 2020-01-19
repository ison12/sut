VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValFormPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' フォームの位置をレジストリに出し入れするクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/06/11　新規作成
'
' 特記事項：
' *********************************************************

' =========================================================
' ▽メンバ
' =========================================================
' フォーム名
Private m_formName As String

' =========================================================
' ▽定数プロパティ
' 　Classモジュールでは、Public Constが定義できないのでプロパティで代用する
' =========================================================
Public Property Get REG_SUB_KEY_FORM_POSITION() As String

    REG_SUB_KEY_FORM_POSITION = "form_position"
End Property

Public Property Get REG_SUB_KEY_FORM_POSITION_X() As String

    REG_SUB_KEY_FORM_POSITION_X = "x"
End Property

Public Property Get REG_SUB_KEY_FORM_POSITION_Y() As String

    REG_SUB_KEY_FORM_POSITION_Y = "y"
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
' ▽初期化メソッド
' =========================================================
Public Sub init(ByVal formName As String)

    m_formName = formName
End Sub

' =========================================================
' ▽レジストリから情報を読み込む
' =========================================================
Public Function readForRegistry(ByRef rect As ValRectPt)

    On Error GoTo err
    
    ' スクリーンサイズを取得する
    Dim ss As ScreenSizePt
    ss = WinAPI_User.getScreenSizePt
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_FORM_POSITION & "\" & m_formName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    Dim varX As String
    Dim varY As String
    
    Dim successedX As Boolean
    Dim successedY As Boolean
    
    successedX = registry.getValue(REG_SUB_KEY_FORM_POSITION_X, varX)
    successedY = registry.getValue(REG_SUB_KEY_FORM_POSITION_Y, varY)
    
    ' 両方の取得に成功した場合
    If _
        successedX = True _
    And successedY = True Then
    
        ' レジストリから情報を取得できた場合は、そちらを利用する
        If _
            IsNumeric(varX) = True _
        And IsNumeric(varY) = True Then
        
            ' 情報を設定処理を抜ける
            rect.Left = CSng(varX)
            rect.Top = CSng(varY)
            
            ' 計算した中心位置がスクリーン内に収まっているか
            If VBUtil.isInnerScreen( _
                                ss.virtualSizeX _
                              , ss.virtualSizeY _
                              , ss.virtualSizeWidth _
                              , ss.virtualSizeHeight _
                              , rect.Left _
                              , rect.Top _
                              , rect.Width _
                              , rect.Height) = True Then

                ' レジストリから取得した情報がスクリーン内に収まっていれば処理を抜ける
                Exit Function
            End If
            
        End If
    
    End If
    
    Dim newX As Single
    Dim newY As Single
    
    ' Application本体のウィンドウの中心位置を計算する
    VBUtil.calcCenterPoint _
                        Application.Left _
                      , Application.Top _
                      , Application.Width _
                      , Application.Height _
                      , newX _
                      , newY _
                      , rect.Width _
                      , rect.Height
    
    ' 計算した中心位置がスクリーン内に収まっているか
    If VBUtil.isInnerScreen( _
                        ss.virtualSizeX _
                      , ss.virtualSizeY _
                      , ss.virtualSizeWidth _
                      , ss.virtualSizeHeight _
                      , newX _
                      , newY _
                      , rect.Width _
                      , rect.Height) = False Then

        ' 収まっていない場合
        ' プライマリスクリーンに対しての中心位置を計算する
        VBUtil.calcCenterPoint _
                            0 _
                          , 0 _
                          , ss.primarySizeWidth _
                          , ss.primarySizeHeight _
                          , newX _
                          , newY _
                          , rect.Width _
                          , rect.Height
        
    End If
    
    rect.Left = newX
    rect.Top = newY
    
    Set registry = Nothing
    
    
    Exit Function
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Function

' =========================================================
' ▽レジストリに情報を書き込む
' =========================================================
Public Sub writeForRegistry(ByRef rect As ValRectPt)

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_FORM_POSITION & "\" & m_formName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
                
    registry.setValue REG_SUB_KEY_FORM_POSITION_X, rect.Left
    registry.setValue REG_SUB_KEY_FORM_POSITION_Y, rect.Top
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub
