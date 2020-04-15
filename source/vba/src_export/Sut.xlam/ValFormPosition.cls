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
' フォームの位置情報を永続化するクラス
'
' 作成者　：Ison
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
' ▽アプリケーションプロパティの生成
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Set createApplicationProperties = New ApplicationProperties
    createApplicationProperties.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_FORM_POSITION & "\" & m_formName & ".ini"
    
End Function

' =========================================================
' ▽情報を読み込む
' =========================================================
Public Function readForData(ByRef rect As ValRectPt)

    On Error GoTo err
    
    ' スクリーンサイズを取得する
    Dim ss As ScreenSizePt
    ss = WinAPI_User.getScreenSizePt
    
    Dim appProp As ApplicationProperties: Set appProp = createApplicationProperties
    
    Dim varX As String
    Dim varY As String
    
    varX = appProp.GetValue(ConstantsApplicationProperties.INI_SECTION_DEFAULT, ConstantsApplicationProperties.INI_KEY_X)
    varY = appProp.GetValue(ConstantsApplicationProperties.INI_SECTION_DEFAULT, ConstantsApplicationProperties.INI_KEY_Y)
    
    ' 情報を取得できた場合は、そちらを利用する
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

            ' 取得した情報がスクリーン内に収まっていれば処理を抜ける
            Exit Function
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
    
    Set appProp = Nothing
    
    Exit Function
    
err:

    ' エラー情報を退避する
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    Set appProp = Nothing

    ' 退避したエラー情報を設定しなおす
    VBUtil.setErr errT

    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽情報を書き込む
' =========================================================
Public Sub writeForData(ByRef rect As ValRectPt)

    On Error GoTo err
    
    Dim appProp As ApplicationProperties: Set appProp = createApplicationProperties
    
    appProp.setValue ConstantsApplicationProperties.INI_SECTION_DEFAULT, ConstantsApplicationProperties.INI_KEY_X, rect.Left
    appProp.setValue ConstantsApplicationProperties.INI_SECTION_DEFAULT, ConstantsApplicationProperties.INI_KEY_Y, rect.Top
    
    Set appProp = Nothing
    
    Exit Sub
    
err:
    
    ' エラー情報を退避する
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    Set appProp = Nothing

    ' 退避したエラー情報を設定しなおす
    VBUtil.setErr errT

    Main.ShowErrorMessage

End Sub
