VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelLongTimeProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' Excelマクロで長時間処理する際に不必要なプロパティを
' 一時的にオフにして、本クラスのオブジェクトが削除される際に元に戻す。
'
' 作成者　：Ison
' 履歴　　：2009/04/12　新規作成

' 使用方法：
' 　本クラスのオブジェクトを New する。
' 　関数を抜けると自動で Class_Terminate で後始末が行われる。
'
'   Dim var As New ExcelLongTimeProcessing: var.init
'
'
' *********************************************************

' 警告表示の有効・無効（シート削除時などの警告ダイアログなど）
Public isSetDisplayAlerts    As Boolean
' マクロ実行中のEscキーの有効・無効
Public isSetEnableCancelKey  As Boolean
' イベントの有効・無効
Public isSetEnableEvents     As Boolean
' マウスカーソルの形状
Public isSetCursor           As Boolean
' 画面更新の停止有無
Public isSetScreenUpdating   As Boolean
' Excelワークシートの計算方式
Public isSetCalculation      As Boolean
' Excelへのキーボード・マウス操作
Public isInteractive         As Boolean

' 計算方式（一時保存）
Private tmpCalculation       As XlCalculation

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    On Error Resume Next
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' アプリケーションの各種プロパティを元に戻す
    restoreApplicationProp
    
End Sub

' =========================================================
' ▽初期化処理
'
' 概要　　　：Applicationオブジェクトの各プロパティをオフに設定する。
' 　　　　　　該当プロパティのオン・オフを切り替えるかどうかは
' 　　　　　　本メソッドの引数によって決める。
' 引数　　　：displayAlerts    警告表示
' 　　　　　　enableCancelKey  ESCキーの有効・無効
' 　　　　　　enableEvents     イベントの有効・無効
' 　　　　　　cursor           マウスカーソル
' 　　　　　　screenUpdating   画面更新の有効・無効
' 　　　　　　calculation      Excelワークシートの計算方式
' 　　　　　　interactive      対話モード（キーボードやマウスからの入力有無）
'
' =========================================================
Public Sub init(Optional ByVal displayAlerts As Boolean = True _
              , Optional ByVal enableCancelKey As Boolean = True _
              , Optional ByVal enableEvents As Boolean = True _
              , Optional ByVal cursor As Boolean = False _
              , Optional ByVal screenUpdating As Boolean = False _
              , Optional ByVal calculation As Boolean = True _
              , Optional ByVal interactive As Boolean = True)
              
    ' メンバに情報を設定する
    isSetDisplayAlerts = displayAlerts
    isSetEnableCancelKey = enableCancelKey
    isSetEnableEvents = enableEvents
    isSetCursor = cursor
    isSetScreenUpdating = screenUpdating
    isSetCalculation = calculation
    isInteractive = interactive
    
    ' アプリケーションの各種プロパティをメンバ変数の値に従って設定する
    setApplicationProp

End Sub

' =========================================================
' ▽破棄処理
'
' 概要　　　：
'
' =========================================================
Public Sub destroy()

    ' アプリケーションの各種プロパティを元に戻す
    restoreApplicationProp

End Sub

' =========================================================
' ▽アプリケーションの各種プロパティをメンバ変数の値に従って設定する
'
' 概要　　　：
'
' =========================================================
Private Sub setApplicationProp()

    On Error Resume Next

    ' 計算方式
    If isSetCalculation = True Then
        ' 計算方式を一時変数に保存しておく
        tmpCalculation = Application.calculation
        ' 計算方式を手動にする
        Application.calculation = xlCalculationManual
    End If
    
    ' 警告表示
    If isSetDisplayAlerts = True Then
        Application.displayAlerts = False
    End If
    
    ' Escキーの有効・無効
    If isSetEnableCancelKey = True Then
        Application.enableCancelKey = xlDisabled
    End If
    
    ' イベントの有効・無効
    If isSetEnableEvents = True Then
        Application.enableEvents = False
    End If
    
    ' カーソル
    If isSetCursor = True Then
    
        ' デバッグ中はカーソルが砂時計になるのを防ぐ（邪魔なので）
        #If (DEBUG_MODE <> 1) Then
            Application.cursor = xlWait
        #End If
        
    End If
    
    ' 画面の自動更新
    If isSetScreenUpdating = True Then
        Application.screenUpdating = False
    End If
    
    If isInteractive = True Then
        Application.interactive = False
    End If
    
    On Error GoTo 0

End Sub

' =========================================================
' ▽アプリケーションの各種プロパティを元に戻す
'
' 概要　　　：
'
' =========================================================
Private Sub restoreApplicationProp()

    On Error Resume Next
    
    ' 計算方式
    If isSetCalculation = True Then
        ' 計算方式を元に戻す
        Application.calculation = tmpCalculation
    End If
    
    ' 警告表示
    If isSetDisplayAlerts = True Then
        Application.displayAlerts = True
    End If
    
    ' Escキーの有効・無効
    If isSetEnableCancelKey = True Then
        Application.enableCancelKey = xlInterrupt
    End If
    
    ' イベントの有効・無効
    If isSetEnableEvents = True Then
        Application.enableEvents = True
    End If
    
    ' カーソル
    If isSetCursor = True Then
    
        Application.cursor = xlDefault
    End If
    
    ' 画面の自動更新
    If isSetScreenUpdating = True Then
        Application.screenUpdating = True
    End If

    '
    If isInteractive = True Then
        Application.interactive = True
    End If
    
    On Error GoTo 0

End Sub

