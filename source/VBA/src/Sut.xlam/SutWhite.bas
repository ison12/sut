Attribute VB_Name = "SutWhite"
Option Explicit
' *********************************************************
' SutWhite.dll関連のモジュール
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/03/14　新規作成
'
' 特記事項：
' *********************************************************

Private Const HOURGLASS_WIDTH  As Single = 55
Private Const HOURGLASS_HEIGHT As Single = 64

' DLLのハンドル
Private libraryHandle As Variant
' DLLのパス
Private libraryPath As String

' =========================================================
' ▽ライブラリをロードする
'
' 概要　　　：
'
' =========================================================
Public Function LoadLibrary()

    ' DLLのパスを設定する
    #If (DEBUG_MODE = 1) Then
    
        #If VBA7 And Win64 Then
            libraryPath = SutWorkbook.path & "\..\CPP\Sut\x64\Debug ASM\SutWhite.dll"
        
        #Else
            libraryPath = SutWorkbook.path & "\..\CPP\Sut\Debug ASM\SutWhite.dll"
        
        #End If
    #Else
        ' DLLのパスを設定
        libraryPath = SutWorkbook.path & "\lib\SutWhite.dll"
        
    #End If
    

    ' モジュールハンドル
    Dim handle As Variant
    
    ' ハンドルを取得する
    handle = WinAPI_Kernel32.GetModuleHandle _
                (libraryPath)

    ' 未ロードの場合
    If handle = 0 Then
    
        ' dllをロードする
        libraryHandle = WinAPI_Kernel32.LoadLibrary _
                                    (libraryPath)
    
        ' 戻り値がNULLの場合
        If libraryHandle = 0 Then
        
            ' エラーを発行する
            err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                    , _
                    , ConstantsError.ERR_DESC_DLL_FUNCTION_FAILED
        End If
    
    ' ロード済みの場合
    Else
    
        libraryHandle = handle
    End If

End Function

' =========================================================
' ▽ライブラリを解放する
'
' 概要　　　：
'
' =========================================================
Public Function freeLibrary()

    ' ハンドルのチェックを行う
    If libraryHandle = 0 Then
    
        ' ハンドルが割り当てられていない場合、終了する
        Exit Function
    End If
    
    ' SutWhite.dll を解放する
    WinAPI_Kernel32.freeLibrary (libraryHandle)
    
    ' ハンドルをゼロクリアする
    libraryHandle = 0
End Function

' =========================================================
' ▽ライブラリを初期化する
'
' 概要　　　：
'
' =========================================================
Public Function initialize()

    ' ハンドルのチェックを行う
    If libraryHandle = 0 Then
    
        ' ハンドルが割り当てられていない場合、終了する
        Exit Function
    End If

    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "Initialize")
    
    ' DLL関数の戻り値
    Dim ret As Long
    
    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtr(procAddr)

    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function

' =========================================================
' ▽ライブラリを破棄する
'
' 概要　　　：
'
' =========================================================
Public Function destroy()

    ' ハンドルのチェックを行う
    If libraryHandle = 0 Then
    
        ' ハンドルが割り当てられていない場合、終了する
        Exit Function
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "Destroy")
    
    ' DLL関数の戻り値
    Dim ret As Long
    
    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtr(procAddr)
    
    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function

' =========================================================
' ▽スプラッシュウィンドウを表示する
'
' 概要　　　：
'
' =========================================================
Public Function showSplashWindow()

    ' ハンドルのチェックを行う
    If libraryHandle = 0 Then
    
        ' ハンドルが割り当てられていない場合
        LoadLibrary     ' ライブラリのロード
        initialize      ' ライブラリの初期化
        
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "ShowSplashWindow")
    
    ' DLL関数の戻り値
    Dim ret As Long
    
    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtrParam(procAddr, ExcelUtil.getApplicationHWnd)
    
    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function

' =========================================================
' ▽スプラッシュウィンドウの表示が完了するまで待機する
'
' 概要　　　：
' 引数　　　：waitTime ミリ秒指定
' 戻り値　　：0  正常
' 　　　　　　10 タイムアウト
'
' =========================================================
Public Function waitSplashWindow(ByVal waitTime As Long) As Long

    ' ハンドルのチェックを行う
    If libraryHandle = 0 Then
    
        ' ハンドルが割り当てられていない場合、終了する
        Exit Function
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "WaitSplashWindow")
    
    ' DLL関数の戻り値
    Dim ret As Long
        
    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtrParamInt(procAddr, waitTime)
    
    If ret <> 0 And ret <> 10 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
    waitSplashWindow = ret
    
End Function

' =========================================================
' ▽砂時計ウィンドウを表示する（フォームの中心に表示する）
'
' 概要　　　：
' 引数　　　：frmObj フォームオブジェクト
'
' =========================================================
Public Function showHourglassWindowOnCenterPt(Optional ByRef frmObj As Object = Nothing, _
                                            Optional ByVal x As Long = 0, _
                                            Optional ByVal y As Long = 0)

    Dim newX As Single
    Dim newY As Single

    ' DPIを取得する
    Dim d As DPI
    d = WinAPI_GDI.getSystemDPI
    
    If Not frmObj Is Nothing Then
    
        VBUtil.calcCenterPoint frmObj.Left _
                             , frmObj.Top _
                             , frmObj.Width _
                             , frmObj.Height _
                             , newX _
                             , newY _
                             , VBUtil.convertPixelToPoint(d.horizontal, HOURGLASS_WIDTH) _
                             , VBUtil.convertPixelToPoint(d.vertical, HOURGLASS_HEIGHT)
    Else
    
        VBUtil.calcCenterPoint Application.Left _
                             , Application.Top _
                             , Application.Width _
                             , Application.Height _
                             , newX _
                             , newY _
                             , VBUtil.convertPixelToPoint(d.horizontal, HOURGLASS_WIDTH) _
                             , VBUtil.convertPixelToPoint(d.vertical, HOURGLASS_HEIGHT)
    End If
                         
    showHourglassWindow VBUtil.convertPointToPixel(d.horizontal, newX) + x _
                      , VBUtil.convertPointToPixel(d.vertical, newY) + y

End Function

' =========================================================
' ▽砂時計ウィンドウを表示する
'
' 概要　　　：座標はピクセル単位で指定する
' 引数　　　：x ウィンドウ表示位置 X
' 　　　　　　y ウィンドウ表示位置 Y
'
' =========================================================
Public Function showHourglassWindow(Optional ByVal x As Long = 0 _
                                  , Optional ByVal y As Long = 0)

    ' ハンドルのチェックを行う
    If libraryHandle = 0 Then
    
        ' ハンドルが割り当てられていない場合
        LoadLibrary     ' ライブラリのロード
        initialize      ' ライブラリの初期化
        
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "ShowHourglassWindow")
    
    ' DLL関数の戻り値
    Dim ret As Long
    
    ' ウィンドウ表示位置
    Dim pt As point
    pt.x = x
    pt.y = y
    
    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtrParam2(procAddr, ExcelUtil.getApplicationHWnd, pt)
    
    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function

' =========================================================
' ▽砂時計ウィンドウを非表示にする
'
' 概要　　　：
'
' =========================================================
Public Function closeHourglassWindow()

    ' ハンドルのチェックを行う
    If libraryHandle = 0 Then
    
        ' ハンドルが割り当てられていない場合、終了する
        Exit Function
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "CloseHourglassWindow")
    
    ' DLL関数の戻り値
    Dim ret As Long
    
    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtr(procAddr)
    
    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function
