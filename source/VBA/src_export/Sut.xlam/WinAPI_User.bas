Attribute VB_Name = "WinAPI_User"
Option Explicit

' *********************************************************
' user32.dllで定義されている関数郡や定数。
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/10/11　新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' ▽ウィンドウ検索
'
' 概要　　　：ウィンドウハンドルを検索する。
' 引数　　　：lpClassName  クラス名
' 　　　　　　lpWindowName ウィンドウタイトル
' 戻り値　　：ウィンドウハンドル
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
            (ByVal lpClassName As String _
           , ByVal lpWindowName As String) As LongPtr
#Else
    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
            (ByVal lpClassName As String _
           , ByVal lpWindowName As String) As Long
#End If

' =========================================================
' ▽ウィンドウ情報取得
'
' 概要　　　：ウィンドウ情報を取得する
' 引数　　　：hWnd   ウィンドウハンドル
' 　　　　　　nIndex 取得する情報
' 戻り値　　：ウィンドウ情報
'
' =========================================================
#If VBA7 And Win64 Then

    Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
           (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) As LongPtr
                                                           
    Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" _
            (ByVal hWnd As LongPtr _
           , ByVal nIndex As Long) As Long
#Else
    Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" _
            (ByVal hWnd As Long _
           , ByVal nIndex As Long) As Long
#End If

' =========================================================
' ▽ウィンドウ情報設定
'
' 概要　　　：ウィンドウ情報を設定する
' 引数　　　：hWnd      ウィンドウハンドル
' 　　　　　　nIndex    取得する情報
' 　　　　　　dwNewLong 新しい情報
' 戻り値　　：結果コード 0の場合エラー
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
            (ByVal hWnd As LongPtr _
           , ByVal nIndex As Long _
           , ByVal dwNewLong As LongPtr) As Long
#Else
    Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
            (ByVal hWnd As Long _
           , ByVal nIndex As Long _
           , ByVal dwNewLong As Long) As Long
#End If

' =========================================================
' ▽クライアント座標からスクリーン座標への変換
'
' 概要　　　：
' 引数　　　：hWnd      ウィンドウハンドル
' 　　　　　　lpPoint   ポイント情報構造体
' 戻り値　　：結果コード 0の場合エラー
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function ClientToScreen Lib "user32.dll" _
            (ByVal hWnd As LongPtr _
           , ByRef lpPoint As Point) As Boolean
#Else
    Public Declare Function ClientToScreen Lib "user32.dll" _
            (ByVal hWnd As Long _
           , ByRef lpPoint As point) As Boolean
#End If

' =========================================================
' ▽システム固有の情報を取得
'
' 概要　　　：
' 引数　　　：nIndex 取得する情報の種類
' 戻り値　　：nIndexに対応するシステム固有の情報
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" ( _
             ByVal nIndex As Long) As Long
#Else
    Public Declare Function GetSystemMetrics Lib "user32.dll" ( _
             ByVal nIndex As Long) As Long
#End If

' =========================================================
' ▽メニュー生成
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function createMenu Lib "user32.dll" Alias "CreateMenu" () As Long
#Else
    Public Declare Function createMenu Lib "user32.dll" Alias "CreateMenu" () As Long
#End If
    

' =========================================================
' ▽メニュー破棄
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function destroyMenu Lib "user32.dll" _
            Alias "DestroyMenu" (ByVal hMenu As LongPtr) As Long
#Else
    Public Declare Function destroyMenu Lib "user32.dll" _
            Alias "DestroyMenu" (ByVal hMenu As Long) As Long

#End If

' =========================================================
' ▽ポップアップメニュー生成
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function CreatePopupMenu Lib "user32.dll" () As Long
#Else
    Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
#End If
       
' =========================================================
' ▽メニュー設定
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SetMenu Lib "user32.dll" _
            (ByVal hWnd As LongPtr, ByVal hMenu As LongPtr) As Boolean
#Else
    Public Declare Function SetMenu Lib "user32.dll" _
            (ByVal hWnd As Long, ByVal hMenu As Long) As Boolean
#End If

' =========================================================
' ▽メニューアイテム追加
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" _
            (ByVal hMenu As LongPtr _
           , ByVal uItem As Long _
           , ByVal fByPosition As Boolean _
           , ByRef lpmii As MENUITEMINFO) As Boolean
#Else
    Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" _
            (ByVal hMenu As Long _
           , ByVal uItem As Long _
           , ByVal fByPosition As Boolean _
           , ByRef lpmii As MENUITEMINFO) As Boolean
#End If
        
' =========================================================
' ▽メニューアイテム設定
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" _
            (ByVal hMenu As LongPtr _
           , ByVal uItem As Long _
           , ByVal fByPosition As Boolean _
           , ByRef lpmii As MENUITEMINFO) As Boolean
#Else
    Public Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" _
            (ByVal hMenu As Long _
           , ByVal uItem As Long _
           , ByVal fByPosition As Boolean _
           , ByRef lpmii As MENUITEMINFO) As Boolean
#End If
        
' =========================================================
' ▽ポップアップメニュー表示
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function TrackPopupMenu Lib "user32.dll" _
            (ByVal hMenu As LongPtr _
           , ByVal uFlags As Long _
           , ByVal X As Long _
           , ByVal Y As Long _
           , ByVal nReserved As Long _
           , ByVal hWnd As LongPtr _
           , ByRef notUserd As Long) As Boolean
#Else
    Public Declare Function TrackPopupMenu Lib "user32.dll" _
            (ByVal hMenu As Long _
           , ByVal uFlags As Long _
           , ByVal x As Long _
           , ByVal y As Long _
           , ByVal nReserved As Long _
           , ByVal hWnd As Long _
           , ByRef notUserd As Long) As Boolean
#End If
        
' =========================================================
' ▽ポップアップメニュー表示
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function TrackPopupMenuEx Lib "user32.dll" _
            (ByVal hMenu As LongPtr _
           , ByVal fuFlags As Long _
           , ByVal X As Long _
           , ByVal Y As Long _
           , ByVal hWnd As LongPtr _
           , ByRef var As LongPtr) As Boolean
#Else
    Public Declare Function TrackPopupMenuEx Lib "user32.dll" _
            (ByVal hMenu As Long _
           , ByVal fuFlags As Long _
           , ByVal x As Long _
           , ByVal y As Long _
           , ByVal hWnd As Long _
           , ByRef var As Long) As Boolean
#End If

' =========================================================
' ▽メニューバー描画
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function DrawMenuBar Lib "user32.dll" _
            (ByVal hWnd As LongPtr) As Long
#Else
    Public Declare Function DrawMenuBar Lib "user32.dll" _
            (ByVal hWnd As Long) As Long
#End If

' =========================================================
' ▽アクセラレータテーブルを作成します。
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function CreateAcceleratorTable Lib "user32.dll" Alias "CreateAcceleratorTableA" _
            (ByRef lpaccl() As ACCEL _
           , ByVal cEntries As Long) As Long
#Else
    Public Declare Function CreateAcceleratorTable Lib "user32.dll" Alias "CreateAcceleratorTableA" _
            (ByRef lpaccl() As ACCEL _
           , ByVal cEntries As Long) As Long
#End If

' =========================================================
' ▽アクセラレータテーブルを破棄します。
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function DestroyAcceleratorTable Lib "user32.dll" (ByVal hAccel As LongPtr) As Boolean
#Else
    Public Declare Function DestroyAcceleratorTable Lib "user32.dll" (ByVal hAccel As Long) As Boolean
#End If

' =========================================================
' ▽メニューコマンドに対応するアクセラレータキー（ ショートカットキー）を処理します。
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function TranslateAccelerator Lib "user32.dll" Alias "TranslateAcceleratorA" _
            (ByVal hWnd As LongPtr _
           , ByVal hAccTable As LongPtr _
           , ByRef lpMsg As LongPtr) As Long
#Else
    Public Declare Function TranslateAccelerator Lib "user32.dll" Alias "TranslateAcceleratorA" _
            (ByVal hWnd As Long _
           , ByVal hAccTable As Long _
           , ByRef lpMsg As Long) As Long
#End If

' =========================================================
' ▽ウィンドウプロシージャ呼出
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As LongPtr _
       , ByVal hWnd As LongPtr _
       , ByVal msg As Long _
       , ByVal wParam As Long _
       , ByVal lParam As Long) As Long
#Else
    Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long _
       , ByVal hWnd As Long _
       , ByVal msg As Long _
       , ByVal wParam As Long _
       , ByVal lParam As Long) As Long
#End If
   
' =========================================================
' ▽デバイスコンテキストハンドル取得
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function GetDC Lib "user32.dll" _
        (ByVal hWnd As LongPtr) As Long
#Else
    Public Declare Function GetDC Lib "user32.dll" _
        (ByVal hWnd As Long) As Long
#End If

' =========================================================
' ▽デバイスコンテキストハンドル解放
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function ReleaseDC Lib "user32.dll" _
        (ByVal hWnd As LongPtr _
       , ByVal hdc As LongPtr) As Long
#Else
    Public Declare Function ReleaseDC Lib "user32.dll" _
        (ByVal hWnd As Long _
       , ByVal hdc As Long) As Long
#End If

' =========================================================
' ▽メッセージボックスの呼び出し
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================

#If VBA7 And Win64 Then
    Public Declare PtrSafe Function MessageBox Lib "user32.dll" Alias "MessageBoxA" _
        (ByVal hWnd As LongPtr _
        , ByVal lpText As String _
        , ByVal lpCaption As String _
        , ByVal uType As Long) As Long
#Else
    Public Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" _
        (ByVal hWnd As Long _
        , ByVal lpText As String _
        , ByVal lpCaption As String _
        , ByVal uType As Long) As Long
#End If

' =========================================================
' ▽ウィンドウ位置変更
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SetWindowPos Lib "user32.dll" _
        (ByVal hWnd As LongPtr _
        , ByVal hWndInsertAfter As LongPtr _
        , ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long _
        , ByVal wFlags As Long) As Long

#Else
    Public Declare Function SetWindowPos Lib "user32.dll" _
        (ByVal hWnd As Long _
        , ByVal hWndInsertAfter As Long _
        , ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long _
        , ByVal wFlags As Long) As Long
#End If

' =========================================================
' ▽Installs an application-defined hook procedure into a hook chain. You would install a hook procedure to monitor the system for certain types of events. These events are associated either with a specific thread or with all threads in the same desktop as the calling thread.
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
                                 ByVal idHook As Long, _
                                 ByVal lpfn As LongPtr, _
                                 ByVal hmod As LongPtr, _
                                 ByVal dwThreadId As Long) As LongPtr
#Else
    Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
                                 ByVal idHook As Long, _
                                 ByVal lpfn As Long, _
                                 ByVal hmod As Long, _
                                 ByVal dwThreadId As Long) As Long
#End If

' =========================================================
' ▽Passes the hook information to the next hook procedure in the current hook chain. A hook procedure can call this function either before or after processing the hook information.
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
                                 ByVal hHook As LongPtr, _
                                 ByVal nCode As Long, _
                                 ByVal wParam As LongPtr, _
                                 lParam As Any) As LongPtr
#Else
    Public Declare Function CallNextHookEx Lib "user32" ( _
                                 ByVal hHook As Long, _
                                 ByVal nCode As Long, _
                                 ByVal wParam As Long, _
                                 lParam As Any) As Long
#End If

' =========================================================
' ▽Removes a hook procedure installed in a hook chain by the SetWindowsHookEx function.
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
                                 ByVal hHook As LongPtr) As Long
#Else
    Public Declare Function UnhookWindowsHookEx Lib "user32" ( _
                                 ByVal hHook As Long) As Long
#End If

' =========================================================
' ▽Places (posts) a message in the message queue associated with the thread that created the specified window and returns without waiting for the thread to process the message.
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function PostMessage Lib "user32.dll" Alias "PostMessageA" ( _
                                 ByVal hWnd As LongPtr, _
                                 ByVal wMsg As Long, _
                                 ByVal wParam As LongPtr, _
                                 ByVal lParam As LongPtr) As Long
#Else
    Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" ( _
                                 ByVal hwnd As Long, _
                                 ByVal wMsg As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long
#End If

' =========================================================
' ▽Retrieves a handle to the window that contains the specified point.
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                 ByVal Point As LongLong) As LongPtr
#Else
    Public Declare Function WindowFromPoint Lib "user32" ( _
                                 ByVal xPoint As Long, _
                                 ByVal yPoint As Long) As Long
#End If

' =========================================================
' ▽Retrieves the position of the mouse cursor, in screen coordinates.
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function GetCursorPos Lib "user32.dll" ( _
                                 ByRef lpPoint As POINTAPI) As Long
#Else
    Public Declare Function GetCursorPos Lib "user32.dll" ( _
                                 ByRef lpPoint As POINTAPI) As Long
#End If


Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2
Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_NOMOVE As Long = &H2&

Public Const MB_OK = &H0
Public Const MB_OKCANCEL = &H1
Public Const MB_ABORTRETRYIGNORE = &H2
Public Const MB_YESNOCANCEL = &H3
Public Const MB_YESNO = &H4
Public Const MB_RETRYCANCEL = &H5

Public Const MB_TOPMOST = &H40000
Public Const MB_ICONHAND = &H10
Public Const MB_ICONQUESTION = &H20
Public Const MB_ICONEXCLAMATION = &H30
Public Const MB_ICONASTERISK = &H40

Public Const MB_ICONERROR = MB_ICONHAND
Public Const MB_ICONWARNING = MB_ICONEXCLAMATION

Public Const MB_DEFBUTTON1 = &H0
Public Const MB_DEFBUTTON2 = &H100
Public Const MB_DEFBUTTON3 = &H200

Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7

' ---------------------------------------------------------
' GetSystemMetrics関連の定数
' ---------------------------------------------------------
Public Const SM_CXSCREEN             As Long = 0
Public Const SM_CYSCREEN             As Long = 1
Public Const SM_XVIRTUALSCREEN       As Long = 76
Public Const SM_YVIRTUALSCREEN       As Long = 77
Public Const SM_CXVIRTUALSCREEN      As Long = 78
Public Const SM_CYVIRTUALSCREEN      As Long = 79

Public Type ScreenSize

    primarySizeWidth  As Long   ' プライマリスクリーンの幅
    primarySizeHeight As Long   ' プライマリスクリーンの高さ
    
    virtualSizeX      As Long   ' 仮想スクリーンの原点X
    virtualSizeY      As Long   ' 仮想スクリーンの原点Y
    virtualSizeWidth  As Long   ' 仮想スクリーンの幅
    virtualSizeHeight As Long   ' 仮想スクリーンの高さ

End Type

Public Type ScreenSizePt

    primarySizeWidth  As Single   ' プライマリスクリーンの幅
    primarySizeHeight As Single   ' プライマリスクリーンの高さ
    
    virtualSizeX      As Single   ' 仮想スクリーンの原点X
    virtualSizeY      As Single   ' 仮想スクリーンの原点Y
    virtualSizeWidth  As Single   ' 仮想スクリーンの幅
    virtualSizeHeight As Single   ' 仮想スクリーンの高さ

End Type

' ---------------------------------------------------------
' メニュー関連　定数
' ---------------------------------------------------------
' Menu item constants
Public Const SC_SIZE         As Long = &HF000&
Public Const SC_SEPARATOR    As Long = &HF00F&
Public Const SC_MOVE         As Long = &HF010&
Public Const SC_MINIMIZE     As Long = &HF020&
Public Const SC_MAXIMIZE     As Long = &HF030&
Public Const SC_CLOSE        As Long = &HF060&
Public Const SC_RESTORE      As Long = &HF120&

' SetMenuItemInfo fMask Constants
Public Const MF_INSERT            As Long = &H0&
Public Const MF_CHANGE            As Long = &H80&
Public Const MF_APPEND            As Long = &H100&
Public Const MF_DELETE            As Long = &H200&
Public Const MF_REMOVE            As Long = &H1000&

Public Const MF_BYCOMMAND         As Long = &H0&
Public Const MF_BYPOSITION        As Long = &H400&

Public Const MF_SEPARATOR         As Long = &H800&

Public Const MF_ENABLED           As Long = &H0&
Public Const MF_GRAYED            As Long = &H1&
Public Const MF_DISABLED          As Long = &H2&

Public Const MF_UNCHECKED         As Long = &H0&
Public Const MF_CHECKED           As Long = &H8&
Public Const MF_USECHECKBITMAPS   As Long = &H200&

Public Const MF_STRING            As Long = &H0&
Public Const MF_BITMAP            As Long = &H4&
Public Const MF_OWNERDRAW         As Long = &H100&

Public Const MF_POPUP             As Long = &H10&
Public Const MF_MENUBARBREAK      As Long = &H20&
Public Const MF_MENUBREAK         As Long = &H40&

Public Const MF_UNHILITE          As Long = &H0&
Public Const MF_HILITE            As Long = &H80&

Public Const MIIM_STATE       As Long = &H1&
Public Const MIIM_ID          As Long = &H2&
Public Const MIIM_SUBMENU     As Long = &H4&
Public Const MIIM_CHECKMARKS  As Long = &H8&
Public Const MIIM_TYPE        As Long = &H10&
Public Const MIIM_DATA        As Long = &H20&

Public Const MIIM_STRING      As Long = &H40&
Public Const MIIM_BITMAP      As Long = &H80&
Public Const MIIM_FTYPE       As Long = &H100&

Public Const TPM_LEFTBUTTON   As Long = &H0&
Public Const TPM_RIGHTBUTTON As Long = &H2&
Public Const TPM_LEFTALIGN    As Long = &H0&
Public Const TPM_CENTERALIGN As Long = &H4&
Public Const TPM_RIGHTALIGN   As Long = &H8&

Public Const TPM_TOPALIGN         As Long = &H0&
Public Const TPM_VCENTERALIGN     As Long = &H10&
Public Const TPM_BOTTOMALIGN      As Long = &H20&

Public Const TPM_HORIZONTAL       As Long = &H0&
Public Const TPM_VERTICAL         As Long = &H40&
Public Const TPM_NONOTIFY         As Long = &H80&
Public Const TPM_RETURNCMD        As Long = &H100&

Public Const TPM_RECURSE          As Long = &H1&
Public Const TPM_HORPOSANIMATION  As Long = &H400&
Public Const TPM_HORNEGANIMATION  As Long = &H800&
Public Const TPM_VERPOSANIMATION  As Long = &H1000&
Public Const TPM_VERNEGANIMATION  As Long = &H2000&

' User-defined Types.
Public Type MENUITEMINFO
    cbSize        As Long
    fMask         As Long
    fType         As Long
    fState        As Long
    wID           As Long
    hSubMenu      As Long
    hbmpChecked   As Long
    hbmpUnchecked As Long
    dwItemData    As Long
    dwTypeData    As String
    cch           As Long
End Type

' ---------------------------------------------------------

' ---------------------------------------------------------
' ウィンドウ関連　定数
' ---------------------------------------------------------
Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)

Public Const WS_OVERLAPPED       As Long = &H0&
Public Const WS_POPUP            As Long = &H80000000
Public Const WS_CHILD            As Long = &H40000000
Public Const WS_MINIMIZE         As Long = &H20000000
Public Const WS_VISIBLE          As Long = &H10000000
Public Const WS_DISABLED         As Long = &H8000000
Public Const WS_CLIPSIBLINGS     As Long = &H4000000
Public Const WS_CLIPCHILDREN     As Long = &H2000000
Public Const WS_CAPTION          As Long = &HC00000
Public Const WS_BORDER           As Long = &H800000
Public Const WS_DLGFRAME         As Long = &H400000
Public Const WS_VSCROLL          As Long = &H200000
Public Const WS_HSCROLL          As Long = &H100000
Public Const WS_SYSMENU          As Long = &H80000
Public Const WS_THICKFRAME       As Long = &H40000
Public Const WS_GROUP            As Long = &H20000
Public Const WS_TABSTOP          As Long = &H10000
Public Const WS_MINIMIZEBOX      As Long = &H20000
Public Const WS_MAXIMIZEBOX      As Long = &H10000
Public Const WS_OVERLAPPEDWINDOW As Long = &HCF0000
Public Const WS_POPUPWINDOW      As Long = &H80880000

Public Const WH_MOUSE_LL         As Long = 14

Public Const WM_MOUSEWHEEL       As Long = &H20A
Public Const WM_KEYDOWN          As Long = &H100
Public Const WM_KEYUP            As Long = &H101
Public Const WM_LBUTTONDOWN      As Long = &H201

Public Const HC_ACTION           As Long = 0
Public Const GWL_HINSTANCE       As Long = (-6)

Public Const VK_UP               As Long = &H26
Public Const VK_DOWN             As Long = &H28
' ---------------------------------------------------------

' ---------------------------------------------------------
' ウィンドウメッセージ　定数
' ---------------------------------------------------------
Public Const WM_COMMAND          As Long = &H111
Public Const WM_INITMENUPOPUP    As Long = &H117 ' WPARAM には、ポップアップメニューのハンドルが返される
Public Const WM_SETCURSOR        As Long = &H20
' ---------------------------------------------------------

Public Type Point
    X As Long
    Y As Long
End Type

#If Win64 Then
    Public Type POINTAPI
        XY As LongLong
    End Type
#Else
    Public Type POINTAPI
        X As Long
        Y As Long
    End Type
#End If

Public Type ACCEL
    fVirt As Byte
    key   As Long
    cmd   As Long
End Type

Public Const FVIRTKEY  As Long = &H1
Public Const FNOINVERT As Long = &H2
Public Const FSHIFT    As Long = &H4
Public Const FCONTROL  As Long = &H8
Public Const FALT      As Long = &H10


Public Type MOUSEHOOKSTRUCT
    pt           As POINTAPI
    hWnd         As LongPtr
    wHitTestCode As Long
    dwExtraInfo  As Long
End Type

' =========================================================
' ▽スクリーンサイズの取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：スクリーンサイズ情報
' =========================================================
Public Function getScreenSize() As ScreenSize

    With getScreenSize
    
        .primarySizeWidth = GetSystemMetrics(SM_CXSCREEN)
        If .primarySizeWidth = 0 Then
            err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                    , "" _
                    , ConstantsError.ERR_DESC_DLL_FUNCTION_FAILED
        End If
        
        .primarySizeHeight = GetSystemMetrics(SM_CYSCREEN)
        If .primarySizeHeight = 0 Then
            err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                    , "" _
                    , ConstantsError.ERR_DESC_DLL_FUNCTION_FAILED
        End If
        
        .virtualSizeX = GetSystemMetrics(SM_XVIRTUALSCREEN)
        
        .virtualSizeY = GetSystemMetrics(SM_YVIRTUALSCREEN)
        
        .virtualSizeWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
        If .virtualSizeWidth = 0 Then
            err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                    , "" _
                    , ConstantsError.ERR_DESC_DLL_FUNCTION_FAILED
        End If
        
        .virtualSizeHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
        If .virtualSizeHeight = 0 Then
            err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                    , "" _
                    , ConstantsError.ERR_DESC_DLL_FUNCTION_FAILED
        End If
    
    
    End With
End Function

' =========================================================
' ▽スクリーンサイズの取得（ポイント単位）
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：スクリーンサイズ情報（ポイント単位）
' =========================================================
Public Function getScreenSizePt() As ScreenSizePt

    ' スクリーンサイズを取得する
    Dim ss As ScreenSize
    ss = getScreenSize
    
    ' システムのDPIを取得する
    Dim systemDPi As DPI
    systemDPi = WinAPI_GDI.getSystemDPI

    ' ピクセルからポイントに情報を変換する
    getScreenSizePt.primarySizeWidth = VBUtil.convertPixelToPoint(systemDPi.horizontal, ss.primarySizeWidth)
    getScreenSizePt.primarySizeHeight = VBUtil.convertPixelToPoint(systemDPi.vertical, ss.primarySizeHeight)
    getScreenSizePt.virtualSizeX = VBUtil.convertPixelToPoint(systemDPi.horizontal, ss.virtualSizeX)
    getScreenSizePt.virtualSizeY = VBUtil.convertPixelToPoint(systemDPi.vertical, ss.virtualSizeY)
    getScreenSizePt.virtualSizeWidth = VBUtil.convertPixelToPoint(systemDPi.horizontal, ss.virtualSizeWidth)
    getScreenSizePt.virtualSizeHeight = VBUtil.convertPixelToPoint(systemDPi.vertical, ss.virtualSizeHeight)

End Function
