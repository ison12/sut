Attribute VB_Name = "ScrollableControl"
Option Explicit

' *********************************************************
' マウスホイールによるコントロールのスクロール化
'
' 作成者　：Hideki Isobe
' 履歴　　：2020/01/17　新規作成
'
' 特記事項：
' *********************************************************

Private scrollableControl As Object
Private isHooked          As Boolean
Private mouseHookHandle   As LongPtr
Private targetHwnd        As LongPtr

' =========================================================
' ▽リストボックスのスクロールイベントをフックする。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub hookScroll(frm As Object, ctl As Object)

    Dim tPT As WinAPI_User.POINTAPI
    
    Dim lngAppInst      As LongPtr
    Dim hwndUnderCursor As LongPtr
        
    WinAPI_User.GetCursorPos tPT
    
    #If VBA7 And Win64 Then
        hwndUnderCursor = WinAPI_User.WindowFromPoint(tPT.XY)
    #Else
        hwndUnderCursor = WinAPI_User.WindowFromPoint(tPT.X, tPT.Y)
    #End If
    
    If TypeOf ctl Is UserForm Then
        If Not frm Is ctl Then
            ctl.SetFocus
        End If
    Else
        If Not frm.ActiveControl Is ctl Then
            ctl.SetFocus
        End If
    End If
    
    If targetHwnd <> hwndUnderCursor Then
    
        unhookScroll
        
        Set scrollableControl = ctl
        targetHwnd = hwndUnderCursor
        #If VBA7 And Win64 Then
            lngAppInst = WinAPI_User.GetWindowLongPtr(targetHwnd, WinAPI_User.GWL_HINSTANCE)
        #Else
            lngAppInst = WinAPI_User.GetWindowLong(targetHwnd, WinAPI_User.GWL_HINSTANCE)
        #End If
         
        If Not isHooked Then
            mouseHookHandle = WinAPI_User.SetWindowsHookEx(WinAPI_User.WH_MOUSE_LL, AddressOf mouseProc, lngAppInst, 0)
            isHooked = mouseHookHandle <> 0
        End If
    End If
End Sub

' =========================================================
' ▽リストボックスのスクロールイベントフックを解除する。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub unhookScroll()

    If mouseHookHandle Then
    
        Set scrollableControl = Nothing
        
        WinAPI_User.UnhookWindowsHookEx mouseHookHandle
        
        mouseHookHandle = 0
        targetHwnd = 0
        isHooked = False
        
    End If
    
End Sub

' =========================================================
' ▽マウスプロシージャ
'
' 概要　　　：
' 引数　　　：nCode  コード番号
' 　　    　  wParam パラメータ
'     　　　  lParam パラメータ
' 戻り値　　：次のマウスプロシージャ
'
' =========================================================
Private Function mouseProc( _
    ByVal nCode As Long, _
    ByVal wParam As Long, _
    ByRef lParam As WinAPI_User.MOUSEHOOKSTRUCT) As LongPtr
    
    On Error GoTo err

    Dim idx As Long
    
    Dim hWnd As LongPtr
    
    If nCode = WinAPI_User.HC_ACTION Then
    
        #If VBA7 And Win64 Then
            hWnd = WinAPI_User.WindowFromPoint(lParam.pt.XY)
        #Else
            hWnd = WinAPI_User.WindowFromPoint(lParam.pt.X, lParam.pt.Y)
        #End If

        If hWnd = targetHwnd Then
        
            If wParam = WinAPI_User.WM_MOUSEWHEEL Then
                
                mouseProc = True

                If TypeOf scrollableControl Is Frame Then
                    
                    If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                    idx = idx + scrollableControl.ScrollTop
                    If idx >= 0 And idx < ((scrollableControl.ScrollHeight - scrollableControl.Height) + 17.25) Then
                        scrollableControl.ScrollTop = idx
                    
                    End If
                ElseIf TypeOf scrollableControl Is UserForm Then
                    
                    If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                    idx = idx + scrollableControl.ScrollTop
                    If idx >= 0 And idx < ((scrollableControl.ScrollHeight - scrollableControl.Height) + 17.25) Then
                        scrollableControl.ScrollTop = idx
                    
                    End If
                Else
                    
                    If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                    idx = idx + scrollableControl.ListIndex
                    If idx >= 0 Then scrollableControl.ListIndex = idx
                    
                End If
                                
                Exit Function
                
            End If
        
        Else
            unhookScroll
        End If
        
    End If
     
    mouseProc = WinAPI_User.CallNextHookEx(targetHwnd, nCode, wParam, ByVal lParam)
     
    Exit Function
    
err:
    unhookScroll
     
End Function
