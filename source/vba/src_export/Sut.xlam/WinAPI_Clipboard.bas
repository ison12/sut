Attribute VB_Name = "WinAPI_Clipboard"
Option Explicit


#If VBA7 And Win64 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long
#Else
    Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
    Private Declare Function CloseClipboard Lib "user32.dll" () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
    Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
    
#End If

Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As LongPtr
    Dim iLen As Long
    Dim iLock As LongPtr
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Public Function GetClipboard() As String
    Dim iStrPtr As LongPtr
    Dim iLen As LongPtr
    Dim iLock As LongPtr
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(CLng(iLen) \ 2& - 1&, vbNullChar)
            lstrcpy StrPtr(sUniText), iLock
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If
    CloseClipboard
End Function

