Attribute VB_Name = "WinAPI_OLEACC"
Option Explicit

#If VBA7 And Win64 Then
    Public Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" _
      (ByVal IAcessible As Object, ByRef hwnd As LongPtr) As Long
#Else
    Public Declare Function WindowFromAccessibleObject Lib "oleacc.dll" _
      (ByVal IAcessible As Object, ByRef hWnd As Long) As Long
#End If

