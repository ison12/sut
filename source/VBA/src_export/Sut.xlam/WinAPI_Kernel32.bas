Attribute VB_Name = "WinAPI_Kernel32"
Option Explicit

#If VBA7 And Win64 Then

' ���C�u���������[�h����
Public Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
  ByVal lpLibFileName As String) As LongPtr

' ���C�u�������������
Public Declare PtrSafe Function freeLibrary Lib "kernel32" Alias _
  "FreeLibrary" (ByVal hLibModule As LongPtr) As Long
  
' ���C�u�����̃n���h�����擾����
Public Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
  ByVal lpModuleName As String) As LongPtr

' �֐��̃A�h���X���擾����
Public Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
  ByVal hModule As LongPtr _
, ByVal lpProcName As String) As LongPtr

Public Declare PtrSafe Function CreateEvent Lib "kernel32" Alias "CreateEventA" ( _
  ByVal LpEventAttributes As LongPtr _
, ByVal bManualReset As Long _
, ByVal bInitiaLState As Long _
, ByVal lpName As String) As Long

Public Declare PtrSafe Function SetEvent Lib "kernel32" ( _
  ByVal hEvent As LongPtr) As Long

Public Declare PtrSafe Function ResetEvent Lib "kernel32" ( _
  ByVal hEvent As LongPtr) As Long

Public Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
  ByVal hHandle As LongPtr _
, ByVal dwMilliseconds As Long) As Long

#Else

' ���C�u���������[�h����
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
  ByVal lpLibFileName As String) As Long

' ���C�u�������������
Public Declare Function freeLibrary Lib "kernel32" Alias _
  "FreeLibrary" (ByVal hLibModule As Long) As Long
  
' ���C�u�����̃n���h�����擾����
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
  ByVal lpModuleName As String) As Long

' �֐��̃A�h���X���擾����
Public Declare Function GetProcAddress Lib "kernel32" ( _
  ByVal hModule As Long _
, ByVal lpProcName As String) As Long

Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" ( _
  ByVal LpEventAttributes As Long _
, ByVal bManualReset As Long _
, ByVal bInitiaLState As Long _
, ByVal lpName As String) As Long

Public Declare Function SetEvent Lib "kernel32" ( _
  ByVal hEvent As Long) As Long

Public Declare Function ResetEvent Lib "kernel32" ( _
  ByVal hEvent As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" ( _
  ByVal hHandle As Long _
, ByVal dwMilliseconds As Long) As Long

#End If
