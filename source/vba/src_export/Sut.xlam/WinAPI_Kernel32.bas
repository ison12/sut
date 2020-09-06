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


'��INI�t�@�C���p��WinAPI�֐��錾
#If VBA7 And Win64 Then
    
    ' INI�t�@�C���̎w�肵���Z�N�V�������̂��ׂẴL�[�ƒl���擾
    Public Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" _
        Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long
    
    ' INI�t�@�C���̎w�肵���Z�N�V�������̂��ׂẴL�[�ƒl��ݒ�
    Public Declare PtrSafe Function WritePrivateProfileSection Lib "kernel32" _
        Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, _
         ByVal lpString As String, _
         ByVal lpFileName As String) As Long
    
    ' INI�t�@�C���̕�������擾
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long
    
    ' INI�t�@�C���̕������ύX
    Public Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpString As Any, _
         ByVal lpFileName As String) As Long

#Else
    
    ' INI�t�@�C���̎w�肵���Z�N�V�������̂��ׂẴL�[�ƒl���擾
    Public Declare Function GetPrivateProfileSection Lib "kernel32" _
        Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long
    
    ' INI�t�@�C���̎w�肵���Z�N�V�������̂��ׂẴL�[�ƒl��ݒ�
    Public Declare Function WritePrivateProfileSection Lib "kernel32" _
        Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, _
         ByVal lpString As String, _
         ByVal lpFileName As String) As Long
    
    ' INI�t�@�C���̕�������擾
    Public Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long
    
    ' INI�t�@�C���̕������ύX
    Public Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpString As Any, _
         ByVal lpFileName As String) As Long

#End If

