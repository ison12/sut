Attribute VB_Name = "WinAPI_Kernel32"
Option Explicit

#If VBA7 And Win64 Then
    
    ' ライブラリをロードする
    Public Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
      ByVal lpLibFileName As String) As LongPtr
    
    ' ライブラリを解放する
    Public Declare PtrSafe Function freeLibrary Lib "kernel32" Alias _
      "FreeLibrary" (ByVal hLibModule As LongPtr) As Long
      
    ' ライブラリのハンドルを取得する
    Public Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
      ByVal lpModuleName As String) As LongPtr
    
    ' 関数のアドレスを取得する
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
    
    ' ライブラリをロードする
    Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
      ByVal lpLibFileName As String) As Long
    
    ' ライブラリを解放する
    Public Declare Function freeLibrary Lib "kernel32" Alias _
      "FreeLibrary" (ByVal hLibModule As Long) As Long
      
    ' ライブラリのハンドルを取得する
    Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
      ByVal lpModuleName As String) As Long
    
    ' 関数のアドレスを取得する
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


'▼INIファイル用のWinAPI関数宣言
#If VBA7 And Win64 Then
    
    ' INIファイルの指定したセクション内のすべてのキーと値を取得
    Public Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" _
        Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long
    
    ' INIファイルの指定したセクション内のすべてのキーと値を設定
    Public Declare PtrSafe Function WritePrivateProfileSection Lib "kernel32" _
        Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, _
         ByVal lpString As String, _
         ByVal lpFileName As String) As Long
    
    ' INIファイルの文字列を取得
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long
    
    ' INIファイルの文字列を変更
    Public Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpString As Any, _
         ByVal lpFileName As String) As Long

#Else
    
    ' INIファイルの指定したセクション内のすべてのキーと値を取得
    Public Declare Function GetPrivateProfileSection Lib "kernel32" _
        Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long
    
    ' INIファイルの指定したセクション内のすべてのキーと値を設定
    Public Declare Function WritePrivateProfileSection Lib "kernel32" _
        Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, _
         ByVal lpString As String, _
         ByVal lpFileName As String) As Long
    
    ' INIファイルの文字列を取得
    Public Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long
    
    ' INIファイルの文字列を変更
    Public Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpString As Any, _
         ByVal lpFileName As String) As Long

#End If

