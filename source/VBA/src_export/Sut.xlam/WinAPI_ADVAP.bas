Attribute VB_Name = "WinAPI_ADVAP"
Option Explicit

' *********************************************************
' advap32.dllで定義されている関数郡や定数。
'
' 作成者　：Ison
' 履歴　　：2009/04/21　新規作成
'
' *********************************************************

Public Const STANDARD_RIGHTS_READ As Long = &H20000
Public Const STANDARD_RIGHTS_WRITE  As Long = &H20000
Public Const STANDARD_RIGHTS_EXECUTE  As Long = &H20000
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000

Public Const SYNCHRONIZE As Long = &H100000

Public Enum RegKeyConstants
    ' ソフトウェアの設定情報
    HKEY_CLASS_ROOT = &H80000000
    ' 特定のユーザー環境の設定情報
    HKEY_CURRENT_USER = &H80000001
    ' ローカルコンピュータの設定情報
    HKEY_LOCAL_MACHINE = &H80000002
    ' 全てのユーザー環境の設定情報
    HKEY_USERS = &H80000004
    ' 現在のハードウェアの設定情報
    HKEY_CURRNET_CONFIG = &H80000005
    ' デバイスのステータス情報
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum RegOptionConstants

    ' 設定内容をレジストリファイルに保存する
    REG_OPTION_NON_VOLATILE = 0
    ' 設定内容をレジストリファイルに保存しない（OSを再起動すると情報が失われる）
    REG_OPTIONS_VOLATILE = 1
    ' バックアップ
    REG_OPTIONS_BACKUP_RESTORE = 4

End Enum

Public Enum RegAccessConstants
    ' レジストリの値取得の許可
    KEY_QUERY_VALUE = &H1
    ' レジストリの値設定の許可
    KEY_SET_VALUE = &H2
    ' サブキーの作成を許可
    KEY_CREATE_SUB_KEY = &H4
    ' サブキーの列挙を許可
    KEY_ENUMERATE_SUB_KEYS = &H8
    ' レジストリの変更通知を許可
    KEY_NOTIFY = &H10
    ' シンボリックリンクの作成を許可
    KEY_CREATE_LINK = &H20
    ' レジストリのサブキーや値を読み込む
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    ' レジストリのサブキーや値を作成する
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    
    ' レジストリに対する全ての操作を実施する
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
End Enum

' バイナリデータ
Public Const REG_BINARY As Long = 3
' 32ビット数値
Public Const REG_DWORD As Long = 4
' バイトの並び順
Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4
' バイトの並びがWindowsとは逆の32ビット値
Public Const REG_DWORD_BIG_ENDIAN As Long = 5
' 展開前の環境変数(例えば%PATH%)
Public Const REG_EXPAND_SZ As Long = 2
' 二つのvbNullStringで終わる文字列
Public Const REG_MULTI_SZ As Long = 7
' 未定義のタイプ
Public Const REG_NONE As Long = 0
' ドライバのリソースリスト
Public Const REG_RESOUCE_LIST As Long = 8
' 文字列
Public Const REG_SZ As Long = 1

' APIの戻り値（正常時）
Public Const ERROR_SUCCESS As Long = 0

Public Const ERROR_MORE_DATA     As Long = 234
Public Const ERROR_NO_MORE_ITEMS As Long = 259

' RegCreateKeyExの結果（キーを新規作成）
Public Const REG_CREATED_NEW_KEY  As Long = &H1
' RegCreateKeyExの結果（既存キーをオープンする）
Public Const REG_OPENED_EXISTING_KEY As Long = &H2

Public Type FILETIME
    
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

#If VBA7 And Win64 Then
Public Declare PtrSafe Function RegOpenKeyEx Lib "ADVAPI32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As LongPtr _
       , ByVal lpSubKey As String _
       , ByVal ulOptions As Long _
       , ByVal samDesired As Long _
       , ByRef phkResult As LongPtr) As Long

Public Declare PtrSafe Function RegCreateKeyEx Lib "ADVAPI32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey As LongPtr _
       , ByVal lpSubKey As String _
       , ByVal Reserved As Long _
       , ByVal lpClass As String _
       , ByVal dwOptions As Long _
       , ByVal samDesired As Long _
       , ByVal lpSecurityAttributes As LongPtr _
       , ByRef phkResult As LongPtr _
       , ByRef lpdwDisposition As LongPtr) As Long
       
Public Declare PtrSafe Function RegCloseKey Lib "ADVAPI32.dll" _
        (ByVal hKey As LongPtr) As Long
       
Public Declare PtrSafe Function RegEnumKeyEx Lib "ADVAPI32.dll" Alias "RegEnumKeyExA" _
        (ByVal hKey As LongPtr _
       , ByVal dwIndex As Long _
       , ByVal lpName As String _
       , ByRef lpcName As LongPtr _
       , ByVal lpReserved As LongPtr _
       , ByVal lpClass As String _
       , ByRef lpcClass As LongPtr _
       , ByRef lpftLastWriteTime As FILETIME) As Long
       
Public Declare PtrSafe Function RegEnumValue Lib "ADVAPI32.dll" Alias "RegEnumValueA" _
        (ByVal hKey As LongPtr _
       , ByVal dwIndex As Long _
       , ByVal lpValueName As String _
       , ByRef lpCbValueName As LongPtr _
       , ByRef lpReserved As LongPtr _
       , ByRef lpType As LongPtr _
       , ByRef lpData As Any _
       , ByVal lpCbData As LongPtr) As Long
       
Public Declare PtrSafe Function RegEnumValueW Lib "ADVAPI32.dll" _
        (ByVal hKey As LongPtr _
       , ByVal dwIndex As Long _
       , ByVal lpValueName As LongPtr _
       , ByRef lpCbValueName As LongPtr _
       , ByVal lpReserved As LongPtr _
       , ByRef lpType As LongPtr _
       , ByVal lpData As LongPtr _
       , ByRef lpCbData As LongPtr) As Long
       
Public Declare PtrSafe Function RegQueryInfoKey Lib "ADVAPI32.dll" Alias "RegQueryInfoKeyA" _
        (ByVal hKey As LongPtr _
       , ByVal lpClass As String _
       , ByRef lpcbClass As String _
       , ByRef lpReserved As LongPtr _
       , ByRef lpcSubKeys As LongPtr _
       , ByRef lpcbMaxSubKeyLen As LongPtr _
       , ByRef lpcbMaxClassLen As LongPtr _
       , ByRef lpcValues As LongPtr _
       , ByRef lpMaxValueNameLen As LongPtr _
       , ByRef lpcbMaxValueLen As LongPtr _
       , ByVal lpcbSecurityDescriptor As LongPtr _
       , ByRef lpftLastWriteTime As FILETIME) As Long

Public Declare PtrSafe Function RegQueryValueEx Lib "ADVAPI32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As LongPtr _
       , ByVal lpValueName As String _
       , ByVal lpReserved As LongPtr _
       , ByRef lpType As LongPtr _
       , ByRef lpData As Any _
       , ByRef lpCbData As LongPtr) As Long

' UNICODEバージョン
Public Declare PtrSafe Function RegQueryValueExW Lib "ADVAPI32.dll" _
        (ByVal hKey As LongPtr _
       , ByVal lpValueName As LongPtr _
       , ByVal lpReserved As LongPtr _
       , ByRef lpType As LongPtr _
       , ByVal lpData As LongPtr _
       , ByRef lpCbData As LongPtr) As Long

Public Declare PtrSafe Function RegSetValueEx Lib "ADVAPI32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As LongPtr _
       , ByVal lpValueName As String _
       , ByVal Reserved As LongPtr _
       , ByVal dwType As LongPtr _
       , ByRef lpData As Any _
       , ByVal cbdata As Long) As Long

' UNICODEバージョン
Public Declare PtrSafe Function RegSetValueExW Lib "ADVAPI32.dll" _
        (ByVal hKey As LongPtr _
       , ByVal lpValueName As LongPtr _
       , ByVal Reserved As Long _
       , ByVal dwType As Long _
       , ByVal lpData As LongPtr _
       , ByVal cbdata As Long) As Long

Public Declare PtrSafe Function RegDeleteKey Lib "ADVAPI32.dll" Alias "RegDeleteKeyA" _
        (ByVal hKey As LongPtr _
       , ByVal lpSubKey As String) As Long

' UNICODEバージョン
Public Declare PtrSafe Function RegDeleteKeyW Lib "ADVAPI32.dll" _
        (ByVal hKey As LongPtr _
       , ByVal lpSubKey As LongPtr) As Long

Public Declare PtrSafe Function RegDeleteValue Lib "ADVAPI32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As LongPtr _
       , ByVal lpValueName As String) As Long

' UNICODEバージョン
Public Declare PtrSafe Function RegDeleteValueW Lib "ADVAPI32.dll" _
        (ByVal hKey As LongPtr _
       , ByVal lpValueName As LongPtr) As Long
#Else

Public Declare Function RegOpenKeyEx Lib "ADVAPI32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long _
       , ByVal lpSubKey As String _
       , ByVal ulOptions As Long _
       , ByVal samDesired As Long _
       , ByRef phkResult As Long) As Long

Public Declare Function RegCreateKeyEx Lib "ADVAPI32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey As Long _
       , ByVal lpSubKey As String _
       , ByVal Reserved As Long _
       , ByVal lpClass As String _
       , ByVal dwOptions As Long _
       , ByVal samDesired As Long _
       , ByVal lpSecurityAttributes As Long _
       , ByRef phkResult As Long _
       , ByRef lpdwDisposition As Long) As Long
       
Public Declare Function RegCloseKey Lib "ADVAPI32.dll" _
        (ByVal hKey As Long) As Long
       
Public Declare Function RegEnumKeyEx Lib "ADVAPI32.dll" Alias "RegEnumKeyExA" _
        (ByVal hKey As Long _
       , ByVal dwIndex As Long _
       , ByVal lpName As String _
       , ByRef lpcName As Long _
       , ByVal lpReserved As Long _
       , ByVal lpClass As String _
       , ByRef lpcClass As Long _
       , ByRef lpftLastWriteTime As FILETIME) As Long
       
Public Declare Function RegEnumValue Lib "ADVAPI32.dll" Alias "RegEnumValueA" _
        (ByVal hKey As Long _
       , ByVal dwIndex As Long _
       , ByVal lpValueName As String _
       , ByRef lpCbValueName As Long _
       , ByRef lpReserved As Long _
       , ByRef lpType As Long _
       , ByRef lpData As Any _
       , ByVal lpCbData As Long) As Long
       
Public Declare Function RegEnumValueW Lib "ADVAPI32.dll" _
        (ByVal hKey As Long _
       , ByVal dwIndex As Long _
       , ByVal lpValueName As Long _
       , ByRef lpCbValueName As Long _
       , ByVal lpReserved As Long _
       , ByRef lpType As Long _
       , ByVal lpData As Long _
       , ByRef lpCbData As Long) As Long
       
Public Declare Function RegQueryInfoKey Lib "ADVAPI32.dll" Alias "RegQueryInfoKeyA" _
        (ByVal hKey As Long _
       , ByVal lpClass As String _
       , ByRef lpcbClass As String _
       , ByRef lpReserved As Long _
       , ByRef lpcSubKeys As Long _
       , ByRef lpcbMaxSubKeyLen As Long _
       , ByRef lpcbMaxClassLen As Long _
       , ByRef lpcValues As Long _
       , ByRef lpMaxValueNameLen As Long _
       , ByRef lpcbMaxValueLen As Long _
       , ByVal lpcbSecurityDescriptor As Long _
       , ByRef lpftLastWriteTime As FILETIME) As Long

Public Declare Function RegQueryValueEx Lib "ADVAPI32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long _
       , ByVal lpValueName As String _
       , ByVal lpReserved As Long _
       , ByRef lpType As Long _
       , ByRef lpData As Any _
       , ByRef lpCbData As Long) As Long

' UNICODEバージョン
Public Declare Function RegQueryValueExW Lib "ADVAPI32.dll" _
        (ByVal hKey As Long _
       , ByVal lpValueName As Long _
       , ByVal lpReserved As Long _
       , ByRef lpType As Long _
       , ByVal lpData As Long _
       , ByRef lpCbData As Long) As Long

Public Declare Function RegSetValueEx Lib "ADVAPI32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long _
       , ByVal lpValueName As String _
       , ByVal Reserved As Long _
       , ByVal dwType As Long _
       , ByRef lpData As Any _
       , ByVal cbdata As Long) As Long

' UNICODEバージョン
Public Declare Function RegSetValueExW Lib "ADVAPI32.dll" _
        (ByVal hKey As Long _
       , ByVal lpValueName As Long _
       , ByVal Reserved As Long _
       , ByVal dwType As Long _
       , ByVal lpData As Long _
       , ByVal cbdata As Long) As Long

Public Declare Function RegDeleteKey Lib "ADVAPI32.dll" Alias "RegDeleteKeyA" _
        (ByVal hKey As Long _
       , ByVal lpSubKey As String) As Long

' UNICODEバージョン
Public Declare Function RegDeleteKeyW Lib "ADVAPI32.dll" _
        (ByVal hKey As Long _
       , ByVal lpSubKey As Long) As Long

Public Declare Function RegDeleteValue Lib "ADVAPI32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As Long _
       , ByVal lpValueName As String) As Long

' UNICODEバージョン
Public Declare Function RegDeleteValueW Lib "ADVAPI32.dll" _
        (ByVal hKey As Long _
       , ByVal lpValueName As Long) As Long
#End If
