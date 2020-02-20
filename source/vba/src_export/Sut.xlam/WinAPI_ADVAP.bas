Attribute VB_Name = "WinAPI_ADVAP"
Option Explicit

' *********************************************************
' advap32.dll�Œ�`����Ă���֐��S��萔�B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/21�@�V�K�쐬
'
' *********************************************************

Public Const STANDARD_RIGHTS_READ As Long = &H20000
Public Const STANDARD_RIGHTS_WRITE  As Long = &H20000
Public Const STANDARD_RIGHTS_EXECUTE  As Long = &H20000
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000

Public Const SYNCHRONIZE As Long = &H100000

Public Enum RegKeyConstants
    ' �\�t�g�E�F�A�̐ݒ���
    HKEY_CLASS_ROOT = &H80000000
    ' ����̃��[�U�[���̐ݒ���
    HKEY_CURRENT_USER = &H80000001
    ' ���[�J���R���s���[�^�̐ݒ���
    HKEY_LOCAL_MACHINE = &H80000002
    ' �S�Ẵ��[�U�[���̐ݒ���
    HKEY_USERS = &H80000004
    ' ���݂̃n�[�h�E�F�A�̐ݒ���
    HKEY_CURRNET_CONFIG = &H80000005
    ' �f�o�C�X�̃X�e�[�^�X���
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum RegOptionConstants

    ' �ݒ���e�����W�X�g���t�@�C���ɕۑ�����
    REG_OPTION_NON_VOLATILE = 0
    ' �ݒ���e�����W�X�g���t�@�C���ɕۑ����Ȃ��iOS���ċN������Ə�񂪎�����j
    REG_OPTIONS_VOLATILE = 1
    ' �o�b�N�A�b�v
    REG_OPTIONS_BACKUP_RESTORE = 4

End Enum

Public Enum RegAccessConstants
    ' ���W�X�g���̒l�擾�̋���
    KEY_QUERY_VALUE = &H1
    ' ���W�X�g���̒l�ݒ�̋���
    KEY_SET_VALUE = &H2
    ' �T�u�L�[�̍쐬������
    KEY_CREATE_SUB_KEY = &H4
    ' �T�u�L�[�̗񋓂�����
    KEY_ENUMERATE_SUB_KEYS = &H8
    ' ���W�X�g���̕ύX�ʒm������
    KEY_NOTIFY = &H10
    ' �V���{���b�N�����N�̍쐬������
    KEY_CREATE_LINK = &H20
    ' ���W�X�g���̃T�u�L�[��l��ǂݍ���
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    ' ���W�X�g���̃T�u�L�[��l���쐬����
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    
    ' ���W�X�g���ɑ΂���S�Ă̑�������{����
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
End Enum

' �o�C�i���f�[�^
Public Const REG_BINARY As Long = 3
' 32�r�b�g���l
Public Const REG_DWORD As Long = 4
' �o�C�g�̕��я�
Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4
' �o�C�g�̕��т�Windows�Ƃ͋t��32�r�b�g�l
Public Const REG_DWORD_BIG_ENDIAN As Long = 5
' �W�J�O�̊��ϐ�(�Ⴆ��%PATH%)
Public Const REG_EXPAND_SZ As Long = 2
' ���vbNullString�ŏI��镶����
Public Const REG_MULTI_SZ As Long = 7
' ����`�̃^�C�v
Public Const REG_NONE As Long = 0
' �h���C�o�̃��\�[�X���X�g
Public Const REG_RESOUCE_LIST As Long = 8
' ������
Public Const REG_SZ As Long = 1

' API�̖߂�l�i���펞�j
Public Const ERROR_SUCCESS As Long = 0

Public Const ERROR_MORE_DATA     As Long = 234
Public Const ERROR_NO_MORE_ITEMS As Long = 259

' RegCreateKeyEx�̌��ʁi�L�[��V�K�쐬�j
Public Const REG_CREATED_NEW_KEY  As Long = &H1
' RegCreateKeyEx�̌��ʁi�����L�[���I�[�v������j
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

' UNICODE�o�[�W����
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

' UNICODE�o�[�W����
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

' UNICODE�o�[�W����
Public Declare PtrSafe Function RegDeleteKeyW Lib "ADVAPI32.dll" _
        (ByVal hKey As LongPtr _
       , ByVal lpSubKey As LongPtr) As Long

Public Declare PtrSafe Function RegDeleteValue Lib "ADVAPI32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As LongPtr _
       , ByVal lpValueName As String) As Long

' UNICODE�o�[�W����
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

' UNICODE�o�[�W����
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

' UNICODE�o�[�W����
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

' UNICODE�o�[�W����
Public Declare Function RegDeleteKeyW Lib "ADVAPI32.dll" _
        (ByVal hKey As Long _
       , ByVal lpSubKey As Long) As Long

Public Declare Function RegDeleteValue Lib "ADVAPI32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As Long _
       , ByVal lpValueName As String) As Long

' UNICODE�o�[�W����
Public Declare Function RegDeleteValueW Lib "ADVAPI32.dll" _
        (ByVal hKey As Long _
       , ByVal lpValueName As Long) As Long
#End If
