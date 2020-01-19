Attribute VB_Name = "WinAPI_ODBC"
Option Explicit

' *********************************************************
' user32.dll�Œ�`����Ă���֐��S��萔�B
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2008/02/11�@�V�K�쐬
'
' ���L�����FWindowsAPI�𗘗p���ăf�[�^�\�[�X���ɃA�N�Z�X����
' *********************************************************

' =========================================================
' ��ODBC���ϐ��擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' �߂�l�@�@�F
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SQLAllocEnv Lib "odbc32.dll" (ByRef phEnv As LongPtr) As Integer
#Else
    Public Declare Function SQLAllocEnv Lib "odbc32.dll" (ByRef phEnv As Long) As Integer
#End If

' =========================================================
' ��ODBC���ϐ����
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' �߂�l�@�@�F
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As LongPtr) As Integer
#Else
    Public Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer
#End If

' =========================================================
' ��ODBC�f�[�^�\�[�X�擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' �߂�l�@�@�F
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SQLDataSources Lib "odbc32.dll" Alias "SQLDataSourcesA" _
       (ByVal hEnv As LongPtr, _
        ByVal fDirection As Integer, _
        ByVal dataSourceName As String, _
        ByVal dataSourceNameMax As Integer, _
        ByRef dataSourceNameLength As Integer, _
        ByVal dataSourceDesc As String, _
        ByVal dataSourceDescMax As Integer, _
        ByRef dataSourceDescLength As Integer) As Integer
#Else
    Public Declare Function SQLDataSources Lib "odbc32.dll" Alias "SQLDataSourcesA" _
       (ByVal hEnv As Long, _
        ByVal fDirection As Integer, _
        ByVal dataSourceName As String, _
        ByVal dataSourceNameMax As Integer, _
        ByRef dataSourceNameLength As Integer, _
        ByVal dataSourceDesc As String, _
        ByVal dataSourceDescMax As Integer, _
        ByRef dataSourceDescLength As Integer) As Integer
#End If

Public Const DATA_SOURCE_INFO_KEY_NAME As String = "name"
Public Const DATA_SOURCE_INFO_KEY_DESC As String = "description"

Public Const SQL_SUCCESS           As Long = 0
Public Const SQL_SUCCESS_WITH_INFO As Long = 1
Public Const SQL_NO_DATA_FOUND     As Long = 100

Public Const SQL_FETCH_FIRST       As Long = 2
Public Const SQL_FETCH_NEXT        As Long = 1

Public Const ERROR_CODE As Long = 1500
Public Const ERROR_MSG  As String = "�f�[�^�\�[�X�擾���ɃG���[���������܂����B"

' =========================================================
' ���f�[�^�\�[�X���X�g�擾
'
' �T�v�@�@�@�F�f�[�^�\�[�X���X�g���擾����
' �߂�l�@�@�F�f�[�^�\�[�X���X�g���i�[���ꂽ�ڸ��ݵ�޼ު��
'
'             �v�f�ɂ̓f�[�^�\�[�X��񂪊i�[���ꂽ�ڸ��ݵ�޼ު��
'             ���ȉ��̌`���Ŋi�[�����y�L�[�z���y�l�z
'             name = �f�[�^�\�[�X��
' �@�@�@�@�@�@desc = �f�[�^�\�[�X����
'
' =========================================================
Public Function getDataSourceList() As ValCollection
    
    ' ���߂�l
    Dim dataSourceList As ValCollection
    Dim dataSourceInfo As ValCollection
    
#If VBA7 And Win64 Then
    Dim hEnv            As LongPtr
#Else
    Dim hEnv            As Long
#End If
    Dim szDSN           As String * 256
    Dim cbDSN           As Integer
    Dim szDescription   As String * 256
    Dim cbDescription   As Integer
    
    Dim retCode         As Integer
    
    On Error GoTo err
    
    retCode = SQLAllocEnv(hEnv)
    
    ' �G���[����
    If retCode < 0 Then
    
        err.Raise ERROR_CODE, err.Source, ERROR_MSG
    End If
    
    ' ���f�[�^�\�[�X���i�[���X�g������������
    Set dataSourceList = New ValCollection
    
    Do While True
    
        ' �f�[�^�\�[�X�����擾����
        retCode = SQLDataSources( _
                                hEnv, _
                                SQL_FETCH_NEXT, _
                                szDSN, _
                                256, _
                                cbDSN, _
                                szDescription, _
                                256, _
                                cbDescription)
                            
        If retCode < 0 Or retCode = SQL_NO_DATA_FOUND Then
        
            GoTo loop_end
        End If
                            
        ' ���f�[�^�\�[�X��������������
        Set dataSourceInfo = New ValCollection
        
        ' �f�[�^�\�[�X����ݒ肷��
        dataSourceInfo.setItem LeftB(szDSN, InStrB(szDSN, Chr$(0))), DATA_SOURCE_INFO_KEY_NAME
        dataSourceInfo.setItem LeftB(szDescription, InStrB(szDescription, Chr$(0))), DATA_SOURCE_INFO_KEY_DESC
        
        ' �f�[�^�\�[�X���X�g�ɏ���ݒ肷��
        dataSourceList.setItem dataSourceInfo
    Loop
    
loop_end:

    ' �G���[����
    If retCode < 0 Then
    
        err.Raise ERROR_CODE, err.Source, ERROR_MSG
    End If
            
    retCode = SQLFreeEnv(hEnv)

    Set getDataSourceList = dataSourceList

    Exit Function
err:

    retCode = SQLFreeEnv(hEnv)

    err.Raise ERROR_CODE, err.Source, ERROR_MSG

End Function

