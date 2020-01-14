Attribute VB_Name = "SutGray"
Option Explicit
' *********************************************************
' SutGray.dll�֘A�̃��W���[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2009/05/18�@�V�K�쐬
'
' ���L�����F
' *********************************************************

#If (DEBUG_MODE = 1) Then
    
    #If VBA7 And Win64 Then
        Public Declare PtrSafe Function Encrypt Lib ".\..\CPP\Sut\x64\Debug ASM\SutGray.dll" _
            (ByVal password As String _
           , ByVal passwordLen As Long _
           , ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
    
        Public Declare PtrSafe Function Decrypt Lib ".\..\CPP\Sut\x64\Debug ASM\SutGray.dll" _
            (ByVal password As String _
           , ByVal passwordLen As Long _
           , ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
                                                                                                            
        Public Declare PtrSafe Function ConvertBinaryDataToHex Lib ".\..\CPP\Sut\x64\Debug ASM\SutGray.dll" _
            (ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByVal resultBuffer As String _
           , ByRef resultBufferLen As Long) As Long
                                                                                                            
        Public Declare PtrSafe Function ConvertHexToBinaryData Lib ".\..\CPP\Sut\x64\Debug ASM\SutGray.dll" _
            (ByVal buffer As String _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
    #Else
        Public Declare Function Encrypt Lib ".\..\CPP\Sut\Debug ASM\SutGray.dll" _
            (ByVal password As String _
           , ByVal passwordLen As Long _
           , ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
    
        Public Declare Function Decrypt Lib ".\..\CPP\Sut\Debug ASM\SutGray.dll" _
            (ByVal password As String _
           , ByVal passwordLen As Long _
           , ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
                                                                                                            
        Public Declare Function ConvertBinaryDataToHex Lib ".\..\CPP\Sut\Debug ASM\SutGray.dll" _
            (ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByVal resultBuffer As String _
           , ByRef resultBufferLen As Long) As Long
                                                                                                            
        Public Declare Function ConvertHexToBinaryData Lib ".\..\CPP\Sut\Debug ASM\SutGray.dll" _
            (ByVal buffer As String _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
    #End If

#Else

    #If VBA7 And Win64 Then
        Public Declare PtrSafe Function Encrypt Lib "lib\SutGray.dll" _
            (ByVal password As String _
           , ByVal passwordLen As Long _
           , ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
    
        Public Declare PtrSafe Function Decrypt Lib "lib\SutGray.dll" _
            (ByVal password As String _
           , ByVal passwordLen As Long _
           , ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
                                                                                                            
        Public Declare PtrSafe Function ConvertBinaryDataToHex Lib "lib\SutGray.dll" _
            (ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByVal resultBuffer As String _
           , ByRef resultBufferLen As Long) As Long
                                                                                                            
        Public Declare PtrSafe Function ConvertHexToBinaryData Lib "lib\SutGray.dll" _
            (ByVal buffer As String _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
    #Else
        Public Declare Function Encrypt Lib "lib\SutGray.dll" _
            (ByVal password As String _
           , ByVal passwordLen As Long _
           , ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
    
        Public Declare Function Decrypt Lib "lib\SutGray.dll" _
            (ByVal password As String _
           , ByVal passwordLen As Long _
           , ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
                                                                                                            
        Public Declare Function ConvertBinaryDataToHex Lib "lib\SutGray.dll" _
            (ByRef buffer As Byte _
           , ByVal bufferLen As Long _
           , ByVal resultBuffer As String _
           , ByRef resultBufferLen As Long) As Long
                                                                                                            
        Public Declare Function ConvertHexToBinaryData Lib "lib\SutGray.dll" _
            (ByVal buffer As String _
           , ByVal bufferLen As Long _
           , ByRef resultBuffer As Byte _
           , ByRef resultBufferLen As Long) As Long
    #End If

#End If

' =========================================================
' ���Í������s
'
' �T�v�@�@�@�FEncryptWrapper�̃��b�p�[�֐�
' �����@�@�@�Fpassword �p�X���[�h�i�Z�b�V�������̃L�[�ƂȂ�l�j
' �@�@�@�@�@�@buffer �o�b�t�@
' �߂�l�@�@�F�Í����f�[�^
'
' =========================================================
Public Function EncryptWrapper(ByVal password As String _
                              , ByRef buffer() As Byte) As Byte()
                              
    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook

    ' API�̖߂�l���󂯎��ϐ�
    ' �o�b�t�@
    Dim resultBuffer() As Byte
    ' �o�b�t�@�̒���
    Dim resultLen      As Long
    
    ' �o�b�t�@�̒������擾����
    resultLen = 0 ' 0��ݒ肵���������擾����
    SutGray.Encrypt password _
                    , Len(password) _
                    , buffer(0) _
                    , VBUtil.arraySize(buffer) _
                    , 0 _
                    , resultLen

    If resultLen = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If

    ' �߂�l���i�[����o�b�t�@���m�ۂ���
    ReDim resultBuffer(0 To resultLen - 1)

    ' �Í��������s����
    If SutGray.Encrypt(password _
                      , Len(password) _
                      , buffer(0) _
                      , VBUtil.arraySize(buffer) _
                      , resultBuffer(0) _
                      , resultLen) = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If
    
    ' �߂�l��ݒ肷��
    EncryptWrapper = resultBuffer

End Function

' =========================================================
' �����������s
'
' �T�v�@�@�@�FDecrypt�̃��b�p�[�֐�
' �����@�@�@�Fpassword �p�X���[�h�i�Z�b�V�������̃L�[�ƂȂ�l�j
' �@�@�@�@�@�@buffer �o�b�t�@
' �߂�l�@�@�F�������f�[�^
'
' =========================================================
Public Function DecryptWrapper(ByVal password As String _
                              , ByRef buffer() As Byte) As Byte()

    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ' API�̖߂�l���󂯎��ϐ�
    ' �o�b�t�@
    Dim resultBuffer() As Byte
    ' �o�b�t�@�̒���
    Dim resultLen      As Long
    
    ' �o�b�t�@�̒������擾����
    resultLen = 0 ' 0��ݒ肵���������擾����
    SutGray.Decrypt password _
                    , Len(password) _
                    , buffer(0) _
                    , VBUtil.arraySize(buffer) _
                    , 0 _
                    , resultLen
    
    If resultLen = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If

    ' �߂�l���i�[����o�b�t�@���m�ۂ���
    ReDim resultBuffer(0 To resultLen - 1)

    ' �Í��������s����
    If SutGray.Decrypt(password _
                      , Len(password) _
                      , buffer(0) _
                      , VBUtil.arraySize(buffer) _
                      , resultBuffer(0) _
                      , resultLen) = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If
    
    ' �߂�l��ݒ肷��
    DecryptWrapper = resultBuffer

End Function

' =========================================================
' ��16�i�������񁨃o�C�i���f�[�^
'
' �T�v�@�@�@�F16�i�������񂩂�o�C�i���f�[�^�i�o�C�g�z��j�ɕϊ�����
' �����@�@�@�Fbuffer �o�b�t�@
' �߂�l�@�@�F�o�C�i���f�[�^
'
' =========================================================
Public Function ConvertHexToBinaryDataWrapper(ByRef buffer As String) As Byte()

    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ' API�̖߂�l���󂯎��ϐ�
    ' �o�b�t�@
    Dim resultBuffer() As Byte
    ' �o�b�t�@��
    Dim resultLen      As Long
    
    ' �o�b�t�@�̒������擾����
    resultLen = 0 ' 0��ݒ肵���������擾����
    SutGray.ConvertHexToBinaryData buffer _
                                , Len(buffer) _
                                , 0 _
                                , resultLen
    
    If resultLen = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If

    ' �߂�l���i�[����o�b�t�@���m�ۂ���
    ReDim resultBuffer(0 To resultLen - 1)

    ' �Í��������s����
    If SutGray.ConvertHexToBinaryData(buffer _
                                    , Len(buffer) _
                                    , resultBuffer(0) _
                                    , resultLen) = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If


    ' �߂�l��ݒ肷��
    ConvertHexToBinaryDataWrapper = resultBuffer

End Function

' =========================================================
' ���o�C�i���f�[�^��16�i��������
'
' �T�v�@�@�@�F�o�C�i���f�[�^�i�o�C�g�z��j����16�i��������ɕϊ�����
' �����@�@�@�Fbuffer �o�b�t�@
' �߂�l�@�@�F16�i��������
'
' =========================================================
Public Function ConvertBinaryDataToHexWrapper(ByRef buffer() As Byte) As String

    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ' API�̖߂�l���󂯎��ϐ�
    ' �o�b�t�@
    Dim resultBuffer   As String
    ' �o�b�t�@��
    Dim resultLen      As Long
    
    ' �o�b�t�@�̒������擾����
    resultLen = 0 ' 0��ݒ肵���������擾����
    SutGray.ConvertBinaryDataToHex buffer(0) _
                                    , VBUtil.arraySize(buffer) _
                                    , 0 _
                                    , resultLen
    
    If resultLen = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If

    ' �߂�l���i�[����o�b�t�@���m�ۂ���
    resultBuffer = Space(resultLen)

    ' �����������s����
    If SutGray.ConvertBinaryDataToHex(buffer(0) _
                                    , VBUtil.arraySize(buffer) _
                                    , resultBuffer _
                                    , resultLen) = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If

    ' �߂�l��ݒ肷��
    ConvertBinaryDataToHexWrapper = resultBuffer
    
End Function


