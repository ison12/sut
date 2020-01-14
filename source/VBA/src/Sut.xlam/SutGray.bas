Attribute VB_Name = "SutGray"
Option Explicit
' *********************************************************
' SutGray.dll関連のモジュール
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/05/18　新規作成
'
' 特記事項：
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
' ▽暗号化実行
'
' 概要　　　：EncryptWrapperのラッパー関数
' 引数　　　：password パスワード（セッション鍵のキーとなる値）
' 　　　　　　buffer バッファ
' 戻り値　　：暗号化データ
'
' =========================================================
Public Function EncryptWrapper(ByVal password As String _
                              , ByRef buffer() As Byte) As Byte()
                              
    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook

    ' APIの戻り値を受け取る変数
    ' バッファ
    Dim resultBuffer() As Byte
    ' バッファの長さ
    Dim resultLen      As Long
    
    ' バッファの長さを取得する
    resultLen = 0 ' 0を設定し長さだけ取得する
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

    ' 戻り値を格納するバッファを確保する
    ReDim resultBuffer(0 To resultLen - 1)

    ' 暗号化を実行する
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
    
    ' 戻り値を設定する
    EncryptWrapper = resultBuffer

End Function

' =========================================================
' ▽複合化実行
'
' 概要　　　：Decryptのラッパー関数
' 引数　　　：password パスワード（セッション鍵のキーとなる値）
' 　　　　　　buffer バッファ
' 戻り値　　：複合化データ
'
' =========================================================
Public Function DecryptWrapper(ByVal password As String _
                              , ByRef buffer() As Byte) As Byte()

    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ' APIの戻り値を受け取る変数
    ' バッファ
    Dim resultBuffer() As Byte
    ' バッファの長さ
    Dim resultLen      As Long
    
    ' バッファの長さを取得する
    resultLen = 0 ' 0を設定し長さだけ取得する
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

    ' 戻り値を格納するバッファを確保する
    ReDim resultBuffer(0 To resultLen - 1)

    ' 暗号化を実行する
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
    
    ' 戻り値を設定する
    DecryptWrapper = resultBuffer

End Function

' =========================================================
' ▽16進数文字列→バイナリデータ
'
' 概要　　　：16進数文字列からバイナリデータ（バイト配列）に変換する
' 引数　　　：buffer バッファ
' 戻り値　　：バイナリデータ
'
' =========================================================
Public Function ConvertHexToBinaryDataWrapper(ByRef buffer As String) As Byte()

    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ' APIの戻り値を受け取る変数
    ' バッファ
    Dim resultBuffer() As Byte
    ' バッファ長
    Dim resultLen      As Long
    
    ' バッファの長さを取得する
    resultLen = 0 ' 0を設定し長さだけ取得する
    SutGray.ConvertHexToBinaryData buffer _
                                , Len(buffer) _
                                , 0 _
                                , resultLen
    
    If resultLen = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If

    ' 戻り値を格納するバッファを確保する
    ReDim resultBuffer(0 To resultLen - 1)

    ' 暗号化を実行する
    If SutGray.ConvertHexToBinaryData(buffer _
                                    , Len(buffer) _
                                    , resultBuffer(0) _
                                    , resultLen) = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If


    ' 戻り値を設定する
    ConvertHexToBinaryDataWrapper = resultBuffer

End Function

' =========================================================
' ▽バイナリデータ→16進数文字列
'
' 概要　　　：バイナリデータ（バイト配列）から16進数文字列に変換する
' 引数　　　：buffer バッファ
' 戻り値　　：16進数文字列
'
' =========================================================
Public Function ConvertBinaryDataToHexWrapper(ByRef buffer() As Byte) As String

    ' ディレクトリを一時的に変更する
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ' APIの戻り値を受け取る変数
    ' バッファ
    Dim resultBuffer   As String
    ' バッファ長
    Dim resultLen      As Long
    
    ' バッファの長さを取得する
    resultLen = 0 ' 0を設定し長さだけ取得する
    SutGray.ConvertBinaryDataToHex buffer(0) _
                                    , VBUtil.arraySize(buffer) _
                                    , 0 _
                                    , resultLen
    
    If resultLen = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If

    ' 戻り値を格納するバッファを確保する
    resultBuffer = Space(resultLen)

    ' 複合化を実行する
    If SutGray.ConvertBinaryDataToHex(buffer(0) _
                                    , VBUtil.arraySize(buffer) _
                                    , resultBuffer _
                                    , resultLen) = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_WARNING _
                , _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_WARNING
    End If

    ' 戻り値を設定する
    ConvertBinaryDataToHexWrapper = resultBuffer
    
End Function


