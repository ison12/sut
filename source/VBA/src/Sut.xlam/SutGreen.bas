Attribute VB_Name = "SutGreen"
Option Explicit
' *********************************************************
' SutGreen.dll�֘A�̃��W���[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2009/03/14�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' GetProcAddr�Ŏ擾�����֐��|�C���^����֐������s����h�֐��h
#If (DEBUG_MODE = 1) Then
    #If VBA7 And Win64 Then
        Public Declare PtrSafe Function CallByFuncPtr Lib ".\..\CPP\Sut\x64\Debug ASM\SutGreen.dll" (ByVal funcAddr As LongPtr) As Long
        Public Declare PtrSafe Function CallByFuncPtrParam Lib ".\..\CPP\Sut\x64\Debug ASM\SutGreen.dll" (ByVal funcAddr As LongPtr, ByRef param As Any) As Long
        Public Declare PtrSafe Function CallByFuncPtrParam2 Lib ".\..\CPP\Sut\x64\Debug ASM\SutGreen.dll" (ByVal funcAddr As LongPtr, ByRef param1 As Any, ByRef param2 As Any) As Long
        Public Declare PtrSafe Function CallByFuncPtrParamInt Lib ".\..\CPP\Sut\x64\Debug ASM\SutGreen.dll" (ByVal funcAddr As LongPtr, ByVal param As Long) As Long
    #Else
        Public Declare Function CallByFuncPtr Lib ".\..\CPP\Sut\Debug ASM\SutGreen.dll" (ByVal funcAddr As Long) As Long
        Public Declare Function CallByFuncPtrParam Lib ".\..\CPP\Sut\Debug ASM\SutGreen.dll" (ByVal funcAddr As Long, ByRef param As Any) As Long
        Public Declare Function CallByFuncPtrParam2 Lib ".\..\CPP\Sut\Debug ASM\SutGreen.dll" (ByVal funcAddr As Long, ByRef param1 As Any, ByRef param2 As Any) As Long
        Public Declare Function CallByFuncPtrParamInt Lib ".\..\CPP\Sut\Debug ASM\SutGreen.dll" (ByVal funcAddr As Long, ByVal param As Long) As Long
    #End If
#Else
    #If VBA7 And Win64 Then
        Public Declare PtrSafe Function CallByFuncPtr Lib "lib\SutGreen.dll" (ByVal funcAddr As LongPtr) As Long
        Public Declare PtrSafe Function CallByFuncPtrParam Lib "lib\SutGreen.dll" (ByVal funcAddr As LongPtr, ByRef param As Any) As Long
        Public Declare PtrSafe Function CallByFuncPtrParam2 Lib "lib\SutGreen.dll" (ByVal funcAddr As LongPtr, ByRef param1 As Any, ByRef param2 As Any) As Long
        Public Declare PtrSafe Function CallByFuncPtrParamInt Lib "lib\SutGreen.dll" (ByVal funcAddr As LongPtr, ByVal param As Long) As Long
    #Else
        Public Declare Function CallByFuncPtr Lib "lib\SutGreen.dll" (ByVal funcAddr As Long) As Long
        Public Declare Function CallByFuncPtrParam Lib "lib\SutGreen.dll" (ByVal funcAddr As Long, ByRef param As Any) As Long
        Public Declare Function CallByFuncPtrParam2 Lib "lib\SutGreen.dll" (ByVal funcAddr As Long, ByRef param1 As Any, ByRef param2 As Any) As Long
        Public Declare Function CallByFuncPtrParamInt Lib "lib\SutGreen.dll" (ByVal funcAddr As Long, ByVal param As Long) As Long
    #End If
#End If
