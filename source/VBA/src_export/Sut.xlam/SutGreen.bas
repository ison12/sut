Attribute VB_Name = "SutGreen"
Option Explicit
' *********************************************************
' SutGreen.dll関連のモジュール
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/03/14　新規作成
'
' 特記事項：
' *********************************************************

' GetProcAddrで取得した関数ポインタから関数を実行する”関数”
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
