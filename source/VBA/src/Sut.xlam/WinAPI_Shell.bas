Attribute VB_Name = "WinAPI_Shell"
Option Explicit

' *********************************************************
' shell32.dllで定義されている関数郡や定数。
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/10/11　新規作成
'
' 特記事項：
'
' *********************************************************

' ▼WinAPIの定義
' 外部プログラム起動関数
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As LongPtr _
       , ByVal lpOperation As String _
       , ByVal lpFile As String _
       , ByVal lpParameters As String _
       , ByVal lpDirectory As String _
       , ByVal nShowCmd_ As Long) As Long
#Else
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long _
       , ByVal lpOperation As String _
       , ByVal lpFile As String _
       , ByVal lpParameters As String _
       , ByVal lpDirectory As String _
       , ByVal nShowCmd_ As Long) As Long
#End If

' =========================================================
' ▽環境変数を取得する
'
' 概要　　　：
' 引数　　　：key 環境変数のキー値
' 戻り値　　：環境変数の値
'
' =========================================================
Public Function getEnvironmentVariable(ByVal key As String) As String

    ' 環境変数
    Dim environmentString As String
    ' 環境変数のキー値部分
    Dim environmentKey    As String
    ' 環境変数の値部分
    Dim environmentVal    As String
    
    ' 環境変数のキーと値を区切り文字（=）の位置を保持する変数
    Dim envKeyValSeparate As Long
    
    ' インデックス
    Dim i As Long
    
    ' インデックスの初期値を1とする
    i = 1
    
    ' ループを実施する
    Do
        ' 環境変数
        environmentString = Environ(i)
        
        ' キーと値の区切り文字の位置を格納する
        envKeyValSeparate = InStr(environmentString, "=")
        
        ' 区切り文字が見つかった場合
        If envKeyValSeparate <> 0 Then
        
            ' キーを格納
            environmentKey = Mid$(environmentString, 1, envKeyValSeparate - 1)
            ' 値を格納
            environmentVal = Mid$(environmentString, envKeyValSeparate + 1, Len(environmentString))
            
            If UCase$(environmentKey) = UCase$(key) Then
            
                ' 環境変数の値を戻り値に格納する
                getEnvironmentVariable = environmentVal
                
                Exit Function
            End If
            
        End If
        
        i = i + 1
        
    Loop Until Environ(i) = ""

    ' 空文字列を返す
    getEnvironmentVariable = ""

End Function
