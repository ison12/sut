VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniFileManipulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' INIファイル操作クラス
'
' 作成者　：Ison
' 履歴　　：2008/09/06　新規作成
'
' 特記事項：
' *********************************************************

'▼INIファイル用のWinAPI関数宣言

#If VBA7 And Win64 Then

' INIファイルの指定したセクション内のすべてのキーと値を取得
Private Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" _
    Alias "GetPrivateProfileSectionA" _
    (ByVal lpAppName As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

' INIファイルの指定したセクション内のすべてのキーと値を設定
Private Declare PtrSafe Function WritePrivateProfileSection Lib "kernel32" _
    Alias "WritePrivateProfileSectionA" _
    (ByVal lpAppName As String, _
     ByVal lpString As String, _
     ByVal lpFileName As String) As Long

' INIファイルの文字列を取得
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

' INIファイルの文字列を変更
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpString As Any, _
     ByVal lpFileName As String) As Long

#Else

' INIファイルの指定したセクション内のすべてのキーと値を取得
Private Declare Function GetPrivateProfileSection Lib "kernel32" _
    Alias "GetPrivateProfileSectionA" _
    (ByVal lpAppName As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

' INIファイルの指定したセクション内のすべてのキーと値を設定
Private Declare Function WritePrivateProfileSection Lib "kernel32" _
    Alias "WritePrivateProfileSectionA" _
    (ByVal lpAppName As String, _
     ByVal lpString As String, _
     ByVal lpFileName As String) As Long

' INIファイルの文字列を取得
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

' INIファイルの文字列を変更
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpString As Any, _
     ByVal lpFileName As String) As Long

#End If

' ファイル名
Public fileName As String

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ▽INIファイルから任意のキー値を取得する
'
' 概要　　　：
' 引数　　　：section セクション
' 　　　　　　key     キー
' 戻り値　　：セクションとキーに紐づく値
'
' =========================================================
Public Function getValue _
            (ByVal section As String, _
             ByVal key As String) As String

    Dim returnedValue As String * 32767
    
    Dim rc As Long


    ' ファイルが存在しない場合は終了
    If (dir(fileName, vbNormal) = "") Then
        
        Exit Function
        
    End If

    On Error Resume Next
    
    rc = GetPrivateProfileString(section _
                               , key _
                               , vbNullString _
                               , returnedValue _
                               , Len(returnedValue) _
                               , fileName)
                 
    
    On Error GoTo 0
    
    If (rc = 0) Then
    
        Exit Function
    End If

    getValue = Left$(returnedValue, InStr(1, returnedValue, vbNullChar) - 1)
    
End Function

' =========================================================
' ▽INIファイルに任意のキー値を設定する
'
' 概要　　　：
' 引数　　　：section セクション
' 　　　　　　key     キー
'             value   値
' 戻り値　　：Trueの場合、正常に書き込み完了
'
' =========================================================
Public Function setValue _
            (ByVal section As String, _
             ByVal key As String, _
             ByVal value As String) As Boolean
             
    Dim rc As Long

    'INIファイルを作成する（既に存在している場合は作成しない）
    createFile
    
    On Error Resume Next
    
    ' INIファイルに指定されたキーと値で書き込みを行う
    rc = WritePrivateProfileString(section, key, value, fileName)
    
    On Error GoTo 0
    
    If (rc = 0) Then
        
        setValue = False
    Else
    
        setValue = True
    End If

End Function

' =========================================================
' ▽INIファイルの任意のセクションに複数のキー値を設定する
'
' 概要　　　：引数valuesはN×2の2次元配列であること。
' 　　　　　　以下に引数valuesの構成を示す。
' 　　　　　　-------------------
' 　　　　　　     0    1
'             0 / [key1][value1]
'             1 / [key2][value2]
'             2 / [key3][value3]
'             3 / [key4][value4]

' 引数　　　：section セクション
'             values  複数の値
' 戻り値　　：Trueの場合、正常に書き込み完了
'
' =========================================================
Public Function setValues _
            (ByVal section As String, _
             ByRef values() As Variant) As Boolean
             
    Dim rc As Long

    'INIファイルを作成する（既に存在している場合は作成しない）
    createFile
    
    ' --------------------------------------------------------
    ' WritePrivateProfileSection関数に渡す文字列を生成する
    
    Dim i As Long
    Dim j As Long
    
    ' 配列の値を結合した値
    Dim joinValues As String
    
    For i = LBound(values, 1) To UBound(values, 1)
    
        ' イコール(=)で結合し、終わりにNULL文字列を付加する
        joinValues = joinValues & values(i, 0) & "=" & values(i, 1) & vbNullChar
        
    Next
    
    ' NULL文字列を付加する
    joinValues = joinValues & vbNullChar
    
    On Error Resume Next
    
    ' INIファイルに指定されたキーと値で書き込みを行う
    rc = WritePrivateProfileSection(section, joinValues, fileName)
    
    On Error GoTo 0
    
    If (rc = 0) Then
        
        setValues = False
    Else
    
        setValues = True
    End If

End Function

' =========================================================
' ▽INIファイルの任意のセクションに複数のキー値を設定する
'
' 概要　　　：戻り値はN×2の2次元配列となる。
' 　　　　　　以下に戻り値の構成を示す。
' 　　　　　　-------------------
' 　　　　　　     0    1
'             0 / [key1][value1]
'             1 / [key2][value2]
'             2 / [key3][value3]
'             3 / [key4][value4]

' 引数　　　：section セクション
' 戻り値　　：2次元配列
'
' =========================================================
Public Function getValues _
            (ByVal section As String) As Variant()

    Dim result() As Variant
    
    Dim rc As Long

    ' WinAPIからの戻り値を格納する文字列（32767文字のバッファを確保）
    Dim returnedValue  As String * 32767
    
    ' returnedValuenから有効な文字列部分を取り出して格納する文字列
    Dim returnedValue2 As String

    On Error Resume Next
    
    ' INIファイルに指定されたセクションに紐づく全てのキー値を取得する
    rc = GetPrivateProfileSection(section, returnedValue, Len(returnedValue), fileName)
    
    On Error GoTo 0
    
    If (rc = 0) Then
        
        Exit Function
    
    End If

    
    returnedValue2 = Left$(returnedValue, InStr(1, returnedValue, vbNullChar & vbNullChar) - 1)
    

    ' --------------------------------------------------------
    ' returnedValue2からvbNullCharを区切り文字と見なし配列に格納する変数
    Dim splitValues() As Variant
    
    splitValues = Split(returnedValue2, vbNullChar)
    
    ' 戻り値の配列のサイズを確保
    ReDim result( _
              0 To UBound(splitValues) - LBound(splitValues) _
            , 0 To 1)
    
    
    ' --------------------------------------------------------
    ' 戻り値の配列にキーと値を格納
    
    Dim i As Long
    
    Dim keyValSeparate As String
    Dim key As String
    Dim val As String
    
    
    ' 配列の値を結合した値
    Dim joinValues As String
    
    For i = LBound(splitValues, 1) To UBound(splitValues, 1)
    
        ' キーと値の区切り文字(=)の位置を格納する
        keyValSeparate = InStr(splitValues(i), "=")
        
        ' 区切り文字が見つかった場合
        If keyValSeparate <> 0 Then
        
            ' キーを格納
            key = Mid$(splitValues(i), 1, keyValSeparate - 1)
            ' 値を格納
            val = Mid$(splitValues(i), keyValSeparate + 1, Len(splitValues(i)))
            
            result(i, 0) = key
            result(i, 1) = val
            
        End If
        
    Next
    
    getValues = result
    
End Function

' =========================================================
' ▽INIファイルの任意のセクションまたはキーを削除する
'
' 概要　　　：引数keyを省略した場合、引数section配下の全てのキー値が削除される。
' 　　　　　　引数keyを省略しなかった場合、キー値が削除される。
'
' 引数　　　：section セクション
' 　　　　　　key     キー
' 戻り値　　：Trueの場合、正常に書き込み完了
'
' =========================================================
Public Function delete _
            (ByVal section As String, _
             Optional ByVal key As String = "") As Boolean

    Dim rc As Long

    On Error Resume Next
    
    If (key = "") Then
        
        'Sectionの削除
        delete = setValue(section, vbNullString, vbNullString)
        
    Else
        
        'Keyの削除
        delete = setValue(section, key, vbNullString)
        
    End If
    
End Function

'--- INIファイルの作成 ------------------------------
Private Sub createFile()

    ' ファイル番号
    Dim fileNum As Long
    
    ' ファイルが存在しない場合
    If (dir(fileName, vbNormal) = "") Then
    
        '指定ファイルが無いので作成
        On Error Resume Next
        
        ' ファイル番号を取得
        fileNum = FreeFile
        ' ファイルを開く
        Open fileName For Output As #fileNum
        ' ファイルを閉じる
        Close #fileNum
        
        
        On Error GoTo 0
    
    End If
    
End Sub

