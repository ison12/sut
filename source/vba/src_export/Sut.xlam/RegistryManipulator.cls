VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistryManipulator"
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

' レジストリのハンドル
#If VBA7 And Win64 Then
Private handle As LongPtr
#Else
Private handle As Long
#End If

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    handle = 0
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' 戻り値
    Dim ret As Long
    
    ' レジストリを開いていない場合
    If handle = 0 Then
        ' 何もせずに終了する
        Exit Sub
    End If
    
    ' ハンドルをクローズする
    ret = WinAPI_ADVAP.RegCloseKey(handle)
    ' 戻り値をチェック
    If ret = WinAPI_ADVAP.ERROR_SUCCESS Then
        
        ' 成功
        #If (DEBUG_MODE <> 1) Then
            Debug.Print "RegestryManipulator Class_Terminate Success"
        #End If

    Else
    
        ' 失敗
        #If (DEBUG_MODE <> 1) Then
            Debug.Print "RegestryManipulator Class_Terminate Failed. error code " & ret
        #End If
    End If
    
    ret = 0
    
End Sub

' =========================================================
' ▽初期化メソッド
'
' 概要　　　：レジストリの任意のキーを開く
' 引数　　　：hKey        ルートキー
' 　　　　　　lpSubKey    サブキー
' 　　　　　　samDesired  セキュリティアクセス権
' 　　　　　　isCreate    任意のキーを開く際に新規作成するかどうかを表すフラグ
' 戻り値　　：成功した場合 True
'
' =========================================================
Public Function init(ByVal hKey As RegKeyConstants _
                   , ByVal lpSubKey As String _
                   , ByVal samDesired As RegAccessConstants _
                   , Optional ByVal isCreate As Boolean = False) As Boolean

    ' APIの戻り値
    Dim ret  As Long
    
    ' キーを作成する
    If isCreate = True Then
    
#If VBA7 And Win64 Then
        Dim lpdwDisposition As LongPtr
#Else
        Dim lpdwDisposition As Long
#End If
        
        ret = WinAPI_ADVAP.RegCreateKeyEx(hKey _
                                       , lpSubKey & vbNullChar _
                                       , 0 _
                                       , 0 _
                                       , WinAPI_ADVAP.REG_OPTION_NON_VOLATILE _
                                       , samDesired _
                                       , 0 _
                                       , handle _
                                       , lpdwDisposition)
        
        ' 戻り値をチェック
        If ret = WinAPI_ADVAP.ERROR_SUCCESS Then
            
            ' キーが新規に作成された場合
            If lpdwDisposition = WinAPI_ADVAP.REG_CREATED_NEW_KEY Then
            
                ' 成功
                #If (DEBUG_MODE <> 1) Then
                    Debug.Print "RegestryManipulator init Success (Created new key)"
                #End If

            ' キーが既に存在する場合
            Else
            
                ' 成功
                #If (DEBUG_MODE <> 1) Then
                    Debug.Print "RegestryManipulator init Success (Opened existing key)"
                #End If
            End If
            
        Else
        
            ' 失敗
            #If (DEBUG_MODE <> 1) Then
                Debug.Print "RegestryManipulator init Failed. error code " & ret
            #End If
            
            err.Raise ConstantsError.ERR_NUMBER_REGISTRY_ACCESS_FAILED _
                    , _
                    , ConstantsError.ERR_DESC_REGISTRY_ACCESS_FAILED & vbNewLine & ret
        End If
    
    ' キーをオープンする
    Else
    
        ret = WinAPI_ADVAP.RegOpenKeyEx(hKey _
                                     , lpSubKey _
                                     , 0 _
                                     , samDesired _
                                     , handle)
    
        ' 戻り値をチェック
        If ret = WinAPI_ADVAP.ERROR_SUCCESS Then
            
            ' 成功
            #If (DEBUG_MODE <> 1) Then
                Debug.Print "RegestryManipulator init Success"
            #End If
        Else
        
            ' 失敗
            #If (DEBUG_MODE <> 1) Then
                Debug.Print "RegestryManipulator init Failed. error code " & ret
            #End If
            
'            err.Raise ConstantsError.ERR_NUMBER_REGISTRY_ACCESS_FAILED _
'                    , _
'                    , ConstantsError.ERR_DESC_REGISTRY_ACCESS_FAILED & vbNewLine & ret

            init = False
            
            Exit Function
        End If
    
    End If

    init = True
End Function

' =========================================================
' ▽値を取得する
'
' 概要　　　：
' 引数　　　：valueKey
' 戻り値　　：キー値
'
' =========================================================
Public Function GetValue _
            (ByVal valueKey As String _
           , ByRef value As String) As Boolean

    ' byteバッファサイズ
#If VBA7 And Win64 Then
    Dim lpCbData As LongPtr
#Else
    Dim lpCbData As Long
#End If
    
    ' 取得したデータタイプの型
#If VBA7 And Win64 Then
    Dim lpType As LongPtr
#Else
    Dim lpType As Long
#End If
    
    ' APIの戻り値
    Dim ret As Long
    
    ' キーの終端にNULL文字を付加する
    valueKey = valueKey & vbNullChar
    
    ' 値のサイズを取得する（UNICODEバージョンで情報を取得）
    ret = WinAPI_ADVAP.RegQueryValueExW(handle, StrPtr(valueKey), 0, lpType, 0, lpCbData)
    
    ' 関数の呼び出しに失敗
    If _
        ret <> WinAPI_ADVAP.ERROR_MORE_DATA _
    And ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        ' 失敗
        GetValue = False
        Exit Function
    End If
    
    Dim lpCbDataVar As Variant
    lpCbDataVar = lpCbData
    
    ' バッファを確保する
    value = String(lpCbDataVar, vbNullChar)
    
    ' 値を取得する（UNICODEバージョンで情報を取得）
    ret = WinAPI_ADVAP.RegQueryValueExW(handle, StrPtr(valueKey), 0, lpType, StrPtr(value), LenB(value))
    
    ' 関数の呼び出しに失敗
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        ' 失敗
        GetValue = False
    Else
    
        GetValue = True
    
        ' 文字列として設定する
        ' ※UnicodeバージョンのAPIを使用するため、LeftBやInStrBは使わない
        If InStr(value, vbNullChar) <> 0 Then
        
            ' NULL文字が存在する場合、NULL文字を除いた前方の文字列を取得する
            value = Left(value, InStr(value, vbNullChar) - 1)
        End If
    End If
    
    
End Function

' =========================================================
' ▽値を取得する
'
' 概要　　　：
' 引数　　　：valueKey
' 戻り値　　：キー値
'
' =========================================================
Public Function getValueByByte _
            (ByVal valueKey As String _
           , ByRef value() As Byte) As Boolean

    ' byteバッファサイズ
#If VBA7 And Win64 Then
    Dim lpCbData As LongPtr
#Else
    Dim lpCbData As Long
#End If
    
    ' 取得したデータタイプの型
#If VBA7 And Win64 Then
    Dim lpType As LongPtr
#Else
    Dim lpType As Long
#End If
    
    ' APIの戻り値
    Dim ret As Long
    
    ' キーの終端にNULL文字を付加する
    valueKey = valueKey & vbNullChar
    
    ' 値のサイズを取得する（UNICODEバージョンで情報を取得）
    ret = WinAPI_ADVAP.RegQueryValueExW(handle, StrPtr(valueKey), 0, lpType, 0, lpCbData)
    
    ' 関数の呼び出しに失敗
    If _
        ret <> WinAPI_ADVAP.ERROR_MORE_DATA _
    And ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        ' 失敗
        getValueByByte = False
        Exit Function
    End If
    
    Dim lpCbDataVar As Variant
    lpCbDataVar = lpCbData

    ' バッファを確保する
    ReDim value(0 To lpCbDataVar - 1)
    
    ' 値を取得する（UNICODEバージョンで情報を取得）
    ret = WinAPI_ADVAP.RegQueryValueExW(handle, StrPtr(valueKey), 0, lpType, VarPtr(value(0)), lpCbData)
    
    ' 関数の呼び出しに失敗
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        ' 失敗
        getValueByByte = False
    Else
    
        getValueByByte = True
    
    End If
    
End Function

' =========================================================
' ▽値を設定する
'
' 概要　　　：
' 引数　　　：valueKey 値のキー
' 　　　　　　value    値
'
' =========================================================
Public Function setValue _
            (ByVal valueKey As String, _
             ByVal value As String) As Boolean
             
    ' APIの戻り値
    Dim ret As Long
    
    ' キーの終端にNULL文字を付加する
    valueKey = valueKey & vbNullChar
    ' 値の終端にNULL文字を付加する
    value = value & vbNullChar
    
    ' 値を設定する（UNICODEバージョンで情報を設定）
    ret = WinAPI_ADVAP.RegSetValueExW(handle, StrPtr(valueKey), 0, WinAPI_ADVAP.REG_SZ, StrPtr(value), LenB(value))
    
    ' エラーの場合
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        setValue = False
        
    Else
    
        setValue = True
    End If
    
End Function

' =========================================================
' ▽値を設定する (Byte型バージョン)
'
' 概要　　　：
' 引数　　　：valueKey 値のキー
' 　　　　　　value    値
'
' =========================================================
Public Function setValueByByte _
            (ByVal valueKey As String, _
             ByRef value() As Byte) As Boolean
             
    ' APIの戻り値
    Dim ret As Long
    
    ' キーの終端にNULL文字を付加する
    valueKey = valueKey & vbNullChar
    
    ' 値を設定する（UNICODEバージョンで情報を設定）
    ret = WinAPI_ADVAP.RegSetValueExW(handle _
                                    , StrPtr(valueKey) _
                                    , 0 _
                                    , WinAPI_ADVAP.REG_BINARY _
                                    , VarPtr(value(0)) _
                                    , VBUtil.arraySize(value))
    
    ' エラーの場合
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        setValueByByte = False
        
    Else
    
        setValueByByte = True
    End If
    
End Function

' =========================================================
' ▽キーリストを取得する
'
' 概要　　　：現在開いているキーのサブキーを全て取得する。
' 引数　　　：
' 戻り値　　：キーリスト、失敗した場合はnothing。
'
' =========================================================
Public Function getKeyList() As ValCollection

    ' キーを格納するバッファ文字列のサイズ（定数）
    Const STR_SIZE As Long = 256
    
    ' キーを格納するバッファ
    Dim str As String
    
    ' 取得したデータタイプの型
#If VBA7 And Win64 Then
    Dim strSize As LongPtr
#Else
    Dim strSize As Long
#End If
    
    ' APIの戻り値
    Dim ret As Long
    ' インデックス変数
    Dim i As Long
    ' 更新時間
    Dim ft As WinAPI_ADVAP.FILETIME
    
    ' 結果リスト
    Dim resultList As New ValCollection
    
    Do While True
    
        ' 文字列を初期化する
        str = String(STR_SIZE, vbNullChar)
        ' 文字列のサイズを初期化する
        strSize = STR_SIZE
        
        ' キーを列挙する
        ret = WinAPI_ADVAP.RegEnumKeyEx(handle, i, str, strSize, 0, 0, 0, ft)
    
        ' アイテムが存在する
        If ret = WinAPI_ADVAP.ERROR_MORE_DATA Then
        
            ' Delete NULL
            ' ※UnicodeバージョンのAPIを使用するため、LeftBやInStrBは使わない
            If InStr(str, vbNullChar) <> 0 Then
            
                ' NULL文字が存在する場合、NULL文字を除いた前方の文字列を取得する
                str = Left(str, InStr(str, vbNullChar) - 1)
            End If

            ' 結果リストにキーを格納する
            resultList.setItem str
        
        ' アイテムが存在しない場合
        ElseIf ret = WinAPI_ADVAP.ERROR_NO_MORE_ITEMS Then
        
            ' ループを終了する
            Exit Do
            
        ' その他の場合
        Else
        
            ' エラーと判定する
            Set resultList = Nothing
            Exit Function
        End If
        
        i = i + 1
    Loop
    
    ' 戻り値を返す
    Set getKeyList = resultList
    
End Function

' =========================================================
' ▽キーリストを取得する
'
' 概要　　　：現在開いているキーのサブキーを全て取得する。
' 引数　　　：valueNameList 値名称リスト
' 　　　　　　valueList     値リスト
' 戻り値　　：
'
' =========================================================
Public Sub getValueList(ByRef valueNameList As ValCollection _
                      , ByRef valueList As ValCollection)

    ' キーを格納するバッファ文字列のサイズ（定数）
    Const STR_SIZE As Long = 256

    ' キーを格納するバッファ
    Dim valueName As String
    ' キーのbyteバッファサイズ
    Dim lpCbValueName As Long
    
    ' 値
    Dim value As String
    ' 値のバッファサイズ
#If VBA7 And Win64 Then
    Dim lpCbData As LongPtr
#Else
    Dim lpCbData As Long
#End If
    
    
    ' 取得したデータタイプの型
#If VBA7 And Win64 Then
    Dim lpType As LongPtr
#Else
    Dim lpType As Long
#End If
    
    ' APIの戻り値
    Dim ret As Long
    ' インデックス変数
    Dim i As Long
    
    ' 結果リスト
    Set valueNameList = New ValCollection
    Set valueList = New ValCollection
    
    Do While True
    
        valueName = String(STR_SIZE, vbNullChar)
        
        ' 値を列挙する
        ret = WinAPI_ADVAP.RegEnumValueW(handle, i, StrPtr(valueName), STR_SIZE, 0, lpType, 0, lpCbData)
    
        ' 正常の場合
        If ret = WinAPI_ADVAP.ERROR_SUCCESS Then
        
            Dim lpCbDataVar As Variant
            lpCbDataVar = lpCbData
        
            ' バッファを確保する
            value = String(lpCbDataVar, vbNullChar)
    
            ' 値を取得する
            ret = WinAPI_ADVAP.RegEnumValueW(handle, i, StrPtr(valueName), STR_SIZE, 0, lpType, StrPtr(value), lpCbData)
    
            ' 文字列として設定する
            ' ※UnicodeバージョンのAPIを使用するため、LeftBやInStrBは使わない
            If InStr(valueName, vbNullChar) <> 0 Then
            
                ' NULL文字が存在する場合、NULL文字を除いた前方の文字列を取得する
                valueName = Left(valueName, InStr(valueName, vbNullChar) - 1)
            End If
            
            If InStr(value, vbNullChar) <> 0 Then
            
                ' NULL文字が存在する場合、NULL文字を除いた前方の文字列を取得する
                value = Left(value, InStr(value, vbNullChar) - 1)
            End If

            ' 結果リストにキーを格納する
            valueNameList.setItem valueName, valueName
            valueList.setItem value, valueName
        
            
        End If

        ' アイテムが存在しない場合
        If ret = WinAPI_ADVAP.ERROR_NO_MORE_ITEMS Then
        
            ' ループを終了する
            Exit Do
            
        ' エラーが発生した場合
        ElseIf ret <> WinAPI_ADVAP.ERROR_SUCCESS _
           And ret <> WinAPI_ADVAP.ERROR_MORE_DATA Then
        
            Exit Sub
        End If
        
        i = i + 1
    Loop
    
End Sub

' =========================================================
' ▽現在のキーに複数の値を設定する
'
' 概要　　　：引数valuesはN×2の2次元配列であること。
' 　　　　　　以下に引数valuesの構成を示す。
' 　　　　　　-------------------
' 　　　　　　     0    1
'             0 / [key1][value1]
'             1 / [key2][value2]
'             2 / [key3][value3]
'             3 / [key4][value4]

' 引数　　　：values  複数の値
' 戻り値　　：
'
' =========================================================
Public Sub setValues _
            (ByRef values() As Variant)

    Dim i As Long
    
    Dim valueKey As String
    Dim value    As String
    
    For i = LBound(values) To UBound(values)
    
        valueKey = CStr(values(i, 0))
        value = CStr(values(i, 1))
    
        setValue valueKey, value
    Next
    
End Sub

' =========================================================
' ▽キーを削除する
'
' 概要　　　：
' 引数　　　：key キー
'
' =========================================================
Public Function delete(ByVal key As String) As Boolean

    Dim ret As Long
    
    key = key & vbNullChar
    ret = WinAPI_ADVAP.RegDeleteKeyW(handle, StrPtr(key))
    
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        delete = False
        
    Else
    
        delete = True
    End If
    
End Function

' =========================================================
' ▽値を削除する
'
' 概要　　　：
' 引数　　　：valueKey 値のキー
'
' =========================================================
Public Function deleteValue(ByVal valueKey As String) As Boolean

    Dim ret As Long
    
    valueKey = valueKey & vbNullChar
    ret = WinAPI_ADVAP.RegDeleteValueW(handle, StrPtr(valueKey))
    
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        deleteValue = False
    Else
    
        deleteValue = True
    End If
    
End Function




