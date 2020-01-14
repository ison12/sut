Attribute VB_Name = "WinAPI_ODBC"
Option Explicit

' *********************************************************
' user32.dllで定義されている関数郡や定数。
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/02/11　新規作成
'
' 特記事項：WindowsAPIを利用してデータソース情報にアクセスする
' *********************************************************

' =========================================================
' ▽ODBC環境変数取得
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SQLAllocEnv Lib "odbc32.dll" (ByRef phEnv As LongPtr) As Integer
#Else
    Public Declare Function SQLAllocEnv Lib "odbc32.dll" (ByRef phEnv As Long) As Integer
#End If

' =========================================================
' ▽ODBC環境変数解放
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As LongPtr) As Integer
#Else
    Public Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer
#End If

' =========================================================
' ▽ODBCデータソース取得
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
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
Public Const ERROR_MSG  As String = "データソース取得時にエラーが発生しました。"

' =========================================================
' ▽データソースリスト取得
'
' 概要　　　：データソースリストを取得する
' 戻り値　　：データソースリストが格納されたｺﾚｸｼｮﾝｵﾌﾞｼﾞｪｸﾄ
'
'             要素にはデータソース情報が格納されたｺﾚｸｼｮﾝｵﾌﾞｼﾞｪｸﾄ
'             ※以下の形式で格納される【キー】＝【値】
'             name = データソース名
' 　　　　　　desc = データソース説明
'
' =========================================================
Public Function getDataSourceList() As ValCollection
    
    ' ■戻り値
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
    
    ' エラー発生
    If retCode < 0 Then
    
        err.Raise ERROR_CODE, err.Source, ERROR_MSG
    End If
    
    ' ◇データソース情報格納リストを初期化する
    Set dataSourceList = New ValCollection
    
    Do While True
    
        ' データソース情報を取得する
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
                            
        ' ◇データソース情報を初期化する
        Set dataSourceInfo = New ValCollection
        
        ' データソース情報を設定する
        dataSourceInfo.setItem LeftB(szDSN, InStrB(szDSN, Chr$(0))), DATA_SOURCE_INFO_KEY_NAME
        dataSourceInfo.setItem LeftB(szDescription, InStrB(szDescription, Chr$(0))), DATA_SOURCE_INFO_KEY_DESC
        
        ' データソースリストに情報を設定する
        dataSourceList.setItem dataSourceInfo
    Loop
    
loop_end:

    ' エラー発生
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

