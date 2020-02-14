Attribute VB_Name = "PostgreSQL"
Option Explicit

' *********************************************************
' PostgreSQLに関連したユーティリティモジュール
'
' 作成者　：Ison
' 履歴　　：2008/04/27　新規作成
'
' 特記事項　：PostgreSQLに依存した関数郡を定義。
' *********************************************************

' =========================================================
' ▽データ値のチェック関数（特別な値）
'
' 概要　　　：INSERTのVALUES句等の値をチェックする。
' 引数　　　：value データ値
'
' 戻り値　　：True NULLもしくはNOW関数の場合
'
' =========================================================
Public Function isSpecialValue(ByVal value As String) As Boolean

    isSpecialValue = False
    
    If _
           value = "NULL" _
        Or UCase$(value) = "NOW()" _
        Or UCase$(value) = "CURRENT_DATE" _
        Or UCase$(value) = "CURRENT_TIME" _
        Or UCase$(value) = "CURRENT_TIMESTAMP" _
    Then
    
        isSpecialValue = True
    
    End If

End Function

' =========================================================
' ▽データ値のチェック関数（文字列型）
'
' 概要　　　：INSERTのVALUES句等の値が文字列型であるかをチェックする。
' 引数　　　：dataType データ型
'
' 戻り値　　：True 文字列型の場合
'
' =========================================================
Public Function isChar(ByVal dataType As String) As Boolean

    isChar = False

    If _
           InStr(UCase$(dataType), "BIT") <> 0 _
        Or InStr(UCase$(dataType), "BIT VARYING") <> 0 _
        Or InStr(UCase$(dataType), "BOX") <> 0 _
        Or InStr(UCase$(dataType), "CHARACTER VARYING") <> 0 _
        Or InStr(UCase$(dataType), "CHARACTER") <> 0 _
        Or InStr(UCase$(dataType), "CHAR") <> 0 _
        Or InStr(UCase$(dataType), "CIDR") <> 0 _
        Or InStr(UCase$(dataType), "CIRCLE") <> 0 _
        Or InStr(UCase$(dataType), "INET") <> 0 _
        Or InStr(UCase$(dataType), "INTERVAL") <> 0 _
        Or InStr(UCase$(dataType), "LINE") <> 0 _
        Or InStr(UCase$(dataType), "LSEG") <> 0 _
        Or InStr(UCase$(dataType), "MACADDR") <> 0 _
        Or InStr(UCase$(dataType), "MONEY") <> 0 _
        Or InStr(UCase$(dataType), "PATH") <> 0 _
        Or InStr(UCase$(dataType), "POINT") <> 0 _
        Or InStr(UCase$(dataType), "POLYGON") <> 0 _
        Or InStr(UCase$(dataType), "TEXT") <> 0 _
    Then
    
        isChar = True
            
    End If

End Function

' =========================================================
' ▽データ値のチェック関数（日付・時間型）
'
' 概要　　　：INSERTのVALUES句等の値が日付・時間型であるかをチェックする。
' 引数　　　：dataType データ型
'
' 戻り値　　：True 日付・時間型の場合
'
' =========================================================
Public Function isTime(ByVal dataType As String) As Boolean

    isTime = False

    If _
           InStr(UCase$(dataType), "DATE") <> 0 _
        Or InStr(UCase$(dataType), "TIME WITHOUT TIME ZONE") <> 0 _
        Or InStr(UCase$(dataType), "TIME WITH TIME ZONE") <> 0 _
        Or InStr(UCase$(dataType), "TIMESTAMP WITHOUT TIME ZONE") <> 0 _
        Or InStr(UCase$(dataType), "TIMESTAMP WITH TIME ZONE") <> 0 _
    Then
    
        isTime = True
            
    End If

End Function

Public Function escapeValue(ByRef val As String) As String

    escapeValue = replace(val, "'", "''")

End Function

