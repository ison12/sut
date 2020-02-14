Attribute VB_Name = "MySQL"
Option Explicit

' *********************************************************
' MySQLに関連したユーティリティモジュール
'
' 作成者　：Ison
' 履歴　　：2007/12/01　新規作成
'
' 特記事項　：MySQLに依存した関数郡を定義。
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
           InStr(UCase$(dataType), "CHAR") <> 0 _
        Or InStr(UCase$(dataType), "VARCHAR") <> 0 _
        Or InStr(UCase$(dataType), "BLOB") <> 0 _
        Or InStr(UCase$(dataType), "TEXT") <> 0 _
        Or InStr(UCase$(dataType), "ENUM") <> 0 _
        Or InStr(UCase$(dataType), "SET") <> 0 _
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
           InStr(UCase$(dataType), "DATETIME") <> 0 _
        Or InStr(UCase$(dataType), "DATE") <> 0 _
        Or InStr(UCase$(dataType), "TIMESTAMP") <> 0 _
        Or InStr(UCase$(dataType), "TIME") <> 0 _
        Or InStr(UCase$(dataType), "YEAR") <> 0 _
    Then
    
        isTime = True
            
    End If

End Function

Public Function escapeValue(ByRef val As String) As String

    escapeValue = replace(val, "'", "''")

End Function

