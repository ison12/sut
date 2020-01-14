Attribute VB_Name = "ConstantsEnum"
Option Explicit

' *********************************************************
' 列挙型定数モジュール
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' ▽テーブル制約種類
'
' 概要　　　：テーブル制約種類
'
' =========================================================
Public Enum TABLE_CONSTANTS_TYPE

    tableConstPk = 0
    tableConstUk = 1
    tableConstFk = 2
    tableConstUnknown = -1

End Enum

' =========================================================
' ▽行フォーマット種類
'
' 概要　　　：行フォーマット種類
'
' =========================================================
Public Enum REC_FORMAT

    recFormatToUnder = 0
    recFormatToRight = 1

End Enum

' =========================================================
' ▽一括クエリ実行種類
'
' 概要　　　：一括クエリ実行種類
'
' =========================================================
Public Enum BATCH_QUERY_TYPE

    none = 0
    insertUpdate = 1
    insert = 2
    update = 3
    deleteOnSheet = 4
    deleteAll = 5
    selectAll = 6
    selectCondition = 7
    selectReExec = 8

End Enum


