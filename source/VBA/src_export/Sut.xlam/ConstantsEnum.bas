Attribute VB_Name = "ConstantsEnum"
Option Explicit

' *********************************************************
' 列挙型定数モジュール
'
' 作成者　：Ison
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
' ▽DBクエリバッチ種類
'
' 概要　　　：DBクエリバッチ種類
'
' =========================================================
Public Enum DB_QUERY_BATCH_TYPE

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

' =========================================================
' ▽アイコンリソース
'
' 概要　　　：アイコンリソース
'
' =========================================================
Public Enum RESOURCE_ICON

    Add = 1
    addFile = 2
    addFolder = 3
    addFolder2 = 4
    alert = 5
    alertMessage = 6
    book = 7
    buttonHelp = 8
    database = 9
    databaseSetting = 10
    databaseSearch = 11
    delete = 12
    deleteDatabase = 13
    devil = 14
    Edit = 15
    remove = 16
    Run = 17
    SaveAs = 18
    Search = 19
    searchWindow = 20
    settings = 21
    smile = 22
    windowImport = 23
    flagGreen = 24
    flagBlue = 25
    flagRed = 26
    areaAdd = 27
    areaEdit = 28
    areaRemove = 29
    areaSearch = 30
    bug = 31
    Paste = 32
    Forward = 33
    
End Enum

' =========================================================
' ▽DB接続情報種類
' =========================================================
Public Enum DB_CONNECT_INFO_TYPE

    favorite = 1
    history = 2

End Enum

' 一括クエリ実行種類名称
Private dbQueryTypeNames As ValCollection

' =========================================================
' ▽DBクエリバッチ種類名称を取得する。
'
' 概要　　　：
' 引数　　　：d 一括クエリ実行種類
' 戻り値　　：一括クエリ実行種類名称
' 特記事項　：
'
' =========================================================
Public Function getDbQueryBatchTypeName(ByVal d As DB_QUERY_BATCH_TYPE) As String

    If dbQueryTypeNames Is Nothing Then
        ' 初回時のみ実行
    
        Set dbQueryTypeNames = New ValCollection
        
        ' 種類名称の設定
        dbQueryTypeNames.setItem "", DB_QUERY_BATCH_TYPE.none
        dbQueryTypeNames.setItem "INSERT + UPDATE", DB_QUERY_BATCH_TYPE.insertUpdate
        dbQueryTypeNames.setItem "INSERT", DB_QUERY_BATCH_TYPE.insert
        dbQueryTypeNames.setItem "UPDATE", DB_QUERY_BATCH_TYPE.update
        dbQueryTypeNames.setItem "DELETE", DB_QUERY_BATCH_TYPE.deleteOnSheet
        dbQueryTypeNames.setItem "DELETE テーブル上の全レコード", DB_QUERY_BATCH_TYPE.deleteAll
        dbQueryTypeNames.setItem "SELECT", DB_QUERY_BATCH_TYPE.selectAll
        dbQueryTypeNames.setItem "SELECT 条件指定", DB_QUERY_BATCH_TYPE.selectCondition
        dbQueryTypeNames.setItem "SELECT 再実行", DB_QUERY_BATCH_TYPE.selectReExec
        
    End If
    
    ' 種類名称の特定
    getDbQueryBatchTypeName = dbQueryTypeNames.getItem(d, vbVariant)

End Function
