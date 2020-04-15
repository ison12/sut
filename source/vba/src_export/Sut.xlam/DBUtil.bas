Attribute VB_Name = "DBUtil"
Option Explicit

' *********************************************************
' DBに関連したユーティリティモジュール
'
' 作成者　：Ison
' 履歴　　：2009/03/21　新規作成
'
' 特記事項：
' *********************************************************

' =========================================================
' ▽DBMS種類
'
' 概要　　　：DBMS（データベースマネージメントシステム）の種類をあらわす列挙型
'
' =========================================================
Public Enum DbmsType

    MySQL = 0
    PostgreSQL = 1
    Oracle = 2
    MicrosoftSqlServer = 4
    MicrosoftAccess = 5
    Symfoware = 10
    Other = 3

End Enum

' =========================================================
' ▽クエリーのリテラル種類
'
' 概要　　　：クエリーのリテラルの種類をあらわす列挙型
'
' =========================================================
Public Enum QueryLiteralType
    
    Number = 0
    String_ = 1
    Date = 2
    Lob = 3
    Any_ = 4
    
End Enum

' カラム書式情報　置換文字　value
Public Const COLUMN_FORMAT_REPLACE_CHAR_VALUE       As String = "$value"
' カラム書式情報　置換文字　column
Public Const COLUMN_FORMAT_REPLACE_CHAR_COLUMN      As String = "$column"
' カラム書式情報　置換文字　exclude
Public Const COLUMN_FORMAT_REPLACE_CHAR_EXCLUDE     As String = "$exclude"
' カラム書式情報　置換文字　direct
Public Const COLUMN_FORMAT_REPLACE_CHAR_DIRECT      As String = "$direct"

' =========================================================
' ▽DBMS名取得
'
' 概要　　　：DBMS名を取得する。
' 引数　　　：dbms DBMS種類
' 戻り値　　：DBMS名
' 特記事項　：
'
' =========================================================
Public Function getDbmsTypeName(ByVal dbms As DbmsType) As String

    Select Case dbms
    
        Case DbmsType.MySQL
            getDbmsTypeName = "MySQL"
    
        Case DbmsType.PostgreSQL
            getDbmsTypeName = "PostgreSQL"
        
        Case DbmsType.Oracle
            getDbmsTypeName = "Oracle"
        
        Case DbmsType.MicrosoftSqlServer
            getDbmsTypeName = "MicrosoftSqlServer"
        
        Case DbmsType.MicrosoftAccess
            getDbmsTypeName = "MicrosoftAccess"
        
        Case DbmsType.Symfoware
            getDbmsTypeName = "Symfoware"
        
        Case DbmsType.Other
            getDbmsTypeName = "Other"
        
    End Select

End Function

' =========================================================
' ▽スキーマ名＋テーブル名の結合
'
' 概要　　　：スキーマ名＋テーブル名を結合する
' 引数　　　：schemaName        スキーマ名
' 　　　　　　tableName         テーブル名
'             schemaUse         スキーマ使用有無（1：スキーマ未使用、2：スキーマ使用）
' 特記事項　：
'
' =========================================================
Public Function concatSchemaTable(ByVal schemaName As String _
                                , ByRef tableName As String _
                                , ByVal schemaUse As Long) As String


    If schemaUse = 1 Then
    
        concatSchemaTable = tableName
        Exit Function
    End If

    If schemaName = "" Then
        concatSchemaTable = tableName
    Else
        concatSchemaTable = schemaName & "." & tableName
    End If

End Function

' =========================================================
' ▽""→NULL
'
' 概要　　　：引数 val が""空文字列の場合、"NULL"に変換する。
' 引数　　　：val 値
' 戻り値　　：変換された文字列
'
' =========================================================
Public Function convertEmptyToNull(ByRef val As String) As String

    If val = "" Then
    
        convertEmptyToNull = "NULL"
    Else
    
        convertEmptyToNull = val
    End If

End Function

' =========================================================
' ▽文字列の両端をシングルクォートで囲む
'
' 概要　　　：引数 val の両端に'(シングルクォート)を付加し戻り値として返す。
' 引数　　　：val 値
' 戻り値　　：'(シングルクォート)が付加された文字列
'
' =========================================================
Public Function encloseSingleQuart(ByRef val As String) As String

    encloseSingleQuart = "'" & val & "'"
End Function

' =========================================================
' ▽エスケープ文字のエスケープ
'
' 概要　　　：引数 val にDB固有のエスケープ文字が含まれている場合
' 　　　　　　エスケープする。
' 引数　　　：dbms        DBMS種類
' 　　　　　　val         値
' 特記事項　：MySQLやPostgresqlの場合、"\"がデフォルトの
' 　　　　　　　エスケープ文字として指定されているが
' 　　　　　　Oracle等は、特にデフォルトのエスケープ文字はない。
'
' =========================================================
Public Function escapeValueForEscapeChar(ByVal dbms As DbmsType _
                                       , ByRef val As String) As String

    Select Case dbms
    
        Case DbmsType.MySQL
            escapeValueForEscapeChar = replace(val, "\", "\\")
    
        Case DbmsType.PostgreSQL
            escapeValueForEscapeChar = replace(val, "\", "\\")
        
        Case DbmsType.Oracle
            escapeValueForEscapeChar = val
        
        Case DbmsType.MicrosoftSqlServer
            escapeValueForEscapeChar = val
        
        Case DbmsType.MicrosoftAccess
            escapeValueForEscapeChar = val
        
        Case DbmsType.Symfoware
            escapeValueForEscapeChar = val
        
        Case DbmsType.Other
            escapeValueForEscapeChar = val
        
    End Select

End Function

' =========================================================
' ▽シングルクォートのエスケープ
'
' 概要　　　：引数 val にシングルクォート文字が含まれている場合
' 　　　　　　エスケープする。
' 引数　　　：dbms        DBMS種類
' 　　　　　　val         値
' 特記事項　：
'
' =========================================================
Public Function escapeValueForSinglequart(ByVal dbms As DbmsType _
                                        , ByRef val As String) As String

    Select Case dbms
    
        Case DbmsType.MySQL
            escapeValueForSinglequart = replace(val, "'", "''")
    
        Case DbmsType.PostgreSQL
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.Oracle
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.MicrosoftSqlServer
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.MicrosoftAccess
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.Symfoware
            escapeValueForSinglequart = replace(val, "'", "''")
        
        Case DbmsType.Other
            escapeValueForSinglequart = replace(val, "'", "''")
        
    End Select

End Function

' =========================================================
' ▽テーブル・カラム名のエスケープ
'
' 概要　　　：テーブル・カラム名をエスケープする
' 引数　　　：dbms        DBMS種類
' 　　　　　　val         値
'             isEscape    エスケープの実行有無
' 特記事項　：
'
' =========================================================
Public Function escapeTableColumn(ByVal dbms As DbmsType _
                                , ByRef val As String _
                                , ByVal isEscape As Boolean) As String

    ' エスケープしない場合は、そのまま返却する
    If Not isEscape Then
        escapeTableColumn = val
        Exit Function
    End If

    ' 値が空文字列の場合、そのまま返却する
    If val = "" Then
        escapeTableColumn = val
        Exit Function
    End If

    Select Case dbms
    
        Case DbmsType.MySQL
            escapeTableColumn = "`" & replace(val, "`", "``") & "`"
    
        Case DbmsType.PostgreSQL
            escapeTableColumn = """" & replace(val, """", """""") & """"
        
        Case DbmsType.Oracle
            escapeTableColumn = """" & replace(val, """", """""") & """"
        
        Case DbmsType.MicrosoftSqlServer
            ' 閉じかっこ "]" のみをエスケープする
            escapeTableColumn = "[" & replace(val, "]", "]]") & "]"
        
        Case DbmsType.MicrosoftAccess
            ' 閉じかっこ "]" のみをエスケープする
            escapeTableColumn = "[" & replace(val, "]", "]]") & "]"
        
        Case DbmsType.Symfoware
        
    End Select

End Function

' =========================================================
' ▽LIKE関数のESCAPE付加
'
' 概要　　　：クエリー式にLIKE関数のESCAPEを付加する。
' 引数　　　：dbms        DBMS種類
' 　　　　　　escapeChar  エスケープ文字
' 戻り値　　：変換後のクエリー式
' 特記事項　：
'
' =========================================================
Public Function addLikeEscape(ByVal dbms As DbmsType _
                            , Optional ByRef escapeChar As String = "\") As String

    addLikeEscape = " ESCAPE '" & escapeChar & "'"

End Function

' =========================================================
' ▽文字データ型かを判定
'
' 概要　　　：
' 引数　　　：dbms        DBMS種類
' 　　　　　　dataType    データ型名
' 戻り値　　：True 文字データ型、False それ以外
' 特記事項　：
'
' =========================================================
Public Function isCharType(ByVal dbms As DbmsType, ByVal dataType As String) As Boolean

    If InStr(dataType, "CHAR") > 0 Or InStr(dataType, "TEXT") > 0 Then
        isCharType = True
    Else
        isCharType = False
    End If

End Function

' =========================================================
' ▽クエリー値変換
'
' 概要　　　：クエリーの値（カラムに対する代入値）を変換する。
' 　　　　　　column = value のvalue部分は
' 　　　　　　文字列の場合はシングルクォート(')で囲む必要があり
' 　　　　　　そういった場合に、変換を実施する。
'
' 引数　　　：dbms            DBMS種類
' 　　　　　　literalType     リテラル種類
' 　　　　　　value           値
' 　　　　　　isEscapeChar    エスケープ文字をエスケープするフラグ
' 　　　　　　directInputChar 直接入力文字
' 戻り値　　：変換後の値
'
' =========================================================
Public Function convertQueryLiteral(ByVal dbms As DbmsType _
                                  , ByVal literalType As QueryLiteralType _
                                  , ByVal value As String _
                                  , Optional ByVal isEscapeChar As Boolean = True _
                                  , Optional ByVal directInputChar As String = "") As String

    ' 直接入力文字の判定
    If directInputChar <> "" And InStr(value, directInputChar) = 1 Then
    
        ' 先頭 1文字目が directInputChar と一致する場合、2文字目以降を取得し戻り値として設定
        convertQueryLiteral = Mid$(value, 2)
    
    ' 特殊文字
    ElseIf isSpecialValue(dbms, value) = True Then
    
        convertQueryLiteral = value
    
    ' 文字列型
    ElseIf literalType = String_ Then

        ' シングルクォートをエスケープする
        convertQueryLiteral = DBUtil.escapeValueForSinglequart(dbms, value)
        ' エスケープ文字をエスケープする
        If isEscapeChar = True Then
            convertQueryLiteral = DBUtil.escapeValueForEscapeChar(dbms, convertQueryLiteral)
        End If
        ' シングルクォートで囲む
        convertQueryLiteral = DBUtil.encloseSingleQuart(convertQueryLiteral)

    ' 時間型
    ElseIf literalType = Date Then

        ' シングルクォートをエスケープする
        convertQueryLiteral = DBUtil.escapeValueForSinglequart(dbms, value)
        ' エスケープ文字をエスケープする
        If isEscapeChar = True Then
            convertQueryLiteral = DBUtil.escapeValueForEscapeChar(dbms, convertQueryLiteral)
        End If
        ' シングルクォートで囲む
        convertQueryLiteral = DBUtil.encloseSingleQuart(convertQueryLiteral)

    ' 数値型
    ElseIf literalType = Number Then

        convertQueryLiteral = value
    
    ' 上記以外
    Else

        convertQueryLiteral = value
    
    End If

End Function

' =========================================================
' ▽クエリー値 更新系変換
'
' 概要　　　：クエリーの値（カラムへの代入値）を変換する。
' 　　　　　　column = value のvalue部分は
' 　　　　　　SQLでは文字列型の場合、値の両端をシングルクォート(')で囲む必要がある。
' 　　　　　　そういった場合に、本メソッドを用いて変換を実施する。
'
' 　　　　　　本メソッドにおける変換の仕組みは
' 　　　　　　書式情報 updateFormat の置換変数を置換することで実現する。
' 　　　　　　updateFormatは、TO_DATE($value, 'xxxxx') といった、内部に置換変数を含んだ文字列になっている。
'
' 引数　　　：dbms                  DBMS種類
' 　　　　　　updateFormat          更新書式情報
' 　　　　　　value                 値
' 　　　　　　isEscapeChar          エスケープ文字をエスケープするフラグ
' 　　　　　　directInputCharPrefix 直接入力文字接頭辞
' 　　　　　　directInputCharSuffix 直接入力文字接尾辞
' 　　　　　　nullInputChar         NULL入力文字
' 戻り値　　：変換後の値
'
' =========================================================
Public Function convertUpdateFormat(ByVal dbms As DbmsType _
                                  , ByVal updateFormat As String _
                                  , ByVal value As String _
                                  , Optional ByVal isEscapeChar As Boolean = True _
                                  , Optional ByVal directInputCharPrefix As String = "" _
                                  , Optional ByVal directInputCharSuffix As String = "" _
                                  , Optional ByVal nullInputChar As String = "") As String

    ' 直接入力文字の判定
    If directInputCharPrefix <> "" And _
       directInputCharSuffix <> "" And _
       InStr(value, directInputCharPrefix) = 1 And _
       InStrRev(value, directInputCharSuffix) = Len(value) Then
    
        ' 先頭 1文字目と最後の文字が directInputChar と一致する場合、囲まれた文字を取り出して設定
        convertUpdateFormat = Mid$(value, 2, Len(value) - 2)
    
    ElseIf directInputCharPrefix <> "" And _
           InStr(value, directInputCharPrefix) = 1 Then
    
        ' 先頭 1文字目が directInputChar と一致する場合、2文字目以降を取得し戻り値として設定
        convertUpdateFormat = Mid$(value, 2)
    
    ' NULL入力文字
    ElseIf nullInputChar <> "" And _
           UCase$(nullInputChar) = UCase$(value) Then
    
        convertUpdateFormat = "NULL"
        
    ' 直接入力形式
    ElseIf updateFormat = COLUMN_FORMAT_REPLACE_CHAR_DIRECT Then
    
        convertUpdateFormat = value
    
    Else
    
        ' エスケープ文字をエスケープする
        If isEscapeChar = True Then
            value = DBUtil.escapeValueForEscapeChar(dbms, value)
        End If
        
        ' シングルクォートが文字列に含まれている場合、エスケープする
        If InStr(value, "'") <> 0 Then
        
            value = DBUtil.escapeValueForSinglequart(dbms, value)
        End If
    
        ' 書式情報の置換変数をvalue値で変換する
        convertUpdateFormat = replace(updateFormat, COLUMN_FORMAT_REPLACE_CHAR_VALUE, value)
    
    End If
    
End Function

' =========================================================
' ▽カラム値 参照系変換
'
' 概要　　　：SELECT SQLにおける、カラム句の変換を実施する
'
' 　　　　　　本メソッドにおける変換の仕組みは
' 　　　　　　書式情報 selectFormat の置換変数を置換することで実現する。
' 　　　　　　selectFormatは、TO_DATE($column, 'xxxxx') といった、内部に置換変数を含んだ文字列になっている。
'
' 引数　　　：dbms            DBMS種類
' 　　　　　　selectFormat    更新書式情報
' 　　　　　　column          値
' 戻り値　　：変換後の値
'
' =========================================================
Public Function convertSelectFormat(ByVal dbms As DbmsType _
                                  , ByVal selectFormat As String _
                                  , ByVal column As String) As String

    ' 書式情報の置換変数をcolumn値で変換する
    convertSelectFormat = replace(selectFormat, COLUMN_FORMAT_REPLACE_CHAR_COLUMN, column)
    
End Function

' =========================================================
' ▽特殊文字判定
'
' 概要　　　：NULL等の特殊な文字であるかを判定する。
' 引数　　　：dbms    DBMS種類
' 　　　　　　value   データ値
'
' 戻り値　　：True 特殊文字
'
' =========================================================
Public Function isSpecialValue(ByVal dbms As DbmsType _
                              , ByVal value As String) As Boolean

    isSpecialValue = False
    
    ' "NULL"という文字列の場合
    If UCase(value) = "NULL" Then
    
        isSpecialValue = True
    
    End If

End Function

' =========================================================
' ▽リストの要素をカンマ区切りの文字列に変換する
'
' 概要　　　：リストの要素をカンマ区切りの文字列に変換する
'
' 引数　　　：dbms DB種類
' 　　　　　　list リスト
'
' =========================================================
Public Function convertListToCommaStr(ByVal dbms As DbmsType _
                                    , ByRef list As collection) As String

    ' 区切り文字
    Const DELIM_STR As String = ", "

    ' 戻り値
    Dim ret As String
    
    ' リストの要素
    Dim value     As Variant
    Dim valueConv As String
    
    For Each value In list
    
        valueConv = CStr(value)
        
        ' シングルクォートをエスケープする
        valueConv = DBUtil.escapeValueForSinglequart(dbms, valueConv)
        ' エスケープ文字をエスケープする
        valueConv = DBUtil.escapeValueForEscapeChar(dbms, valueConv)
        ' シングルクォートで囲む
        valueConv = DBUtil.encloseSingleQuart(valueConv)

        ret = ret & DELIM_STR & valueConv
    
    Next
    
    ' 前方に付加された冗長な文字列を除去する
    ret = replace(ret, DELIM_STR, "", , 1)
    
    convertListToCommaStr = ret

End Function

' =========================================================
' ▽スキーマ・テーブル名抽出
'
' 概要　　　：任意の文字列からスキーマ・テーブル名を抽出する。
'
' 　　　　　　文字列はドット(.)で区切られることを前提としており
' 　　　　　　ドットの左側・右側がそれぞれスキーマ・テーブル名として抽出される。
' 　　　　　　ドットが存在しない場合、テーブル名のみ抽出される。
'
' 引数　　　：val    入力　スキーマ・テーブル名
' 　　　　　　schema 出力　スキーマ名
' 　　　　　　table  出力　テーブル名
' 戻り値　　：
'
' =========================================================
Public Sub extractSchemaTable(ByVal val As String _
                            , ByRef schema As String _
                            , ByRef table As String)

    ' ドット(.)で区切られた文字列を取得する
    '
    ' 例：スキーマとテーブルがドットで連結された文字列を分解する。[schema].[table]
    ' 　schema = [schema]
    ' 　table  = [table]
    '
    Dim splitStr() As String
    
    splitStr = Split(val, ".")
    
    ' 分割された配列が2つ以上の場合
    If VBUtil.arraySize(splitStr) >= 2 Then
    
        ' ドットの左側
        schema = splitStr(LBound(splitStr))
        ' ドットの右側
        table = splitStr(UBound(splitStr))
        
    Else
    
        schema = ""
        table = val
        
    End If


End Sub

' =========================================================
' ▽レコードから情報を取得する。
'
' 概要　　　：
'
' 引数　　　：val    値
' 戻り値　　：値を取得する
'
' =========================================================
Public Function GetRecordValue(ByRef val As Variant) As Variant

    ' チャンクサイズ
    Const CHUNK_SIZE As Long = 1024

    Dim actualSize As Long ' 実際のサイズ
    Dim offset     As Long ' オフセット

    If val.Attributes And &H80 Then
        ' GetChunkで情報を取得すべき場合（adFldLongの場合）
        
        ' 合計サイズを取得する
        actualSize = val.actualSize
        
        ' GetChunkで情報を全て取得する
        Do While offset < actualSize
            GetRecordValue = GetRecordValue & val.GetChunk(CHUNK_SIZE)
            offset = offset + CHUNK_SIZE
        Loop
        
    Else
        ' 通常の場合
        GetRecordValue = val
    
    End If

End Function
