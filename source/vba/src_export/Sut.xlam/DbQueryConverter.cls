VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbQueryConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのクエリを生成するインターフェース。
' 全DBで共通する汎用的な実装インターフェース。
'
' 作成者　：Ison
' 履歴　　：2009/03/28　新規作成
'
' 特記事項：
'
' *********************************************************
' DBコネクション
Public conn          As Object

' 囲み文字 開始
Public encloseStart As String
' 囲み文字 終了
Public encloseEnd   As String

' クエリパラメータリスト
Private queryParamList As ValCollection
' 変換対象パラメータリスト
Private convertedParamList As ValCollection
' ワークシート
Private sheet As Worksheet
' 初期化フラグ
Private isInit As Boolean

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    encloseStart = ConstantsTable.QUERY_PARAMETER_ENCLOSE_START
    encloseEnd = ConstantsTable.QUERY_PARAMETER_ENCLOSE_END
    
    Set queryParamList = New ValCollection
    Set convertedParamList = New ValCollection
    
    isInit = False

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ▽初期化処理メソッド
' =========================================================
Public Sub init(ByRef conn As Object, ByRef queryParamList_ As ValQueryParameterList)

    ' DBコネクション
    Set Me.conn = conn
    ' クエリパラメータリスト
    Dim queryParameter As ValQueryParameter
    For Each queryParameter In queryParamList_.list.col
    
        queryParamList.setItem queryParameter, queryParameter.name
    
    Next
    
    isInit = True
    
End Sub

' =========================================================
' ▽初期化処理メソッド（初回のみ実行）
' =========================================================
Public Sub initIfFirst(ByRef conn As Object, ByRef queryParamList_ As ValQueryParameterList)

    If isInit = False Then
        Call init(conn, queryParamList_)
    End If
    
End Sub


' =========================================================
' ▽SQL変換実行
'
' 概要　　　：SQL変換を実行する。
' 引数　　　：sql SQL
'
' 戻り値　　：変換後のSQL
'
' 特記事項　：
'
'
' =========================================================
Public Function convertSql(ByRef sql As String) As String

    If queryParamList.count <= 0 Then
        ' クエリパラメータリストがない場合は、何もしない
        convertSql = sql
        Exit Function
    End If

    ' 変換対象パラメータ
    Dim convertedParam           As Variant
    Dim convertedQueryParamName  As String
    Dim convertedQueryParamValue As String
    
    Dim queryParameter  As ValQueryParameter
    
    Dim mm         As Object
    
    Dim i As Long
    Dim match      As Object
    Dim matches    As Object
    Dim matcheSub  As Object
    
    Dim matchParamName As String

    ' 正規表現関連オブジェクト
    ' RegExpオブジェクトの作成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    ' 正規表現の指定
    With reg
        ' パターンを指定
        .Pattern = "\" & encloseStart & "(.+?)" & "\" & encloseEnd
        ' 大文字と小文字を区別するか(False)、しないか(True)
        .IgnoreCase = False
        ' 文字列全体を検索するか(True)、しないか(False)
        .Global = True
    End With
    
    ' 正規表現でのマッチングを実行
    Set matches = reg.execute(sql)
    
    If matches.count <= 0 Then
        convertSql = sql
        Exit Function
    End If
    
    For i = 0 To matches.count - 1
    
        Set match = matches.item(i)
        Set matcheSub = match.submatches
        matchParamName = matcheSub.item(0)
        
        If queryParamList.exist(matchParamName) = False Then
            ' 該当するクエリパラメータが存在しない
            GoTo continue
        End If
    
        Set queryParameter = queryParamList.getItem(matchParamName)
    
        If convertedParamList.exist(matchParamName) = False Then
            ' 変換済みクエリパラメータが存在しない
        
            If isShouldExecuteSql(queryParameter.value) Then
                convertedQueryParamValue = executeSql(queryParameter.value)
            Else
                convertedQueryParamValue = queryParameter.value
            End If
            
            convertedParam = Array(queryParameter.name, convertedQueryParamValue)
            convertedParamList.setItem convertedParam, queryParameter.name
            
        Else
            ' 既に変換済みクエリパラメータが存在する
            convertedParam = convertedParamList.getItem(matchParamName, vbVariant)
            
        End If
        
        ' 変換対象パラメータで変換を実行する
        ' SQLを変換する
        convertedQueryParamName = convertedParam(0)
        convertedQueryParamValue = convertedParam(1)
        
        sql = replace(sql, encloseStart & convertedQueryParamName & encloseEnd, convertedQueryParamValue)
        
continue:
    
    Next
    
    convertSql = sql
    
End Function

' =========================================================
' ▽SQL実行
'
' 概要　　　：SQLを実行する。
' 引数　　　：sql SQL
'
' 戻り値　　：実行結果
'
' 特記事項　：
'
'
' =========================================================
Public Function executeSql(ByRef sql As String) As String

    ' レコードセット
    Dim recordSet   As Object
    ' クエリーを実行する
    Set recordSet = ADOUtil.querySelect(conn, sql, 0)
    
    ' 戻り値
    Dim result         As Variant

    ' レコードセットがEOFではない場合
    If Not recordSet.EOF Then
    
        ' レコードセットから全レコードを取得する
        result = recordSet.getRows(1)
    
        executeSql = result(0, 0)
    Else
        ' 空を返す
        executeSql = Empty
    End If

End Function

' =========================================================
' ▽SELECTを実行するかどうかの判定処理
'
' 概要　　　：SELECTを実行するかどうかの判定処理を行う。
'             先頭文字が"SELECT"で始まる場合
' 引数　　　：sql SQL
'
' 戻り値　　：SELECTを実行するかどうかの判定フラグ
'
' 特記事項　：
'
'
' =========================================================
Public Function isShouldExecuteSql(ByRef sql As String) As Boolean

    ' 正規表現関連オブジェクト
    ' RegExpオブジェクトの作成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    ' 正規表現の指定
    With reg
        .Pattern = "SELECT "         ' パターンを指定
        .IgnoreCase = True           ' 大文字と小文字を区別するか(False)、しないか(True)
        .Global = True               ' 文字列全体を検索するか(True)、しないか(False)
    End With
    
    Dim matches    As Object
    Dim matcheSub  As Object
    Dim matcheText As String
    
    Set matches = reg.execute(sql) ' 正規表現でのマッチングを実行
    
    If matches.count > 0 Then
        isShouldExecuteSql = True
    Else
        isShouldExecuteSql = False
    End If

End Function
