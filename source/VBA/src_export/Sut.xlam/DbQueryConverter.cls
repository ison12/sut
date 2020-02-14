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
Private queryParamList As ValQueryParameterList
' 変換対象パラメータリスト
Private convertedParamList As ValCollection
' ワークシート
Private sheet As Worksheet

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    encloseStart = ConstantsTable.QUERY_PARAMETER_ENCLOSE_START
    encloseEnd = ConstantsTable.QUERY_PARAMETER_ENCLOSE_END
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ▽初期化処理メソッド
' =========================================================
Public Function init(ByRef conn As Object, ByVal queryParamList_ As ValQueryParameterList)

    ' DBコネクション
    Set Me.conn = conn
    ' クエリパラメータリスト
    Set queryParamList = queryParamList_
    
End Function

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

    ' 変換対象パラメータ
    Dim convertedParam As Variant
    
    ' クエリパラメータの定義情報
    Dim bookPropVal As ValCollection
    
    Dim queryParameter  As ValQueryParameter
    Dim queryParamName  As String
    Dim queryParamValue As String

    If convertedParamList Is Nothing Then
    
        Set convertedParamList = New ValCollection
        
        For Each queryParameter In queryParamList.list.col
        
            If isShouldExecuteSql(queryParameter.value) Then
                queryParameter.value = executeSql(queryParameter.value)
            End If
        
            convertedParamList.setItem Array(queryParameter.name, queryParameter.value), queryParameter.name
        
        Next
        
    End If

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
    
    Dim mm         As Object
    Dim matches    As Object
    Dim matcheSub  As Object
    Dim matcheText As String
    
    ' 正規表現でのマッチングを実行
    Set matches = reg.execute(sql)
    
    If matches.count <= 0 Then
        convertSql = sql
        Exit Function
    End If

    ' 変換対象パラメータで変換を実行する
    For Each convertedParam In convertedParamList.col
        ' SQLを変換する
        queryParamName = convertedParam(0)
        queryParamValue = convertedParam(1)
        
        sql = replace(sql, encloseStart & queryParamName & encloseEnd, queryParamValue)
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
