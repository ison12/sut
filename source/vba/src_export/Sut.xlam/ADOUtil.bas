Attribute VB_Name = "ADOUtil"
Option Explicit

' *********************************************************
' ADOを簡便に利用するためのユーティリティモジュール
'
' 作成者　：Ison
' 履歴　　：2007/12/01　新規作成
'
' 特記事項：Microsoft ActiveX Data Objects Libraryを参照
' 　　　　　Ver 2.5 で動作確認
' *********************************************************

' ObjectStateEnumのコピー
' オブジェクトを開いているか閉じているか､データ ソースに接続中か､コマンドを実行中か､またはデータを取得中かどうかを表します｡
' http://msdn.microsoft.com/ja-jp/library/cc389847.aspx
Public Enum ADOConnectStatusConstants

    adStateClosed = 0        ' オブジェクトが閉じていることを示します。
    adStateOpen = 1          ' オブジェクトが開いていることを示します。
    adStateConnecting = 2    ' オブジェクトが接続していることを示します。
    adStateExecuting = 4     ' オブジェクトがコマンドを実行中であることを示します。
    adStateFetching = 8      ' オブジェクトの行が取得されていることを示します。

End Enum

' CursorTypeEnumのコピー
' http://msdn.microsoft.com/ja-jp/library/cc389787.aspx
Public Enum ADOCursorTypeEnum

    adOpenDynamic = 2         ' 動的カーソルを使います。ほかのユーザーによる追加、変更、および削除を確認できます。プロバイダがブックマークをサポートしていない場合を除き、Recordset 内でのすべての動作を許可します。
    adOpenForwardOnly = 0     ' 既定値です。前方専用カーソルを使います。レコードのスクロール方向が前方向に限定されていることを除き、静的カーソルと同じ働きをします。Recordset のスクロールが 1 回だけで十分な場合は、これによってパフォーマンスを向上できます。
    adOpenKeyset = 1          ' キーセット カーソルを使います。ほかのユーザーが追加したレコードは表示できない点を除き、動的カーソルと同じく、自分の Recordset からほかのユーザーが削除したレコードはアクセスできません。ほかのユーザーが変更したデータは表示できます。
    adOpenStatic = 3          ' キーセット カーソルを開きます。データの検索またはレポートの作成に使用するための、レコードの静的コピーです。ほかのユーザーによる追加、変更、または削除は表示されません。
    adOpenUnspecified = -1    '  カーソルの種類を指定しません。

End Enum

' =========================================================
' ▽DB接続関数
'
' 概要　　　：DBに接続する
' 引数　　　：connString 接続文字列
'
' 戻り値　　：コネクションオブジェクト
'
' =========================================================
Public Function connectDb(ByVal connString As String) As Object

    Dim conn As Object
    
    Set conn = CreateObject("ADODB.Connection")
    
    conn.ConnectionString = connString
        
    conn.Open
    
    Set connectDb = conn
    
End Function

' =========================================================
' ▽DB切断関数
'
' 概要　　　：アクティブなDBを切断する
' 引数　　　：conn コネクションオブジェクト
'
' =========================================================
Public Sub closeDB(ByRef conn As Object)

    If Not conn Is Nothing Then
    
        conn.Close
    End If
    
    Set conn = Nothing

End Sub

' =========================================================
' ▽DB名取得
'
' 概要　　　：データソースからDB名を取得する
' 引数　　　：connStr 接続文字列
'
' 戻り値　　：DB名
'
' =========================================================
Public Function getDBName(ByRef conn As Object) As String

    ' データベース名を取得する
    getDBName = conn.defaultdatabase
    
End Function

' =========================================================
' ▽クエリ一括実行
'
' 概要　　　：クエリを一括実行する
' 引数　　　：conn コネクションオブジェクト
' 　　　　　　sql  SQLステートメント
'             cnt  取得件数
'             cursorType カーソルタイプ
'
' 戻り値　　：レコードセットオブジェクト
'
' =========================================================
Public Function queryBatch(ByRef conn As Object _
                          , ByVal sql As String _
                          , Optional ByRef cnt As Long _
                          , Optional ByVal cursorType As ADOCursorTypeEnum = ADOCursorTypeEnum.adOpenForwardOnly) As Object

    On Error GoTo err
    
    ' ▽変数定義
    Dim cmd As Object
    Dim rec As Object
    
    ' ■コマンドオブジェクトを初期化
    Set cmd = CreateObject("ADODB.Command")
    
    cmd.ActiveConnection = conn
    cmd.CommandType = 1 ' adCmdText
    cmd.CommandText = sql
    
    ' クエリーを実行する
    Set rec = cmd.execute(cnt)
    
    ' ■レコードセットを初期化
    'Set rec = CreateObject("ADODB.Recordset")
    
    ' クエリーを実行する
    'rec.Open Source:=cmd, cursorType:=cursorType
    
    ' ○戻り値を設定する
    Set queryBatch = rec
   
    Set cmd = Nothing
    Set rec = Nothing

    Exit Function
err:

    ' エラーハンドラで別の関数を呼び出すとエラー情報が消えてしまうことがあるので
    ' 構造体にエラー情報を保存しておく
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    ADOUtil.closeRecordSet rec
    Set rec = Nothing
    Set cmd = Nothing
        
    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ▽Select実行
'
' 概要　　　：Select文を実行する
' 引数　　　：conn コネクションオブジェクト
' 　　　　　　sql  SQLステートメント
'             cnt  取得件数
'             cursorType カーソルタイプ
'
' 戻り値　　：レコードセットオブジェクト
'
' =========================================================
Public Function querySelect(ByRef conn As Object _
                          , ByVal sql As String _
                          , Optional ByRef cnt As Long _
                          , Optional ByVal cursorType As ADOCursorTypeEnum = ADOCursorTypeEnum.adOpenForwardOnly) As Object

    On Error GoTo err
    
    ' ▽変数定義
    Dim cmd As Object
    Dim rec As Object
    
    ' ■コマンドオブジェクトを初期化
    Set cmd = CreateObject("ADODB.Command")
    
    cmd.ActiveConnection = conn
    cmd.CommandType = 1 ' adCmdText
    cmd.CommandText = sql
    
    ' クエリーを実行する
    'Set rec = cmd.execute(cnt)
    
    ' ■レコードセットを初期化
    Set rec = CreateObject("ADODB.Recordset")
    
    ' クエリーを実行する
    rec.Open Source:=cmd, cursorType:=cursorType
    
    ' ○戻り値を設定する
    Set querySelect = rec
   
    Set cmd = Nothing
    Set rec = Nothing

    Exit Function
err:

    ' エラーハンドラで別の関数を呼び出すとエラー情報が消えてしまうことがあるので
    ' 構造体にエラー情報を保存しておく
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    ADOUtil.closeRecordSet rec
    Set rec = Nothing
    Set cmd = Nothing
        
    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ▽アクションクエリー実行
'
' 概要　　　：Insert・Update・Delete文を実行する
' 引数　　　：conn コネクションオブジェクト
' 　　　　　　sql  SQLステートメント
'
' 戻り値　　：更新件数
'
' =========================================================
Public Function queryAction(ByRef conn As Object, ByVal sql As String) As Long

    ' ▽変数定義
    Dim cmd As Object
    Dim rec As Object
    
    Dim cnt As Long
    
    ' ■コマンドオブジェクトを初期化
    Set cmd = CreateObject("ADODB.Command")
    
    cmd.ActiveConnection = conn
    cmd.CommandType = 1 ' adCmdText
    cmd.CommandText = sql
    
    
    ' ■クエリーを実行する
    Set rec = cmd.execute(cnt)
    
    ' ○戻り値を設定する
    queryAction = cnt
    
    Set cmd = Nothing
    Set rec = Nothing

 End Function

' =========================================================
' ▽レコードセット解放
'
' 概要　　　：アクティブなレコードセットを解放する
' 引数　　　：rec レコードセットオブジェクト
'
' =========================================================
Public Sub closeRecordSet(ByRef rec As Object)

    If Not rec Is Nothing Then
    
        ' レコードセットが開いている場合のみクローズを行う
        If rec.state <> ADOConnectStatusConstants.adStateClosed Then
        
            rec.Close
        End If
        
    End If
    
    Set rec = Nothing
    
    Exit Sub
    
End Sub

' =========================================================
' ▽DBMS種類取得
'
' 概要　　　：コネクションオブジェクトからDBMSの種類を取得する
' 引数　　　：conn コネクションオブジェクト
'
' =========================================================
Public Function getDBMSType(ByRef conn As Object) As DbmsType

    ' データベース名
    Dim dbmsName As String
    
    ' データベース名を取得
    dbmsName = conn.properties.item("DBMS Name")
    
    ' MySQLデータベース
    If InStr(LCase$(dbmsName), "mysql") > 0 Then
    
        getDBMSType = DbmsType.MySQL
    
    ' PostgreSQLデータベース
    ElseIf InStr(LCase$(dbmsName), "postgresql") > 0 Then
    
        getDBMSType = DbmsType.PostgreSQL
    
    ' Oracleデータベース
    ElseIf InStr(LCase$(dbmsName), "oracle") > 0 Then
    
        getDBMSType = DbmsType.Oracle
    
    ' SQL Serverデータベース
    ElseIf InStr(LCase$(dbmsName), "microsoft sql server") > 0 Then
    
        getDBMSType = DbmsType.MicrosoftSqlServer
        
    
    ' Accessデータベース
    ElseIf InStr(LCase$(dbmsName), "access") > 0 Or InStr(LCase$(dbmsName), "ms jet") > 0 Then
    
        getDBMSType = DbmsType.MicrosoftAccess
        
    ' Symfowareデータベース
    ElseIf InStr(LCase$(dbmsName), "symfoware") > 0 Then
    
        getDBMSType = DbmsType.Symfoware
    ' 判別できない場合
    Else
    
        getDBMSType = DbmsType.Other
    
    End If

End Function

' =========================================================
' ▽DBMS種類取得
'
' 概要　　　：コネクションオブジェクトからDBMSの種類を取得する
' 引数　　　：conn コネクションオブジェクト
'
' =========================================================
Public Function getDBMSTypeByConnStr(ByVal connStr As String) As DbmsType

    On Error GoTo err
    
    Dim conn As Object
    
    ' DBに接続する
    Set conn = ADOUtil.connectDb(connStr)
    
    getDBMSTypeByConnStr = ADOUtil.getDBMSType(conn)
    
    ' DBを切断する
    ADOUtil.closeDB conn
    
    Exit Function
    
err:

    ' ■後始末
    ' DBを切断する
    ADOUtil.closeDB conn
    
    Set conn = Nothing
    
End Function

' =========================================================
' ▽コネクションのステータス取得
'
' 概要　　　：コネクションオブジェクトからステータスを取得する
' 引数　　　：conn コネクションオブジェクト
' 戻り値　　：ADOConnectStatusConstants
'
' =========================================================
Public Function getConnectionStatus(ByRef conn As Object) As ADOConnectStatusConstants

    If conn Is Nothing Then
    
        getConnectionStatus = adStateClosed
        Exit Function
    End If

    ' コネクションのステータスを取得する
    getConnectionStatus = conn.state

End Function
