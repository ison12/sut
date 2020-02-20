VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSnapQueryExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スナップショットのクエリを実行するクラス
'
' 作成者　：Ison
' 履歴　　：2019/01/05　新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' イベント
' ________________________________________________________
' =========================================================
' ▽SQLリスト実行イベント（開始）
'
' 概要　　　：SQL実行時に呼び出されるイベント
' 引数　　　：count 件数
'
' =========================================================
Public Event ExecuteSqlListBegin(ByVal count As Long)

' =========================================================
' ▽SQLリスト実行イベント（完了）
'
' 概要　　　：SQL実行時に呼び出されるイベント
' 引数　　　：
'
' =========================================================
Public Event ExecuteSqlListComplete()

' =========================================================
' ▽SQL実行イベント
'
' 概要　　　：SQL実行する前に呼び出されるイベント
' 引数　　　：sql        SQLステートメント
'
' =========================================================
Public Event ExecuteSqlBegin(ByRef sql As String)

' =========================================================
' ▽SQL実行成功イベント
'
' 概要　　　：SQL実行に成功した際に呼び出されるイベント
' 引数　　　：sql        SQLステートメント
'
' =========================================================
Public Event ExecuteSqlSuccessful(ByRef sql As String)

' =========================================================
' ▽SQL実行時のエラーイベント
'
' 概要　　　：SQL実行時にエラーが発生した際に呼び出されるイベント
' 引数　　　：sheet          ワークシート
' 　　　　　：recordSetIndex レコードセットインデックス
' 　　　　　：sql            SQLステートメント
' 　　　　　：sqlstate       SQLステータス
' 　　　　　：message        メッセージ
'
' =========================================================
Public Event ExecuteSqlError(ByRef sheet As Worksheet, ByVal recordSetIndex As Long, ByRef sql As String, ByRef sqlState As String, ByRef message As String)

' ---------------------------------------------------------
' □プロパティ
' ---------------------------------------------------------
' DBコネクション
Public conn          As Object
' キャンセルフラグ
Public canceled      As Boolean

' SQL実行中フラグ
Private executingSql      As Boolean

' アプリケーション設定情報
Private m_applicationSetting As ValApplicationSetting

' クエリ変換
Private queryConverter As New DbQueryConverter
' クエリパラメータリスト
Private queryParameterList_ As ValQueryParameterList


' =========================================================
' ▽アプリケーション設定情報
' =========================================================
Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
End Property

' =========================================================
' ▽クエリパラメータリスト取得プロパティ
' =========================================================
Public Property Get queryParameterList(ByRef sheet As Worksheet) As ValQueryParameterList

    If Not queryParameterList_ Is Nothing Then
        Set queryParameterList = queryParameterList_
        Exit Property
    End If

    Set queryParameterList_ = Nothing
        
    If sheet Is Nothing Then
        Exit Property
    End If
    
    Set queryParameterList_ = New ValQueryParameterList
    queryParameterList_.init sheet.parent
    queryParameterList_.readForData
    
    Set queryParameterList = queryParameterList_
    
End Property

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ▽クエリ実行
'
' 概要　　　：対象となるクエリを実行する
' 引数　　　：sqlDefineWorkSheet  クエリリスト
' 戻り値　　：レコード取得情報
'
' =========================================================
Public Function execute(ByRef sqlDefineWorkSheet As ValSnapSqlDefineWorkSheet) As ValSnapRecordsSet

    On Error GoTo err

    ' 影響を与えたレコード
    Dim affectedRecord As Long
    
    Dim i, j As Long
    
    ' ヘッダ配列
    Dim headerList As Variant
    ' フィールド
    Dim field   As Object
    
    ' レコードセット
    Dim recordSet   As Object
    ' レコード配列
    Dim recordList  As Variant
    ' レコード配列（行サイズ）
    Dim recordListSize As Long
    ' レコード総件数
    Dim recordAllCount As Long

    ' レコードセットインデックス
    Dim recordSetIndex As Long
    
    If sqlDefineWorkSheet.snapSqlDefine.sqlList.count <= 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_SQL_EMPTY _
                , err.Source _
                , ConstantsError.ERR_DESC_SQL_EMPTY
    End If

    ' レコードセットインデックス
    recordSetIndex = 1
    
    ' キャンセルフラグをチェックする
    checkCancel
    
    ' トランザクションオブジェクト
    Dim transaction As DBTransaction
    Set transaction = Nothing
    
    Set transaction = New DBTransaction
    transaction.init conn
        
    Dim varOfSqlList   As Variant
    Dim varOfParamList As Variant
    
    Dim sqlDefineSql   As ValSnapSqlDefineSql
    
    Dim sqlTemplate As String
    Dim sql As String
    
    executingSql = False
    
    RaiseEvent ExecuteSqlListBegin(sqlDefineWorkSheet.snapSqlDefine.sqlList.count)
    
    ' 戻り値を初期化
    Set execute = New ValSnapRecordsSet
    recordAllCount = 0
    
    Dim snapRecords As ValSnapRecords
    
    For Each varOfSqlList In sqlDefineWorkSheet.snapSqlDefine.sqlList.col
    
        ' キャンセルフラグをチェックする
        checkCancel
    
        Set sqlDefineSql = varOfSqlList
        
        ' -------------------------------------------------------------
        ' SQL（文字列）を生成する（パラメータの置き換え）
        ' -------------------------------------------------------------
        sqlTemplate = sqlDefineSql.sql
        sql = sqlTemplate
        
        queryConverter.initIfFirst Me.conn, queryParameterList(sqlDefineWorkSheet.sheet)
        sql = queryConverter.convertSql(sql)
    
        ' -------------------------------------------------------------
        ' SQLを実行する
        ' -------------------------------------------------------------
        ' キャンセルフラグをチェックする
        checkCancel
        
        ' 処理のはじめにフラグをオフにしておく
        executingSql = False
    
        RaiseEvent ExecuteSqlBegin(sql)
    
        ' SQL実行中フラグをONにする
        executingSql = True
        
        ' クエリーを実行する
        Set recordSet = ADOUtil.queryBatch(conn, sql, affectedRecord)
        
        ' SQL実行中フラグをOFFにする
        executingSql = False
        
        ' -------------------------------------------------------------
        ' レコードを取得する
        ' -------------------------------------------------------------
        
        ' UPDATEなどの、レコードを返さないクエリの場合 State = 0 (adStateClosed) になる
        If recordSet.state = 0 Then
        
        Else
        
            ' ---------------------------------------------------
            ' ヘッダ情報の取得
            ' ---------------------------------------------------
            ReDim headerList(recordSet.Fields.count - 1)
            
            ' ヘッダ情報の取得
            i = 0
            For Each field In recordSet.Fields
            
                headerList(i) = field.name
                
                i = i + 1
                
            Next field
            
            ' ---------------------------------------------------
            
            ' ---------------------------------------------------
            ' レコード情報の取得
            ' ---------------------------------------------------
            ' 取得レコードが0件になるまでレコードセットを処理する
            recordList = Empty
            Do While Not recordSet.EOF
            
                ' キャンセルフラグをチェックする
                checkCancel
                
                ' レコードセットからレコードリストを取得する
                recordList = getRecordList(recordSet, 0)
                
                ' レコードリストのサイズを取得する
                ' ADOのレコードセットからGetRowsで取得した配列は(列×行)でデータが格納されている
                recordListSize = VBUtil.arraySize(recordList, 2)
        
                ' レコードリストが配列ではない場合
                If IsArray(recordList) = False Then
                
                    ' このブロックに入ることはあるのか？
                    GoTo loop_end
                End If
                
                ' レコードリストを変換する
                convertRecordList recordList
                
                recordAllCount = recordAllCount + recordListSize
            Loop
    
            Set snapRecords = New ValSnapRecords
            snapRecords.sql = sql
            snapRecords.memo = sqlDefineSql.memo
            snapRecords.headerList = headerList
            snapRecords.recordList = recordList
            Set snapRecords.pkList = sqlDefineSql.pkList

            execute.list.setItem snapRecords

        End If
            
loop_end:
            
        recordSetIndex = recordSetIndex + 1
        
        ' レコードセットを解放する
        ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
        RaiseEvent ExecuteSqlSuccessful(sql)

    Next
    
    recordSetIndex = 0
    
    ' 念のためレコードセットを解放する
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    ' トランザクションをコミットする
    If Not transaction Is Nothing Then
        transaction.commit
    End If
    
    ' -------------------------------------------------------
    
    execute.name = sqlDefineWorkSheet.snapSqlDefine.name
    execute.recordCount = recordAllCount
    execute.getDate = Now
    
    RaiseEvent ExecuteSqlListComplete
    
    Exit Function

err:

    ' エラーハンドラで別の関数を呼び出すとエラー情報が消えてしまうことがあるので
    ' 構造体にエラー情報を保存しておく
    Dim errT As errInfo: errT = VBUtil.swapErr

    ' レコードセットを解放する
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    ' SQL実行時にエラーが発生
    If executingSql = True Then

        ' SQLステータス
        Dim sqlState As String
        
        If conn.Errors.count > 0 Then
        
            ' エラーオブジェクトを取得する
            Dim error As Object
            Set error = conn.Errors.item(0)
            sqlState = error.sqlState
        End If

        ' ▼SQL実行時エラー
        RaiseEvent ExecuteSqlError(sqlDefineWorkSheet.sheet, recordSetIndex, sql, sqlState, err.Description)
        
    Else
    
        If recordSetIndex <= 0 Then
            ' 前処理時
        Else
            ' ▼SQL実行時エラー
            RaiseEvent ExecuteSqlError(sqlDefineWorkSheet.sheet, recordSetIndex, sql, Empty, err.Description)
        End If
    
    End If

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ▽レコード取得
'
' 概要　　　：レコードをRecordSetオブジェクトから取得し、配列に格納して返す。
'
' 引数　　　：recordSet レコードセット
' 　　　　　　fetchSize フェッチサイズ
' 戻り値　　：レコード配列
'
' =========================================================
Private Function getRecordList(ByRef recordSet As Object _
                             , Optional ByVal fetchSize As Long = 0) As Variant

    ' 戻り値
    Dim result         As Variant
    
    ' レコードセットがEOFではない場合
    If Not recordSet.EOF Then
    
        ' フェッチサイズをチェックする
        If fetchSize = 0 Then
        
            ' レコードセットから全レコードを取得する
            result = recordSet.getRows
        Else
        
            ' レコードセットからレコードをフェッチサイズ分取得する
            result = recordSet.getRows(fetchSize)
        End If
    
        getRecordList = result
    Else
    
        ' 空を返す
        getRecordList = Empty
    End If

End Function

' =========================================================
' ▽レコードリスト変換
'
' 概要　　　：レコードリストのデータを変換する。
' 引数　　　：recordList   レコードリスト
'
' =========================================================
Private Function convertRecordList(ByRef recordList As Variant)

    ' レコード配列の要素入れ替えによって新たに生成される配列
    Dim temp As Variant
    
    ' レコード配列のデータ変換処理
    ' RecordSet.GetRowsより返された配列は、(col,rec)となっているため(rec,col)となるように入れ替えを実施する
    temp = convertArrayValues(recordList)
    
    ' レコード配列を解放する
    Erase recordList

    ' レコードリストを入れ替える
    recordList = temp

End Function

' =========================================================
' ▽配列要素の変換
'
' 概要　　　：2次元配列の要素を(x,y)から(y,x)に設定しなおす。
' 　　　　　　さらに要素をセル形式の値に変換する。
' 引数　　　：v 2次元配列
'
' 戻り値　　：2次元配列
' 特記事項　：VBUtil.transposeDim関数をカスタマイズした関数。
' 　　　　　　処理はほぼ同じ。
'
' =========================================================
Private Function convertArrayValues(ByRef v As Variant) As Variant
    
    Dim X As Long
    Dim Y As Long
    
    Dim Xlower As Long
    Dim Xupper As Long
    
    Dim Ylower As Long
    Dim Yupper As Long
    
    Dim tempArray As Variant
    
    Xlower = LBound(v, 2)
    Xupper = UBound(v, 2)
    Ylower = LBound(v, 1)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xlower To Xupper, Ylower To Yupper)
    
    Dim value As Variant
    
    For X = Xlower To Xupper
        For Y = Ylower To Yupper
        
            ' 一旦値を取得する
            value = v(Y, X)
            
            ' 文字列の場合、セル形式の値へと変換する
            tempArray(X, Y) = ExcelUtil.convertCellStrValue(value)
        
        Next Y
    Next X
    
    convertArrayValues = tempArray

End Function

' =========================================================
' ▽キャンセルフラグチェック
'
' 概要　　　：キャンセルフラグをチェックしキャンセルされていた場合に
' 　　　　　　エラーを発行する。
'
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub checkCancel()

    If canceled = True Then

        err.Raise ConstantsError.ERR_NUMBER_PROC_CANCEL _
                , err.Source _
                , ConstantsError.ERR_DESC_PROC_CANCEL
    End If

End Sub


