VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' レコードを更新するクラス
'
' 作成者　：Ison
' 履歴　　：2008/02/01　新規作成
'
' 特記事項：
' *********************************************************
'
' ________________________________________________________
' イベント
' ________________________________________________________

' =========================================================
' ▽ExeTableSheetReaderで定義しているイベントを再定義
' 　上位モジュールに、発生したイベントをそのまま伝播するために再定義
' =========================================================
Public Event readColumnListBegin(ByVal cnt As Long)
Public Event readColumn(ByRef column As ValDbDefineColumn)
Public Event readColumnList(ByRef columnList As ValCollection)
Public Event readTableInfoBegin()
Public Event readTableInfo(ByRef table As ValTableWorksheet)
Public Event readRecordListSize(ByVal cnt As Long)
Public Event readRecordListBegin(ByVal cnt As Long)
Public Event readRecordList(ByRef recordList As ValRecordList)
Public Event readRecordListBySelectionBegin(ByVal cnt As Long)
Public Event readRecordListBySelectionOne(ByRef recordList As ValRecordList)
Public Event readRecordListBySelection(ByRef recordLists As ValCollection)

' =========================================================
' ▽SQL実行開始イベント
'
' 概要　　　：SQL実行開始前に呼び出されるイベント
' 引数　　　：count 実行件数
'
' =========================================================
Public Event ExecuteSqlBegin(ByVal count As Long)

' =========================================================
' ▽SQL実行イベント
'
' 概要　　　：SQL実行時に呼び出されるイベント
' 　　　　　　正常に実行されたとしても結果が0件の場合があるため
' 　　　　　　そういった場合には、countを参照して何件更新されているかを確認する。
'
' 引数　　　：tableSheet       テーブルシート
' 　　　　　　recordList       レコードリスト
' 　　　　　　recordListIndex  レコードリストインデックス
'             sql              SQLステートメント
' 　　　　　　count            結果件数
'
' =========================================================
Public Event executeSql(ByRef tableSheet As ValTableWorksheet _
                      , ByRef recordList As ValRecordList _
                      , ByVal recordListIndex As Long _
                      , ByRef sql As String _
                      , ByVal count As Long)

' =========================================================
' ▽SQL実行成功イベント
'
' 概要　　　：SQL実行に成功した際に呼び出されるイベント
' 引数　　　：count 実行件数
'
' =========================================================
Public Event ExecuteSqlSuccessful(ByVal count As Long)

' =========================================================
' ▽SQL実行時エラーイベント
'
' 概要　　　：SQL実行時にエラーが発生した際に呼び出されるイベント
' 引数　　　：tableSheet       テーブルシート
' 　　　　　　recordList       レコードリスト
' 　　　　　　recordListIndex  レコードリストインデックス
'             sql              SQLステートメント
'             sqlstate         SQLステータス
' 　　　　　　errorDesc        エラー情報
'
' =========================================================
Public Event ExecuteSqlError(ByRef tableSheet As ValTableWorksheet _
                           , ByRef recordList As ValRecordList _
                           , ByVal recordListIndex As Long _
                           , ByRef sql As String _
                           , ByRef sqlState As String _
                           , ByRef errorDesc As String)

' =========================================================
' ▽SQL実行キャンセルイベント
'
' 概要　　　：SQLの実行キャンセル時に呼び出されるイベント
' 引数　　　：count 実行件数
'
' =========================================================
Public Event ExecuteSqlCancel(ByVal count As Long)


' ---------------------------------------------------------
' □プロパティ
' ---------------------------------------------------------
' テーブルシート読み込み
Private WithEvents tableSheetReader As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1

' DBコネクション
Public conn          As Object
' ワークシート
Public sheet         As Worksheet
' キャンセルフラグ
Public canceled      As Boolean

' アプリケーション設定情報
Public applicationSetting As ValApplicationSetting
' アプリケーション設定（カラム書式情報）
Public applicationSettingColFmt As ValApplicationSettingColFormat

' カラム書式リスト
Public columnFormatList As ValCollection

' クエリー種類
Public queryType     As Long

' クエリー生成オブジェクト
Private queryCreator As IDbQueryCreator
' クエリ変換
Private DbQueryConverter As New DbQueryConverter
' クエリパラメータリスト
Private queryParameterList_ As ValQueryParameterList

' SQL実行中フラグ
Private executingSql      As Boolean
' SQL実行時にエラーが発生した場合
Private executingSqlError As Boolean

' ファイル出力
Public isFileOutput  As Boolean
' ファイルパス
Public filePath      As String
' 文字コード
Public charSet       As String
' 改行コード
Public newline       As String

' 戻り値用　読み込まれたテーブルシート
' delete all 時に呼び出し元に情報を返却するために使用する
Public readTableSheet As ValTableWorksheet
' 戻り値用　生成されたSQL
' delete all 時に呼び出し元に情報を返却するために使用する
Public createdSql As String

' ファイル出力オブジェクト
Private writer As FileWriter

' =========================================================
' ▽クエリー種類 INSERT + UPDATE
' =========================================================
Public Property Get QUERY_TYPE_INSERT_UPDATE() As Variant
    
    QUERY_TYPE_INSERT_UPDATE = 4
End Property

' =========================================================
' ▽クエリー種類 INSERT
' =========================================================
Public Property Get QUERY_TYPE_INSERT() As Variant
    
    QUERY_TYPE_INSERT = 0
End Property

' =========================================================
' ▽クエリー種類 UPDATE
' =========================================================
Public Property Get QUERY_TYPE_UPDATE() As Variant

    QUERY_TYPE_UPDATE = 1
End Property

' =========================================================
' ▽クエリー種類 DELETE
' =========================================================
Public Property Get QUERY_TYPE_DELETE() As Variant

    QUERY_TYPE_DELETE = 2
End Property

' =========================================================
' ▽クエリー種類 DELETE（全レコード削除）
' =========================================================
Public Property Get QUERY_TYPE_DELETE_ALL() As Variant

    QUERY_TYPE_DELETE_ALL = 3
End Property

' =========================================================
' ▽クエリパラメータリスト取得プロパティ
' =========================================================
Public Property Get queryParameterList() As ValQueryParameterList

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
' ▽更新実行
'
' 概要　　　：テーブルシートからレコード情報を読み込み
' 　　　　　　レコードリストの生成とSQLの生成・実行を繰り返し行う。
' 引数　　　：
' 戻り値　　：実行件数
'
' =========================================================
Public Function execute() As Long

    On Error GoTo err
    
    ' テーブルシート
    Dim tableSheet As ValTableWorksheet
    ' レコードリスト
    Dim recordList As ValRecordList
    
    ' レコード処理件数
    Dim processCount As Long: processCount = applicationSetting.getRecProcessCount

    ' ファイルを出力する場合
    If isFileOutput = True Then
        
        ' ファイルの存在チェックを行い、存在する場合ファイルを削除する
        If VBUtil.isExistFile(filePath) = True Then
        
            Kill filePath
        End If
        
        ' ファイル書き込みオブジェクトを初期化する
        Set writer = New FileWriter
        
        writer.init filePath _
                  , charSet _
                  , newline _
                  , False
    
    End If

    ' テーブルシート読み込みオブジェクトを初期化する
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = sheet
    
    Set tableSheet = tableSheetReader.readTableInfo

    ' 実行件数
    Dim ret     As Long
    
    ' レコードサイズ
    Dim recSize As Long
    ' レコードサイズを取得する
    recSize = tableSheetReader.getRecordSize(tableSheet)

    ' レコードリストのサイズをイベントとして送信する
    RaiseEvent readRecordListSize(recSize)
    
    ' 処理単位が1件以上の場合
    If processCount > 0 Then
    
        ' 処理単位（一時変数）
        Dim tmpProcessCount As Long: tmpProcessCount = processCount
        
        Dim i As Long
    
        ' 初期値は1とする
        i = 1
        
        ' 処理単位ごとに分割して処理
        Do While i <= recSize
            
            ' 処理単位がレコードサイズを上回ると
            ' 後続の処理でエラーが発生してしまうので、処理単位の数を調整する
            If i + tmpProcessCount - 1 > recSize Then
            
                tmpProcessCount = recSize - i + 1
            End If
        
            
            ' レコードリストを取得する
            Set recordList = tableSheetReader.readRecordList(tableSheet _
                                                           , i _
                                                           , tmpProcessCount)
                                                                 
            ' クエリーを実行する
            If isFileOutput = False Then
                ' クエリーを実行する
                ret = ret + executeQuery(tableSheet, recordList)
            
            ' ファイルを出力する
            Else
            
                ret = ret + executeOutputQueryFile(tableSheet, recordList)
            End If
            
            i = i + tmpProcessCount
            
        Loop
        
    ' 処理単位が上記以外
    Else
    
        ' 全レコードまとめて処理する
        
        ' レコードリストを取得する
        Set recordList = tableSheetReader.readRecordListAll(tableSheet)
                                                             
        ' クエリーを実行する
        If isFileOutput = False Then
        
            ret = executeQuery(tableSheet, recordList)
        
        ' ファイルを出力する
        Else
            
            ret = executeOutputQueryFile(tableSheet, recordList)
        End If
    
    End If
    
    ' 戻り値に実行件数を設定
    execute = ret

    ' ファイルを出力する
    If Not writer Is Nothing Then
    
        writer.destroy
        Set writer = Nothing
    End If
    
    Exit Function
err:

    ' ファイルの出力をキャンセルする
    If Not writer Is Nothing Then
    
        writer.cancel
        Set writer = Nothing
    End If
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Function

' =========================================================
' ▽選択領域のレコードの更新を実行する
'
' 概要　　　：テーブルシートからレコード情報を読み込み
' 　　　　　　レコードリストの生成とSQLの生成・実行を繰り返し行う。
' 引数　　　：
' 戻り値　　：実行件数
'
' =========================================================
Public Function executeSelection() As Long

    On Error GoTo err
    
    ' 実行件数
    Dim ret     As Long
    
    ' テーブルシート
    Dim tableSheet  As ValTableWorksheet

    ' レコードリスト【コレクション】
    Dim recordLists As ValCollection
    ' レコードリスト
    Dim recordList As ValRecordList

    ' ファイルを出力する場合
    If isFileOutput = True Then
        
        ' ファイルの存在チェックを行い、存在する場合ファイルを削除する
        If VBUtil.isExistFile(filePath) = True Then
        
            Kill filePath
        End If
        
        ' ファイル書き込みオブジェクトを初期化する
        Set writer = New FileWriter
        
        writer.init filePath _
                  , charSet _
                  , newline _
                  , False
    
    End If

    ' テーブルシート読み込みオブジェクトを初期化する
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = sheet
    
    Set tableSheet = tableSheetReader.readTableInfo

    ' レコードリスト【コレクション】を取得する
    Set recordLists = tableSheetReader.readRecordListBySelectionArea(tableSheet)
    
    ' レコードリスト【コレクション】を1件ずつ処理する
    For Each recordList In recordLists.col
    
        ' クエリーを実行する
        If isFileOutput = False Then
        
            ret = ret + executeQuery(tableSheet, recordList)
        
        ' ファイルを出力する
        Else
            
            ret = ret + executeOutputQueryFile(tableSheet, recordList)
        End If
    
    Next
        
    ' 戻り値に実行件数を設定
    executeSelection = ret

    ' ファイルを出力する
    If Not writer Is Nothing Then

        writer.destroy
        Set writer = Nothing
    End If
    
    Exit Function
err:

    ' ファイルの出力をキャンセルする
    If Not writer Is Nothing Then
        
        writer.cancel
        Set writer = Nothing
    End If
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
        
End Function

' =========================================================
' ▽レコード更新
'
' 概要　　　：テーブル定義とレコードリストからクエリーを生成し実行する。
' 　　　　　　本メソッド内でトランザクションが完結しており
' 　　　　　　正常終了した場合は、コミット。異常終了した場合は、ロールバックされる。
' 引数　　　：tableSheet テーブルシート
' 　　　　　　recordList レコードリスト
' 戻り値　　：実行件数
'
' =========================================================
Private Function executeQuery(ByRef tableSheet As ValTableWorksheet _
                            , ByRef recordList As ValRecordList) As Long

    On Error GoTo err

    ' テーブル定義
    Dim table As ValDbDefineTable
    ' テーブル定義を取得する
    Set table = tableSheet.table
    
    executingSql = False
    executingSqlError = False
    
    ' クエリー生成オブジェクトを生成するFactoryオブジェクト
    Dim dbObjFactory As New DbObjectFactory
    Set queryCreator = dbObjFactory.createQueryCreator(conn _
                                                            , applicationSetting.emptyCellReading _
                                                            , applicationSetting.getDirectInputChar _
                                                            , applicationSettingColFmt.getDbColFormatListByDbConn(conn) _
                                                            , applicationSetting.schemaUse _
                                                            , applicationSetting.getTableColumnEscapeByDbConn(conn))

    ' トランザクションオブジェクト
    Dim transaction As New DBTransaction: transaction.init conn
    
    ' レコード総数
    Dim recordCount As Long

    ' レコードリスト配列
    Dim records    As Variant
    ' レコード配列
    Dim record     As Variant
    
    ' SQL
    Dim sql       As String
    ' 結果件数
    Dim resultCnt   As Long
    ' 結果件数（1クエリー）
    Dim resultCntOne As Long

    ' インデックス
    Dim i   As Long
    
    ' キャンセルフラグをチェックする
    checkCancel

    ' レコード配列を取得する
    records = recordList.values
    
    ' レコード配列（Variant型）が本当に配列かどうかをチェックする
    If IsArray(records) = True Then
    
        ' レコード総数を取得する
        recordCount = VBUtil.arraySize(records)
    
        ' ▼SQL開始イベント
        RaiseEvent ExecuteSqlBegin(recordCount)
    
        ' 最小インデックス
        Dim min As Long: min = LBound(records, 1)
        ' 最大インデックス
        Dim max As Long: max = UBound(records, 1)

        ' 1件ずつ処理する
        For i = min To max

            ' キャンセルフラグをチェックする
            checkCancel

            ' 2次元配列から i の部分を取り出し1次元配列に変換する
            record = VBUtil.convert2to1Array(records, i)

            ' SQLを生成
            sql = createSql(table, record)
    
            ' SQLを実行
            If sql <> "" Then
            
                ' SQL実行フラグをONにする
                executingSql = True
                executingSqlError = False

                resultCntOne = ADOUtil.queryAction(conn, sql)
                
                ' SQL実行時エラーフラグがONの場合
                ' エラーフラグは、queryAction実行時にエラーが発生した場合に
                ' 本メソッドのエラーハンドリングでONに設定される。
                If executingSqlError = False Then
                
                    ' エラーが発生しなかった場合にカウントする
                    resultCnt = resultCnt + resultCntOne
                    
                    ' ▼SQL実行
                    RaiseEvent executeSql(tableSheet, recordList, i, sql, resultCntOne)
    
                End If

                ' SQL実行フラグをOFFにする
                executingSql = False

            End If

        Next

        ' ▼SQL完了イベント
        RaiseEvent ExecuteSqlSuccessful(resultCnt)

    End If

    ' 確認を行う
    If applicationSetting.commitConfirm = applicationSetting.COMMIT_CONFIRM_YES Then
    
        ' メッセージボックスの戻り値
        Dim resultMsg As Long
        
        resultMsg = VBUtil.showMessageBoxForYesNoCancel("コミットしてもよろしいですか？" _
                           & vbNewLine _
                           & vbNewLine & "・[はい] → コミット" _
                           & vbNewLine & "・[いいえ] → ロールバック" _
                           & vbNewLine & "・[キャンセル] → ロールバックして処理を中断" _
                        , ConstantsCommon.APPLICATION_NAME)

        ' キャンセル押下（ロールバックして中断）
        If resultMsg = WinAPI_User.IDCANCEL Then
        
            ' ▼SQLキャンセルイベント
            RaiseEvent ExecuteSqlCancel(resultCnt)
            
            ' 処理をキャンセルする
            canceled = True: checkCancel
            
        ' いいえ押下（ロールバックして継続）
        ElseIf resultMsg = WinAPI_User.IDNO Then
        
            ' ▼SQLキャンセルイベント
            RaiseEvent ExecuteSqlCancel(resultCnt)
            
            ' 戻り値に実行件数を設定して関数を抜ける
            executeQuery = 0
            
            Exit Function
        End If
        
    End If

    ' トランザクションをコミットする
    transaction.commit

    ' 戻り値に実行件数を設定
    executeQuery = resultCnt

    Exit Function

err:

    ' エラーハンドラで別の関数を呼び出すとエラー情報が消えてしまうことがあるので
    ' 構造体にエラー情報を保存しておく
    Dim errT As errInfo: errT = VBUtil.swapErr

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
        RaiseEvent ExecuteSqlError(tableSheet, recordList, i, sql, sqlState, err.Description)

        ' SQLエラーが発生した場合に処理を継続する
        If applicationSetting.sqlErrorHandling = applicationSetting.SQL_ERROR_HANDLING_RESUME Then
        
            ' SQL実行時エラーフラグを立てる
            executingSqlError = True
            Resume Next
        End If
        
        ' エラーメッセージを設定
        errT.Description = ConstantsError.ERR_DESC_SQL_EXECUTE_FAILED _
                        & vbNewLine _
                        & vbNewLine _
                        & "[詳細情報]" & vbNewLine _
                        & errT.Number & vbNewLine & errT.Description
                          
        ' エラー番号を設定
        errT.Number = ConstantsError.ERR_NUMBER_SQL_EXECUTE_FAILED
        
    End If
    
    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ▽レコード更新
'
' 概要　　　：テーブル定義とレコードリストからクエリーを生成しファイルに出力する。
' 引数　　　：tableSheet テーブルシート
' 　　　　　　recordList レコードリスト
' 戻り値　　：実行件数
'
' =========================================================
Private Function executeOutputQueryFile(ByRef tableSheet As ValTableWorksheet _
                                      , ByRef recordList As ValRecordList) As Long

    On Error GoTo err

    ' テーブル定義
    Dim table As ValDbDefineTable
    ' テーブル定義を取得する
    Set table = tableSheet.table

    ' クエリー生成オブジェクトを生成するFactoryオブジェクト
    Dim dbObjFactory As New DbObjectFactory
    Set queryCreator = dbObjFactory.createQueryCreator(conn _
                                                            , applicationSetting.emptyCellReading _
                                                            , applicationSetting.getDirectInputChar _
                                                            , applicationSettingColFmt.getDbColFormatListByDbConn(conn) _
                                                            , applicationSetting.schemaUse _
                                                            , applicationSetting.getTableColumnEscapeByDbConn(conn))

    ' レコード総数
    Dim recordCount As Long

    ' レコードリスト配列
    Dim records    As Variant
    ' レコード配列
    Dim record     As Variant
    
    ' SQL
    Dim sql       As String
    ' 結果件数
    Dim resultCnt   As Long
    ' 結果件数（1クエリー）
    Dim resultCntOne As Long

    ' インデックス
    Dim i   As Long
    
    ' キャンセルフラグをチェックする
    checkCancel

    ' レコード配列を取得する
    records = recordList.values
    
    ' レコード配列（Variant型）が本当に配列かどうかをチェックする
    If IsArray(records) = True Then
    
        ' レコード総数を取得する
        recordCount = VBUtil.arraySize(records)
    
        ' ▼SQL開始イベント
        RaiseEvent ExecuteSqlBegin(recordCount)
    
        ' 最小インデックス
        Dim min As Long: min = LBound(records, 1)
        ' 最大インデックス
        Dim max As Long: max = UBound(records, 1)

        ' 1件ずつ処理する
        For i = min To max

            ' キャンセルフラグをチェックする
            checkCancel

            ' 2次元配列から i の部分を取り出し1次元配列に変換する
            record = VBUtil.convert2to1Array(records, i)

            ' SQLを生成
            sql = createSql(table, record)

            ' SQLを実行
            If sql <> "" Then
            
                ' カウントする
                resultCnt = resultCnt + 1
                
                ' 改行コードを変換する
                sql = replace(sql, vbLf, newline)
                ' セミコロンを付加する
                sql = sql & ";" & newline
                
                ' ファイルを出力する
                writer.writeText sql, True
                
                
                ' ▼SQL実行
                RaiseEvent executeSql(tableSheet, recordList, i, sql, resultCntOne)

            End If

        Next

        ' ▼SQL完了イベント
        RaiseEvent ExecuteSqlSuccessful(resultCnt)

    End If

    ' 戻り値に実行件数を設定
    executeOutputQueryFile = resultCnt

    Exit Function

err:

    ' エラーハンドラで別の関数を呼び出すとエラー情報が消えてしまうことがあるので
    ' 構造体にエラー情報を保存しておく
    Dim errT As errInfo: errT = VBUtil.swapErr

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ▽レコード削除更新
'
' 概要　　　：テーブル上の全レコードを削除するSQLを発行する。
' 　　　　　　本メソッドでトランザクションが完結しており
' 　　　　　　正常終了した場合は、コミット。異常終了した場合は、ロールバックされる。
' 引数　　　：
' 戻り値　　：実行件数
'
' =========================================================
Public Function executeQueryDeleteOfAllTable() As Long

    On Error GoTo err

    ' テーブルシート
    Dim tableSheet  As ValTableWorksheet

    ' テーブルシート読み込みオブジェクトを初期化する
    Dim tsr As New ExeTableSheetReader
    Set tsr.conn = conn
    Set tsr.sheet = sheet
    
    Set tableSheet = tsr.readTableInfo

    Set readTableSheet = tableSheet

    executingSql = False
    
    ' クエリー生成オブジェクトを生成するFactoryオブジェクト
    Dim dbObjFactory As New DbObjectFactory
    Set queryCreator = dbObjFactory.createQueryCreator(conn _
                                                            , applicationSetting.emptyCellReading _
                                                            , applicationSetting.getDirectInputChar _
                                                            , applicationSettingColFmt.getDbColFormatListByDbConn(conn) _
                                                            , applicationSetting.schemaUse _
                                                            , applicationSetting.getTableColumnEscapeByDbConn(conn))

    ' SQL
    Dim sql       As String
    ' 結果件数
    Dim resultCnt As Long

    ' トランザクションオブジェクト
    Dim transaction As New DBTransaction: transaction.init conn
    
    ' キャンセルフラグをチェックする
    checkCancel

    ' SQLを生成
    sql = createSql(tableSheet.table, Nothing)
    createdSql = sql
    
    ' SQLを実行
    If sql <> "" Then
    
        ' SQL実行フラグをONにする
        executingSql = True

        resultCnt = resultCnt + ADOUtil.queryAction(conn, sql)

        ' SQL実行フラグをOFFにする
        executingSql = False

    End If

    ' 削除確認を行う
    If VBUtil.showMessageBoxForYesNo("全てのレコードを削除しました。コミットしてもよろしいですか？" _
     , ConstantsCommon.APPLICATION_NAME) = WinAPI_User.IDNO Then
    
        ' 処理をキャンセルする
        canceled = True: checkCancel
    End If
    
    ' トランザクションをコミットする
    transaction.commit

    ' 戻り値に実行件数を設定
    executeQueryDeleteOfAllTable = resultCnt

    Exit Function

err:

    ' エラーハンドラで別の関数を呼び出すとエラー情報が消えてしまうことがあるので
    ' 構造体にエラー情報を保存しておく
    Dim errT As errInfo: errT = VBUtil.swapErr

    ' SQL実行時にエラーが発生
    If executingSql = True Then

        ' エラーメッセージを設定
        errT.Description = ConstantsError.ERR_DESC_SQL_EXECUTE_FAILED _
                        & vbNewLine _
                        & vbNewLine _
                        & "[詳細情報]" & vbNewLine _
                        & errT.Number & vbNewLine & errT.Description
                          
        ' エラー番号を設定
        errT.Number = ConstantsError.ERR_NUMBER_SQL_EXECUTE_FAILED
        
    End If
    
    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ▽SQL生成
'
' 概要　　　：クエリー種類に対応するSQLを生成する。
'
' 引数　　　：table  テーブル定義
' 　　　　　　record レコード配列
' 戻り値　　：SQL
'
' =========================================================
Public Function createSql(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    Select Case queryType
    
        Case QUERY_TYPE_INSERT_UPDATE
        
            On Error GoTo errExecuteSelect
            
            Dim recordSet As Object
            
            ' ---------------------------------------------------------
            ' PK条件を作成する
            ' ---------------------------------------------------------
            Dim condition As New ValSelectCondition
            
            Dim i As Long: i = 1
            
            ' カラム
            Dim column     As ValDbDefineColumn
            ' カラムリスト
            Dim columnList As ValCollection
            
            ' テーブル制約情報(PK)
            Dim tableConstPk    As New ValDbDefineTableConstraints
            ' PKカラムであるかをあらわすフラグ
            Dim isColumnPk      As Boolean
            
            Dim tableConstTmp   As ValDbDefineTableConstraints
            ' テーブル制約リストからPK制約を取得する
            For Each tableConstTmp In table.constraintsList.col
            
                If tableConstTmp.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk Then
                
                    Set tableConstPk = tableConstTmp
                    Exit For
                End If
            Next
            
            ' カラムリストを取得する
            Set columnList = table.columnList
            
            ' カラムリストを1件ずつ処理する
            For Each column In columnList.col
                    
                ' PK制約であるかどうかを判定する
                If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
                
                    isColumnPk = False
                Else
                
                    isColumnPk = True
                End If
                
                ' カラムがPKである場合
                If isColumnPk = True Then
                
                    If record(i) = "" Then
                    
                        condition.setCondition column.columnName, "NULL", Empty
                    Else
                    
                        condition.setCondition column.columnName, record(i), Empty
                    End If
                    
                End If
                
                i = i + 1
            Next
            ' ---------------------------------------------------------
            
            Dim selectSql As String
            selectSql = queryCreator.createSelect(table, condition)
            
            DbQueryConverter.init Me.conn, queryParameterList
            selectSql = DbQueryConverter.convertSql(selectSql)
            
            Set recordSet = ADOUtil.querySelect(conn, selectSql)

            ' 取得レコードが0件の場合
            If recordSet.EOF Then
                
                createSql = queryCreator.createInsert(table, record)
                
            ' 取得レコードが1件の場合
            Else
                
                createSql = queryCreator.createUpdate(table, record)
                
            End If
            
            ' レコードセットを解放する
            ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing

            Exit Function
            
errExecuteSelect:

            ' レコードセットを解放する
            ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
            
            createSql = queryCreator.createInsert(table, record)
    
        Case QUERY_TYPE_INSERT
        
            createSql = queryCreator.createInsert(table, record)
        
        Case QUERY_TYPE_UPDATE
        
            createSql = queryCreator.createUpdate(table, record)
        
        Case QUERY_TYPE_DELETE
        
            createSql = queryCreator.createDelete(table, record)
        
        Case QUERY_TYPE_DELETE_ALL
        
            createSql = queryCreator.createDeleteOfAllTable(table)
    
    End Select
    
    DbQueryConverter.init Me.conn, queryParameterList
    createSql = DbQueryConverter.convertSql(createSql)

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

' --------------------↓↓↓
' テーブルシート読み込みクラスのイベントを本クラスのイベントとしてそのまま発行する
Private Sub tableSheetReader_readColumnListBegin(ByVal cnt As Long)

    RaiseEvent readColumnListBegin(cnt)
End Sub

Private Sub tableSheetReader_readColumnList(columnList As ValCollection)

    RaiseEvent readColumnList(columnList)
End Sub

Private Sub tableSheetReader_readColumn(column As ValDbDefineColumn)

    RaiseEvent readColumn(column)
End Sub

Private Sub tableSheetReader_readTableInfoBegin()

    RaiseEvent readTableInfoBegin
End Sub

Private Sub tableSheetReader_readTableInfo(table As ValTableWorksheet)

    RaiseEvent readTableInfo(table)
End Sub

Private Sub tableSheetReader_readRecordListBegin(ByVal cnt As Long)

    RaiseEvent readRecordListBegin(cnt)
End Sub

Private Sub tableSheetReader_readRecordList(recordList As ValRecordList)

    RaiseEvent readRecordList(recordList)
End Sub

Private Sub tableSheetReader_readRecordListBySelectionBegin(ByVal cnt As Long)

    RaiseEvent readRecordListBySelectionBegin(cnt)
End Sub

Private Sub tableSheetReader_readRecordListBySelectionOne(recordList As ValRecordList)

    RaiseEvent readRecordListBySelectionOne(recordList)
End Sub

Private Sub tableSheetReader_readRecordListBySelection(recordLists As ValCollection)

    RaiseEvent readRecordListBySelection(recordLists)
End Sub
' --------------------↑↑↑

