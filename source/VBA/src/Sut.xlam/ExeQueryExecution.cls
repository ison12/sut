VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeQueryExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリを実行するクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2013/06/19　新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' イベント
' ________________________________________________________

' =========================================================
' ▽SQL実行イベント
'
' 概要　　　：SQL実行時に呼び出されるイベント
' 引数　　　：sql        SQLステートメント
'
' =========================================================
Public Event executeSql(ByRef sql As String)

' =========================================================
' ▽SQL実行キャンセルイベント
'
' 概要　　　：SQLの実行キャンセル時に呼び出されるイベント
' 引数　　　：count 実行件数
'
' =========================================================
Public Event ExecuteSqlCancel(ByVal count As Long)

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
' 引数　　　：sql        SQLステートメント
' 　　　　　：sqlstate   SQLステータス
' 　　　　　：message    メッセージ
'
' =========================================================
Public Event ExecuteSqlError(ByRef sql As String, ByRef sqlState As String, ByRef message As String)

' =========================================================
' ▽SQL実行結果反映時のイベント（レコード取得）
'
' 概要　　　：
' 引数　　　：recordListSize レコードリストサイズ
' 　　　　　　recordList     レコードリスト
'
' =========================================================
Public Event ReflectionRecordGetBegin(ByVal recordListSize As Long)

' =========================================================
' ▽SQL実行結果反映時のイベント（レコード取得）
'
' 概要　　　：
' 引数　　　：recordListSize レコードリストサイズ
' 　　　　　　recordList     レコードリスト
'
' =========================================================
Public Event ReflectionRecordGet(ByVal recordListSizeActual As Long _
                               , ByVal recordListSize As Long _
                               , ByRef recordList As Variant)

' =========================================================
' ▽SQL実行結果反映後のイベント（取得レコード編集）
'
' 概要　　　：
' 引数　　　：recordListSize レコードリストサイズ
' 　　　　　　recordList     レコードリスト
'
' =========================================================
Public Event ReflectionRecordConvertBegin(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ▽SQL実行結果反映後のイベント（取得レコード編集）
'
' 概要　　　：
' 引数　　　：recordListSize レコードリストサイズ
' 　　　　　　recordList     レコードリスト
'
' =========================================================
Public Event ReflectionRecordConvert(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ▽SQL実行結果反映時のイベント（シート反映）
'
' 概要　　　：
' 引数　　　：recordListSize レコードリストサイズ
' 　　　　　　recordList     レコードリスト
'
' =========================================================
Public Event ReflectionRecordBegin(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ▽SQL実行結果反映後のイベント（シート反映）
'
' 概要　　　：
' 引数　　　：recordListSize レコードリストサイズ
' 　　　　　　recordList     レコードリスト
'
' =========================================================
Public Event ReflectionRecord(ByVal recordListSize As Long, ByRef recordList As Variant)


' ---------------------------------------------------------
' □プロパティ
' ---------------------------------------------------------
' DBコネクション
Public conn          As Object
' キャンセルフラグ
Public canceled      As Boolean
' ワークシートに追加するかどうか
Public appendSheet   As Boolean
' トランザクション開始有無フラグ
Public isBeginTran   As Boolean

' SQL実行中フラグ
Private executingSql      As Boolean
' SQL実行時にエラーが発生した場合
Private executingSqlError As Boolean

' クエリ結果シート生成
Private queryResultSheetCreator As ExeQueryResultSheetCreator

' ワークシート
Private m_sheet       As Worksheet
' アプリケーション設定情報
Private m_applicationSetting As ValApplicationSetting

' =========================================================
' ▽プロパティ シート
' =========================================================
Public Property Let sheet(ByVal vNewValue As Variant)

    Set m_sheet = vNewValue
    queryResultSheetCreator.sheet = vNewValue
    queryResultSheetCreator.book = vNewValue.parent
End Property

' =========================================================
' ▽アプリケーション設定情報
' =========================================================
Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
    queryResultSheetCreator.applicationSetting = vNewValue
End Property

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set queryResultSheetCreator = New ExeQueryResultSheetCreator
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set queryResultSheetCreator = Nothing
End Sub

' =========================================================
' ▽クエリ実行
'
' 概要　　　：対象となるクエリの実行結果をシートに反映する
' 引数　　　：queryList  クエリリスト
' 戻り値　　：レコード取得件数
'
' =========================================================
Public Function execute(ByRef queryList As sutredlib.GenericCollection) As Long

    On Error GoTo err

    ' 影響を与えたレコード
    Dim affectedRecord As Long
    
    Dim i, j As Long
    
    ' ヘッダ配列
    Dim headerList As Variant
    ' フィールド
    Dim field   As Object
    
    ' レコードの総数
    Dim recordCount As Long
    ' レコードセット
    Dim recordSet   As Object
    ' レコード配列
    Dim recordList  As Variant
    ' レコード配列（行サイズ）
    Dim recordListSize As Long

    ' レコードセットインデックス
    Dim recordSetIndex As Long
    ' レコード反映オフセット
    Dim recordOffset  As Long
    
    ' レコード処理単位
    Dim processCount As Long: processCount = m_applicationSetting.getRecProcessCount

    If queryList.count <= 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_SQL_EMPTY _
                , err.Source _
                , ConstantsError.ERR_DESC_SQL_EMPTY
    End If

    ' レコードセットインデックス
    recordSetIndex = 1
    ' レコード反映オフセット位置を取得する
    recordOffset = queryResultSheetCreator.getRecordOffset(appendSheet)
    
    ' 対象シートがクエリ結果シートではない場合
    If Not queryResultSheetCreator.isQueryResultSheet(m_sheet) Then
    
        Set m_sheet = queryResultSheetCreator.createSheet
    End If

    If appendSheet = False Then
        
        ' 行を削除する
        ExcelUtil.deleteRowEndOfLastInputted m_sheet, ConstantsTable.QUERY_COLUMN_OFFSET_ROW, ConstantsTable.QUERY_COLUMN_TITLE_COL
        
        ' アウトラインをクリアする
        m_sheet.Cells.ClearOutline
        
    End If
    
    m_sheet.activate
    m_sheet.Cells(recordOffset _
              , ConstantsTable.QUERY_COLUMN_OFFSET_COL).activate
    
    ' キャンセルフラグをチェックする
    checkCancel
    
    ' トランザクションオブジェクト
    Dim transaction As DBTransaction
    Set transaction = Nothing
    
    If isBeginTran = False Then
    
        Set transaction = New DBTransaction
        transaction.init conn
        
    End If
        
    Dim var As Variant
    Dim sql As String
    
    executingSql = False
    executingSqlError = False
    
    For Each var In queryList
    
        sql = var.query
    
        ' -------------------------------------------------------------
        ' SQLを実行する
        ' -------------------------------------------------------------
        ' キャンセルフラグをチェックする
        checkCancel
        
        ' 処理のはじめにフラグをオフにしておく
        executingSql = False
    
        RaiseEvent executeSql(sql)
    
        ' SQL実行中フラグをONにする
        executingSql = True
    
        ' SQL実行フラグをONにする
        executingSql = True
        executingSqlError = False
        
        ' クエリーを実行する
        Set recordSet = ADOUtil.queryBatch(conn, sql, affectedRecord)
        
        ' SQL実行中フラグをOFFにする
        executingSql = False
                
        ' SQL実行時エラーフラグがONの場合
        ' エラーフラグは、クエリ実行時にエラーが発生した場合に
        ' 本メソッドのエラーハンドリングでONに設定される。
        If executingSqlError = True Then
        
            GoTo loop_end

        End If
    
        RaiseEvent ExecuteSqlSuccessful(sql)
        
        ' -------------------------------------------------------------
        ' レコードを取得する
        ' -------------------------------------------------------------
        
        ' UPDATEなどの、レコードを返さないクエリの場合 State = 0 (adStateClosed) になる
        If recordSet.State = 0 Then
        
            ' レコードセットの番号をセルに埋め込む
            queryResultSheetCreator.insertEmptyCellForTitle recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
            m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "No." & recordSetIndex
            
            ' sqlをコメントとして埋め込む
            ExcelUtil.addComment m_sheet.name, recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL, sql
        
            queryResultSheetCreator.insertEmptyCellForResultCount recordOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL
            m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL).value = "正常に実行されました（" & affectedRecord & "件のレコードに影響を与えました）。"
            recordOffset = recordOffset + 1
            
            recordCount = recordCount + affectedRecord
        Else
        
            ' レコードセットの番号をセルに埋め込む
            queryResultSheetCreator.insertEmptyCellForTitle recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
            m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "No." & recordSetIndex
            
            ' sqlをコメントとして埋め込む
            ExcelUtil.addComment m_sheet.name, recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL, sql
        
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

            ' ヘッダリストの枠組みを挿入する
            queryResultSheetCreator.insertEmptyCellForColumn recordOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL, VBUtil.arraySize(headerList)
            ' ワークシートにレコードリストの内容を反映する
            queryResultSheetCreator.reflectSheetForColumn headerList, recordOffset
            
            ' ヘッダを反映したのでオフセット位置を加算する
            recordOffset = recordOffset + 1
            ' ---------------------------------------------------
            
            ' ---------------------------------------------------
            ' レコード情報の取得
            ' ---------------------------------------------------
            ' 取得レコードが0件になるまでレコードセットを処理する
            Do While Not recordSet.EOF
            
                ' キャンセルフラグをチェックする
                checkCancel
                
                ' ▼レコード取得イベント
                RaiseEvent ReflectionRecordGetBegin(processCount)
                
                ' レコードセットからレコードリストを取得する
                recordList = getRecordList(recordSet, processCount)
                
                ' レコードリストのサイズを取得する
                ' ADOのレコードセットからGetRowsで取得した配列は(列×行)でデータが格納されている
                recordListSize = VBUtil.arraySize(recordList, 2)
        
                ' ▲レコード取得イベント
                RaiseEvent ReflectionRecordGet(processCount, recordListSize, recordList)
                
                ' レコードリストが配列ではない場合
                If IsArray(recordList) = False Then
                
                    ' このブロックに入ることはあるのか？
                    GoTo loop_end
                End If
                
                ' ▼レコード変換イベント
                RaiseEvent ReflectionRecordConvertBegin(recordListSize, recordList)
                
                ' レコードリストを変換する
                convertRecordList recordList
                
                ' ▲レコード変換イベント
                RaiseEvent ReflectionRecordConvert(recordListSize, recordList)
                
                ' ▼レコード反映イベント
                RaiseEvent ReflectionRecordBegin(recordListSize, recordList)
                
                ' レコードリストの枠組みを挿入する
                queryResultSheetCreator.insertEmptyCellForRecord recordOffset, recordListSize, ConstantsTable.QUERY_COLUMN_OFFSET_COL, VBUtil.arraySize(headerList)
                ' ワークシートにレコードリストの内容を反映する
                queryResultSheetCreator.reflectSheetForRecord recordList, recordOffset
                
                ' ▲レコード反映イベント
                RaiseEvent ReflectionRecord(recordListSize, recordList)
                
                ' 次回のレコードリスト反映時のオフセット位置の計算
                ' レコードリストの反映先にレコードリストのサイズを加算する
                recordOffset = recordOffset + recordListSize
                
                ' レコード取得件数をカウントする
                recordCount = recordCount + recordListSize
            Loop
    
        End If
            
loop_end:
            
        recordSetIndex = recordSetIndex + 1
        
        ' レコードセットを解放する
        ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    Next
    
    ' 念のためレコードセットを解放する
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    ' トランザクションをコミットする
    If Not transaction Is Nothing Then
        transaction.commit
    End If
    
    ' -------------------------------------------------------

    ' レコード取得件数を返す
    execute = recordCount
    
    ' レコードリストの枠組みを挿入する
    queryResultSheetCreator.insertEmptyCellForResultSet recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
    ' 最終行に現在日付を埋め込む
    m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).WrapText = False
    m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "'" & Now & " : クエリ実行件数…" & (recordSetIndex - 1)

    ' グループ化された領域を全て閉じる
    m_sheet.Outline.ShowLevels 1

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
        RaiseEvent ExecuteSqlError(sql, sqlState, err.Description)
        
        ' レコードセットの番号をセルに埋め込む
        queryResultSheetCreator.insertEmptyCellForTitle recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
        m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "No." & recordSetIndex
            
        ' sqlをコメントとして埋め込む
        ExcelUtil.addComment m_sheet.name, recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL, sql

        queryResultSheetCreator.insertEmptyCellForError recordOffset, QUERY_COLUMN_OFFSET_COL
        m_sheet.Cells(recordOffset, QUERY_COLUMN_OFFSET_COL).value = sqlState & " " & errT.Description
        m_sheet.Cells(recordOffset, QUERY_COLUMN_OFFSET_COL).WrapText = False
        
        ' エラー情報を埋め込んだのでオフセット位置を加算する
        recordOffset = recordOffset + 1
        
        ' グループ化された領域を全て閉じる
        m_sheet.Outline.ShowLevels 1
        
        ' SQL実行時エラーフラグを立てる
        executingSqlError = True
        Resume Next
        
        ' ----------------------------------------------------
        ' Resume next でエラー発生場所に戻るので、以下の処理は実行しない
        ' ----------------------------------------------------
'        ' レコードリストの枠組みを挿入する
'        queryResultSheetCreator.insertEmptyCellForResultSet recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
'        ' 最終行に現在日付を埋め込む
'        m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).WrapText = False
'        m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "'" & Now & " : クエリが中断されました。"
'
'        ' グループ化された領域を全て閉じる
'        m_sheet.Outline.ShowLevels 1
        
    Else

        err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext
    
    End If

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
' 引数　　　：tableSheet   テーブルシート
'             recordList   レコードリスト
'
' =========================================================
Private Function convertRecordList(ByRef recordList As Variant)

    ' レコード配列の要素入れ替えによって新たに生成される配列
    Dim temp As Variant
    
    ' レコード配列のデータ変換処理
    ' RecordSet.GetRowsより返された配列は、(col,rec)となっているため(rec,col)となるように入れ替えを実施する
    temp = convertArrayValuesToUnder(recordList)
    
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
Private Function convertArrayValuesToUnder(ByRef v As Variant) As Variant
    
    Dim x As Long
    Dim y As Long
    
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
    
    For x = Xlower To Xupper
        For y = Ylower To Yupper
        
            ' 一旦値を取得する
            value = v(y, x)
            
            ' 文字列の場合、セル形式の値へと変換する
            tempArray(x, y) = ExcelUtil.convertCellValue(value)
        
        Next y
    Next x
    
    convertArrayValuesToUnder = tempArray

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
