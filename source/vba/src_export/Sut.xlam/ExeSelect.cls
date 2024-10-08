VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' レコードを選択するクラス
'
' 作成者　：Ison
' 履歴　　：2008/02/01　新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' イベント
' ________________________________________________________
' =========================================================
' ▽ExeTableSheetReaderで定義しているイベントを再定義
' 　上位モジュールに発生したイベントをそのまま伝播するために再定義している
' =========================================================
Public Event readColumnListBegin(ByVal cnt As Long)
Public Event readColumn(ByRef column As ValDbDefineColumn)
Public Event readColumnList(ByRef columnList As ValCollection)
Public Event readTableInfoBegin()
Public Event readTableInfo(ByRef table As ValTableWorksheet)
Public Event readRecordListBegin(ByVal cnt As Long)
Public Event readRecordList(ByRef recordList As ValRecordList)
Public Event readRecordListBySelectionBegin(ByVal cnt As Long)
Public Event readRecordListBySelectionOne(ByRef recordList As ValRecordList)
Public Event readRecordListBySelection(ByRef recordLists As ValCollection)

' =========================================================
' ▽SQL実行イベント
'
' 概要　　　：SQL実行時に呼び出されるイベント
' 引数　　　：sql        SQLステートメント
'
' =========================================================
Public Event executeSql(ByRef sql As String)

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

' =========================================================
' ▽SQL実行結果反映時のエラーイベント
'
' 概要　　　：SQL実行時にエラーが発生した際に呼び出されるイベント
' 引数　　　：message    メッセージ
'     　　　：cell       セル
'
' =========================================================
Public Event ReflectionRecordError(ByRef message As String, ByVal cell As String)

' ---------------------------------------------------------
' □プロパティ
' ---------------------------------------------------------
' テーブルシート読み込み
Private WithEvents tableSheetReader As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1

' テーブルシート生成
Private tableSheetCreator As New ExeTableSheetCreator
' クエリ変換
Private queryConverter As New DbQueryConverter
' クエリパラメータリスト
Private queryParameterList_ As ValQueryParameterList

' DBコネクション
Public conn          As Object
' ワークシート
Public sheet         As Worksheet
' キャンセルフラグ
Public canceled      As Boolean

' ワークシートに追加するかどうか
Public appendSheet   As Boolean

' SQL実行中フラグ
Private executingSql As Boolean

' アプリケーション設定情報
Public applicationSetting As ValApplicationSetting
' アプリケーション設定（カラム書式情報）
Public applicationSettingColFmt As ValApplicationSettingColFormat

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
' ▽レコード取得実行（条件指定なし）
'
' 概要　　　：テーブル定義を基にクエリーを生成、実行しレコードを取得する。
' 引数　　　：
' 戻り値　　：レコード取得件数
'
' =========================================================
Public Function executeSelect() As Long

    ' テーブルシート
    Dim tableSheet As ValTableWorksheet

    ' テーブルシート読み込みオブジェクトを初期化する
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = sheet
    
    Set tableSheet = tableSheetReader.readTableInfo

    ' クエリー生成オブジェクトを生成するFactoryオブジェクト
    Dim dbObjFactory As New DbObjectFactory
    ' クエリー生成オブジェクト
    Dim queryCreator        As IDbQueryCreator
    
    ' クエリー生成オブジェクトを取得（第2,3引数は適当）
    Set queryCreator = dbObjFactory.createQueryCreator(conn, 0, "", "", "", applicationSettingColFmt.getDbColFormatListByDbConn(conn), applicationSetting.schemaUse, applicationSetting.getTableColumnEscapeByDbConn(conn))

    ' 条件
    Dim emptyCondition As New ValSelectCondition

    ' SELECT SQL
    Dim sql As String
    ' SELECTを生成する
    sql = queryCreator.createSelect(tableSheet.table, emptyCondition)

    ' SELECTを実行する
    executeSelect = execute(tableSheet, sql)
    
End Function

' =========================================================
' ▽レコード取得実行（クエリー指定）
'
' 概要　　　：指定されたクエリーを実行しレコードを取得する。
' 引数　　　：sql    SELECT
' 戻り値　　：レコード取得件数
'
' =========================================================
Public Function executeSelectByQuery(ByRef sql As String) As Long

    ' テーブルシート
    Dim tableSheet As ValTableWorksheet

    ' テーブルシート読み込みオブジェクトを初期化する
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = sheet
    
    Set tableSheet = tableSheetReader.readTableInfo

    ' SELECTを実行する
    executeSelectByQuery = execute(tableSheet, sql)

End Function

' =========================================================
' ▽レコード取得実行
'
' 概要　　　：対象となるテーブルの全レコードをシートに反映する
' 引数　　　：tableSheet  テーブルシート
' 　　　　　　sql    SELECT
' 戻り値　　：レコード取得件数
'
' =========================================================
Private Function execute(ByRef tableSheet As ValTableWorksheet, ByRef sql As String) As Long

    On Error GoTo err

    ' レコードの総数
    Dim recordCount As Long
    ' レコードセット
    Dim recordSet   As Object
    ' レコード配列
    Dim recordList  As Variant
    ' レコード配列（行サイズ）
    Dim recordListSize As Long

    ' レコード反映オフセット
    Dim recordOffset  As Long
    
    ' セルの削除・挿入用に利用するオブジェクト
    Set tableSheetCreator = New ExeTableSheetCreator
    tableSheetCreator.applicationSetting = applicationSetting
    tableSheetCreator.book = sheet.parent
    
    ' レコード処理単位
    Dim processCount As Long: processCount = applicationSetting.getRecProcessCount

    ' 処理のはじめにフラグをオフにしておく
    executingSql = False

    ' -------------------------------------------------------------
    ' SQLを実行する
    ' -------------------------------------------------------------
    ' キャンセルフラグをチェックする
    checkCancel
    
    RaiseEvent executeSql(sql)

    ' SQL実行中フラグをONにする
    executingSql = True
    
    queryConverter.initIfFirst Me.conn, queryParameterList, applicationSetting.queryParameterEncloseCustomPrefix, applicationSetting.queryParameterEncloseCustomSuffix
    sql = queryConverter.convertSql(sql)
    
    ' クエリーを実行する
    Set recordSet = ADOUtil.querySelect(conn, sql, 0)
    
    ' SQL実行中フラグをOFFにする
    executingSql = False

    RaiseEvent ExecuteSqlSuccessful(sql)


    ' -------------------------------------------------------------
    ' レコードを取得する
    ' -------------------------------------------------------------
    
    ' レコード反映オフセット位置を取得する
    recordOffset = getRecordOffset(tableSheet)
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
        
        sheet.activate
        sheet.Cells(recordOffset _
                  , ConstantsTable.U_COLUMN_OFFSET_COL).activate
    
    Else
    
        sheet.activate
        sheet.Cells(ConstantsTable.R_COLUMN_OFFSET_ROW _
                  , recordOffset).activate
    End If
    
    ' キャンセルフラグをチェックする
    checkCancel

    ' 取得レコードが0件の場合
    If recordSet.EOF Then
    
        If appendSheet = False Then
        
            ' ワークシートの入力領域を削除する
            tableSheetCreator.deleteCellOfRecord tableSheet, recordOffset
            ' 空の入力領域を生成する
            tableSheetCreator.insertEmptyCell tableSheet, recordOffset, 1
            
        End If
            
        ' 戻り値として0を返す
        execute = 0
        Exit Function
        
    End If
        
    If appendSheet = False Then
    
        ' ワークシートの入力領域を削除する
        tableSheetCreator.deleteCellOfRecord tableSheet, recordOffset
        
    End If
    
    ' セル最大文字数チェックエラー（Trueの場合はエラー）
    Dim isMaxLengthOverCell As Boolean
        
    ' 取得レコードが0件になるまでレコードセットを処理する
    Do While Not recordSet.EOF
    
        ' キャンセルフラグをチェックする
        checkCancel
        
        ' ▽レコード取得イベント
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
        
            GoTo loop_end
        End If
        
        ' ▽レコード変換イベント
        RaiseEvent ReflectionRecordConvertBegin(recordListSize, recordList)
        
        ' レコードリストを変換する
        convertRecordList tableSheet, recordList
        
        ' ▲レコード変換イベント
        RaiseEvent ReflectionRecordConvert(recordListSize, recordList)
        
        ' ▽レコード反映イベント
        RaiseEvent ReflectionRecordBegin(recordListSize, recordList)
        
        ' ワークシートにレコードリストの内容を反映する
        reflectSheet tableSheet, recordList, recordOffset, isMaxLengthOverCell
        
        ' ▲レコード反映イベント
        RaiseEvent ReflectionRecord(recordListSize, recordList)
        
        ' 次回のレコードリスト反映時のオフセット位置の計算
        ' レコードリストの反映先にレコードリストのサイズを加算する
        recordOffset = recordOffset + recordListSize
        
        ' レコード取得件数をカウントする
        recordCount = recordCount + recordListSize
    Loop
    
loop_end:
    
    ' レコードセットを解放する
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing

    ' レコード取得件数を返す
    execute = recordCount
    
    If isMaxLengthOverCell Then
        ' セルの最大文字数超過エラー
        err.Raise ConstantsError.ERR_NUMBER_CELL_MAX_LENGTH_OVER_REFLECT, , ConstantsError.ERR_DESC_CELL_MAX_LENGTH_OVER_REFLECT
    End If

    Exit Function

err:

    ' エラーハンドラで別の関数を呼び出すとエラー情報が消えてしまうことがあるので
    ' 構造体にエラー情報を保存しておく
    Dim errT As errInfo: errT = VBUtil.swapErr

    ' レコードセットを解放する
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
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
        
        ' SQLステータス
        Dim sqlState As String
        
        If conn.Errors.count > 0 Then
        
            ' エラーオブジェクトを取得する
            Dim error As Object
            Set error = conn.Errors.item(0)
            sqlState = error.sqlState
        End If

        RaiseEvent ExecuteSqlError(sql, sqlState, CStr(errT.Description))

    End If

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ▽レコードのオフセット位置取得
'
' 概要　　　：取得したレコードを反映するオフセット位置を取得する。
'
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：レコードのオフセット位置
'
' =========================================================
Private Function getRecordOffset(ByRef tableSheet As ValTableWorksheet) As Long

    ' 戻り値
    Dim recordOffset As Long
    
    ' 既存の行を削除せずに、最後の位置からレコードを追加する
    If appendSheet = True Then
        
        If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
        
            ' 最終入力行を取得し＋１する
            recordOffset = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsTable.U_RECORD_NUM_COL) + 1
            
            ' 上記で取得した値がシートの範囲を超えているかをチェックする
            If ExcelUtil.checkOverMaxRow(sheet, recordOffset) = False Then
            
                ' エラーを発行する
                err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SHEET _
                        , _
                        , ConstantsError.ERR_DESC_OUT_OF_RANGE_SHEET
            End If
        Else
        
            ' 最終入力列を取得し＋１する
            recordOffset = ExcelUtil.getCellEndOfLastInputtedCol(sheet, ConstantsTable.R_RECORD_NUM_ROW) + 1
        
            ' 上記で取得した値がシートの範囲を超えているかをチェックする
            If ExcelUtil.checkOverMaxCol(sheet, recordOffset) = False Then
            
                ' エラーを発行する
                err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SHEET _
                        , _
                        , ConstantsError.ERR_DESC_OUT_OF_RANGE_SHEET
            End If
            
        End If
                
    ' 既存の行を削除して、最初の位置からレコードを追加する
    Else
    
        If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
        
            ' 定数からレコードオフセットを取得する
            recordOffset = ConstantsTable.U_RECORD_OFFSET_ROW
                        
        Else
        
            ' 定数からレコードオフセットを取得する
            recordOffset = ConstantsTable.R_RECORD_OFFSET_COL
        End If
    
    End If

    ' 戻り値を設定する
    getRecordOffset = recordOffset

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
Private Function convertRecordList(ByRef tableSheet As ValTableWorksheet, ByRef recordList As Variant)

    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ' レコード配列の要素入れ替えによって新たに生成される配列
        Dim temp As Variant
        
        ' レコード配列のデータ変換処理
        ' RecordSet.GetRowsより返された配列は、(col,rec)となっているため(rec,col)となるように入れ替えを実施する
        temp = convertArrayValuesToUnder(recordList)
        
        ' レコード配列を解放する
        Erase recordList

        ' レコードリストを入れ替える
        recordList = temp
        
    Else
        
        ' レコード配列のデータ変換処理
        recordList = convertArrayValuesToRight(recordList)
    End If
        
    

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
            tempArray(X, Y) = ExcelUtil.convertCellValue(value)
        
        Next Y
    Next X
    
    convertArrayValuesToUnder = tempArray

End Function

' =========================================================
' ▽配列要素の変換
'
' 概要　　　：要素をセル形式の値に変換する。
' 引数　　　：v 2次元配列
'
' 戻り値　　：2次元配列
' 特記事項　：
'
' =========================================================
Private Function convertArrayValuesToRight(ByRef v As Variant) As Variant
    
    Dim X As Long
    Dim Y As Long
    
    Dim Xlower As Long
    Dim Xupper As Long
    
    Dim Ylower As Long
    Dim Yupper As Long
    
    Xlower = LBound(v, 2)
    Xupper = UBound(v, 2)
    Ylower = LBound(v, 1)
    Yupper = UBound(v, 1)
    
    Dim value As Variant
    
    For X = Xlower To Xupper
        For Y = Ylower To Yupper
        
            ' 一旦値を取得する
            value = v(Y, X)
            
            ' 文字列の場合、セル形式の値へと変換する
            v(Y, X) = ExcelUtil.convertCellValue(value)
        
        Next Y
    Next X
    
    convertArrayValuesToRight = v
    
End Function

' =========================================================
' ▽レコードリストのシート反映
'
' 概要　　　：レコードリストをシートに反映する
' 引数　　　：tableSheet   テーブルシート
'             recordList   レコードリスト
' 　　　　　　recordOffset レコード反映時のオフセット位置
' 　　　　　　outIsMaxLengthOverCell セルの最大文字数オーバーフラグ
'
' =========================================================
Private Sub reflectSheet(ByRef tableSheet As ValTableWorksheet _
                       , ByRef recordList As Variant _
                       , ByVal recordOffset As Long _
                       , ByRef outIsMaxLengthOverCell As Boolean)

    Dim ret As Boolean
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ret = reflectSheetToUnder(tableSheet, recordList, recordOffset, outIsMaxLengthOverCell)
        
    Else
    
        ret = reflectSheetToRight(tableSheet, recordList, recordOffset, outIsMaxLengthOverCell)
    End If
    
    ' 戻り値が False の場合、エラーを発行する
    If ret = False Then
    
        err.Raise ERR_NUMBER_OUT_OF_RANGE_SHEET _
                , _
                , ERR_DESC_OUT_OF_RANGE_SHEET
    End If

End Sub

' =========================================================
' ▽レコードリストのシート反映
'
' 概要　　　：レコードリストをシートに反映する
' 引数　　　：tableSheet   テーブルシート
'             recordList   レコードリスト
' 　　　　　　recordOffset レコード反映時のオフセット位置
' 　　　　　　outIsMaxLengthOverCell セルの最大文字数オーバーフラグ
'
' 戻り値　　：True  正常
' 　　　　　　False シートへの反映失敗
'
' =========================================================
Private Function reflectSheetToUnder(ByRef tableSheet As ValTableWorksheet _
                                   , ByRef recordList As Variant _
                                   , ByVal recordOffset As Long _
                                   , ByRef outIsMaxLengthOverCell As Boolean) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = True

    ' テーブル定義
    Dim table As ValDbDefineTable
    ' テーブル定義を取得する
    Set table = tableSheet.table
    
    ' レコードサイズ
    Dim rowSize As Long
    ' カラムサイズ
    Dim colSize As Long

    ' 配列からレコードサイズを取得する
    rowSize = VBUtil.arraySize(recordList)
    ' 配列からカラムサイズを取得する
    colSize = VBUtil.arraySize(recordList, 2)

    ' ------------------------------------------------------------
    ' シートの範囲チェック
    ' ------------------------------------------------------------
    If ExcelUtil.checkOverMaxRow(sheet, recordOffset, rowSize) = False Then
    
        ' サイズをオーバーしているので調整する
        rowSize = ExcelUtil.getSizeOfSheetRow(sheet) - recordOffset + 1
    
        ' 全レコード反映できないので戻り値をFalseに設定
        ret = False
    End If
        
    ' 空の入力領域を生成する
    tableSheetCreator.insertEmptyCell tableSheet, recordOffset, rowSize

    ' ------------------------------------------------------------
    ' 配列の内容をシートに反映する
    ' ------------------------------------------------------------
    ExcelUtil.copyArrayToCells recordList, sheet, recordOffset, ConstantsTable.U_COLUMN_OFFSET_COL, rowSize, colSize
    
    ' ------------------------------------------------------------
    ' セルの最大文字数チェック
    ' ------------------------------------------------------------
    If applicationSetting.selectCheckCellMaxLength = True Then
        
        Dim maxLengthOverCell As Variant
        Dim maxLengthCellList As ValCollection
        Set maxLengthCellList = checkCellMaxLength(recordList)
        
        For Each maxLengthOverCell In maxLengthCellList.col
            RaiseEvent ReflectionRecordError( _
                                    ConstantsError.ERR_DESC_CELL_MAX_LENGTH_OVER, _
                                    sheet.Cells(recordOffset + maxLengthOverCell(0), ConstantsTable.U_COLUMN_OFFSET_COL + maxLengthOverCell(1)).Address)
            outIsMaxLengthOverCell = True
        Next
        
    End If

    reflectSheetToUnder = ret
    
End Function

' =========================================================
' ▽レコードリストのシート反映
'
' 概要　　　：レコードリストをシートに反映する
' 引数　　　：tableSheet   テーブルシート
'             recordList   レコードリスト
' 　　　　　　recordOffset レコード反映時のオフセット位置
' 　　　　　　outIsMaxLengthOverCell セルの最大文字数オーバーフラグ
'
' 戻り値　　：True  正常
' 　　　　　　False シートへの反映失敗
'
' =========================================================
Private Function reflectSheetToRight(ByRef tableSheet As ValTableWorksheet _
                                   , ByRef recordList As Variant _
                                   , ByVal recordOffset As Long _
                                   , ByRef outIsMaxLengthOverCell As Boolean) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = True

    ' テーブル定義
    Dim table As ValDbDefineTable
    ' テーブル定義を取得する
    Set table = tableSheet.table
    
    ' レコードサイズ
    Dim rowSize As Long
    ' カラムサイズ
    Dim colSize As Long

    ' 配列からレコードサイズを取得する
    rowSize = VBUtil.arraySize(recordList, 2)
    ' 配列からカラムサイズを取得する
    colSize = VBUtil.arraySize(recordList)

    ' ------------------------------------------------------------
    ' シートの範囲チェック
    ' ------------------------------------------------------------
    If ExcelUtil.checkOverMaxCol(sheet, recordOffset, rowSize) = False Then
    
        ' サイズをオーバーしているので調整する
        rowSize = ExcelUtil.getSizeOfSheetCol(sheet) - recordOffset + 1
    
        ' 全レコード反映できないので戻り値をFalseに設定
        ret = False
    End If

    ' 空の入力領域を生成する
    tableSheetCreator.insertEmptyCell tableSheet, recordOffset, rowSize

    ' ------------------------------------------------------------
    ' 配列の内容をシートに反映する
    ' ------------------------------------------------------------
    ExcelUtil.copyArrayToCells recordList, sheet, ConstantsTable.R_COLUMN_OFFSET_ROW, recordOffset, colSize, rowSize
    
    ' ------------------------------------------------------------
    ' セルの最大文字数チェック
    ' ------------------------------------------------------------
    If applicationSetting.selectCheckCellMaxLength = True Then
        
        Dim maxLengthOverCell As Variant
        Dim maxLengthCellList As ValCollection
        Set maxLengthCellList = checkCellMaxLength(recordList)
        
        For Each maxLengthOverCell In maxLengthCellList.col
            RaiseEvent ReflectionRecordError( _
                                    ConstantsError.ERR_DESC_CELL_MAX_LENGTH_OVER, _
                                    sheet.Cells(ConstantsTable.R_COLUMN_OFFSET_ROW + maxLengthOverCell(0), recordOffset + maxLengthOverCell(1)).Address)
            outIsMaxLengthOverCell = True
        Next
        
    End If

    reflectSheetToRight = ret

End Function

' =========================================================
' ▽セルの最大文字数を超えているかをチェックする。
'
' 概要　　　：
'
' 引数　　　：recordList レコードリスト
' 戻り値　　：最大文字数を超えているセルのリスト
'
' =========================================================
Private Function checkCellMaxLength(ByRef recordList As Variant) As ValCollection
    
    Set checkCellMaxLength = New ValCollection
    
    Dim i As Long
    Dim j As Long
    
    For i = LBound(recordList) To UBound(recordList)
    
        For j = LBound(recordList, 2) To UBound(recordList, 2)
        
            If Not isNull(recordList(i, j)) Then
                ' NOT NULL
            
                If Len(CStr(recordList(i, j))) > 32767 Then
                    ' セルの最大長を超えているかをチェックする
                    checkCellMaxLength.setItem Array(i, j)
                End If
                
            Else
                ' NULL
            
            End If
        
        Next
        
    Next

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

