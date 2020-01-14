VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeTableSheetCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' テーブルシートを生成するクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2007/12/01 新規作成
'           2008/05/05 31文字超シート名対応
'           2008/05/05 カラム256超対応
' 　　　　　2009/03/31 行フォーマット、Xバージョン・Yバージョンの追加
'           2010/02/22 テーブル定義クラスを変更
'           2013/05/18 カラム＋制約情報未ロード時に読み込み処理が走るように変更
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' イベント
' ________________________________________________________

' =========================================================
' ▽複数テーブルシート生成処理開始前イベント
'
' 概要　　　：複数のテーブルシートの生成前に呼び出されるイベント
' 引数　　　：tableInfoList   テーブル情報リスト
'
' =========================================================
Public Event createTableSheetsBegin(ByRef tableInfoList As sutredlib.GenericCollection)

' =========================================================
' ▽複数テーブルシート生成イベント
'
' 概要　　　：複数のテーブルシートの生成完了時に呼び出されるイベント
' 引数　　　：tableInfoList   テーブル情報リスト
'
' =========================================================
Public Event createTableSheets(ByRef tableInfoList As sutredlib.GenericCollection)

' =========================================================
' ▽テーブルシート生成処理開始前イベント
'
' 概要　　　：テーブルシートの生成前に呼び出されるイベント
' 引数　　　：tableInfo   テーブル情報
'
' =========================================================
Public Event createTableSheetBegin(ByRef table As ValTableWorksheet)

' =========================================================
' ▽テーブルシート生成イベント
'
' 概要　　　：テーブルシートの生成完了時に呼び出されるイベント
' 引数　　　：tableInfo   テーブル情報
'
' =========================================================
Public Event createTableSheet(ByRef tableSheet As ValTableWorksheet)

' =========================================================
' プロパティ
' =========================================================
' アプリケーション設定情報
Private m_applicationSetting As ValApplicationSetting
' 処理対象ワークブック
Private m_book               As Workbook
' テーブル定義の論理名リスト
Private m_namingList         As ValCollection
' キャンセルフラグ
Private m_canceled           As Boolean

Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
End Property

Public Property Let book(ByVal vNewValue As Variant)

    Set m_book = vNewValue
    
End Property

Public Property Let NamingList(ByVal vNewValue As Variant)

    Set m_namingList = vNewValue
    
End Property

Public Property Let canceled(ByVal vNewValue As Variant)

    m_canceled = vNewValue
    
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
' ▽テーブルシートを複数生成する
'
' 概要　　　：テーブルシートを複数生成する
' 引数　　　：tableSheetList  テーブル定義リスト
'
' =========================================================
Public Sub createSheets(ByRef tableSheetList As sutredlib.GenericCollection)

    ' DBメタ情報取得用オブジェクト生成クラス
    Dim dbObjFactory As New DbObjectFactory
    ' DBメタ情報取得用クラス
    Dim DbMetaInfoGetter  As IDbMetaInfoGetter
    Set DbMetaInfoGetter = dbObjFactory.createMetaInfoGetterObject(dbConn)
    
    Dim reloadTable As ValDbDefineTable
    
    ' テーブル情報
    Dim tableSheet As ValTableWorksheet

    ' ◇テーブルシート生成前イベント
    RaiseEvent createTableSheetsBegin(tableSheetList)

    ' 処理がキャンセルされている場合中断する
    checkCanceled

    ' テーブルの数だけシートを作成する
    For Each tableSheet In tableSheetList
        ' 未ロード時、ロードを実行する
        If Not tableSheet.table.isLoad Then
        
            ' DBにアクセスするのでロード画像を表示する
            SutWhite.showHourglassWindowOnCenterPt Nothing, 0, -50
        
            Set reloadTable = DbMetaInfoGetter.getTable(tableSheet.table.schemaName, tableSheet.table.tableName)
            
            If reloadTable Is Nothing Then
            
                ' DBにアクセスしたのでロード画像を非表示にする
                SutWhite.closeHourglassWindow
            
                ' テーブル情報が取得できない場合エラーを発行する
                err.Raise ERR_NUMBER_NOT_EXIST_TABLE_INFO _
                        , err.Source _
                        , ERR_DESC_NOT_EXIST_TABLE_INFO & vbNewLine & tableSheet.table.schemaTableName _
                        , err.HelpFile _
                        , err.HelpContext
                
            End If
            
            Set tableSheet.table = reloadTable
            
            ' DBにアクセスしたのでロード画像を非表示にする
            SutWhite.closeHourglassWindow
        
        End If
        ' テーブルシートを生成する
        createSheet tableSheet
        
        ' 処理がキャンセルされている場合中断する
        checkCanceled
        
    Next
    
    ' ◇テーブルシート生成完了イベント
    RaiseEvent createTableSheets(tableSheetList)

End Sub

' =========================================================
' ▽シートを生成する
'
' 概要　　　：指定されたテーブルのシートを生成する
' 引数　　　：tableSheet  テーブル定義情報
'
' =========================================================
Public Sub createSheet(ByRef tableSheet As ValTableWorksheet)

    ' ◇テーブルシート生成イベント
    RaiseEvent createTableSheetBegin(tableSheet)

    ' テーブルシートのヘッダ部分を生成する
    createSheetHeader tableSheet
    
    ' テーブルシートの本体部分を生成する
    If tableSheet.recFormat = sutredlib.recFormatToUnder Then
    
        createSheetToUnder tableSheet
    Else
    
        createSheetToRight tableSheet
    End If
    
    ' ◇テーブルシート生成完了イベント
    RaiseEvent createTableSheet(tableSheet)

End Sub

' =========================================================
' ▽シートのヘッダ部分を生成する
'
' 概要　　　：指定されたテーブルのヘッダ部分を生成する
' 引数　　　：tableSheet  テーブル定義情報
'
' =========================================================
Public Sub createSheetHeader(ByRef tableSheet As ValTableWorksheet)

    ' テーブル名
    Dim tableName     As String
    ' シート名
    Dim sheetName     As String
    ' シートオブジェクト
    Dim sheet         As Worksheet
    
    ' テーブル名を取得
    If tableSheet.omitsSchema = True Then
        ' スキーマを省略する場合、テーブル名のみとする
        tableName = tableSheet.table.tableName
    Else
        ' スキーマを省略しない場合、スキーマ＋テーブル名とする
        tableName = tableSheet.table.schemaTableName
    End If
    
    ' テーブル情報の取得元シート名をチェックする
    If tableSheet.sheetName = "" Then
    
        ' 取得元シート名が設定されていない場合
        ' 新規に生成するワークシートの名前をテーブル名とする
        sheetName = tableName
        
        ' ワークシートの名前に禁則文字が使用されているかチェックする
        If ExcelUtil.checkProhibitionCharOfSheetName(sheetName) = False Then
        
            ' 禁則文字を変換する
            sheetName = ExcelUtil.convertProhibitionCharOfSheetName(sheetName)
        
        End If
        
        ' ワークシートの最大サイズをチェックする
        If ExcelUtil.checkMaxLengthOfSheetName(sheetName) = False Then
        
            ' 規定サイズを超えている場合、ワークシート名を変換する
            sheetName = truncateExceededSheetName(sheetName)
        
            ' ワークシートが既に存在している場合
            If ExcelUtil.existSheet(m_book, sheetName) = True Then
            
                ' ワークシートの名称を一意な名称に変換する
                sheetName = ExcelUtil.convertUniqSheetName(m_book, sheetName)
                
            End If
        
        End If
        
    Else
    
        ' 取得元シート名が設定されている場合
        sheetName = tableSheet.sheetName
    End If
    
    ' シートの削除（存在していたら）と挿入を実行する
    deleteAndInsertSheet tableSheet, sheetName
    
    ' 新たにコピーされたシートオブジェクトを取得する
    Set sheet = m_book.Worksheets(sheetName)
    
    ' テーブル情報に新しいシートオブジェクトを設定する
    Set tableSheet.sheet = sheet
    
    ' ▽新シートにテーブル定義情報を埋め込む
    With sheet

        ' シートを可視にする
        .visible = xlSheetVisible

        ' スキーマ名を省略する場合
        If tableSheet.omitsSchema = True Then
        
            ' スキーマ名を隠しセルに設定する
            .Cells(ConstantsTable.SCHEMA_NAME_ROW _
                 , ConstantsTable.SCHEMA_NAME_COL) = ExcelUtil.convertCellValue(tableSheet.table.schemaName)
        End If

        ' テーブル名を設定する
        .Cells(ConstantsTable.TABLE_NAME_ROW _
             , ConstantsTable.TABLE_NAME_COL) = ExcelUtil.convertCellValue(tableName)
        ' テーブルコメントを設定する
        .Cells(ConstantsTable.TABLE_NAME_LOG_ROW _
             , ConstantsTable.TABLE_NAME_LOG_COL) = ExcelUtil.convertCellValue(tableSheet.table.tableComment)
        
    End With
    
End Sub

' =========================================================
' ▽シートを生成する
'
' 概要　　　：指定されたテーブルのシートを生成する
' 引数　　　：tableSheet  テーブル定義情報
'
' =========================================================
Public Sub createSheetToUnder(ByRef tableSheet As ValTableWorksheet)

    ' テーブル情報
    Dim table  As ValDbDefineTable
    ' テーブル情報を取得する
    Set table = tableSheet.table
    ' シートオブジェクト
    Dim sheet         As Worksheet
    
    ' カラム情報配列
    Dim columnArray As Variant
    
    ' カラムサイズ
    Dim colSize  As Long
    
    ' テーブルカラム数
    Dim tblColumnCount As Long
    ' テーブルのカラム数を取得する
    tblColumnCount = table.columnList.count
    
    ' シートのカラムサイズ
    Dim worksheetColSize As Long
    
    ' 新たにコピーされたシートオブジェクトを取得する
    Set sheet = tableSheet.sheet
    
    ' ▽新シートにテーブル定義情報を埋め込む
    With sheet

        ' シートのカラムサイズを取得する
        ' Excel 2003 以前では 256
        ' Excel 2007 以降はそれ以上
        worksheetColSize = ExcelUtil.getSizeOfSheetCol(sheet)

        ' テーブルのカラムがワークシートの規定カラムの範囲内であるかをチェックする
        If tblColumnCount > worksheetColSize - ConstantsTable.U_COLUMN_OFFSET_COL + 1 Then
        
            ' 規定カラムを超えた場合
            
            ' テーブルのカラムサイズにワークシートのカラムサイズを強制的に設定
            colSize = worksheetColSize - ConstantsTable.U_COLUMN_OFFSET_COL + 1
            
            ' エラー情報をセルに埋め込む
            outCellErrorInfoByOverColumnMaxSize sheet
        
        Else
        
            ' 規定カラムの範囲内である場合
            colSize = tblColumnCount
            
        End If
        
        ' カラム定義欄のテンプレートセル
        Dim srcCell As String
        ' カラム定義欄の反映先セル
        Dim desCell As String
        
        ' カラム定義欄を反映先にコピーする
        srcCell = .Cells(ConstantsTable.U_COLUMN_NAME_ROW _
                       , ConstantsTable.U_COLUMN_OFFSET_COL).Address(RowAbsolute:=False, columnAbsolute:=False) & ":" & _
                  .Cells(ConstantsTable.U_COLUMN_REFER_ROW _
                       , ConstantsTable.U_COLUMN_OFFSET_COL).Address(RowAbsolute:=False, columnAbsolute:=False)
                       
        desCell = .Cells(ConstantsTable.U_COLUMN_NAME_ROW _
                       , ConstantsTable.U_COLUMN_OFFSET_COL).Address(RowAbsolute:=False, columnAbsolute:=False) & ":" & _
                  .Cells(ConstantsTable.U_COLUMN_REFER_ROW _
                       , ConstantsTable.U_COLUMN_OFFSET_COL + colSize - 1).Address(RowAbsolute:=False, columnAbsolute:=False)
                  
        .Range(srcCell).Copy Destination:=.Range(desCell)
                      
        ' 空の入力領域を挿入する
        insertEmptyCell tableSheet

        ' カラム情報配列を生成する
        columnArray = createColumnInfoArrayToUnder(tableSheet, colSize)

        ' カラム情報配列をシートに反映する
        ExcelUtil.copyArrayToCells columnArray _
                                , sheet _
                                , ConstantsTable.U_COLUMN_NAME_ROW _
                                , ConstantsTable.U_COLUMN_OFFSET_COL

        ' 行の高さを自動調節する
        .Range(desCell).Rows.AutoFit
        
        On Error Resume Next
        
        ' シートをアクティブにした後にセルをアクティブにする
        .activate
        ' セルをアクティブにする
        .Cells(ConstantsTable.U_RECORD_OFFSET_ROW, ConstantsTable.U_COLUMN_OFFSET_COL).activate
        
        On Error GoTo 0
        
    End With ' sheet
    ' ====================================================

End Sub

' =========================================================
' ▽シートを生成する
'
' 概要　　　：指定されたテーブルのシートを生成する
' 引数　　　：tableSheet  テーブル定義情報
'
' =========================================================
Public Sub createSheetToRight(ByRef tableSheet As ValTableWorksheet)

    ' テーブル情報
    Dim table  As ValDbDefineTable
    ' テーブル情報を取得する
    Set table = tableSheet.table
    ' シートオブジェクト
    Dim sheet         As Worksheet
    
    ' カラム情報配列
    Dim columnArray As Variant
    
    ' カラムサイズ
    Dim colSize  As Long
    
    ' テーブルカラム数
    Dim tblColumnCount As Long
    ' テーブルのカラム数を取得する
    tblColumnCount = table.columnList.count
    
    ' シートの行サイズ
    Dim worksheetRowSize As Long
    
    ' 新たにコピーされたシートオブジェクトを取得する
    Set sheet = tableSheet.sheet
    
    ' ▽新シートにテーブル定義情報を埋め込む
    With sheet

        ' シートの行サイズを取得する
        ' Excel 2003 以前では 256
        ' Excel 2007 以降はそれ以上
        worksheetRowSize = ExcelUtil.getSizeOfSheetRow(sheet)

        ' テーブルのカラム数がワークシートの最大行数の範囲内であるかをチェックする
        If tblColumnCount > worksheetRowSize - ConstantsTable.R_COLUMN_OFFSET_ROW + 1 Then
        
            ' 規定カラムを超えた場合
            
            ' テーブルのカラムサイズにワークシートのカラムサイズを強制的に設定
            colSize = worksheetRowSize - ConstantsTable.R_COLUMN_OFFSET_ROW + 1
            
            ' エラー情報をセルに埋め込む
            outCellErrorInfoByOverColumnMaxSize sheet
            
        Else
        
            ' 規定カラムの範囲内である場合
            colSize = tblColumnCount
            
        End If
        
        ' カラム定義欄のテンプレートセル
        Dim srcCell As String
        ' カラム定義欄の反映先セル
        Dim desCell As String
        
        ' カラム定義欄を反映先にコピーする
        srcCell = .Cells(ConstantsTable.R_COLUMN_OFFSET_ROW _
                       , ConstantsTable.R_COLUMN_NAME_COL).Address(RowAbsolute:=False, columnAbsolute:=False) & ":" & _
                  .Cells(ConstantsTable.R_COLUMN_OFFSET_ROW _
                       , ConstantsTable.R_COLUMN_REFER_COL).Address(RowAbsolute:=False, columnAbsolute:=False)
                       
        desCell = .Cells(ConstantsTable.R_COLUMN_OFFSET_ROW _
                       , ConstantsTable.R_COLUMN_NAME_COL).Address(RowAbsolute:=False, columnAbsolute:=False) & ":" & _
                  .Cells(ConstantsTable.R_COLUMN_OFFSET_ROW + colSize - 1 _
                       , ConstantsTable.R_COLUMN_REFER_COL).Address(RowAbsolute:=False, columnAbsolute:=False)
                  
        .Range(srcCell).Copy Destination:=.Range(desCell)
                      
        ' 空の入力領域を挿入する
        insertEmptyCell tableSheet
        
        ' カラム情報配列を生成する
        columnArray = createColumnInfoArrayToRight(tableSheet, colSize)

        ' カラム情報配列をシートに反映する
        ExcelUtil.copyArrayToCells columnArray _
                                , sheet _
                                , ConstantsTable.R_COLUMN_OFFSET_ROW _
                                , ConstantsTable.R_COLUMN_NAME_COL
        
        ' 行の高さを自動調節する
        .Range(desCell).Rows.AutoFit
        
        On Error Resume Next
        
        ' シートをアクティブにした後にセルをアクティブにする
        .activate
        ' セルをアクティブにする
        .Cells(ConstantsTable.R_COLUMN_OFFSET_ROW, ConstantsTable.R_RECORD_OFFSET_COL).activate
        
        On Error GoTo 0
        
    End With ' sheet
    ' ====================================================

End Sub

' =========================================================
' ▽カラム情報配列生成（行フォーマット↓）
'
' 概要　　　：カラム情報配列を生成する。
' 　　　　　　カラム情報配列の数はcolSizeによって決定する。
'
' 　　　　　　カラムの数はExcelワークシートのサイズを
' 　　　　　　オーバーすることがあるため。
' 　　　　　　シートサイズをオーバーしているかどうかの判定は
' 　　　　　　呼び出し元で行いcolSizeに反映してもらう。
'
'
' 引数　　　：tableSheet テーブル情報
' 　　　　　　colSize    カラムサイズ
'
' =========================================================
Private Function createColumnInfoArrayToUnder(ByRef tableSheet As ValTableWorksheet, ByVal colSize As Long) As Variant

    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf
    
    ' テーブル情報
    Dim table  As ValDbDefineTable
    ' テーブル情報を取得する
    Set table = tableSheet.table
    ' カラム情報
    Dim column As ValDbDefineColumn
    
    ' カラムとセルのマッピング情報
    Dim columnCellMap As New ValCollection
    
    ' カラム情報配列
    Dim ret() As String
    
    Dim i, j As Long
    
    ' カラム情報配列を初期化する
    ReDim ret(0 To 8 _
            , 0 To colSize - 1)

    For Each column In table.columnList.col
    
        ' インデックスがカラムの最大数を超えていないかチェックする
        If i >= colSize Then
        
            ' 超えている場合、処理を終了する
            Exit For
        End If

        ' カラムコメント
        ret( _
            ConstantsTable.U_COLUMN_NAME_LOG_ROW _
          - ConstantsTable.U_COLUMN_NAME_ROW, i) = ExcelUtil.convertCellValue(column.columnComment)
        
        ' カラム名
        ret( _
            ConstantsTable.U_COLUMN_NAME_ROW _
          - ConstantsTable.U_COLUMN_NAME_ROW, i) = ExcelUtil.convertCellValue(column.columnName)
        
        ' 型
        ret( _
            ConstantsTable.U_COLUMN_TYPE_ROW _
          - ConstantsTable.U_COLUMN_NAME_ROW, i) = ExcelUtil.convertCellValue(column.dataTypeFormal)
        
        ' NULL可
        If column.isNull = True Then
            
            ret( _
                ConstantsTable.U_COLUMN_NULL_ROW _
              - ConstantsTable.U_COLUMN_NAME_ROW, i) = "Y"
        Else
            
            ret( _
                ConstantsTable.U_COLUMN_NULL_ROW _
              - ConstantsTable.U_COLUMN_NAME_ROW, i) = "N"
        End If
            
        ' デフォルト値
        ret( _
            ConstantsTable.U_COLUMN_DEF_ROW _
          - ConstantsTable.U_COLUMN_NAME_ROW, i) = ExcelUtil.convertCellValue(column.default)
        
        ' PK制約を設定
        ret( _
            ConstantsTable.U_COLUMN_PK_ROW _
          - ConstantsTable.U_COLUMN_NAME_ROW, i) = ""
        
        ' UK制約を設定
        ret( _
            ConstantsTable.U_COLUMN_UK_ROW _
          - ConstantsTable.U_COLUMN_NAME_ROW, i) = ""
        
        ' FK制約
        ret( _
            ConstantsTable.U_COLUMN_REFER_ROW _
          - ConstantsTable.U_COLUMN_NAME_ROW, i) = ""
        
        ' カラム名からセル位置が分かるようにマップ変数に情報を格納する
        columnCellMap.setItem i, column.columnName
        
        i = i + 1
        
    Next
    
    ' テーブル制約情報
    Dim tableConst As ValDbDefineTableConstraints
    ' テーブル制約情報に含まれるカラム情報
    Dim tableConstColumn As ValDbDefineColumn
    ' テーブル制約情報に含まれるカラム情報(参照先)
    Dim tableConstReferColumn As ValDbDefineColumn
    ' セルの値を格納する変数
    Dim val  As String
    Dim val2 As String
    
    ' セル位置
    Dim findCellIndex As Long
    
    Dim ukSequence As Long: ukSequence = 1 ' テーブル内でのUKのシーケンス
    
    For Each tableConst In table.constraintsList.col
    
        Select Case tableConst.constraintType
        
            ' PKの場合
            Case tableConstPk
            
                For Each tableConstColumn In tableConst.columnList.col
                
                    If columnCellMap.exist(tableConstColumn.columnName) = True Then
                    
                        findCellIndex = columnCellMap.getItem(tableConstColumn.columnName, vbLong)
                        
                        ' PK制約を設定
                        ret( _
                            ConstantsTable.U_COLUMN_PK_ROW _
                          - ConstantsTable.U_COLUMN_NAME_ROW, findCellIndex) = "○"
                    End If
                                    
                Next
                
            ' UKの場合
            Case tableConstUk
            
                For Each tableConstColumn In tableConst.columnList.col
                
                    If columnCellMap.exist(tableConstColumn.columnName) = True Then
                    
                        findCellIndex = columnCellMap.getItem(tableConstColumn.columnName, vbLong)
                    
                        ' UK制約を設定
                        
                        val = ret( _
                            ConstantsTable.U_COLUMN_UK_ROW _
                          - ConstantsTable.U_COLUMN_NAME_ROW, findCellIndex)
                          
                        ' 既に何らかの値が設定されている場合、改行コードを付与する
                        If val <> Empty Then
                        
                            val = val & newline
                        End If
                        
                        ret( _
                            ConstantsTable.U_COLUMN_UK_ROW _
                          - ConstantsTable.U_COLUMN_NAME_ROW, findCellIndex) = val & ukSequence
                          
                    End If
                Next
                
                ukSequence = ukSequence + 1
            ' FKの場合
            Case tableConstFk
            
                j = 1
                
                For Each tableConstColumn In tableConst.columnList.col
                
                    Set tableConstReferColumn = tableConst.fkColumnList.getItem(j)
                    
                    If columnCellMap.exist(tableConstColumn.columnName) = True Then
                    
                        findCellIndex = columnCellMap.getItem(tableConstColumn.columnName, vbLong)
                    
                        ' FK制約を設定
                        
                        val = ret( _
                            ConstantsTable.U_COLUMN_REFER_ROW _
                          - ConstantsTable.U_COLUMN_NAME_ROW, findCellIndex)
                          
                        ' 既に何らかの値が設定されている場合、改行コードを付与する
                        If val <> Empty Then
                        
                            val = val & newline
                        End If
                        
                        If tableConstReferColumn.schemaName = table.schemaName Then
                            val2 = tableConstReferColumn.tableName _
                                    & "." & tableConstReferColumn.columnName
                        Else
                            val2 = tableConstReferColumn.schemaName _
                                    & "." & tableConstReferColumn.tableName _
                                    & "." & tableConstReferColumn.columnName
                        End If
                        
                        ret( _
                            ConstantsTable.U_COLUMN_REFER_ROW _
                          - ConstantsTable.U_COLUMN_NAME_ROW, findCellIndex) = val & val2
                          
                    End If
                    
                    j = j + 1
                Next
        
        End Select
        
    Next

    ' 戻り値を設定する
    createColumnInfoArrayToUnder = ret
    
End Function

' =========================================================
' ▽カラム情報配列生成（行フォーマット→）
'
' 概要　　　：カラム情報配列を生成する。
' 　　　　　　カラム情報配列の数はcolSizeによって決定する。
'
' 　　　　　　カラムの数はExcelワークシートのサイズを
' 　　　　　　オーバーすることがあるため。
' 　　　　　　シートサイズをオーバーしているかどうかの判定は
' 　　　　　　呼び出し元で行いcolSizeに反映してもらう。
'
'
' 引数　　　：tableSheet テーブル情報
' 　　　　　　colSize    カラムサイズ
'
' =========================================================
Private Function createColumnInfoArrayToRight(ByRef tableSheet As ValTableWorksheet, ByVal colSize As Long) As Variant

    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf
    
    ' テーブル情報
    Dim table  As ValDbDefineTable
    ' テーブル情報を設定する
    Set table = tableSheet.table
    ' カラム情報
    Dim column As ValDbDefineColumn
    
    ' カラムとセルのマッピング情報
    Dim columnCellMap As New ValCollection
    
    ' カラム情報配列
    Dim ret() As String
    
    Dim i, j As Long
    
    ' カラム情報配列を初期化する
    ReDim ret(0 To colSize - 1 _
            , 0 To 8)

    For Each column In table.columnList.col
    
        ' インデックスがカラムの最大数を超えていないかチェックする
        If i >= colSize Then
        
            ' 超えている場合、処理を終了する
            Exit For
        End If

        ' カラムコメント
        ret( _
            i _
          , ConstantsTable.R_COLUMN_NAME_LOG_COL _
          - ConstantsTable.R_COLUMN_NAME_COL) = ExcelUtil.convertCellValue(column.columnComment)
        
        ' カラム名
        ret( _
            i _
          , ConstantsTable.R_COLUMN_NAME_COL _
          - ConstantsTable.R_COLUMN_NAME_COL) = ExcelUtil.convertCellValue(column.columnName)
        
        ' 型
        ret( _
            i _
          , ConstantsTable.R_COLUMN_TYPE_COL _
          - ConstantsTable.R_COLUMN_NAME_COL) = ExcelUtil.convertCellValue(column.dataTypeFormal)
        
        ' NULL可
        If column.isNull = True Then
            
            ret( _
                i _
              , ConstantsTable.R_COLUMN_NULL_COL _
              - ConstantsTable.R_COLUMN_NAME_COL) = "Y"
        Else
            
            ret( _
                i _
              , ConstantsTable.R_COLUMN_NULL_COL _
              - ConstantsTable.R_COLUMN_NAME_COL) = "N"
        End If
            
        ' デフォルト値
        ret( _
            i _
          , ConstantsTable.R_COLUMN_DEF_COL _
          - ConstantsTable.R_COLUMN_NAME_COL) = ExcelUtil.convertCellValue(column.default)
        
        ' PK制約
        ret( _
            i _
          , ConstantsTable.R_COLUMN_PK_COL _
          - ConstantsTable.R_COLUMN_NAME_COL) = ""
          
        ' UK制約
        ret( _
            i _
          , ConstantsTable.R_COLUMN_UK_COL _
          - ConstantsTable.R_COLUMN_NAME_COL) = ""
        
        ' FK制約
        ret( _
            i _
          , ConstantsTable.R_COLUMN_REFER_COL _
          - ConstantsTable.R_COLUMN_NAME_COL) = ""
        
        ' カラム名からセル位置が分かるようにマップ変数に情報を格納する
        columnCellMap.setItem i, column.columnName
        
        i = i + 1
        
    Next

    ' テーブル制約情報
    Dim tableConst As ValDbDefineTableConstraints
    ' テーブル制約情報に含まれるカラム情報
    Dim tableConstColumn As ValDbDefineColumn
    ' テーブル制約情報に含まれるカラム情報(参照先)
    Dim tableConstReferColumn As ValDbDefineColumn
    ' セルの値を格納する変数
    Dim val  As String
    Dim val2 As String
    
    ' セル位置
    Dim findCellIndex As Long
    
    Dim ukSequence As Long: ukSequence = 1 ' テーブル内でのUKのシーケンス
    
    For Each tableConst In table.constraintsList.col
    
        Select Case tableConst.constraintType
        
            ' PKの場合
            Case tableConstPk
            
                For Each tableConstColumn In tableConst.columnList.col
                
                    If columnCellMap.exist(tableConstColumn.columnName) = True Then
                    
                        findCellIndex = columnCellMap.getItem(tableConstColumn.columnName, vbLong)
                        
                        ' PK制約を設定
                        ret( _
                            findCellIndex _
                          , ConstantsTable.R_COLUMN_PK_COL _
                          - ConstantsTable.R_COLUMN_NAME_COL) = "○"
                    End If
                                    
                Next
                
            ' UKの場合
            Case tableConstUk
            
                For Each tableConstColumn In tableConst.columnList.col
                
                    If columnCellMap.exist(tableConstColumn.columnName) = True Then
                    
                        findCellIndex = columnCellMap.getItem(tableConstColumn.columnName, vbLong)
                        
                        ' UK制約を設定
                        
                        val = ret( _
                            findCellIndex _
                          , ConstantsTable.R_COLUMN_UK_COL _
                          - ConstantsTable.R_COLUMN_NAME_COL)
                          
                        ' 既に何らかの値が設定されている場合、改行コードを付与する
                        If val <> Empty Then
                        
                            val = val & newline
                        End If
                        
                        ' UK制約
                        ret( _
                            findCellIndex _
                          , ConstantsTable.R_COLUMN_UK_COL _
                          - ConstantsTable.R_COLUMN_NAME_COL) = val & ukSequence
                        
                    End If
                Next
                
                ukSequence = ukSequence + 1
            ' FKの場合
            Case tableConstFk
            
                j = 1
                
                For Each tableConstColumn In tableConst.columnList.col
                
                    Set tableConstReferColumn = tableConst.fkColumnList.getItem(j)
                    
                    If columnCellMap.exist(tableConstColumn.columnName) = True Then
                    
                        findCellIndex = columnCellMap.getItem(tableConstColumn.columnName, vbLong)
                    
                        ' FK制約を設定
                        
                        val = ret( _
                            findCellIndex _
                          , ConstantsTable.R_COLUMN_REFER_COL _
                          - ConstantsTable.R_COLUMN_NAME_COL)
                          
                        ' 既に何らかの値が設定されている場合、改行コードを付与する
                        If val <> Empty Then
                        
                            val = val & newline
                        End If
                        
                        If tableConstReferColumn.schemaName = table.schemaName Then
                            val2 = tableConstReferColumn.tableName _
                                    & "." & tableConstReferColumn.columnName
                        Else
                            val2 = tableConstReferColumn.schemaName _
                                    & "." & tableConstReferColumn.tableName _
                                    & "." & tableConstReferColumn.columnName
                        End If
                        
                        ret( _
                            findCellIndex _
                          , ConstantsTable.R_COLUMN_REFER_COL _
                          - ConstantsTable.R_COLUMN_NAME_COL) = val & val2
                        
                    End If
                    
                    j = j + 1
                Next
        
        End Select
        
    Next
    
    ' 戻り値を設定する
    createColumnInfoArrayToRight = ret
    
End Function

' =========================================================
' ▽シートの削除・挿入
'
' 概要　　　：シートが既に存在する場合は削除し、挿入を行う。
' 引数　　　：tableSheet      テーブル情報
' 　　　　　　targetSheetName 対象シート名
'
' =========================================================
Private Sub deleteAndInsertSheet(ByRef tableSheet As ValTableWorksheet, ByRef targetSheetName As String)

    ' コピーシート
    Dim copySheetName As String

    ' 行フォーマットによってコピーシートを変更する
    If tableSheet.recFormat = sutredlib.recFormatToUnder Then
    
        copySheetName = ConstantsTable.U_SHEET_NAME_TEMPLATE
    
    Else
    
        copySheetName = ConstantsTable.R_SHEET_NAME_TEMPLATE
    End If

    ' 削除されたシートの位置に再びシートを作成したいので
    ' 削除対象シートの位置を調べ隣合わせになっているシートオブジェクトを取得する。
    ' （シートのコピー関数が
    ' 　あるシートを基点としてコピーするといった方法にしか対応していないためこのような処理が必要になる）
    On Error Resume Next

    ' 削除対象シート
    Dim targetSheet     As Worksheet
    ' 削除対象シートの左隣にあるシート
    Dim targetSheetPrev As Worksheet
    ' 削除対象シートの右隣にあるシート
    Dim targetsheetNext As Worksheet

    ' 削除対象シートを取得する
    Set targetSheet = m_book.Worksheets(targetSheetName)

    ' 削除対象シートが取得できた場合（元々存在していない場合もある）
    If Not targetSheet Is Nothing Then

        ' ブック内にシートの数が1つだけの場合（削除対象シート含めて）
        If m_book.Worksheets.count = 1 Then
        
            ' 何もしない
        ' 削除対象シートが左端の場合
        ElseIf targetSheet.index = 1 Then

            Set targetsheetNext = m_book.Worksheets(2)

        ' 削除対象シートが右端の場合
        ElseIf targetSheet.index = m_book.Worksheets.count Then

            Set targetSheetPrev = m_book.Worksheets(targetSheet.index - 1)

        ' 削除対象シートが中間（端にはない）
        Else
        
            Set targetSheetPrev = m_book.Worksheets(targetSheet.index - 1)
        End If


    End If

    ' Excelではシートの数が0になるのを許さないので、この状態でシートを削除すると
    ' ダミーのシートを勝手に挿入してしまう。
    ' したがって、対象のブックのシート数が1つの場合は、前もってシートを挿入しておき、処理終了後に削除するようにする
    Dim dummySheetName As String
    
    If m_book.Worksheets.count = 1 Then
    
        m_book.Worksheets.Add
        dummySheetName = m_book.ActiveSheet.name
    End If
    
    ' シート削除実行
    ExcelUtil.deleteSheet m_book, targetSheetName
    
    ' 左隣のシートが取得できている場合
    If Not targetSheetPrev Is Nothing Then
    
        ExcelUtil.copySheet SutWorkbook _
                          , copySheetName _
                          , m_book _
                          , targetSheetName _
                          , targetSheetPrev.name _
                          , "after"
    
    ' 右隣のシートが取得できている場合
    ElseIf Not targetsheetNext Is Nothing Then
    
        ExcelUtil.copySheet SutWorkbook _
                          , copySheetName _
                          , m_book _
                          , targetSheetName _
                          , targetsheetNext.name _
                          , "before"
    
    ' 削除対象シートが存在していない場合
    Else
    
        ExcelUtil.copySheetAppend SutWorkbook _
                                , copySheetName _
                                , m_book _
                                , targetSheetName
        
    
    End If
    
    ' ダミーシートが挿入されている場合、ダミーシートを削除する
    If dummySheetName <> "" Then
    
        m_book.Worksheets(dummySheetName).delete
    End If
    
    On Error GoTo 0

End Sub

' =========================================================
' ▽エラー情報をセルに埋め込む
'
' 概要　　　：
' 引数　　　：sheet シートオブジェクト
'
' =========================================================
Private Sub outCellErrorInfoByOverColumnMaxSize(ByRef sheet As Worksheet)

    With sheet
    
        ' カラムが最大数超えた旨をメッセージとして埋め込む
        .Cells(ConstantsTable.TABLE_OVER_MAX_COL_SIZE_ROW _
             , ConstantsTable.TABLE_OVER_MAX_COL_SIZE_COL).Cells.value = ConstantsError.ERR_DESC_COLUMN_SIZE_OVER_SHEET_SIZE
        ' フォントサイズも変更する
        ExcelUtil.changeFontSize sheet.name _
                              , ConstantsTable.TABLE_OVER_MAX_COL_SIZE_ROW _
                              , ConstantsTable.TABLE_OVER_MAX_COL_SIZE_COL, 12, m_book
    
        ' イメージシートからエラー時のアイコンを取得し貼り付ける
        Dim imageSheet As Worksheet
        Set imageSheet = ThisWorkbook.Worksheets(ConstantsImage.SHEET_NAME_IMAGE)
        
        imageSheet.Shapes(ConstantsImage.ERROR_ICON).Copy
        
        .Paste .Cells(TABLE_ERROR_ICON_ROW _
                    , TABLE_ERROR_ICON_COL)

    End With
    
End Sub

' =========================================================
' ▽キャンセルチェック
'
' 概要　　　：処理がキャンセルされているかを確認して
'             キャンセルされていれば処理を中断する
'
' =========================================================
Private Sub checkCanceled()

    ' キャンセルフラグがTrueの場合、処理が中断されたと判断して例外をスローする
    If m_canceled = True Then
    
        err.Raise ERR_NUMBER_PROC_CANCEL, err.Source, ERR_DESC_PROC_CANCEL
    End If

End Sub

' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：tableSheet   テーブル情報
' 　　　　　　recordOffset レコード反映時のオフセット位置
' 　　　　　　recordSize   レコードサイズ
'
' 戻り値　　：
' =========================================================
Public Sub insertEmptyCell(ByRef tableSheet As ValTableWorksheet _
                         , Optional ByVal recordOffset As Long = 0 _
                         , Optional ByVal recordSize As Long = 0)

    ' レコードサイズが0の場合、デフォルト値を10に設定する
    If recordSize = 0 Then
    
        recordSize = 10
    End If
        
    If tableSheet.recFormat = sutredlib.recFormatToUnder Then
    
        ' オフセット値が0の場合、デフォルトの位置を設定する
        If recordOffset = 0 Then
        
            recordOffset = ConstantsTable.U_RECORD_OFFSET_ROW
        End If
    
        insertEmptyCellToUnder tableSheet, recordOffset, recordSize
        
    Else
    
        ' オフセット値が0の場合、デフォルトの位置を設定する
        If recordOffset = 0 Then
        
            recordOffset = ConstantsTable.R_RECORD_OFFSET_COL
        End If
    
        insertEmptyCellToRight tableSheet, recordOffset, recordSize
    End If
    

End Sub

' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：tableSheet   テーブル情報
' 　　　　　　recordOffset レコード反映時のオフセット位置
' 　　　　　　recordSize   レコードサイズ
'
' 戻り値　　：
' =========================================================
Private Sub insertEmptyCellToUnder(ByRef tableSheet As ValTableWorksheet, ByVal recordOffset As Long, ByVal recordSize As Long)

    ' カラムの数
    Dim columnCount As Long
    ' シートの最大列数
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(tableSheet.sheet)
    
    ' カラムの数がシートの最大列数を超えていないかをチェックする
    If tableSheet.table.columnList.count > sheetColSize - ConstantsTable.U_COLUMN_OFFSET_COL + 1 Then
    
        ' 最大列数を超えている場合、列数を調整する
        columnCount = sheetColSize - ConstantsTable.U_COLUMN_OFFSET_COL + 1
        
    Else
    
        ' 最大列数を超えていない場合、そのまま設定
        columnCount = tableSheet.table.columnList.count
    End If

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.U_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.U_RECORD_OFFSET_ROW _
                          , ConstantsTable.U_COLUMN_OFFSET_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = tableSheet.sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsTable.U_COLUMN_OFFSET_COL).Address & ":" & _
                        Cells(recordOffset + recordSize - 1 _
                            , ConstantsTable.U_COLUMN_OFFSET_COL + columnCount - 1).Address)

    ' 入力領域をコピーする
    srcRange.Copy Destination:=desRange

    ' ------------------------------------------------------------
    ' セル書式の設定
    ' ------------------------------------------------------------
    ' フォントを設定
    desRange.Font.name = m_applicationSetting.cellFontName
    ' フォントサイズを設定
    desRange.Font.size = m_applicationSetting.cellFontSize
    ' 折り返しを設定
    desRange.WrapText = m_applicationSetting.cellWordwrap
    ' 列幅を設定
    ExcelUtil.changeColWidth desRange, m_applicationSetting.cellWidth
    ' 行高を設定
    If m_applicationSetting.lineHeightAutoAdjust Then
        ExcelUtil.changeRowHeight desRange, -1
    Else
        ExcelUtil.changeRowHeight desRange, m_applicationSetting.cellHeight
    End If

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.U_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.U_RECORD_OFFSET_ROW _
                          , ConstantsTable.U_RECORD_NUM_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = tableSheet.sheet _
                     .Range( _
                        Cells(recordOffset _
                            , ConstantsTable.U_RECORD_NUM_COL).Address & ":" & _
                        Cells(recordOffset + recordSize - 1 _
                            , ConstantsTable.U_RECORD_NUM_COL).Address)

    ' レコード番号をコピーする
    srcRange.Copy Destination:=desRange

End Sub

' =========================================================
' ▽空の入力領域を挿入
'
' 概要　　　：
' 引数　　　：table        テーブル情報
' 　　　　　　recordOffset レコード反映時のオフセット位置
' 　　　　　　recordSize   レコードサイズ
'
' 戻り値　　：
' =========================================================
Private Sub insertEmptyCellToRight(ByRef tableSheet As ValTableWorksheet, ByVal recordOffset As Long, ByVal recordSize As Long)

    ' カラムの数
    Dim columnCount As Long
    ' シートの最大列数
    Dim sheetRowSize As Long: sheetRowSize = ExcelUtil.getSizeOfSheetRow(tableSheet.sheet)
    
    ' カラムの数がシートの最大列数を超えていないかをチェックする
    If tableSheet.table.columnList.count > sheetRowSize - ConstantsTable.R_COLUMN_OFFSET_ROW + 1 Then
    
        ' 最大列数を超えている場合、列数を調整する
        columnCount = sheetRowSize - ConstantsTable.R_COLUMN_OFFSET_ROW + 1
        
    Else
    
        ' 最大列数を超えていない場合、そのまま設定
        columnCount = tableSheet.table.columnList.count
    End If

    ' ------------------------------------------------------------
    ' テンプレートシートのレコード領域をコピーする（書式等をコピー）
    ' ------------------------------------------------------------
    ' コピー元セル
    Dim srcRange As Range
    ' コピー先セル
    Dim desRange As Range

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.R_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.R_COLUMN_OFFSET_ROW _
                          , ConstantsTable.R_RECORD_OFFSET_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = tableSheet.sheet _
                     .Range( _
                        Cells(ConstantsTable.R_COLUMN_OFFSET_ROW _
                            , recordOffset).Address & ":" & _
                        Cells(ConstantsTable.R_COLUMN_OFFSET_ROW + columnCount - 1 _
                            , recordOffset + recordSize - 1).Address)

    ' 入力領域をコピーする
    srcRange.Copy Destination:=desRange

    ' ------------------------------------------------------------
    ' セル書式の設定
    ' ------------------------------------------------------------
    ' フォントを設定
    desRange.Font.name = m_applicationSetting.cellFontName
    ' フォントサイズを設定
    desRange.Font.size = m_applicationSetting.cellFontSize
    ' 折り返しを設定
    desRange.WrapText = m_applicationSetting.cellWordwrap
    ' 列幅を設定
    ExcelUtil.changeColWidth desRange, m_applicationSetting.cellWidth
    ' 行高を設定
    If m_applicationSetting.lineHeightAutoAdjust Then
        ExcelUtil.changeRowHeight desRange, -1
    Else
        ExcelUtil.changeRowHeight desRange, m_applicationSetting.cellHeight
    End If

    ' コピー元のRangeオブジェクトを取得する
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.R_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.R_RECORD_NUM_ROW _
                          , ConstantsTable.R_RECORD_OFFSET_COL)
                  
                  
    ' コピー先のRangeオブジェクトを取得する
    Set desRange = tableSheet.sheet _
                     .Range( _
                        Cells(ConstantsTable.R_RECORD_NUM_ROW _
                            , recordOffset).Address & ":" & _
                        Cells(ConstantsTable.R_RECORD_NUM_ROW _
                            , recordOffset + recordSize - 1).Address)

    ' レコード番号をコピーする
    srcRange.Copy Destination:=desRange

End Sub


' =========================================================
' ▽レコードを削除する
'
' 概要　　　：レコードオフセット位置から最後尾の位置までのレコードを削除する。
' 引数　　　：tableSheet   テーブル定義
' 　　　　　　recordOffset レコードオフセット位置
'
' =========================================================
Public Sub deleteCellOfRecord(ByRef tableSheet As ValTableWorksheet, Optional ByVal recordOffset As Long = 0)

    ' 削除対象範囲
    Dim targetRange As Range
    ' 削除終了位置
    Dim length As Long
    
    If tableSheet.recFormat = sutredlib.recFormatToUnder Then
    
        ' レコードオフセットをデフォルト値に設定する
        If recordOffset = 0 Then
        
            recordOffset = ConstantsTable.U_RECORD_OFFSET_ROW
        End If
        
        ' 最後尾の入力箇所を取得する
        length = ExcelUtil.getCellEndOfLastInputtedRow(tableSheet.sheet, ConstantsTable.U_RECORD_NUM_COL)
        
        If length < ConstantsTable.U_RECORD_OFFSET_ROW Then
        
            length = ConstantsTable.U_RECORD_OFFSET_ROW
        End If
        
        ' 削除対象範囲を取得
        Set targetRange = tableSheet.sheet _
                            .Range( _
                               Cells(recordOffset _
                                   , 1).Address & ":" & _
                               Cells(length _
                                   , 1).Address)
    
        ' 削除する（行単位で削除）
        targetRange.EntireRow.delete
        
    Else
    
        ' レコードオフセットをデフォルト値に設定する
        If recordOffset = 0 Then
        
            recordOffset = ConstantsTable.R_RECORD_OFFSET_COL
        End If
        
        ' 最後尾の入力箇所を取得する
        length = ExcelUtil.getCellEndOfLastInputtedCol(tableSheet.sheet, ConstantsTable.R_RECORD_NUM_ROW)
    
        If length < ConstantsTable.R_RECORD_OFFSET_COL Then
        
            length = ConstantsTable.R_RECORD_OFFSET_COL
        End If
        
        ' 削除対象範囲を取得
        Set targetRange = tableSheet.sheet _
                            .Range( _
                               Cells(1 _
                                   , recordOffset).Address & ":" & _
                               Cells(1 _
                                   , length).Address)
    
        ' 削除する（列単位で削除）
        targetRange.EntireColumn.delete
        
    End If
    
End Sub
