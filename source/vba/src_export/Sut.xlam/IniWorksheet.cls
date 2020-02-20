VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' INIファイル操作クラス（Excelワークシート版）
'
' 制限事項
' ・セクションの省略不可
' ・コメント設定不可
' ・データ閲覧時、改行コードやイコール記号が特殊な制御コードに変換される
'
' 作成者　：Ison
' 履歴　　：2020/02/12　新規作成
'
' 特記事項：
' *********************************************************
' ブック
Private book As Workbook
' シート
Private sheet As Worksheet
' シート名
Private sheetName As String
' ファイル名
Private fileName As String
' ファイル名に関連した列番号
Private fileNameColumn As Long
' INIデータ
Private inid As IniData

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Public Sub Class_Initialize()
    
    Set inid = New IniData
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Public Sub Class_Terminate()
    
    destroy
    
End Sub

' =========================================================
' ▽初期化処理。
'
' 概要　　　：
' 引数　　　：book_      ブック
'     　　　：sheetName_ シート名
'     　　　：fileName_  ファイル名
' 戻り値　　：
'
' =========================================================
Public Sub init(ByVal book_ As Workbook, ByVal sheetName_ As String, ByVal fileName_ As String)

    ' メンバ設定
    Set book = book_
    sheetName = sheetName_
    fileName = fileName_
    
    ' 既にシートが存在する場合は、シートオブジェクトを取得しメンバに情報を設定する
    initSheet True
    
    readSheet

End Sub

' =========================================================
' ▽シートの初期化処理。
'
' 概要　　　：
' 引数　　　：isGetOnly シートオブジェクトの取得のみフラグ
' 戻り値　　：
'
' =========================================================
Public Sub initSheet(Optional ByVal isGetOnly As Boolean = False)

    ' シートの生成
    Set sheet = createSheet(book, sheetName, isGetOnly)
    If sheet Is Nothing Then
        Exit Sub
    End If
    
    ' シート内の列番号取得
    fileNameColumn = findColumnFilePath(sheet, fileName, False)
    
End Sub

' =========================================================
' ▽破棄処理。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub destroy()

    writeSheet

End Sub

' =========================================================
' ▽シートの読み込み処理。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub readSheet()
    
    Dim contents As String
    If sheet Is Nothing Then
        contents = ""
    Else
        contents = readContents(sheet, fileNameColumn)
    End If
    
    inid.datas = inid.readContents(contents)
        
    ' 読み込み直後なので変更済みをオフにする
    inid.isModified = False

End Sub

' =========================================================
' ▽シートの書き込み処理。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub writeSheet()
    
    If inid.isModified = False Then
        ' 変更なしなので書き込まない
        Exit Sub
    End If

    Dim contents As String
    contents = inid.createWriteContents(inid.datas)
    
    If sheet Is Nothing Then
        ' シートオブジェクトが初期化されていない場合は、生成を試みる
        initSheet
    End If
    
    ' ヘッダとしてファイル名を記入
    sheet.Cells(ConstantsIniData.ROW_FILE_NAME, fileNameColumn) = fileName
    
    writeContents sheet, fileNameColumn, contents
        
    ' 書き込んだので変更済みをオフにする
    inid.isModified = False

End Sub

' =========================================================
' ▽データが存在するかを確認するメソッド
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：True データが存在する、False データが存在しない
'
' =========================================================
Public Function isExistsData() As Boolean

    isExistsData = False

    ' 新しいデータの場合
    Dim isNew As Boolean
    
    ' シートを初期化（取得を試みる）
    initSheet True

    If Not sheet Is Nothing Then
        ' 対象シートが取得できた場合
        
        findColumnFilePath sheet, fileName, isNew
        If isNew = False Then
            ' 既にデータが存在する
            isExistsData = True
        End If
        
    End If

End Function

' =========================================================
' ▽設定情報シートを生成するメソッド
'
' 概要　　　：シートが既に存在する場合は何もしない。存在しない場合は挿入を行う。
' 引数　　　：book      ブック
'     　　　：sheetName シート名
'     　　　：isGetOnly シートオブジェクトの取得のみフラグ
' 戻り値　　：挿入したシート
'
' =========================================================
Private Function createSheet(ByVal book As Workbook, ByVal sheetName As String, Optional ByVal isGetOnly As Boolean = False) As Worksheet

    On Error Resume Next

    ' ------------------------------------------
    ' 画面の更新を抑止する
    Application.screenUpdating = False

    Dim returnSheet As Worksheet
    Set returnSheet = ActiveSheet
    ' ------------------------------------------

    ' 対象シート
    Dim targetSheet     As Worksheet
    ' 対象シートを取得する
    Set targetSheet = book.Worksheets(sheetName)
    
    If isGetOnly = True Then
        ' 取得のみ
        ' シートを保護する
        targetSheet.Protect UserInterfaceOnly:=True
        Set createSheet = targetSheet
        Exit Function
    End If

    ' 対象シートが取得できた場合（元々存在していない場合もある）
    If targetSheet Is Nothing Then
        Set targetSheet = book.Worksheets.Add
        targetSheet.name = sheetName
    End If

    ' ▽新シートにテーブル定義情報を埋め込む
    With targetSheet

        ' シートを可視にする
        .visible = xlSheetVisible
        ' シート見出しに色を付ける
        .Tab.colorIndex = ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_COLOR_INDEX
        ' シートを保護する
        .Protect UserInterfaceOnly:=True
        ' メッセージを設定する
        .Range(ConstantsApplicationProperties.BOOK_PROPERTIES_WARN_MESSAGE_CELL).value = ConstantsApplicationProperties.BOOK_PROPERTIES_WARN_MESSAGE
    End With
    
    ' ------------------------------------------
    ' 画面の更新抑止を解除して、直前のシートをアクティブにし直す
    returnSheet.activate
    Application.screenUpdating = True
    ' ------------------------------------------
    
    On Error GoTo 0

    If targetSheet Is Nothing Then
        ' シートの作成に失敗した場合
        err.Raise ConstantsError.ERR_NUMBER_CREATE_WORKSHEET_FAILED, err.Source, ConstantsError.ERR_DESC_CREATE_WORKSHEET_FAILED, err.HelpFile, err.HelpContext
    End If

    Set createSheet = targetSheet

End Function

' =========================================================
' ▽ファイル名から列番号を検索する。
'
' 概要　　　：
' 引数　　　：sheet    シート名
'     　　　：fileName ファイル名
'     　　　：outIsNew （戻り値）新規かどうかのフラグ
' 戻り値　　：列番号
'
' =========================================================
Private Function findColumnFilePath(ByVal sheet As Worksheet, ByVal fileName As String, ByRef outIsNew As Boolean) As Long

    outIsNew = False
    findColumnFilePath = 0

    ' 既に設定済みの場合にはFindで検索する
    Dim findRange As Range
    Set findRange = sheet.Range(ConstantsIniData.ROW_FILE_NAME & ":" & ConstantsIniData.ROW_FILE_NAME).Find(what:=fileName, MatchCase:=False, MatchByte:=True, SearchFormat:=False)
    
    If Not findRange Is Nothing Then
        ' 見つかった場合
        findColumnFilePath = findRange.column
        Exit Function
    End If
    
    ' 存在しない場合は、新しい場所の列番号を返却する
    findColumnFilePath = 1
    Do While sheet.Cells(2, findColumnFilePath) <> ""
        findColumnFilePath = findColumnFilePath + 2
    Loop
    
    outIsNew = True

End Function

' =========================================================
' ▽コンテンツを読み込む。
'
' 概要　　　：
' 引数　　　：sheet    シート名
'     　　　：column   列番号
' 戻り値　　：コンテンツ
'
' =========================================================
Private Function readContents(ByVal sheet As Worksheet, ByVal column As Long) As String

    Dim ret As New StringBuilder

    If sheet.Cells(ConstantsIniData.ROW_FIRST, column) = "" Then
        ' 先頭が未入力の場合は、空の内容とみなす
        readContents = ret.str
        Exit Function
    End If
    
    ' 最終入力行を取得
    Dim lastInputtedRow As Long
    lastInputtedRow = ExcelUtil.getCellEndOfLastInputtedRow(sheet, column)

    Dim i As Long
    Dim findContentCells As Variant

    ' データ範囲を取得する
    Dim findRange As Range
    Set findRange = sheet.Range(sheet.Cells(ConstantsIniData.ROW_FIRST, column), sheet.Cells(lastInputtedRow, column + 1))

    If Not findRange Is Nothing Then
    
        ' 検索範囲から配列を取得する
        findContentCells = ExcelUtil.copyCellsToArray(sheet, ConstantsIniData.ROW_FIRST, findRange.End(xlDown).row - findRange.row + 1, column, 2)
        
        ' 配列からコンテンツを取得する
        For i = LBound(findContentCells, 1) To UBound(findContentCells, 1)
            
            If Mid(findContentCells(i, 1), 1, 1) = "[" And _
               Mid(findContentCells(i, 1), Len(findContentCells(i, 1)), 1) = "]" And _
               findContentCells(i, 2) = "" Then
               ' セクションの場合
                ret.append findContentCells(i, 1) & vbNewLine
               
            Else
                ' 通常の値
                ret.append findContentCells(i, 1) & "=" & findContentCells(i, 2) & vbNewLine
            
            End If

        Next
        
    End If
    
    readContents = ret.str

End Function

' =========================================================
' ▽コンテンツを書き込む。
'
' 概要　　　：
' 引数　　　：sheet    シート名
'     　　　：column   列番号
'     　　　：contents コンテンツ
' 戻り値　　：
'
' =========================================================
Private Sub writeContents(ByVal sheet As Worksheet, ByVal column As Long, ByRef contents As String)

    If sheet.Cells(ConstantsIniData.ROW_FIRST, column) = "" Then
        ' 先頭が未入力の場合は、空の内容とみなす
    
    Else
        
        ' 最終入力行を取得
        Dim lastInputtedRow As Long
        lastInputtedRow = ExcelUtil.getCellEndOfLastInputtedRow(sheet, column)

        ' 既に設定済みの場合には、一旦セルの内容を空にする
        Dim findRange As Range
        Set findRange = sheet.Range(sheet.Cells(ConstantsIniData.ROW_FIRST, column), sheet.Cells(lastInputtedRow, column + 1))
        findRange.delete xlShiftUp
    
    End If
    
    If contents = "" Then
        Exit Sub
    End If

    Dim i As Long
    Dim contentsArray As Variant
    Dim contentsKeyValueArray As Variant
    
    ' コンテンツを改行区切りにする（一行ずつセルに書き込みたいので）
    contentsArray = Split(contents, vbNewLine)
    
    Dim contentsCells() As Variant
    ReDim contentsCells(0 To UBound(contentsArray) - LBound(contentsArray), 0 To 1)
    
    For i = LBound(contentsArray) To UBound(contentsArray)
    
        ' key=value形式なので配列に分割する（ただしセクションの場合もあるので必ずしもkey=valueではない）
        contentsKeyValueArray = Split(contentsArray(i), "=")
        
        If contentsArray(i) = "" Then
            ' contentsArray(i)が""の場合は、Split関数に失敗しているので何もしない
            
        ElseIf VBUtil.arraySize(contentsKeyValueArray) > 1 Then
            ' key=value形式
            contentsCells(i, 0) = contentsKeyValueArray(0)
            contentsCells(i, 1) = contentsKeyValueArray(1)
        Else
            ' keyのみ（またはセクション）
            contentsCells(i, 0) = contentsKeyValueArray(0)
            contentsCells(i, 1) = ""
        End If
        
    Next
    
    ' 反映前に書式を文字列にする
    ExcelUtil.getArrayRange(contentsCells, sheet, ConstantsIniData.ROW_FIRST, column).NumberFormatLocal = "@"
    ' セルに反映する
    ExcelUtil.copyArrayToCells contentsCells, sheet, ConstantsIniData.ROW_FIRST, column

End Sub

' ------------------------------------------------------------------
' IniDataをラップして各メソッドを呼び出す
' ------------------------------------------------------------------
Public Function getValue _
            (ByVal SECTION As String, _
             ByVal key As String) As String

    getValue = inid.getValue(SECTION, key)
End Function

Public Sub setValue _
            (ByVal SECTION As String, _
             ByVal key As String, _
             ByVal value As String)

    inid.setValue SECTION, key, value
End Sub

Public Function getValues _
            (ByVal SECTION As String) As ValCollection

    Set getValues = inid.getValues(SECTION)
End Function

Public Sub setValues _
            (ByVal SECTION As String, _
             ByRef values As ValCollection)

    inid.setValues SECTION, values
End Sub

Public Sub delete _
            (ByVal SECTION As String, _
             Optional ByVal key As String = "")

    inid.delete SECTION, key
End Sub

