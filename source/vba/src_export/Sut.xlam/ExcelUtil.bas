Attribute VB_Name = "ExcelUtil"
Option Explicit

' *********************************************************
' Excelを簡便に利用するためのユーティリティモジュール
'
' 作成者　：Ison
' 履歴　　：2007/12/01　新規作成
' 　　　　　2009/06/21　Excelのバージョン取得関数を修正
' 　　　　　          　これにより、Excel2002が正常に認識されない(?)バグが修正された。
'
' 特記事項：
'
' *********************************************************

' Excelのクラス名
Private Const XLS_CLASSNAME As String = "XLMAIN"

' Excel2000・2003・2007で確認済み
' コマンドバーコントロールID フォント リスト
Private Const COMMAND_CONTROL_ID_FONT_LIST As Long = 1728

' Excel2000・2003・2007で確認済み
' コマンドバーコントロールID フォントサイズ リスト
Private Const COMMAND_CONTROL_ID_FONT_SIZE As Long = 1731

' Excelワークシートの禁止文字
Private Const EXCEL_SHEET_NAME_PROHIBITION_CHAR As String = "\[:]*/?"

' Excelワークシートのシート名最大長
Private Const EXCEL_SHEET_NAME_MAX_LENGTH As Long = 31

' Excelのバージョン
Public Enum ExcelVersion
    Ver2000 = 9
    Ver2002 = 10
    Ver2003 = 11
    Ver2007 = 12
    Ver2010 = 14
    Ver2013 = 15
    Ver2016 = 16
    VerOver = 99
    VerUnknown = -1
End Enum
    
' =========================================================
' ▽エクセルのバージョン取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：エクセルのバージョン
'
' =========================================================
Public Function getExcelVersion() As ExcelVersion

    Dim ver    As String
    
    ver = Application.version
    
    getExcelVersion = VerUnknown
    
    Select Case ver
    
        Case "9.0"
            getExcelVersion = Ver2000
        
        Case "10.0"
            getExcelVersion = Ver2002
        
        Case "11.0"
            getExcelVersion = Ver2003
        
        Case "12.0"
            getExcelVersion = Ver2007
        
        Case "14.0"
            getExcelVersion = Ver2010
        
        Case "15.0"
            getExcelVersion = Ver2013
        
        Case "16.0"
            getExcelVersion = Ver2016
            
    End Select
    
    ' 数値に変換できるか？
    If IsNumeric(ver) = False Then
    
        Exit Function
    End If
    
    Dim verSin    As Single
    verSin = CSng(ver)
    
    ' マイナーバージョンを考慮して以下の処理を実行
    If getExcelVersion = VerUnknown Then
    
        If verSin >= 17 Then
        
            getExcelVersion = VerOver
        ElseIf verSin >= 16 Then
        
            getExcelVersion = Ver2016
        ElseIf verSin >= 15 Then
        
            getExcelVersion = Ver2013
        ElseIf verSin >= 14 Then
        
            getExcelVersion = Ver2010
        ElseIf verSin >= 12 Then
        
            getExcelVersion = Ver2007
        ElseIf verSin >= 11 Then
        
            getExcelVersion = Ver2003
        ElseIf verSin >= 10 Then
        
            getExcelVersion = Ver2002
        ElseIf verSin >= 9 Then
        
            getExcelVersion = Ver2000
        End If
    
    End If
End Function

' =========================================================
' ▽Excelアプリケーションのウィンドウハンドル取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：ウィンドウハンドル
'
' =========================================================
#If VBA7 And Win64 Then

    Public Function getApplicationHWnd() As LongPtr
    
        Dim e As ExcelVersion
        
        Dim app As Object
        Set app = Excel.Application
        
        e = getExcelVersion
        
        If e >= Ver2002 Then
        
            getApplicationHWnd = app.hwnd
        
        Else
        
            getApplicationHWnd = WinAPI_User.FindWindow(XLS_CLASSNAME, Application.Caption)
        End If
        
    End Function
#Else

    Public Function getApplicationHWnd() As Long
    
        Dim e As ExcelVersion
        
        Dim app As Object
        Set app = Excel.Application
        
        e = getExcelVersion
        
        If e >= Ver2002 Then
        
            getApplicationHWnd = app.hwnd
        
        Else
        
            getApplicationHWnd = WinAPI_User.FindWindow(XLS_CLASSNAME, Application.Caption)
        End If
        
    End Function
#End If

' =========================================================
' ▽ウィンドウ最前面表示
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：
'
' =========================================================
Public Sub setUserFormTopMost(ByVal form As Object, Optional ByVal topmost As Boolean = True)

    ' ウィンドウハンドル
    #If VBA7 And Win64 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd As Long
    #End If

    Dim ret As Long

    ' フォームのウィンドウハンドルを取得する
    'ret = WinAPI_OLEACC.WindowFromAccessibleObject(form, hwnd)
    
    ' フォームキャプションを保存しておく
    Dim formCaption As String: formCaption = form.Caption
    ' 空白を設定してフォームキャプションが重複しないようにする
    form.Caption = formCaption & "                                "
    
    ' フォームのウィンドウハンドルを取得する
    hwnd = WinAPI_User.FindWindow("ThunderDFrame", form.Caption)
    
    ' フォームキャプションを元に戻す
    form.Caption = formCaption
    
    If hwnd <> 0 Then
    
        ' フォームを最前面表示する
        If topmost Then
        
            ret = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            
        Else
            ret = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

        End If
    End If
    
End Sub

' =========================================================
' ▽タイトルバーを外したウィンドウスタイルを設定する。
'
' 概要　　　：
' 引数　　　：uForm ユーザーフォーム
'
' 戻り値　　：true 成功、false 失敗
'
' =========================================================
Public Function setNonTitleBarWindowStyle(ByRef uForm As Object) As Boolean

    Dim ret As Long

    ' ウィンドウハンドル
    #If VBA7 And Win64 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd As Long
    #End If
  
    ' スタイル適用前のフォームのサイズを取得する
    Dim formWidth  As Double
    Dim formHeight As Double
    
    formWidth = uForm.InsideWidth
    formHeight = uForm.InsideHeight
  
    ' ウィンドウハンドルを取得する
    WinAPI_OLEACC.WindowFromAccessibleObject uForm, hwnd

    ' ダイアログの枠を除去
    ret = WinAPI_User.SetWindowLong(hwnd _
                      , GWL_EXSTYLE _
                      , WinAPI_User.GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME)
    If ret = 0 Then
        setNonTitleBarWindowStyle = False
    End If
    
    ' タイトルバーを除去
    ret = WinAPI_User.SetWindowLong(hwnd _
                      , GWL_STYLE _
                      , WinAPI_User.GetWindowLong(hwnd, GWL_STYLE) And Not WS_CAPTION)
    If ret = 0 Then
        setNonTitleBarWindowStyle = False
    End If
    
    ' メニューバーを再描画
    ret = WinAPI_User.DrawMenuBar(hwnd)
    If ret = 0 Then
        setNonTitleBarWindowStyle = False
    End If
    
    ' サイズ調整
    uForm.Width = uForm.Width - uForm.InsideWidth + formWidth
    uForm.Height = uForm.Height - uForm.InsideHeight + formHeight
    
End Function

' =========================================================
' ▽セルをアクティブにする
'
' 概要　　　：任意のシートのセルをアクティブにする
' 引数　　　：sheetName シート名
'             row       行
'             col       列
'
' =========================================================
Public Sub activateCell(sheetName As String, row As Long, col As Long)

    Worksheets(sheetName).Cells(row, col).activate

End Sub

' =========================================================
' ▽シート削除
'
' 概要　　　：任意のブックのシートを削除する。シートが無い場合は何もしない。
' 引数　　　：targetBook  ワークブック
' 　　　　　　targetSheet 削除対象シート
'
'
' 戻り値　　：True 削除成功
'
' =========================================================
Public Function deleteSheet(ByRef targetBook As Workbook, ByVal targetSheet As String) As Boolean

    ' シート
    Dim sheet As Worksheet
    
    ' 削除フラグ
    deleteSheet = False
    
    ' ブック内のシートを走査する
    For Each sheet In targetBook.Worksheets
    
        ' 削除対象シートかどうかを判断する
        If sheet.name = targetSheet Then
        
            ' シートを削除する
            sheet.delete
            
            ' 削除フラグをONにする
            deleteSheet = True
            
            Exit Function
        
        End If
    
    Next

End Function

' =========================================================
' ▽シートをコピーする
'
' 概要　　　：任意のブックのシートをコピーして、新しい名前を付ける。
' 引数　　　：copyBook       コピー元ブック
' 　　　　　　copySheetName  コピー元シート
' 　　　　　　newBook        コピー先ブック
'             newSheetName   コピー先シート
'             baseSheetName  新しいシートを配置する基準となるシート
'             direction      新しいシートを配置する基準となるシートに対して前方に配置するか後方に配置するか
'
' =========================================================
Public Sub copySheet(ByRef copyBook As Workbook, _
                     ByRef copySheetName As String, _
                     ByRef newBook As Workbook, _
                     ByRef newSheetName As String, _
                     ByRef baseSheetName As String, _
                     Optional ByVal direction As String = "after")


    ' 後方に配置
    If direction = "after" Then
    
        copyBook.Worksheets(copySheetName).copy after:=newBook.Worksheets(baseSheetName)
        
    ' 前方に配置
    Else
    
        copyBook.Worksheets(copySheetName).copy before:=newBook.Worksheets(baseSheetName)
    End If
    
    ' コピー後は必ずアクティブシートが新しいシートになる
    ' アクティブシートの名前を変更する
    ActiveSheet.name = newSheetName

End Sub

' =========================================================
' ▽ブックの最後尾にシートをコピーする
'
' 概要　　　：詳細はcopySheetを参照
'
' =========================================================
Public Sub copySheetAppend(ByRef copyBook As Workbook _
                         , ByRef copySheetName As String _
                         , ByRef newBook As Workbook _
                         , ByRef newSheetName As String)
                     
    copySheet _
        copyBook _
      , copySheetName _
      , newBook _
      , newSheetName _
      , newBook.Worksheets(newBook.Worksheets.count).name
                     
End Sub

' =========================================================
' ▽セルコピー
'
' 概要　　　：任意のシートのセル範囲を指定して、同シートの別のセル範囲に貼り付ける
' 引数　　　：sheetName     シート名
'             srcStartRow   コピー元開始行
'             srcStartCol   コピー元開始列
'             srcEndRow     コピー元終了行
'             srcEndCol     コピー元終了列
'             desStartRow   貼付け先開始行
'             desStartCol   貼付け先会熾烈
'             pasteType     コピーする値の種類（値のみや書式のみ等）
'
' =========================================================
Public Sub copyCell(sheetName As String, _
                        srcStartRow As Long, _
                        srcStartCol As Long, _
                        srcEndRow As Long, _
                        srcEndCol As Long, _
                        desStartRow As Long, _
                        desStartCol As Long, _
                        Optional pasteType As Variant = xlPasteAll)

    Dim sheet    As Worksheet
    Dim srcRange As Range
    Dim desRange As Range
    
    
    Set sheet = Worksheets(sheetName)
    sheet.activate
    
    Set srcRange = sheet.Range( _
        sheet.Cells(srcStartRow, srcStartCol), _
        sheet.Cells(srcEndRow, srcEndCol) _
    )
        
    Set desRange = sheet.Range( _
        sheet.Cells(desStartRow, desStartCol), _
        sheet.Cells(desStartRow + srcEndRow - srcStartRow, desStartCol + srcEndCol - srcStartCol) _
    )
    
    srcRange.copy
    
    desRange.PasteSpecial Paste:=pasteType

    Application.CutCopyMode = False
    
End Sub

' =========================================================
' ▽シートが存在しているかどうかの確認
'
' 概要　　　：
' 引数　　　：book          ワークブック
' 　　　　　　sheetName     シート名
' 戻り値　　：True シートが存在する
'
' =========================================================
Public Function existSheet(ByRef book As Workbook _
                         , ByVal sheetName As String) As Boolean
                         
                         
    On Error GoTo err
    
    ' シートオブジェクト
    Dim sheet As Worksheet
    ' シートオブジェクトを取得する
    Set sheet = book.Worksheets(sheetName)
    
    ' 正常に終了できた場合、シートは存在している
    existSheet = True
    
    Exit Function
err:

    ' エラーになった場合、シートは存在していない
    existSheet = False

End Function

' =========================================================
' ▽一意なシート名への変換
'
' 概要　　　：任意のシート名が対象となるブックに既に存在する場合
' 　　　　　　一意なシート名に変換を行う。
' 引数　　　：book          ワークブック
' 　　　　　　sheetName     シート名
' 戻り値　　：一意なシート名への変換
'
' =========================================================
Public Function convertUniqSheetName(ByRef book As Workbook _
                                   , ByVal sheetName As String) As String
                         
                         
    On Error Resume Next
    
    Dim i As Long: i = 1
    
    ' 変換後のシート名
    Dim convertSheetName As String: convertSheetName = sheetName
    ' 変換後のシート名　接尾辞
    Dim convertSheetNameSuffix As String
    
    ' シートオブジェクト
    Dim sheet As Worksheet
    
    ' 999回シート名の変換を行っても、それでも変換できない場合はループを終了する
    Do While i < 1000
    
        i = i + 1
        
        ' シートオブジェクトを取得する
        Set sheet = book.Worksheets(convertSheetName)
    
        ' エラーが発生していないかを確認
        If err.Number <> 0 Then
        
            ' エラーが発生している（＝シートは存在しない）ので処理を抜ける
            convertUniqSheetName = convertSheetName
            
            Exit Function
            
        End If
    
        ' 変換後のシート名 接尾辞を設定する
        convertSheetNameSuffix = " (" & i & ")"
        
        ' シート名の規定の長さを超えていないかを確認する
        If checkMaxLengthOfSheetName(sheetName & convertSheetNameSuffix) = False Then
        
            ' 規定の長さを超えている場合、長さを調整する
            convertSheetName = Mid$(sheetName _
                                  , 1 _
                                  , EXCEL_SHEET_NAME_MAX_LENGTH - Len(convertSheetNameSuffix)) & convertSheetNameSuffix
                                  
        Else
        
            convertSheetName = sheetName & convertSheetNameSuffix
        End If
        
    
    Loop
    
    On Error GoTo 0
    
End Function

' =========================================================
' ▽セルコピー（書式のみ）
'
' 概要　　　：任意のシートのセル範囲を指定して、同シートの別のセル範囲に書式情報のみを貼り付ける
' 引数　　　：sheetName     シート名
'             srcStartRow   コピー元開始行
'             srcStartCol   コピー元開始列
'             srcEndRow     コピー元終了行
'             srcEndCol     コピー元終了列
'             desStartRow   貼付け先開始行
'             desStartCol   貼付け先開始列
'
' =========================================================
Public Sub copyCellFormat(sheetName As String, _
                        srcStartRow As Long, _
                        srcStartCol As Long, _
                        srcEndRow As Long, _
                        srcEndCol As Long, _
                        desStartRow As Long, _
                        desStartCol As Long)

    copyCell sheetName, _
             srcStartRow, _
             srcStartCol, _
             srcEndRow, _
             srcEndCol, _
             desStartRow, _
             desStartCol, _
             xlPasteFormats
    
End Sub

Public Sub fillBgColor(ByVal sheetName As String, _
                       ByVal startRow As Long, _
                       ByVal startCol As Long, _
                       ByVal endRow As Long, _
                       ByVal endCol As Long, _
                       ByVal colorIndex As Long)

    Dim r As Range
    Dim sheet As Worksheet
    Set sheet = Worksheets(sheetName)

    Set r = sheet.Range( _
                sheet.Cells(startRow, startCol), _
                sheet.Cells(endRow, endCol) _
            )

    ' 既存の色を確認
    If r.Interior.colorIndex <> colorIndex Then
    
        ' 既存の色と新しい色が異なる場合、新しい色で上書きする
        r.Interior.colorIndex = colorIndex
    End If
    

End Sub

' =========================================================
' ▽コメントを追加する
'
' 概要　　　：任意のセルにコメントを追加する
' 引数　　　：sheet     シート名
'             row   行
'             col   列
'             text  テキスト
'
' =========================================================
Public Sub addCommentForWorkSheet(ByVal sheet As Worksheet, _
                      ByVal row As Long, _
                      ByVal col As Long, _
                      ByVal text As String)
    
    ' レンジオブジェクト
    Dim r As Range
    
    ' レンジオブジェクトを取得する
    Set r = sheet.Range( _
                sheet.Cells(row, col), _
                sheet.Cells(row, col) _
            )
    
    ' コメントが既に存在するセルに対して、コメントを追加するとエラーが発生する。
    ' そのため、エラー発生時にも処理を継続するようにする
    On Error Resume Next
    
    With r.Cells(1, 1)
    
        .ClearComments
        .addComment
        
        .comment.Shape.TextFrame.AutoSize = True
        .comment.Shape.TextFrame.Characters.Font.name = "ＭＳ ゴシック"
        ' .Comment.Visible = True
        .comment.text text:=text
    
    End With
    
    On Error GoTo 0

End Sub

' =========================================================
' ▽コメントを追加する
'
' 概要　　　：任意のセルにコメントを追加する
' 引数　　　：sheetName     シート名
'             row   行
'             col   列
'             text  テキスト
'
' =========================================================
Public Sub addComment(ByVal sheetName As String, _
                      ByVal row As Long, _
                      ByVal col As Long, _
                      ByVal text As String)
    
    addCommentForWorkSheet Worksheets(sheetName), row, col, text

End Sub

' =========================================================
' ▽コメントを削除する
'
' 概要　　　：任意のセルにコメントを削除する
' 引数　　　：sheetName     シート名
'             row   行
'             col   列
'
' =========================================================
Public Sub deleteComment(sheetName As String, _
                         row As Long, _
                         col As Long)

    ' レンジオブジェクト
    Dim r As Range
    ' コメントオブジェクト
    Dim c As comment
    
    Dim sheet As Worksheet
    Set sheet = Worksheets(sheetName)
    
    ' レンジオブジェクトを取得する
    Set r = sheet.Range( _
                sheet.Cells(row, col), _
                sheet.Cells(row, col) _
            )
            
    ' コメントオブジェクトを取得する
    Set c = r.Cells(1, 1).comment
    
    ' コメントが存在するか確認する
    If Not c Is Nothing Then
    
        r.Cells(1, 1).ClearComments
    End If

End Sub

' =========================================================
' ▽ハイパーリンク追加
'
' 概要　　　：任意のシートのセルにハイパーリンクを追加する
' 引数　　　：sheetName           シート名
'             row                 行
'             col                 列
'             text                テキスト
'             linkTargetSheetName リンク先のシート
'             linkTargetCellRow   リンク先のセル（行）
'             linkTargetCellCol   リンク先のセル（列）
'             book                ワークブック
'
' =========================================================
Public Sub addHyperLinkInBook(ByVal sheetName As String, _
                              ByVal row As Long, _
                              ByVal col As Long, _
                              ByVal text As String, _
                              ByVal linkTargetSheetName As String, _
                              Optional ByVal linkTargetCellRow As Long = 1, _
                              Optional ByVal linkTargetCellCol As Long = 1, _
                              Optional ByRef book As Workbook)

    Dim sheet  As Worksheet
    Dim r      As Range
    Dim cellRc As String
    
    ' 引数bookが省略されていた場合、アクティブなワークブックを対象とする
    If book Is Nothing Then
    
        Set book = ActiveWorkbook
    End If
    
    Set sheet = book.Worksheets(sheetName)
    sheet.activate
    
    Set r = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col))
    
    ' R1C1形式でセル位置を指定する
    cellRc = "R" & linkTargetCellRow & "C" & linkTargetCellCol
    
    sheet.Hyperlinks.Add _
        anchor:=r, _
        Address:="", _
        SubAddress:="#" & linkTargetSheetName & "!" & cellRc, _
        TextToDisplay:=text


End Sub

' =========================================================
' ▽フォントサイズを変更する
'
' 概要　　　：任意のシートのセルのフォントサイズを変更する
' 引数　　　：sheetName  シート名
'             row        行
'             col        列
'             fontSize   フォントサイズ
'             book       ワークブック
'
' =========================================================
Public Sub changeFontSize(ByVal sheetName As String, ByVal row As Long, ByVal col As Long, fontSize As Long, Optional book As Workbook)

    Dim sheet As Worksheet
    Dim r     As Range
    
    ' 引数bookが省略されていた場合、アクティブなワークブックを対象とする
    If book Is Nothing Then
    
        Set book = ActiveWorkbook
    End If
    
    Set sheet = book.Worksheets(sheetName)
    sheet.activate
    
    Set r = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col))

    r.Font.size = fontSize
    
End Sub

' =========================================================
' ▽セル文字列への変換処理
'
' 概要　　　：任意の文字列をセル文字列に変換する
' 引数　　　：val 文字列
'
' =========================================================
Public Function convertCellValue(ByRef val As Variant) As Variant

    ' 戻り値を初期化する
    convertCellValue = val
    
    ' 文字列の先頭がシングルクォーテーションであるかを確認する
    If Mid(val, 1, 1) = "'" Then
    
        ' 文字列の先頭がシングルクォーテーションの場合、さらにシングルクォーテーションを付加する
        convertCellValue = "'" & val
        
    End If
    
End Function

' =========================================================
' ▽セル文字列への変換処理
'
' 概要　　　：任意の文字列をセル文字列に変換する
' 引数　　　：val 文字列
'
' =========================================================
Public Function convertCellStrValue(ByRef val As Variant) As Variant

    ' 戻り値を初期化する
    convertCellStrValue = val

    If isNull(val) Then
        Exit Function
    End If
    
    ' 文字列の先頭がシングルクォーテーションであるかを確認する
    If Mid(val, 1, 1) = "'" Then
    
        ' 文字列の先頭がシングルクォーテーションの場合、さらにシングルクォーテーションを付加する
        convertCellStrValue = "'" & val
        
    Else
    
        convertCellStrValue = CStr(val)
    End If
    
End Function

' =========================================================
' ▽任意の行数から最終行までの行削除
'
' 概要　　　：任意の行数から最終行までの行を削除する。
' 引数　　　：sheet 任意のシート
' 　　　　　　row   任意の行
' 　　　　　　col   任意の列
' 戻り値　　：
'
' =========================================================
Public Sub deleteRowEndOfLastInputted(ByRef sheet As Worksheet, ByVal row As Long, ByVal col As Long)

    ' 削除対象範囲
    Dim targetRange As Range
    ' レコードオフセット位置
    Dim recordOffset As Long
    ' 最後尾の入力箇所
    Dim length As Long
    
    ' レコードオフセット位置
    recordOffset = row
    ' 最後尾の入力箇所を取得する
    length = ExcelUtil.getCellEndOfLastInputtedRow(sheet, col)
    
    If length < recordOffset Then
        length = recordOffset
    End If
    
    ' 削除対象範囲を取得
    Set targetRange = sheet _
                        .Range( _
                           sheet.Cells(recordOffset _
                                     , 1).Address & ":" & _
                           sheet.Cells(length _
                                     , 1).Address)
    
    ' 削除する（行単位で削除）
    targetRange.EntireRow.delete

End Sub

' =========================================================
' ▽有効行取得
'
' 概要　　　：任意のシートの任意の列の有効行を取得する。
' 　　　　　　（一番最後尾に入力されている行）
' 引数　　　：sheet 任意のシート
' 　　　　　　col   任意の列
' 戻り値　　：有効行
'
' =========================================================
Public Function getCellEndOfLastInputtedRow(ByRef sheet As Worksheet, ByVal col As Long) As Long

    ' 最大行
    Dim max As Long
    ' 最大行サイズを取得する
    max = getSizeOfSheetRow(sheet)

    ' 有効行を求める。
    If CStr(sheet.Cells(max, col)) <> "" Then
    
        ' Excelの最大行数のセル位置に設定値がある場合
        ' 最大行数位置を返す
        getCellEndOfLastInputtedRow = max
    
    Else
        ' Excelの最大行数から上方向に空白でないセルを探す
        getCellEndOfLastInputtedRow = sheet.Cells(max, col).End(xlUp).row
        
    End If
    
End Function

' =========================================================
' ▽有効列取得
'
' 概要　　　：任意のシートの任意の列の有効列を取得する。
' 　　　　　　（一番最後尾に入力されている列）
' 引数　　　：sheet 任意のシート
' 　　　　　　row   任意の行
' 戻り値　　：有効列
'
' =========================================================
Public Function getCellEndOfLastInputtedCol(ByRef sheet As Worksheet, ByVal row As Long) As Long
    
    ' 最大列
    Dim max As Long
    ' 最大列サイズを取得する
    max = getSizeOfSheetCol(sheet)

    ' 有効行を求める。
    If CStr(sheet.Cells(row, max)) <> "" Then
    
        ' Excelの最大行数のセル位置に設定値がある場合
        ' 最大行数位置を返す
        getCellEndOfLastInputtedCol = max
    
    Else
        ' Excelの最大列数から左方向に空白でないセルを探す
        getCellEndOfLastInputtedCol = sheet.Cells(row, max).End(xlToLeft).column
        
    End If
    
End Function

' =========================================================
' ▽シートの最大行サイズ取得
'
' 概要　　　：任意のシートの最大行サイズを取得する。
' 引数　　　：sheet 任意のシート
' 戻り値　　：最大行サイズ
'
' =========================================================
Public Function getSizeOfSheetRow(ByRef sheet As Worksheet) As Long

    #If EXCEL_SHEET_ROW_SIZE_256 = 1 Then
    
        getSizeOfSheetRow = 260
    #Else
    
        ' Rangeオブジェクトを参照しカラム全体を選択し（EntireColumn）行のカウントを取得する
        getSizeOfSheetRow = sheet.Range("A1").EntireColumn.Rows.count
    #End If

End Function

' =========================================================
' ▽シートの最大列サイズ取得
'
' 概要　　　：任意のシートの最大列サイズを取得する。
' 引数　　　：sheet 任意のシートa
' 戻り値　　：最大行サイズ
'
' =========================================================
Public Function getSizeOfSheetCol(ByRef sheet As Worksheet) As Long

    ' Rangeオブジェクトを参照し行全体を選択し（EntireRow）列のカウントを取得する
    getSizeOfSheetCol = sheet.Range("A1").EntireRow.Columns.count

End Function

' =========================================================
' ▽シート名の禁止文字チェック
'
' 概要　　　：任意のシートに禁止文字が含まれているかをチェックする。
' 引数　　　：sheetName 任意のシート名
' 戻り値　　：True 正常（禁止文字が含まれていない場合）
'
' =========================================================
Public Function checkProhibitionCharOfSheetName(ByVal sheetName As String) As Boolean

    ' 仮実装
    checkProhibitionCharOfSheetName = True

    ' 禁止文字（1文字）
    Dim char As String
    ' インデックス
    Dim i As Long
    
    ' 禁止文字を1文字ずつ取り出し、シート名に禁止文字が含まれていないかをチェックする
    For i = 1 To Len(EXCEL_SHEET_NAME_PROHIBITION_CHAR)
    
        ' 1文字取り出す
        char = Mid$(EXCEL_SHEET_NAME_PROHIBITION_CHAR, i, 1)
        
        ' 禁止文字を発見した場合
        If InStr(sheetName, char) <> 0 Then
        
            ' 禁止文字が含まれているのでFalseに設定
            checkProhibitionCharOfSheetName = False
            
            Exit Function
        End If
    
    Next

End Function

' =========================================================
' ▽シート名の禁止文字変換
'
' 概要　　　：任意のシートに禁止文字が含まれている場合、変換を実施する。
' 引数　　　：sheetName 任意のシート名
' 戻り値　　：変換後のシート名
'
' =========================================================
Public Function convertProhibitionCharOfSheetName(ByVal sheetName As String) As String

    ' シート名として有効な文字
    Const VALID_CHAR As String = "_"

    ' 禁止文字（1文字）
    Dim char As String
    ' インデックス
    Dim i As Long
    
    ' 禁止文字を1文字ずつ取り出し、シート名に禁止文字が含まれていないかをチェックする
    For i = 1 To Len(EXCEL_SHEET_NAME_PROHIBITION_CHAR)
    
        ' 1文字取り出す
        char = Mid$(EXCEL_SHEET_NAME_PROHIBITION_CHAR, i, 1)
        
        ' 禁止文字を発見した場合
        If InStr(sheetName, char) <> 0 Then
        
            ' 禁止文字を有効な文字に変換する
            sheetName = replace(sheetName, char, VALID_CHAR)
        End If
    
    Next
    
    ' 戻り値として返す
    convertProhibitionCharOfSheetName = sheetName

End Function

' =========================================================
' ▽シート名の変換
'
' 概要　　　：シートの名称が規定値を超えている場合に、規定値に収まるように変換を行う
' 引数　　　：sheetName 任意のシート名
' 戻り値　　：変換後のシート名
'
' =========================================================
Public Function truncateExceededSheetName(ByVal sheetName As String) As String

    truncateExceededSheetName = Mid$(sheetName, 1, 28) & "..."
    
End Function

' =========================================================
' ▽シート名のサイズチェック
'
' 概要　　　：任意のシート名の桁長が規定サイズを超えていないかをチェックする。
' 引数　　　：sheetName 任意のシート名
' 戻り値　　：True 正常（禁止文字が含まれていない場合）
'
' =========================================================
Public Function checkMaxLengthOfSheetName(ByVal sheetName As String) As Boolean

    ' 戻り値
    Dim ret As Boolean

    ret = True
    
    ' シート名の桁長が最大文字長を超えている場合
    If Len(sheetName) > EXCEL_SHEET_NAME_MAX_LENGTH Then
    
        ret = False
    End If

    ' 戻り値に結果を設定
    checkMaxLengthOfSheetName = ret
    
End Function

' =========================================================
' ▽最大行サイズを超えているかをチェック
'
' 概要　　　：任意のシートの最大行サイズが超えているかをチェックする。
' 引数　　　：sheet     シートオブジェクト
' 　　　　　　rowOffset 行オフセット
' 　　　　　　rowSize   行サイズ
' 戻り値　　：True  最大行サイズの範囲内
' 　　　　　　False 最大行サイズの範囲外
'
' =========================================================
Public Function checkOverMaxRow(ByRef sheet As Worksheet _
                              , ByVal rowOffset As Long _
                              , Optional ByVal rowSize As Long = 1) As Boolean

    ' シートの最大行
    Dim max As Long
    ' シートの最大行を取得
    max = getSizeOfSheetRow(sheet)
    
    ' 最大行を超えているかをチェック
    If max < rowOffset + rowSize - 1 Then
    
        ' 最大行を超えているので False
        checkOverMaxRow = False
    
    Else
    
        ' 最大行を超えていないので True
        checkOverMaxRow = True
    End If

End Function

' =========================================================
' ▽最大列サイズを超えているかをチェック
'
' 概要　　　：任意のシートの最大列サイズが超えているかをチェックする。
' 引数　　　：sheet     シートオブジェクト
' 　　　　　　colOffset 列オフセット
' 　　　　　　colSize   列サイズ
' 戻り値　　：True  最大列サイズの範囲内
' 　　　　　　False 最大列サイズの範囲外
'
' =========================================================
Public Function checkOverMaxCol(ByRef sheet As Worksheet _
                              , ByVal colOffset As Long _
                              , Optional ByVal colSize As Long = 1) As Boolean

    ' シートの最大行
    Dim max As Long
    ' シートの最大行を取得
    max = getSizeOfSheetCol(sheet)
    
    ' 最大行を超えているかをチェック
    If max < colOffset + colSize - 1 Then
    
        ' 最大行を超えているので False
        checkOverMaxCol = False
    
    Else
    
        ' 最大行を超えていないので True
        checkOverMaxCol = True
    End If

End Function

' =========================================================
' ▽配列からRangeを取得する
'
' 概要　　　：
' 引数　　　：val       配列
'             sheet     シート
' 　　　　　　offsetRow 行
' 　　　　　　offsetCol 列
' 　　　　　　rowSize   行サイズ
' 　　　　　　colSize   列サイズ
'
' =========================================================
Public Function getArrayRange(ByRef val As Variant _
                            , ByRef sheet As Worksheet _
                            , ByVal rowOffset As Long _
                            , ByVal colOffset As Long _
                            , Optional ByVal rowSize As Long = -1 _
                            , Optional ByVal colSize As Long = -1) As Range

    If IsArray(val) = False Then
    
        Exit Function
        
    End If

    If rowSize = -1 Then
        rowSize = VBUtil.arraySize(val)
    End If
    
    If colSize = -1 Then
        colSize = VBUtil.arraySize(val, 2)
    End If

    Set getArrayRange = sheet.Range(sheet.Cells(rowOffset _
                                              , colOffset) _
                      , sheet.Cells(rowOffset + rowSize - 1 _
                                              , colOffset + colSize - 1))
    
End Function

' =========================================================
' ▽配列内容コピー
'
' 概要　　　：セルに配列内容をコピーする
' 引数　　　：val       配列
'             sheet     シート
' 　　　　　　offsetRow 行
' 　　　　　　offsetCol 列
' 　　　　　　rowSize   行サイズ
' 　　　　　　colSize   列サイズ
'
' =========================================================
Public Sub copyArrayToCells(ByRef val As Variant _
                          , ByRef sheet As Worksheet _
                          , ByVal rowOffset As Long _
                          , ByVal colOffset As Long _
                          , Optional ByVal rowSize As Long = -1 _
                          , Optional ByVal colSize As Long = -1)

    If IsArray(val) = False Then
    
        Exit Sub
        
    End If

    If rowSize = -1 Then
        rowSize = VBUtil.arraySize(val)
    End If
    
    If colSize = -1 Then
        colSize = VBUtil.arraySize(val, 2)
    End If

    sheet.Range(sheet.Cells(rowOffset _
                          , colOffset) _
              , sheet.Cells(rowOffset + rowSize - 1 _
                          , colOffset + colSize - 1)) = val

    'sheet.Cells(rowOffset, colOffset).Resize(rowSize, colSize) = val
    
End Sub

' =========================================================
' ▽配列内容コピー（カラム情報用）
'
' 概要　　　：セルに配列内容をコピーする
' 引数　　　：val       配列
'             sheet     シート
' 　　　　　　rowOffset 行
' 　　　　　　colOffset 列
' 　　　　　　colSize   列サイズ
'
' =========================================================
Public Sub copyArrayToCellsForColumns(ByRef val As Variant _
                                    , ByRef sheet As Worksheet _
                                    , ByVal rowOffset As Long _
                                    , ByVal colOffset As Long _
                                    , Optional ByVal colSize As Long = -1)

    If IsArray(val) = False Then
    
        Exit Sub
        
    End If

    Dim rowSize As Long: rowSize = 1

    If colSize = -1 Then
        colSize = VBUtil.arraySize(val)
    End If

    sheet.Range(sheet.Cells(rowOffset _
                          , colOffset) _
              , sheet.Cells(rowOffset + rowSize - 1 _
                          , colOffset + colSize - 1)) = val

    'sheet.Cells(rowOffset, colOffset).Resize(rowSize, colSize) = val
    
End Sub

' =========================================================
' ▽配列内容コピー
'
' 概要　　　：セルに配列内容をコピーする
' 引数　　　：val       配列
'             sheet     シート
' 　　　　　　rowOffset 行
' 　　　　　　rowSize   行サイズ
' 　　　　　　colOffset 列
' 　　　　　　colSize   列サイズ
'
' =========================================================
Public Function copyCellsToArray(ByRef sheet As Worksheet _
                               , ByVal rowOffset As Long _
                               , ByVal rowSize As Long _
                               , ByVal colOffset As Long _
                               , ByVal colSize As Long) As Variant

    Dim retArray As Variant
    Dim ret      As Variant

    Dim srcCell As String
    
    With sheet
    
        srcCell = .Cells(rowOffset _
                       , colOffset).Address & ":" & _
                  .Cells(rowOffset + rowSize - 1 _
                       , colOffset + colSize - 1).Address
                  
        ret = .Range(srcCell)

    End With
    
    ' 戻り値が配列ではない場合
    If IsArray(ret) = False Then
        
        ' ※Rangeオブジェクトのサイズが1の場合、配列以外のプリミティブ型が返ってくるので
        ' 　変換する必要がある
        ' サイズが1の配列を生成する
        ReDim retArray(1 To 1, 1 To 1)
    
        ' 値を代入する
        retArray(1, 1) = ret
        
    Else
    
        retArray = ret
    End If
    
    copyCellsToArray = retArray
    
End Function

' =========================================================
' ▽行の高さを変更
'
' 概要　　　：行の高さを変更する。
' 　　　　　　高さの単位はポイント。(Excelの仕様に準拠)
' 引数　　　：r    レンジオブジェクト（シートオブジェクト含む）
' 　　　　　　h    高さ
'
' =========================================================
Public Sub changeRowHeight(ByRef r As Range _
                         , ByVal h As Double)

    If h = -1 Then
    
        r.EntireRow.AutoFit
    Else
    
        r.EntireRow.RowHeight = h
    End If

End Sub

' =========================================================
' ▽列の幅を変更
'
' 概要　　　：列の幅を変更
' 　　　　　　幅の単位は文字数。(Excelの仕様に準拠)
' 引数　　　：r    レンジオブジェクト（シートオブジェクト含む）
' 　　　　　　w    幅
'
' =========================================================
Public Sub changeColWidth(ByRef r As Range _
                        , ByVal w As Double)

    If w = -1 Then
    
        r.EntireColumn.AutoFit
    Else
    
        r.EntireColumn.ColumnWidth = w
    End If

End Sub

' =========================================================
' ▽利用可能なフォントリストを取得
'
' 概要　　　：
' 引数　　　：
' 特記事項　：Excelブックが1つ以上開かれていないと取得に失敗するので注意
'
' =========================================================
Public Function getFontList() As ValCollection

    ' 戻り値
    Dim ret As New ValCollection

    Dim i As Long
    
    ' コマンドバーコントロール
    Dim c As commandBarControl

    ' フォントサイズリストを取得する
    Set c = Application.CommandBars.FindControl(Id:=COMMAND_CONTROL_ID_FONT_LIST)
    
    ' コントロールが取得できた場合
    If Not c Is Nothing Then
    
        ' リストの内容を全て戻り値に追加
        For i = 1 To c.ListCount
        
            ret.setItem c.list(i), c.list(i)
        
        Next

    End If
    
    ' 戻り値を返す
    Set getFontList = ret

End Function

' =========================================================
' ▽フォントサイズリストを取得
'
' 概要　　　：
' 引数　　　：
' 特記事項　：Excelブックが1つ以上開かれていないと取得に失敗するので注意
'
' =========================================================
Public Function getFontSizeList() As ValCollection

    ' 戻り値
    Dim ret As New ValCollection

    Dim i As Long
    ' コマンドバーコントロール
'    Dim c As CommandBarControl

'    ' フォントサイズリストを取得する
'    Set c = Application.CommandBars.FindControl(ID:=COMMAND_CONTROL_ID_FONT_SIZE)
'
'    ' コントロールが取得できた場合
'    If Not c Is Nothing Then
'
'        Debug.Print TypeName(c)
'        ' リストの内容を全て戻り値に追加
'        For i = 1 To c.ListCount
'
'            ret.setItem c.list(i), c.list(i)
'
'        Next
'
'    End If
    
    ' Excel2000 - 2007の規定値のサイズをセットする
    ret.setItem 6
    ret.setItem 8
    ret.setItem 9
    ret.setItem 10
    ret.setItem 11
    ret.setItem 12
    ret.setItem 14
    ret.setItem 16
    ret.setItem 18
    ret.setItem 20
    ret.setItem 22
    ret.setItem 24
    ret.setItem 26
    ret.setItem 28
    ret.setItem 36
    ret.setItem 48
    ret.setItem 72
    
    
    ' 戻り値を返す
    Set getFontSizeList = ret

End Function

' =========================================================
' ▽Excel行列の数値をアルファベットに変換する
'
' 概要　　　：Excel行列の数値をアルファベットに変換する
' 引数　　　：row 行
' 引数　　　：col 列
' 戻り値　　：変換結果
'
' =========================================================
Public Function convertExcelNumberToAlpha(ByVal row As Long, ByVal col As Long) As String

    ' 戻り値
    Dim ret As String

    ' ゼロベースとする
    row = row - 1
    col = col - 1

    ' 基数を 26 とする （アルファベットの数）
    Dim base As Long: base = 26
    
    ' 現在基数べき乗数値（baseのn乗数）
    Dim curBase As Long

    ' 基数における数値の長さ
    Dim length As Long
    
    Dim i   As Long
    Dim tmp As Long

    ' 基数の対数を求める
    If col > 0 Then
        length = Application.WorksheetFunction.RoundDown(Application.WorksheetFunction.Log(col, base), 0)
    Else
        length = 0
    End If
    
    For i = length To 0 Step -1
    
        ' 現在の基数をもとにした桁数の、開始数値を求める
        ' 例：length = 1 ⇒ 26
        curBase = Application.WorksheetFunction.Power(base, i)
        If curBase <> 1 Then
        
            tmp = Application.WorksheetFunction.RoundDown(col / curBase, 0) - 1
        Else
            
            tmp = col - curBase + 1
        End If
        
        ' Chr(65) = A なので、Aを基準にアルファベットを算出する
        ret = ret & Chr(65 + tmp)
        
        ' 次の計算に備え列値を減算する
        col = col - (curBase * (tmp + 1))
    
    Next
    
    convertExcelNumberToAlpha = ret & "" & (row + 1)
    
End Function

' =========================================================
' ▽Excel列のアルファベットを数値に変換する
'
' 概要　　　：Excel列のアルファベットを数値に変換する
' 引数　　　：var 値
' 戻り値　　：変換結果
'
' =========================================================
Public Function convertExcelAlphaToNumber(ByVal var As String) As Long

    ' 正規表現オブジェクト
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    
    RE.Pattern = "^[A-Z]$" ' アルファベットを検索対象とする
    RE.IgnoreCase = True   ' 大文字と小文字を区別しない
    RE.Global = True       ' 文字列全体を検索
    
    Dim keta As Long
    
    Dim i As Long
    Dim c As String
    Dim length As Long: length = Len(var)
    
    For i = length To 1 Step -1
    
        ' 1文字取得する
        c = UCase$(Mid$(var, i, 1))
        
        ' アルファベットの場合
        If RE.test(c) Then
        
            convertExcelAlphaToNumber = convertExcelAlphaToNumber + (Asc(c) - Asc("A")) + (26 ^ keta)
            
            keta = keta + 1
        End If
    
    Next
    
    convertExcelAlphaToNumber = convertExcelAlphaToNumber
    
End Function

Public Sub protectSheet(ByVal sheetName As String)

    With Worksheets(sheetName)
    
        ' 一旦、シート保護を解除
        .Unprotect
        ' シート保護を設定
        .Protect _
            UserInterfaceOnly:=True, _
            contents:=True, _
            Scenarios:=True, _
            AllowFiltering:=True

        .EnableSelection = xlUnlockedCells
        
    End With
    
End Sub

Public Sub copyCommandBarControl(ByRef srcControl As Object, ByRef desControl As Object)

    ' エクセルのバージョン
    Static excelVer As ExcelVersion: excelVer = ExcelUtil.getExcelVersion

    With desControl
    
        .Style = srcControl.Style
        .Caption = srcControl.DescriptionText
        .DescriptionText = srcControl.DescriptionText
        .OnAction = srcControl.OnAction
        .Tag = srcControl.Tag
        .ShortcutText = srcControl.ShortcutText
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            .Picture = srcControl.Picture
            .mask = srcControl.mask
        End If
    
    End With
    
End Sub

Public Function showSaveConfirmDialog(book As Workbook) As VbMsgBoxResult

    showSaveConfirmDialog = MsgBox("'" & book.name & "'への変更を保存しますか？", vbYesNoCancel Or vbExclamation, "Microsoft Excel")

End Function

' =========================================================
' ▽アクティブブックを取得する。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：アクティブブック
'
' =========================================================
Public Function getActiveWorkbook() As Workbook

    On Error Resume Next
    
    ' 最初はアクティブシートから取得を試みる
    ' ※アドインマクロの場合に、アドイン自身のブック情報が取得される可能性があるため
    Set getActiveWorkbook = ActiveSheet.parent
    
    If err.Number <> 0 Then
    
        ' 次にアクティブブックから取得を試みる
        Set getActiveWorkbook = ActiveWorkbook
    
    End If
    
    On Error GoTo 0

End Function

