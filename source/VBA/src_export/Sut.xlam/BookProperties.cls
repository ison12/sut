VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ブック設定情報にアクセスするクラス
'
' 作成者　：Ison
' 履歴　　：2019/12/03　新規作成
'
' 特記事項：
' *********************************************************

' ---------------------------------------------------------
' □プロパティ
' ---------------------------------------------------------
' ワークブック
Private m_book As Workbook
' ワークシート
Private m_sheet As Worksheet
' プロパティシート
Private m_propertiesSheet As Worksheet


' =========================================================
' ▽プロパティ シート
' =========================================================
Public Property Let sheet(ByVal vNewValue As Variant)

    ' シートを設定する
    Set m_sheet = vNewValue
    
    If Not vNewValue Is Nothing Then
        ' シートが未設定の可能性があるのでチェックする
        Set m_book = m_sheet.parent
    End If
    
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
' ▽設定情報が存在するかを判定するメソッド
' =========================================================
Public Function isExistsProperties() As Boolean

    If m_book Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        isExistsProperties = False
        Exit Function
    End If

    Dim sheetName As String
    sheetName = ConstantsBookProperties.BOOK_PROPERTIES_SHEET_NAME

    On Error Resume Next

    ' 対象シート
    Dim targetSheet As Worksheet
    ' 対象シートを取得する
    Set targetSheet = m_book.Worksheets(sheetName)

    ' 対象シートが取得できた場合（元々存在していない場合もある）
    If Not targetSheet Is Nothing Then
        isExistsProperties = True
    Else
        isExistsProperties = False
    End If
  
    On Error GoTo 0

End Function

' =========================================================
' ▽設定情報シートを生成するメソッド
'
' 概要　　　：シートが既に存在する場合は何もしない。存在しない場合は挿入を行う。
' 引数　　　：
' 戻り値　　：挿入したシート
'
' =========================================================
Private Function createPropertiesSheet() As Worksheet

    If m_book Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        Set m_propertiesSheet = Nothing
        Set createPropertiesSheet = Nothing
        Exit Function
    End If

    If Not m_propertiesSheet Is Nothing Then
    
        Set createPropertiesSheet = m_propertiesSheet
        Exit Function
    End If

    ' コピーシート
    Dim copySheetName As String
    copySheetName = ConstantsBookProperties.BOOK_PROPERTIES_SHEET_NAME

    On Error Resume Next

    ' 対象シート
    Dim targetSheet     As Worksheet

    ' 対象シートを取得する
    Set targetSheet = m_book.Worksheets(copySheetName)

    ' 対象シートが取得できた場合（元々存在していない場合もある）
    If Not targetSheet Is Nothing Then
        ' 処理を中断
        Set m_propertiesSheet = targetSheet
        Set createPropertiesSheet = targetSheet
        Exit Function
    End If
        
    
    ExcelUtil.copySheetAppend SutWorkbook _
                            , copySheetName _
                            , m_book _
                            , copySheetName
                            
    Set targetSheet = m_book.Worksheets(copySheetName)

    ' ▽新シートにテーブル定義情報を埋め込む
    With targetSheet

        ' シートを可視にする
        .visible = xlSheetVisible

    End With
    
    On Error GoTo 0
    
    Set m_propertiesSheet = targetSheet
    Set createPropertiesSheet = targetSheet

End Function

' =========================================================
' ▽値を設定するメソッド
' =========================================================
Public Sub setValue(ByVal table As Long, ByVal name As String, ByVal value As String)

    If createPropertiesSheet Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        Exit Sub
    End If
    
    Dim findRow    As Long
    
    Dim currentRow As Long
    Dim currentVal As String
    
    currentRow = ConstantsBookProperties.FIRST_ROW
    
    Do While True
    
        currentVal = m_propertiesSheet.Cells(currentRow, table).value
        If currentVal = "" Then
            ' セルの値が空白になったら、行の終端なのでループを終了
            Exit Do
        End If
        
        If currentVal = name Then
            ' 現在のセルの値と名前が一致していたら行が見つかったのでループ終了
            findRow = currentRow
            Exit Do
        End If
        
        currentRow = currentRow + 1
        
    Loop
    
    If findRow = 0 Then
        ' 見つからなかったので名前と値を設定
        m_propertiesSheet.Cells(currentRow, table) = name
        m_propertiesSheet.Cells(currentRow, table + 1) = value
    Else
        ' 見つかったので値を設定
        m_propertiesSheet.Cells(findRow, table + 1) = value
    End If

End Sub

' =========================================================
' ▽値を削除するメソッド
' =========================================================
Public Sub removeValue(ByVal table As Long, ByVal name As String)

    If createPropertiesSheet Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        Exit Sub
    End If
    
    Dim findRow    As Long
    
    Dim currentRow As Long
    Dim currentVal As String
    
    currentRow = ConstantsBookProperties.FIRST_ROW
    
    Do While True
    
        currentVal = m_propertiesSheet.Cells(currentRow, table).value
        If currentVal = "" Then
            ' セルの値が空白になったら、行の終端なのでループを終了
            Exit Do
        End If
        
        If currentVal = name Then
            ' 現在のセルの値と名前が一致していたら行が見つかったのでループ終了
            findRow = currentRow
            Exit Do
        End If
        
        currentRow = currentRow + 1
        
    Loop
    
    If findRow = 0 Then
        ' 見つからないので何もしない
    Else
        ' 見つかったので削除
        With m_propertiesSheet
            .Range(.Cells(findRow, table), .Cells(findRow, table + 1)).delete xlShiftUp
        End With
        
    End If

End Sub

' =========================================================
' ▽値を削除するメソッド
' =========================================================
Public Sub removeValueByPrefixName(ByVal table As Long, ByVal prefixName As String)

    If createPropertiesSheet Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        Exit Sub
    End If
    
    Dim findRow    As Long
    
    Dim currentRow As Long
    Dim currentVal As String
    
    currentRow = ConstantsBookProperties.FIRST_ROW
    
    Do While True
    
        currentVal = m_propertiesSheet.Cells(currentRow, table).value
        If currentVal = "" Then
            ' セルの値が空白になったら、行の終端なのでループを終了
            Exit Do
        End If
        
        If InStr(currentVal, prefixName) = 1 Then
            ' 現在のセルの値と接頭辞名が一致していたら削除
            With m_propertiesSheet
                .Range(.Cells(currentRow, table), .Cells(currentRow, table + 1)).delete xlShiftUp
            End With
            
            ' 一つ削除されたのでカウンタを戻す
            currentRow = currentRow - 1
        End If
        
        currentRow = currentRow + 1
        
    Loop

End Sub

' =========================================================
' ▽値を全て削除するメソッド
' =========================================================
Public Sub removeAllValue(ByVal table As Long)

    If createPropertiesSheet Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        Exit Sub
    End If
    
    Dim currentRow As Long
    Dim currentVal As String
    
    currentRow = ConstantsBookProperties.FIRST_ROW
    
    Do While True
    
        currentVal = m_propertiesSheet.Cells(currentRow, table).value
        If currentVal = "" Then
            ' セルの値が空白になったら、行の終端なのでループを終了
            Exit Do
        End If
        
        currentRow = currentRow + 1
        
    Loop
    
    ' 削除
    With m_propertiesSheet
        .Range(.Cells(FIRST_ROW, table), .Cells(currentRow, table + 1)).delete xlShiftUp
    End With

End Sub

' =========================================================
' ▽値を取得するメソッド
' =========================================================
Public Function getValues(ByVal table As Long) As ValCollection

    If createPropertiesSheet Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        Exit Function
    End If
    
    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim currentRow As Long
    Dim currentVal As String
    
    currentRow = ConstantsBookProperties.FIRST_ROW
    
    Do While True
    
        currentVal = m_propertiesSheet.Cells(currentRow, table).value
        If currentVal = "" Then
            ' セルの値が空白になったら、行の終端なのでループを終了
            Exit Do
        End If
        
        ret.setItem m_propertiesSheet.Cells(currentRow, table + 1), currentVal
        
        currentRow = currentRow + 1
        
    Loop

    Set getValues = ret

End Function

' =========================================================
' ▽キーと値を取得するメソッド（要素が配列）
' =========================================================
Public Function getValuesOfElementArray(ByVal table As Long) As ValCollection

    If createPropertiesSheet Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        Exit Function
    End If
    
    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim currentRow As Long
    Dim currentVal As String
    
    currentRow = ConstantsBookProperties.FIRST_ROW
    
    Do While True
    
        currentVal = m_propertiesSheet.Cells(currentRow, table).value
        If currentVal = "" Then
            ' セルの値が空白になったら、行の終端なのでループを終了
            Exit Do
        End If
        
        ret.setItem Array(currentVal, m_propertiesSheet.Cells(currentRow, table + 1).value), currentVal
        
        currentRow = currentRow + 1
        
    Loop

    Set getValuesOfElementArray = ret

End Function

' =========================================================
' ▽値を取得するメソッド
' =========================================================
Public Function getValuesByPrefixName(ByVal table As Long, ByVal prefixName As String) As ValCollection

    If createPropertiesSheet Is Nothing Then
        ' ブックが未設定の場合は、設定情報が存在しない
        Exit Function
    End If
    
    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim currentRow As Long
    Dim currentVal As String
    
    currentRow = ConstantsBookProperties.FIRST_ROW
    
    Do While True
    
        currentVal = m_propertiesSheet.Cells(currentRow, table).value
        If currentVal = "" Then
            ' セルの値が空白になったら、行の終端なのでループを終了
            Exit Do
        End If
        
        If InStr(currentVal, prefixName) = 1 Then
            ' 接頭辞が一致
            ret.setItem m_propertiesSheet.Cells(currentRow, table + 1).value, currentVal
        End If
        
        currentRow = currentRow + 1
        
    Loop

    Set getValuesByPrefixName = ret

End Function
