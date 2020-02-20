VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeDataTypeReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' 任意のシートからデータ種類の定義を読み込むクラス
'
' 作成者　：Ison
' 履歴　　：2009/08/16　新規作成
'
' 特記事項：
' *********************************************************

' ワークシート
Public sheet         As Worksheet
' DB名
Public dbName        As String

' =========================================================
' ▽定数プロパティ
' 　Classモジュールでは、Public Constが定義できないのでプロパティで代用する
' =========================================================
' セル行番号 DB名
Public Property Get CELL_ROW_DBNAME() As Long

    CELL_ROW_DBNAME = 1
End Property

' セル行番号 オフセット
Public Property Get CELL_ROW_OFFSET() As Long

    CELL_ROW_OFFSET = 3
End Property

' セル列番号 リテラル
Public Property Get CELL_COL_LITERAL() As Long

    CELL_COL_LITERAL = 2
End Property

' セル列番号 データ種類
Public Property Get CELL_COL_DATA_TYPE() As Long

    CELL_COL_DATA_TYPE = 3
End Property

' セル列番号 更新時の書式
Public Property Get CELL_COL_FORMAT_UPDATE() As Long

    CELL_COL_FORMAT_UPDATE = 4
End Property

' セル列番号 参照時の書式
Public Property Get CELL_COL_FORMAT_SELECT() As Long

    CELL_COL_FORMAT_SELECT = 5
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
' ▽DB名から列番号を検索する。
'
' 概要　　　：
' 引数　　　：sheet    シート名
'     　　　：dbName   DB名
' 戻り値　　：列番号
'
' =========================================================
Private Function findColumnDbName(ByVal sheet As Worksheet, ByVal dbName As String) As Long

    findColumnDbName = 0

    ' Findで検索する
    Dim findRange As Range
    Set findRange = sheet.Range(CELL_ROW_DBNAME & ":" & CELL_ROW_DBNAME).Find(what:=dbName, MatchCase:=False, MatchByte:=True, SearchFormat:=False)
    
    If Not findRange Is Nothing Then
        ' 見つかった場合
        findColumnDbName = findRange.column
        Exit Function
    End If
    
End Function

' =========================================================
' ▽任意のシートからデータ種類の定義を読み取る
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Function execute() As Variant

    ' カラム情報リスト
    Dim colList As ValCollection
    ' カラム情報
    Dim colType As ValDbColumnTypeColInfo

    ' カラム情報リストのサイズを確保する
    Set colList = New ValCollection
    
    ' 基準位置となる列番号を検索する
    Dim offsetColumn As Long
    offsetColumn = findColumnDbName(sheet, dbName)
    
    If offsetColumn <= 0 Then
        ' 見つからない場合
        Set execute = colList
        Exit Function
    End If
    
    ' 定義情報のサイズ
    Dim size As Long
    ' 定義情報のサイズを取得する
    size = ExcelUtil.getCellEndOfLastInputtedRow(sheet, offsetColumn - 2 + CELL_COL_LITERAL) - CELL_ROW_OFFSET + 1
    
    Dim i As Long
    
    For i = 0 To size - 1
    
        ' カラム情報を生成する
        Set colType = New ValDbColumnTypeColInfo
    
        ' データ種類を取得する（常に大文字でカラム情報を検索するため、大文字に変換する）
        colType.columnName = UCase$(sheet.Cells(i + CELL_ROW_OFFSET, offsetColumn - 2 + CELL_COL_DATA_TYPE))
        ' 更新時の書式情報を取得する
        colType.formatUpdate = sheet.Cells(i + CELL_ROW_OFFSET, offsetColumn - 2 + CELL_COL_FORMAT_UPDATE)
        ' 参照時の書式情報を取得する
        colType.formatSelect = sheet.Cells(i + CELL_ROW_OFFSET, offsetColumn - 2 + CELL_COL_FORMAT_SELECT)
        
        colList.setItem colType, colType.columnName
    Next
    
    Set execute = colList
    
End Function
