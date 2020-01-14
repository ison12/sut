VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbColumnFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのカラム書式情報を表現するクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/08/01　新規作成
'
' 特記事項：
' *********************************************************

' DB名
Public dbName    As String
' カラムリスト
Public columnList  As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    ' カラムリストを初期化する
    Set columnList = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set columnList = Nothing
    
End Sub

Public Function convertAliasDataType() As sutredlib.DbColumnTypeDbInfo

    ' 戻り値
    Dim ret As New sutredlib.DbColumnTypeDbInfo
    
    ' DB名を設定する
    ret.dbName = dbName
    
    If columnList.count > 0 Then
    
        ' カラムリストを生成する
        Dim columnFmtList() As sutredlib.DbColumnTypeColInfo
        ' カラムリストのサイズを初期化する
        ReDim columnFmtList(0 To columnList.count - 1)
        
        ' コレクションから取得した情報を格納する一時変数
        Dim var As sutredlib.DbColumnTypeColInfo
        
        Dim i As Long
        
        For Each var In columnList.col
        
            Set columnFmtList(i) = New sutredlib.DbColumnTypeColInfo
            
            columnFmtList(i).columnName = var.columnName
            columnFmtList(i).formatUpdate = var.formatUpdate
            columnFmtList(i).formatSelect = var.formatSelect
        
            i = i + 1
        Next
        
    End If
    
    ret.columnList = columnFmtList
    
    Set convertAliasDataType = ret
    
End Function

Public Function setAliasDataType(ByVal info As sutredlib.DbColumnTypeDbInfo)

    ' DB名を設定する
    dbName = info.dbName
    
    ' 引数 info からカラムリストを取得し、本クラス（this）に入れ替えを行う
    Dim colList() As sutredlib.DbColumnTypeColInfo
    colList = info.columnList
    
    Dim i As Long
    
    For i = LBound(colList) To UBound(colList)
    
        columnList.setItem colList(i), colList(i).columnName
    
    Next

End Function
