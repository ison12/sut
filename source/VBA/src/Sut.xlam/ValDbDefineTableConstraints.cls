VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbDefineTableConstraints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBテーブル制約定義クラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public constraintName As String
Public constraintType As TABLE_CONSTANTS_TYPE
Public schemaName As String
Public tableName As String
Public columnList As ValCollection
Public fkColumnList As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    Set columnList = New ValCollection
    Set fkColumnList = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    Set columnList = Nothing
    Set fkColumnList = Nothing
End Sub

Public Sub addColumn(ByVal column As ValDbDefineColumn)
    columnList.setItem column, column.columnName
End Sub

Public Sub addFkColumn(ByVal column As ValDbDefineColumn)
    fkColumnList.setItem column, column.columnName
End Sub

