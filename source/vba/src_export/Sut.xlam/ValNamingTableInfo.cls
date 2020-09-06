VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValNamingTableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' 名称を表現するクラス (テーブル情報)
'
' 作成者　：Ison
' 履歴　　：2008/06/29　新規作成
'
' 特記事項：
' *********************************************************

' テーブル論理名
Private m_tableNameLog  As String
' テーブル物理名
Private m_tableName     As String
' カラムリスト
Private m_columnList As ValCollection

' =========================================================
' ▽プロパティ宣言
Public Property Get columnList() As Variant

    Set columnList = m_columnList
End Property

Public Property Let columnList(ByVal vNewValue As Variant)

    Set m_columnList = vNewValue
End Property

Public Property Get tableNameLog() As Variant

    tableNameLog = m_tableNameLog
End Property

Public Property Let tableNameLog(ByVal vNewValue As Variant)

    m_tableNameLog = vNewValue
End Property

Public Property Get tableName() As Variant

    tableName = m_tableName
End Property

Public Property Let tableName(ByVal vNewValue As Variant)

    m_tableName = vNewValue
End Property
' △プロパティ宣言
' =========================================================


' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set m_columnList = New ValCollection

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set m_columnList = Nothing
End Sub

' =========================================================
' ▽クラス初期化メソッド (引数付き)
' =========================================================
Public Sub initial(ByVal tableName As String, ByVal tableNameLog As String)

    Me.tableName = tableName
    Me.tableNameLog = tableNameLog
End Sub

' =========================================================
' ▽カラム情報を追加
' =========================================================
Public Sub addColumnInfo(ByVal columnName As String, ByVal columnNameLog As String)

    ' カラム情報
    Dim columnInfo As New ValNamingColumnInfo
    
    ' カラム情報を初期化
    columnInfo.initial columnName, columnNameLog, Me
    
    ' カラムリストにカラム情報を追加
    m_columnList.setItem columnInfo, columnName
    
End Sub

' =========================================================
' ▽カラム情報を取得
' =========================================================
Public Function getColumnInfo(ByVal columnName As String) As ValNamingColumnInfo

    ' 戻り値を初期化
    Set getColumnInfo = Nothing
    
    ' カラムリストから情報を取得
    If m_columnList.exist(columnName) = True Then
    
        ' キーが存在している場合に情報を取得する
        Set getColumnInfo = m_columnList.getItem(columnName)
    
    End If
    
End Function

