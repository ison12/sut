VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValNamingColumnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' 名称を表現するクラス (カラム情報)
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/06/29　新規作成
'
' 特記事項：
' *********************************************************

' 親オブジェクト
Private m_parent As ValNamingTableInfo

' カラム論理名
Private m_columnNameLog  As String
' カラム物理名
Private m_columnName     As String

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
' ▽クラス初期化メソッド (引数付き)
' =========================================================
Public Sub initial(ByVal columnName As String, ByVal columnNameLog As String, ByRef parent As ValNamingTableInfo)

    Me.columnName = columnName
    Me.columnNameLog = columnNameLog
    Me.parent = parent
    
End Sub

Public Property Get parent() As Variant

    Set parent = m_parent
End Property

Public Property Let parent(ByVal vNewValue As Variant)

    Set m_parent = vNewValue
End Property

Public Property Get columnNameLog() As Variant

    columnNameLog = m_columnNameLog
End Property

Public Property Let columnNameLog(ByVal vNewValue As Variant)

    m_columnNameLog = vNewValue
End Property

Public Property Get columnName() As Variant

    columnName = m_columnName
End Property

Public Property Let columnName(ByVal vNewValue As Variant)

    m_columnName = vNewValue
End Property
