VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリパラメータを表現するクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/08　新規作成
'
' 特記事項：
' *********************************************************

' 名前
Private m_name    As String
' 値
Private m_value As String

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
' ▽プロパティ
' =========================================================
Public Property Get name() As String

    name = m_name
End Property

Public Property Let name(ByVal vNewValue As String)

    m_name = vNewValue
End Property

Public Property Get value() As String

    value = m_value
End Property

Public Property Let value(ByVal vNewValue As String)

    m_value = vNewValue
End Property
