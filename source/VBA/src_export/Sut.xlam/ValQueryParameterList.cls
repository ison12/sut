VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryParameterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリパラメータリストを表現するクラス
'
' 作成者　：Ison
' 履歴　　：2019/12/08　新規作成
'
' 特記事項：
' *********************************************************

' リスト
Private m_list    As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    Set m_list = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    Set m_list = Nothing
End Sub

' =========================================================
' ▽プロパティ
' =========================================================
Public Property Get list() As ValCollection

    Set list = m_list
End Property

Public Sub setListFromFlatRecords(ByVal flatRecords As ValCollection)

    Dim var As Variant
    Dim index As String
    
    Dim key As String
    Dim keyArray As Variant
    Dim val As String
    Dim tmp As ValQueryParameter

    For Each var In flatRecords.col
    
        key = var(0)
        val = var(1)
        
        keyArray = Split(key, "_")
        index = keyArray(UBound(keyArray))
        
        If Not m_list.exist(index) Then
        
            Set tmp = New ValQueryParameter
            m_list.setItem tmp, index
        Else
            Set tmp = m_list.getItem(index)
        End If
        
        If InStr(key, "name_") <> 0 Then
            tmp.name = val
        ElseIf InStr(key, "value_") <> 0 Then
            tmp.value = val
        End If
        
    Next

End Sub
