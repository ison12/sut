VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryResultDetailInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBクエリ結果詳細情報クラス
'
' 作成者　：Ison
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public cell As String
Public message As String
Public query As String
Public sqlState As String

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

Public Property Get messageWithSqlState() As String

    If sqlState = "" Then
        messageWithSqlState = message
    Else
        messageWithSqlState = sqlState & " - " & message
    End If
    
    messageWithSqlState = replace(messageWithSqlState, vbNewLine, " ")
    messageWithSqlState = replace(messageWithSqlState, vbCr, " ")
    messageWithSqlState = replace(messageWithSqlState, vbLf, " ")
    
End Property

Public Property Get queryWithoutNewLine() As String

    queryWithoutNewLine = query
    queryWithoutNewLine = replace(queryWithoutNewLine, vbNewLine, " ")
    queryWithoutNewLine = replace(queryWithoutNewLine, vbCr, " ")
    queryWithoutNewLine = replace(queryWithoutNewLine, vbLf, " ")
    
End Property

' =========================================================
' ▽タブ区切りでのヘッダ情報の取得
' =========================================================
Public Property Get tabbedInfoHeader() As String

    tabbedInfoHeader = _
        """" & "Cell" & """" & vbTab & _
        """" & "Message" & """" & vbTab & _
        """" & "Query" & """"

End Property

' =========================================================
' ▽タブ区切りでの情報の取得
' =========================================================
Public Property Get tabbedInfo() As String

    tabbedInfo = _
        """" & replace(cell, """", """""") & """" & vbTab & _
        """" & replace(messageWithSqlState, """", """""") & """" & vbTab & _
        """" & replace(query, """", """""") & """"

End Property

