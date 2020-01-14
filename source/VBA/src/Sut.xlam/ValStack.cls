VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スタック
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/04/17　新規作成
'
' 特記事項：
' *********************************************************

' コレクションオブジェクト
Private col As collection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set col = New collection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set col = Nothing
End Sub

' =========================================================
' ▽データ追加
'
' 概要　　　：スタックにデータを追加する。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub push(ByVal value As Variant)

    col.Add value
End Sub

' =========================================================
' ▽データ取得
'
' 概要　　　：スタックからデータを取得する。
' 引数　　　：t データ型
' 戻り値　　：
'
' =========================================================
Public Function pop(Optional t As VbVarType = vbVariant) As Variant

    ' サイズが０の場合
    If col.count = 0 Then
    
        ' 空で返す
        pop = Empty
        
    ' サイズが１以上
    Else
    
        ' データ型がオブジェクト
        If t = vbObject Then
        
            ' Setを使って戻り値に設定
            Set pop = col.item(col.count)
            
        ' データ型がオブジェクト以外
        Else
        
            ' 戻り値に設定
            pop = col.item(col.count)
        End If
        
        ' 最後尾の情報を削除する
        col.remove col.count
    
    End If
    
End Function

' =========================================================
' ▽サイズ取得
'
' 概要　　　：スタックのサイズを取得する。
' 引数　　　：
' 戻り値　　：スタックサイズ
'
' =========================================================
Public Function count() As Long

    count = col.count

End Function
