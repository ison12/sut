VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' Scripting.Dictionaryのラッパークラス
'
' 作成者　：Ison
' 履歴　　：2020/09/06　新規作成
'
' 特記事項：
' *********************************************************

' ディクショナリオブジェクト
Public dic  As Object

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set dic = CreateObject("Scripting.Dictionary")
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set dic = Nothing
    
End Sub

' =========================================================
' ▽存在判定
'
' 概要　　　：キーに紐づくオブジェクトが存在しているかを判定する。
' 引数　　　：key オブジェクト設定時に利用したキー
' 戻り値　　：存在している場合True
'
' =========================================================
Public Function exist(ByRef key As Variant) As Boolean

    exist = dic.Exists(key)
End Function

' =========================================================
' ▽オブジェクト設定
'
' 概要　　　：オブジェクトを設定する。
' 　　　　　　keyを指定するとgetItemでキーを指定してオブジェクトを取得することが可能。
' 引数　　　：item オブジェクト
' 　　　　　　 key キー
' 戻り値　　：無し
'
' =========================================================
Public Sub setItem(ByRef item As Variant, ByRef key As Variant)

    Dim isKeyFlg As Boolean

    ' 引数keyをString型に変換
    Dim keyStr As String: keyStr = CStr(key)

    ' keyに紐づくオブジェクトが存在しているかどうかを判定する
    If exist(keyStr) = True Then
    
        ' ▽アイテムを一旦削除する
        dic.remove (keyStr)
        ' ▽アイテムを追加
        dic.Add keyStr, item
    
    Else
    
        ' ▽アイテムを追加する
        dic.Add keyStr, item
    
    End If

End Sub

' =========================================================
' ▽オブジェクト取得
'
' 概要　　　：オブジェクトを取得する。
' 　　　　　　設定したオブジェクトがObject型かそうでないかを判別して戻り値に適切な型を返す。
'
' 引数　　　：key キー
' 　　　　　　   t オブジェクトの型
' 戻り値　　：無し
'
' =========================================================
Public Function getItem(ByRef key As Variant, Optional t As VbVarType = vbObject) As Variant
    
    ' 引数keyをString型に変換
    Dim keyStr As String
    
    keyStr = CStr(key)
    
    ' keyに紐づくオブジェクトが存在しているかどうかを判定する
    If exist(keyStr) = False Then
    
        If t = vbObject Then
            
            Set getItem = Nothing
            
        Else
            getItem = Empty
        
        End If
    
        Exit Function
    End If
    
    ' オブジェクト型
    If t = vbObject Then
    
        Set getItem = dic.item(keyStr)
    
    ' その他の型
    Else
    
        getItem = dic.item(keyStr)
    
    End If

End Function

' =========================================================
' ▽リストカウント取得
'
' 概要　　　：リストに設定されているオブジェクトの総個数を取得する。
' 引数　　　：無し
' 戻り値　　：総個数
'
' =========================================================
Public Function count() As Long

    count = dic.count

End Function

' =========================================================
' ▽オブジェクト削除
'
' 概要　　　：オブジェクトを削除する
' 引数　　　：key キー
' 戻り値　　：正常に削除した場合True
'
' =========================================================
Public Function remove(ByRef key As Variant) As Boolean
    
    ' 引数keyをString型に変換
    Dim keyStr As String: keyStr = CStr(key)
    
    If exist(keyStr) = True Then
    
        dic.remove keyStr
        
        remove = True
    Else
    
        remove = False
    End If
    
End Function

' =========================================================
' ▽オブジェクト全削除
'
' 概要　　　：オブジェクトを全て削除する
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub removeAll()

    dic.removeAll
    
End Sub

' =========================================================
' ▽配列変換
'
' 概要　　　：ディクショナリを配列に変換する。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Function toArray(Optional t As VbVarType = vbObject) As Variant

    ' 戻り値
    Dim ret As Variant
    
    If dic.count = 0 Then
    
        toArray = Empty
        Exit Function
    End If
    
    ' 配列を確保する（サイズは１～ディクショナリのサイズ）
    ReDim ret(1 To dic.count)
    
    Dim a As Variant
    Dim i As Long
    
    a = dic.items
    
    For i = LBound(a) To UBound(a)
    
        ' 型がオブジェクトの場合
        If t = vbObject Then
        
            Set ret(i + 1) = a(i)
            
        ' 上記以外
        Else
        
            ret(i + 1) = a(i)
        End If
    Next
    
    ' 戻り値を返す
    toArray = ret
    
End Function

' =========================================================
' ▽ディクショナリオブジェクトのコピー
'
' 概要　　　：
' 引数　　　：dic ディクショナリオブジェクト
' 戻り値　　：コピーしたオブジェクト
'
' =========================================================
Public Function copy() As ValDictionary

    Dim ret As New ValDictionary
    
    Dim key As Variant
    
    For Each key In dic
        ret.setItem dic.item(key), key
    Next
    
    Set copy = ret

End Function

