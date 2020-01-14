VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' VB6.0 Collectionのラッパークラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/06/28　新規作成
'
' 特記事項：
' *********************************************************

' コレクションオブジェクト
Public col  As collection

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
' ▽存在判定
'
' 概要　　　：キーに紐づくオブジェクトが存在しているかを判定する。
' 引数　　　：key オブジェクト設定時に利用したキー
' 戻り値　　：存在している場合True
'
' =========================================================
Public Function exist(ByRef key As Variant) As Boolean

    On Error Resume Next

    Dim temp As Variant
    
    col.item key
    
    If err <> 0 Then
    
        exist = False
        
    Else
    
        exist = True
    End If

    On Error GoTo 0

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
Public Sub setItem(ByRef item As Variant, Optional ByRef key As Variant = Empty)

    Dim isKeyFlg As Boolean
    
    ' キーの指定がない場合、そのまま追加する
    If key = Empty Then
    
        col.Add item
    
        Exit Sub
        
    End If

    ' 引数keyをString型に変換
    Dim keyStr As String: keyStr = CStr(key)

    ' keyに紐づくオブジェクトが存在しているかどうかを判定する
    If exist(keyStr) = True Then
    
        ' ▽アイテムを一旦削除する
        col.remove (keyStr)
        ' ▽アイテムを追加
        col.Add item, keyStr
    
    Else
    
        ' ▽アイテムを追加する
        col.Add item, keyStr
    
    End If

End Sub

' =========================================================
' ▽オブジェクト設定（インデックス指定）
'
' 概要　　　：指定したインデックスの前方にオブジェクトを設定する。
' 引数　　　：item オブジェクト
' 　　　　　　index インデックス
' 戻り値　　：True 設定に成功
'
' =========================================================
Public Function setItemByIndexBefore(ByRef item As Variant, ByVal index As Long) As Boolean

    On Error Resume Next
    
    col.Add item:=item, before:=index

    If err.Number <> 0 Then
    
        setItemByIndexBefore = False
    Else
    
        setItemByIndexBefore = True
    End If

    On Error GoTo 0
    
End Function

' =========================================================
' ▽オブジェクト設定（インデックス指定）
'
' 概要　　　：指定したインデックスの後方にオブジェクトを設定する。
' 引数　　　：item オブジェクト
' 　　　　　　index インデックス
' 戻り値　　：True 設定に成功
'
' =========================================================
Public Function setItemByIndexAfter(ByRef item As Variant, ByVal index As Long) As Boolean

    On Error Resume Next
    
    col.Add item:=item, after:=index

    If err.Number <> 0 Then
    
        setItemByIndexAfter = False
    Else
    
        setItemByIndexAfter = True
    End If

    On Error GoTo 0
    
End Function

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
    
        Set getItem = col.item(keyStr)
    
    ' その他の型
    Else
    
        getItem = col.item(keyStr)
    
    End If

End Function

' =========================================================
' ▽オブジェクト取得
'
' 概要　　　：オブジェクトを取得する。
' 　　　　　　インデックス番号indexには[1～count]までの数値を設定する。
'
' 　　　　　　引数 t を基に適切な戻り値を返す。
'
' 引数　　　：index インデックス
' 　　　　　　    t オブジェクトの型
' 戻り値　　：無し
'
' =========================================================
Public Function getItemByIndex(ByVal index As Long, Optional t As VbVarType = vbObject) As Variant
    
    ' keyに紐づくオブジェクトが存在しているかどうかを判定する
    If exist(index) = False Then
    
        If t = vbObject Then
            
            Set getItemByIndex = Nothing
            
        Else
            getItemByIndex = Empty
        
        End If
    
        Exit Function
    End If
    
    ' オブジェクト型
    If t = vbObject Then
    
        Set getItemByIndex = col.item(index)
    
    ' その他の型
    Else
    
        getItemByIndex = col.item(index)
    
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

    count = col.count

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
    
        col.remove keyStr
        
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
    
    Do While col.count > 0
    
        removeByIndex 1
    Loop
    
End Sub

' =========================================================
' ▽オブジェクト削除
'
' 概要　　　：オブジェクトを削除する
' 　　　　　　インデックス番号indexには[1～count]までの数値を設定する。
' 引数　　　：index インデックス
' 戻り値　　：正常に削除した場合True
'
' =========================================================
Public Function removeByIndex(ByVal index As Long) As Boolean
    
    If exist(index) = True Then
    
        col.remove index
        
        removeByIndex = True
    Else
    
        removeByIndex = False
    End If
    
End Function

' =========================================================
' ▽配列変換
'
' 概要　　　：コレクションを配列に変換する。
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Function toArray(Optional t As VbVarType = vbObject) As Variant

    ' 戻り値
    Dim ret As Variant
    
    If col.count = 0 Then
    
        toArray = Empty
        Exit Function
    End If
    
    ' 配列を確保する（サイズは１～コレクションのサイズ）
    ReDim ret(1 To col.count)
    
    Dim i As Long
    
    For i = LBound(ret) To UBound(ret)
    
        ' 型がオブジェクトの場合
        If t = vbObject Then
        
            Set ret(i) = getItemByIndex(i, vbObject)
            
        ' 上記以外
        Else
        
            ret(i) = getItemByIndex(i, vbVariant)
        End If
    Next
    
    ' 戻り値を返す
    toArray = ret
    
End Function
