VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CntListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ListBoxコントロールのラップクラス
'
' リストボックスにはラベルしか追加できないという欠点がある。
' 本クラスでは、これを補うためにコレクションオブジェクトを
' 内部で保持しラベルとオブジェクトを１セットで追加・取得可
' 能なインターフェースを定義している。
'
' 作成者　：Ison
' 履歴　　：2008/09/23　新規作成
'
' 特記事項：
' *********************************************************

' リストボックス
Private m_control     As Object
' コレクションオブジェクト
Private m_collection  As ValCollection

' =========================================================
' ▼プロパティ リストコントロール取得
' =========================================================
Public Property Get control() As Variant

    Set control = m_control
    
End Property

' =========================================================
' ▼プロパティ リストコントロール設定
' =========================================================
Public Property Let control(ByVal vNewValue As Variant)

    Set m_control = vNewValue
    
End Property

' =========================================================
' ▼プロパティ コレクション取得
' =========================================================
Public Property Get collection() As ValCollection

    Set collection = m_collection
End Property

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set m_collection = Nothing
End Sub

' =========================================================
' ▽擬似コンストラクタ
' =========================================================
Public Sub init(ByRef var As Variant)

    ' リストコントロールを設定
    Set m_control = var
    ' コレクションを新たに初期化する
    Set m_collection = New ValCollection
    
End Sub

' =========================================================
' ▽要素追加
'
' 概要　　　：リストコントロールとコレクションに値を追加する。
'
' 引数　　　：label リストコントロールに表示するラベル
' 　　　　　　var   コレクションに追加する要素
' 戻り値　　：無し
' =========================================================
Public Sub addItem(ByVal label As String _
                 , ByVal var As Variant)

    ' リストコントロールにラベルを追加
    m_control.addItem label
    ' コレクションに要素を追加
    m_collection.setItem var

End Sub

' =========================================================
' ▽要素追加
'
' 概要　　　：リストコントロールとコレクションに値を追加する。
'
' 引数　　　：var   コレクションに追加する要素
'             labelProp1  ラベルプロパティ1
'             labelProp2  ラベルプロパティ2
'             labelProp3  ラベルプロパティ3
'             labelProp4  ラベルプロパティ4
'             labelProp5  ラベルプロパティ5
'             labelProp6  ラベルプロパティ6
'             labelProp7  ラベルプロパティ7
'             labelProp8  ラベルプロパティ8
'             labelProp9  ラベルプロパティ9
'             labelProp10 ラベルプロパティ10
' 戻り値　　：無し
' =========================================================
Public Sub addItemByProp(ByVal var As Variant _
                 , Optional ByVal labelProp1 As String = "" _
                 , Optional ByVal labelProp2 As String = "" _
                 , Optional ByVal labelProp3 As String = "" _
                 , Optional ByVal labelProp4 As String = "" _
                 , Optional ByVal labelProp5 As String = "" _
                 , Optional ByVal labelProp6 As String = "" _
                 , Optional ByVal labelProp7 As String = "" _
                 , Optional ByVal labelProp8 As String = "" _
                 , Optional ByVal labelProp9 As String = "" _
                 , Optional ByVal labelProp10 As String = "")

    Dim Label1  As String
    Dim Label2  As String
    Dim Label3  As String
    Dim Label4  As String
    Dim Label5  As String
    Dim Label6  As String
    Dim Label7  As String
    Dim Label8  As String
    Dim Label9  As String
    Dim Label10 As String

    ' リストコントロールにラベルを追加
    m_control.addItem

    If labelProp1 <> "" Then
        Label1 = CallByName(var, labelProp1, VbGet)
        m_control.list(m_control.ListCount - 1, 0) = Label1
    End If

    If labelProp2 <> "" Then
        Label2 = CallByName(var, labelProp2, VbGet)
        m_control.list(m_control.ListCount - 1, 1) = Label2
    End If

    If labelProp3 <> "" Then
        Label3 = CallByName(var, labelProp3, VbGet)
        m_control.list(m_control.ListCount - 1, 2) = Label3
    End If

    If labelProp4 <> "" Then
        Label4 = CallByName(var, labelProp4, VbGet)
        m_control.list(m_control.ListCount - 1, 3) = Label4
    End If

    If labelProp5 <> "" Then
        Label5 = CallByName(var, labelProp5, VbGet)
        m_control.list(m_control.ListCount - 1, 4) = Label5
    End If

    If labelProp6 <> "" Then
        Label6 = CallByName(var, labelProp6, VbGet)
        m_control.list(m_control.ListCount - 1, 5) = Label6
    End If

    If labelProp7 <> "" Then
        Label7 = CallByName(var, labelProp7, VbGet)
        m_control.list(m_control.ListCount - 1, 6) = Label7
    End If

    If labelProp8 <> "" Then
        Label8 = CallByName(var, labelProp8, VbGet)
        m_control.list(m_control.ListCount - 1, 7) = Label8
    End If

    If labelProp9 <> "" Then
        Label9 = CallByName(var, labelProp9, VbGet)
        m_control.list(m_control.ListCount - 1, 8) = Label9
    End If

    If labelProp10 <> "" Then
        Label10 = CallByName(var, labelProp10, VbGet)
        m_control.list(m_control.ListCount - 1, 9) = Label10
    End If

    ' コレクションに要素を追加
    m_collection.setItem var

End Sub

' =========================================================
' ▽要素設定
'
' 概要　　　：コレクションに要素を設定する。
'
' 引数　　　：index 添え字（0を最小値とする）
' 　　　　　　val   値
'             labelProp1 ラベルプロパティ1
'             labelProp2 ラベルプロパティ2
' 戻り値　　：無し
' =========================================================
Public Function setItem(ByVal index As Long _
                , ByVal val As Variant _
                 , Optional ByVal labelProp1 As String = "" _
                 , Optional ByVal labelProp2 As String = "" _
                 , Optional ByVal labelProp3 As String = "" _
                 , Optional ByVal labelProp4 As String = "" _
                 , Optional ByVal labelProp5 As String = "" _
                 , Optional ByVal labelProp6 As String = "" _
                 , Optional ByVal labelProp7 As String = "" _
                 , Optional ByVal labelProp8 As String = "" _
                 , Optional ByVal labelProp9 As String = "" _
                 , Optional ByVal labelProp10 As String = "") As Variant

    Dim item As Variant
    
    Dim Label1 As String
    Dim Label2 As String
    Dim Label3  As String
    Dim Label4  As String
    Dim Label5  As String
    Dim Label6  As String
    Dim Label7  As String
    Dim Label8  As String
    Dim Label9  As String
    Dim Label10 As String

    ' インデックスオーバーフローチェック
    If m_collection.count <= index Then
        Exit Function
    End If

    m_collection.removeByIndex index + 1
    If m_collection.count >= index + 1 Then
        m_collection.setItemByIndexBefore val, index + 1
    Else
        m_collection.setItem val
    End If
    
    If labelProp1 <> "" Then
        Label1 = CallByName(val, labelProp1, VbGet)
        m_control.list(index, 0) = Label1
    End If

    If labelProp2 <> "" Then
        Label2 = CallByName(val, labelProp2, VbGet)
        m_control.list(index, 1) = Label2
    End If

    If labelProp3 <> "" Then
        Label3 = CallByName(val, labelProp3, VbGet)
        m_control.list(index, 2) = Label3
    End If

    If labelProp4 <> "" Then
        Label4 = CallByName(val, labelProp4, VbGet)
        m_control.list(index, 3) = Label4
    End If

    If labelProp5 <> "" Then
        Label5 = CallByName(val, labelProp5, VbGet)
        m_control.list(index, 4) = Label5
    End If

    If labelProp6 <> "" Then
        Label6 = CallByName(val, labelProp6, VbGet)
        m_control.list(index, 5) = Label6
    End If

    If labelProp7 <> "" Then
        Label7 = CallByName(val, labelProp7, VbGet)
        m_control.list(index, 6) = Label7
    End If

    If labelProp8 <> "" Then
        Label8 = CallByName(val, labelProp8, VbGet)
        m_control.list(index, 7) = Label8
    End If

    If labelProp9 <> "" Then
        Label9 = CallByName(val, labelProp9, VbGet)
        m_control.list(index, 8) = Label9
    End If

    If labelProp10 <> "" Then
        Label10 = CallByName(val, labelProp10, VbGet)
        m_control.list(index, 9) = Label10
    End If

End Function

' =========================================================
' ▽要素取得
'
' 概要　　　：コレクションから要素を取得する。
'
' 引数　　　：index 添え字（0を最小値とする）
' 　　　　　　t     コレクションから取得する要素の型
' 　　　　　　　　　省略した場合はObject型。
' 戻り値　　：要素
' =========================================================
Public Function getItem(ByVal index As Long, Optional ByVal t As VbVarType = vbObject) As Variant

    ' 型がオブジェクトの場合
    If t = vbObject Then
    
        ' Setステートメントを使用して戻り値を取得する
        ' （Collectionオブジェクトの添え字は1が最小値となるのでindexに+1している）
        Set getItem = m_collection.getItemByIndex(index + 1, t)
        
    ' 型がオブジェクト以外
    Else
        
        ' Getステートメントを使用して戻り値を取得する
        ' （Collectionオブジェクトの添え字は1が最小値となるのでindexに+1している）
        getItem = m_collection.getItemByIndex(index + 1, t)
    End If

End Function

Public Sub swapItem(ByVal index1 As Long, _
                    ByVal index2 As Long, _
                    Optional ByVal t As VbVarType = vbObject, _
                    Optional ByVal colCount As Long = 1)

    Dim var1 As Variant
    Dim var2 As Variant
    
    Dim indexMin As Long
    Dim indexMax As Long
    
    If index1 > index2 Then
    
        indexMin = index2
        indexMax = index1
    Else
        
        indexMin = index1
        indexMax = index2
    End If
    
    ' ----------------------------------------
    ' Collectionオブジェクトに要素を設定する
    
    ' 型がObject型であるか、そうでないかを比較し適切な方法で
    ' 値を取得する
    If t = vbObject Then
    
        Set var1 = m_collection.getItemByIndex(indexMax + 1)
        Set var2 = m_collection.getItemByIndex(indexMin + 1)
        
    Else
        
        var1 = m_collection.getItemByIndex(indexMax + 1)
        var2 = m_collection.getItemByIndex(indexMin + 1)
        
    End If

    m_collection.removeByIndex indexMax + 1
    m_collection.setItemByIndexAfter var2, indexMax
    
    m_collection.removeByIndex indexMin + 1
    m_collection.setItemByIndexBefore var1, indexMin + 1

    ' ----------------------------------------
    ' コントロールのラベルを入れ替える
    If colCount = 1 Then
        var1 = m_control.list(indexMax)
        var2 = m_control.list(indexMin)
        m_control.list(indexMax) = var2
        m_control.list(indexMin) = var1
        
    ElseIf colCount = 2 Then
    
        var1 = m_control.list(indexMax, 0)
        var2 = m_control.list(indexMin, 0)
        
        If isNull(var2) Then
            m_control.list(indexMax, 0) = ""
        Else
            m_control.list(indexMax, 0) = var2
        End If
        If isNull(var1) Then
            m_control.list(indexMin, 0) = ""
        Else
            m_control.list(indexMin, 0) = var1
        End If
    
        var1 = m_control.list(indexMax, 1)
        var2 = m_control.list(indexMin, 1)
        If isNull(var2) Then
            m_control.list(indexMax, 1) = ""
        Else
            m_control.list(indexMax, 1) = var2
        End If
        If isNull(var1) Then
            m_control.list(indexMin, 1) = ""
        Else
            m_control.list(indexMin, 1) = var1
        End If
    
    End If


End Sub

' =========================================================
' ▽要素追加
'
' 概要　　　：要素をまとめて追加する。
'
' 引数　　　：col       コレクション
' 　　　　　　labelProp1 ラベルとして利用する【col】に格納されている要素のプロパティ
' 　　　　　　labelProp2 〃
' 　　　　　　labelProp3 〃
' 　　　　　　labelProp4 〃
' 　　　　　　labelProp5 〃
' 　　　　　　labelProp6 〃
' 　　　　　　labelProp7 〃
' 　　　　　　labelProp8 〃
' 　　　　　　labelProp9 〃
' 　　　　　　labelProp10 〃
'
' 戻り値　　：無し
' =========================================================
Public Sub addAll(ByVal col As ValCollection _
                , Optional ByVal labelProp1 As String = "" _
                , Optional ByVal labelProp2 As String = "" _
                , Optional ByVal labelProp3 As String = "" _
                , Optional ByVal labelProp4 As String = "" _
                , Optional ByVal labelProp5 As String = "" _
                , Optional ByVal labelProp6 As String = "" _
                , Optional ByVal labelProp7 As String = "" _
                , Optional ByVal labelProp8 As String = "" _
                , Optional ByVal labelProp9 As String = "" _
                , Optional ByVal labelProp10 As String = "" _
                , Optional ByVal isAppend As Boolean = False)

    Dim var As Variant
    
    Dim i As Long
    
    Dim Label1 As String
    Dim Label2 As String
    
    
    ' リストコントロールから要素を全て削除する
    If Not isAppend Then
        m_control.clear
        m_collection.removeAll
        i = 0
    Else
        i = m_control.ListCount
    End If
    
    For Each var In col.col
    
        If labelProp1 <> "" Or _
           labelProp2 <> "" Or _
           labelProp3 <> "" Or _
           labelProp4 <> "" Or _
           labelProp5 <> "" Or _
           labelProp6 <> "" Or _
           labelProp7 <> "" Or _
           labelProp8 <> "" Or _
           labelProp9 <> "" Or _
           labelProp10 <> "" _
        Then
            addItemByProp var, labelProp1, labelProp2, labelProp3, labelProp4, labelProp5, labelProp6, labelProp7, labelProp8, labelProp9, labelProp10
        Else
            addItem var, var
        End If
        
        i = i + 1
    Next
    
End Sub

' =========================================================
' ▽要素追加
'
' 概要　　　：ネストされたプロパティからさらにプロパティを取得しリストに追加する。
'
' 引数　　　：col       コレクション
'             nestedProp ネストされたプロパティ
' 　　　　　　labelProp1 ラベルとして利用する【col】に格納されている要素のプロパティ
' 　　　　　　labelProp2 ラベルとして利用する【col】に格納されている要素のプロパティ
'
' 戻り値　　：無し
' =========================================================
Public Sub addNestedProperty(ByVal col As ValCollection _
                            , ByVal nestedProp As String _
                            , Optional ByVal labelProp1 As String = "" _
                            , Optional ByVal labelProp2 As String = "")

    Dim var  As Variant
    Dim var2 As Variant
    
    Dim i As Long: i = 0
    
    Dim Label1 As String
    Dim Label2 As String
    
    
    ' リストコントロールから要素を全て削除する
    m_control.clear
    
    ' リストコントロールに要素を追加する
    ' ラベルプロパティ1・2が設定されている場合
    If labelProp1 <> "" And labelProp2 <> "" Then
    
        For Each var In col.col
        
            ' ネストされたプロパティオブジェクト
            Set var2 = CallByName(var, nestedProp, VbGet)
            
            If Not var2 Is Nothing Then
            
                Label1 = CallByName(var2, labelProp1, VbGet)
                Label2 = CallByName(var2, labelProp2, VbGet)
                
                m_control.addItem
                m_control.list(i, 0) = Label1
                m_control.list(i, 1) = Label2
                
            Else
            
                m_control.addItem
                m_control.list(i, 0) = ""
                m_control.list(i, 1) = ""
            End If
            
            i = i + 1
        Next
    
    ' ラベルプロパティ1が設定されている場合
    ElseIf labelProp1 <> "" Then
    
        For Each var In col.col
        
            ' ネストされたプロパティオブジェクト
            Set var2 = CallByName(var, nestedProp, VbGet)
            
            If Not var2 Is Nothing Then
            
                Label1 = CallByName(var2, labelProp1, VbGet)
            
                m_control.addItem
                m_control.list(i, 0) = Label1
                
                
            Else
            
                m_control.addItem
                m_control.list(i, 0) = ""
                m_control.list(i, 1) = ""
            End If
            
            i = i + 1
        Next
    
    ' ラベルプロパティが設定されていない場合、colから取得した要素自身をリストに追加する
    Else
    
        For Each var In col.col
        
            m_control.addItem var
        Next
    
    End If
    
    Set m_collection = col
    
End Sub

' =========================================================
' ▽要素削除
'
' 概要　　　：要素を削除する。
'
' 引数　　　：index 添え字（0を最小値とする）
'
' 戻り値　　：無し
' =========================================================
Public Sub removeItem(ByVal index As Long)

    ' リストコントロールから要素を削除
    m_control.removeItem index
    ' コレクションから要素を削除（コレクションの添え字の最小値は1なので+1する）
    m_collection.removeByIndex index + 1

End Sub

' =========================================================
' ▽要素削除
'
' 概要　　　：要素をまとめて削除する。
'
' 引数　　　：無し
' 戻り値　　：無し
' =========================================================
Public Sub removeAll()

    Dim i As Long
    
    ' リストコントロールから要素を削除
    If Not m_control Is Nothing Then
        m_control.clear
    End If
    
    Set m_collection = New ValCollection
    
End Sub

' =========================================================
' ▽選択済み要素取得
'
' 概要　　　：リストコントロールにて選択されている要素を取得する
'
' 引数　　　：無し
' 戻り値　　：無し
' =========================================================
Public Function getSelectedList(Optional ByVal t As VbVarType = vbObject) As ValCollection

    ' 戻り値
    Dim result As ValCollection

    ' 戻り値を初期化する
    Set result = New ValCollection

    Dim i As Long
    
    ' リストコントロールをループさせる
    For i = 0 To m_control.ListCount - 1
    
        ' リストコントロールにて選択されているかを判定する
        If m_control.selected(i) = True Then
        
            ' 戻り値に要素を設定する
            result.setItem m_collection.getItemByIndex(i + 1, t)
        
        End If
        
    Next
    
    ' 戻り値を設定する
    Set getSelectedList = result

End Function

' =========================================================
' ▽選択済みインデックス取得。
'
' 概要　　　：選択済みのインデックスを取得する。
'
' 引数　　　：無し
' 戻り値　　：選択済みインデックス
' =========================================================
Public Function getSelectedIndex() As Long

    getSelectedIndex = m_control.ListIndex

End Function

' =========================================================
' ▽選択済み項目取得。
'
' 概要　　　：選択済みの項目を取得する。
'
' 引数　　　：無し
' 戻り値　　：選択済み項目
' =========================================================
Public Function getSelectedItem(Optional ByVal t As VbVarType = vbObject) As Variant

    Dim selectedIndex As Long
    selectedIndex = getSelectedIndex
    
    If t = vbObject Then
        Set getSelectedItem = getItem(selectedIndex, t)
    Else
        getSelectedItem = getItem(selectedIndex, t)
    End If

End Function

' =========================================================
' ▽任意の選択肢の選択。
'
' 概要　　　：任意の選択肢を選択する。
'
' 引数　　　：index 任意の選択肢のインデックス
' 戻り値　　：無し
' =========================================================
Public Sub setSelectedIndex(ByVal index As Long)

    If m_control.ListCount = 0 Or m_control.ListCount <= index Then
        Exit Sub
    End If

    m_control.ListIndex = index

End Sub

' =========================================================
' ▽全ての選択肢の選択状態を変更。
'
' 概要　　　：
'
' 引数　　　：isSelected 選択状態（true 選択する、false 選択しない）
' 戻り値　　：無し
' =========================================================
Public Sub setSelectedAll(ByVal isSelected As Boolean)

    Dim i As Long
    
    For i = 0 To m_control.ListCount - 1
        m_control.selected(i) = isSelected
    Next

End Sub

' =========================================================
' ▽任意の項目が存在するかを判定する。
'
' 概要　　　：
'
' 引数　　　：value 値
' 戻り値　　：True 項目が存在する、False 項目が存在しない
' =========================================================
Public Function exist(ByVal value As String) As Boolean

    Dim i As Long

    ' リストコントロールをループさせる
    For i = 0 To m_control.ListCount - 1
    
        ' リストコントロールにて選択されているかを判定する
        If m_control.list(i) = value Then
        
            exist = True
            Exit Function
        
        End If
        
    Next

    exist = False

End Function

' =========================================================
' ▽件数取得。
'
' 概要　　　：件数を取得する。
'
' 引数　　　：無し
' 戻り値　　：件数
' =========================================================
Public Function count() As Long

    count = m_control.ListCount

End Function
