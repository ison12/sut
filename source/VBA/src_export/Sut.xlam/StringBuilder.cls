VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' VBA用のStringBuilderクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

' 容量
Private m_capacity As Long
' 文字列の長さ
Private m_length As Long
' 文字列の本体
Private m_buffer As String

Private Sub Class_Initialize()
    m_capacity = 1023
    Me.clear
End Sub

Private Sub Class_Terminate()
    m_buffer = vbNullString
End Sub

Public Property Let capacity(ByVal newValue As Long)

    '現在の容量より大きな容量を確保したい場合
    If newValue > m_capacity Then
        ' 文字列バッファを拡張する
        m_buffer = m_buffer & String$(newValue - m_capacity, vbNullChar)
        ' 容量を置き換える
        m_capacity = newValue
    End If
    
End Property

Public Property Get capacity() As Long
    capacity = m_capacity
End Property

Public Property Get length() As Long
    length = m_length
End Property

Public Function remove(ByVal firstPos As Long, ByVal removeLength As Long) As StringBuilder

    ' 引数チェック
    If firstPos < 1 Then
        Exit Function
    End If

    If removeLength < 1 Then
        Exit Function
    End If

    If m_length < firstPos + removeLength - 1 Then
        ' 指定位置が文字列の長さを超えている場合は、削除位置を調整する
        removeLength = m_length - firstPos + 1
    End If
    
    ' 削除対象文字列の右以降の文字列
    Dim removeRightStr As String
    Dim removeRightStrLen As Long
    
    removeRightStr = Mid(m_buffer, firstPos + removeLength)
    removeRightStrLen = Len(removeRightStr)

    ' 指定位置の文字列を削除する
    Mid(m_buffer, firstPos) = removeRightStr
    ' 残りの文字列をNULLで埋める
    Mid(m_buffer, firstPos + removeRightStrLen) = String$(removeRightStrLen, vbNullChar)
    
    ' 削除した文字列長で長さを再設定
    m_length = m_length - removeLength
    
    Set remove = Me
        
End Function

Public Function clear() As StringBuilder

    m_length = 0
    m_buffer = String$(m_capacity, vbNullChar)
    
    Set clear = Me
    
End Function

Public Function append(ByRef str As String) As StringBuilder

    Dim pos As Long
    
    ' 文字列の挿入位置
    pos = m_length + 1
    
    ' 文字列の長さを再設定
    m_length = m_length + Len(str)
    
    ' 容量の拡張
    expandCapacity
    
    ' 末尾に文字列を追加する
    Mid(m_buffer, pos) = str
    
    Set append = Me
    
End Function

Public Function insert(ByRef str As String, _
                       ByVal position As Long) As StringBuilder
    
    ' 引数チェック
    If position < 1 Then
        Exit Function
    End If

    If position > m_length + 1 Then
        Exit Function
    End If

    ' 変更前の長さを保持しておく
    Dim tmpLen As Long
    tmpLen = m_length
    
    ' 追加する文字列を加味した長さを設定
    m_length = m_length + Len(str)
    
    ' 容量の拡張
    expandCapacity
    
    ' 挿入位置から末尾までの文字列を、追加する文字列の長さ分、右にスライドする
    Mid(m_buffer, position + Len(str)) = Mid$(m_buffer, position, tmpLen)
    
    ' 挿入位置に新しい文字列を追加する
    Mid(m_buffer, position) = str
    
    Set insert = Me
    
End Function

Public Function replace(ByRef strFind As String, ByRef strReplace As String) As StringBuilder
    
    Dim strFindLen  As Long
    strFindLen = Len(strFind)
    
    ' InStr関数の検索開始位置（置き換え後の文字列を再度検索対象に含め無限ループに陥らないように、調整するインデックス位置）
    Dim instrStartIndex As Long
    instrStartIndex = 1
    
    Dim findIndex As Long
    
    Do While True
    
        findIndex = InStr(instrStartIndex, m_buffer, strFind)
        
        If findIndex = 0 Then
            ' 検索文字列が見つからない場合は、処理を終了する
            Exit Do
        End If
        
        ' 検索文字列を一度削除する
        Me.remove findIndex, strFindLen
        ' 置換文字列で置き換える
        Me.insert strReplace, findIndex
        
        instrStartIndex = findIndex + strFindLen
        
    Loop
    
    Set replace = Me
    
End Function

Public Function str() As String
    str = Left$(m_buffer, m_length)
End Function

Private Sub expandCapacity()

    Dim tmpCap As Long

    ' 容量オーバーチェック
    If m_length > m_capacity Then
        
        ' 容量を2倍にする
        tmpCap = m_capacity
        Do While tmpCap < m_length
            tmpCap = tmpCap * 2
        Loop
        
        m_capacity = tmpCap
        
    End If
    
End Sub
