VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' CSV解析クラス
'
' 作成者　：Ison
' 履歴　　：2019/12/03　新規作成
'
' 特記事項：
' *********************************************************

' ---------------------------------------------------------
' □プロパティ
' ---------------------------------------------------------
' 区切り文字
Private m_separator As String

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
' ▽初期化処理
' =========================================================
Public Sub init(Optional ByVal separator As String = ",")
    m_separator = separator
End Sub

' =========================================================
' ▽CSV解析処理
' =========================================================
Public Function parse(ByVal csvStr As String) As ValCollection

    Dim ret    As New ValCollection
    Dim retRec As New ValCollection

    Dim i As Long
    
    Dim outNextPos As Long
    Dim outFieldStr As String
    Dim outIsNextRec As Boolean
    
    Dim csvChar   As String
    Dim csvStrLen As Long
    csvStrLen = Len(csvStr)

    i = 1
    Do While i <= csvStrLen
    
        ' フィールドを解析する
        parseOneField csvStr, i, outNextPos, outFieldStr, outIsNextRec
        
        ' レコードにフィールドを追加する
        retRec.setItem outFieldStr
        ' インデックスを進める
        i = outNextPos
        
        If outIsNextRec Then
            ' 次の行に移動するのでレコードリストにレコードを追加する
            ret.setItem retRec
            ' レコードを生成する
            Set retRec = New ValCollection
        End If
        
    Loop
    
    If retRec.count > 0 Then
        ' 次の行に移動するのでレコードリストにレコードを追加する
        ret.setItem retRec
    End If
    
    Set parse = ret

End Function

' =========================================================
' ▽CSVフィールド解析処理
' =========================================================
Private Sub parseOneField(ByVal csvStr As String _
                        , ByVal first As Long _
                        , ByRef outNextPos As Long _
                        , ByRef outFieldStr As String _
                        , ByRef outIsNextRec As Boolean)

    Dim isEnclosed As Boolean

    Dim i As Long
    
    Dim csvCharBefore   As String
    Dim csvCharBefore2  As String
    Dim csvCharNext     As String
    Dim csvChar         As String
    Dim csvStrLen As Long
    csvStrLen = Len(csvStr)
    
    outNextPos = first
    outFieldStr = ""
    outIsNextRec = False


    For i = first To csvStrLen
    
        csvChar = Mid$(csvStr, i, 1)
        
        If i = first Then
            ' 最初の文字
        
            If csvChar = """" Then
                ' 最初の文字が囲み文字であるかを判定する
                isEnclosed = True

            ElseIf csvChar = m_separator Then
                ' 現在の文字が区切り文字の場合
                ' フィールド解析終了
                outFieldStr = Mid$(csvStr, first, i - first)
                
                isNextRec csvStr, i + 1, outNextPos, outIsNextRec
                
                Exit Sub

            Else
            
                isNextRec csvStr, i, outNextPos, outIsNextRec
                If outIsNextRec Then
                    Exit Sub
                End If
            
            End If
            
        Else
            ' 2文字目以降
        
            If isEnclosed Then
                ' 囲み文字の場合
                
                csvCharNext = Mid$(csvStr, i + 1, 1)
                
                If ( _
                    ( _
                      csvCharBefore <> """" Or _
                     (csvCharBefore = """" And csvCharBefore2 = """") Or _
                     (csvCharBefore = """" And csvCharBefore2 = "") _
                    )) And _
                     csvChar = """" And _
                    (csvCharNext = m_separator Or csvCharNext = vbCr Or csvCharNext = vbLf Or csvCharNext = "") Then
                    ' 以下の条件を満たすこと
                    ' ・現在の文字が囲み文字である
                    ' ・直前の文字が囲み文字ではない または 囲み文字であっても2文字連続で囲み文字である
                    ' ・直後の文字が区切り文字で終了、改行コードで終了、文字の終端である
                    
                    ' フィールド解析終了
                    
                    outNextPos = i + 1 ' 次回文字、区切り文字と続くはずなので、1文字飛ばす
                    outFieldStr = replace(Mid$(csvStr, first + 1, i - 1 - first), """""", """")
                    
                    isNextRec csvStr, i + 2, outNextPos, outIsNextRec

                    Exit Sub
                
                End If
            Else
                ' 囲み文字なし

                If csvChar = m_separator Then
                    ' 現在の文字が区切り文字の場合
                    ' フィールド解析終了
                    outFieldStr = Mid$(csvStr, first, i - first)
                    
                    isNextRec csvStr, i + 1, outNextPos, outIsNextRec
                    
                    Exit Sub
                    
                Else
                
                    isNextRec csvStr, i, outNextPos, outIsNextRec
                    If outIsNextRec Then
                        outFieldStr = Mid$(csvStr, first, i - first)
                        Exit Sub
                    End If
                
                End If
                
            
            End If
            
        End If
        
        csvCharBefore2 = csvCharBefore
        csvCharBefore = csvChar
        
    Next
    
    outNextPos = i
    outFieldStr = Mid$(csvStr, first, i - 1)
    outIsNextRec = True

End Sub

Private Sub isNextRec(ByVal csvStr As String _
                        , ByVal first As Long _
                        , ByRef outNextPos As Long _
                        , ByRef outIsNextRec As Boolean)

    outIsNextRec = False
    outNextPos = first

    Dim csvCharNext     As String
    Dim csvChar         As String
    
    csvChar = Mid$(csvStr, first, 1)
    csvCharNext = Mid$(csvStr, first + 1, 1)

    If csvChar = vbCr And csvCharNext = vbLf Then
        ' CRLFで終わった場合
        outIsNextRec = True
        outNextPos = first + 2
    ElseIf csvChar = vbCr Or _
           csvChar = vbLf Then
        ' CR or LFで終わった場合
        outIsNextRec = True
        outNextPos = first + 1
    End If

End Sub


Private Function excludeEncloseChar(ByVal field As String) As String

    If InStr(field, """") = 1 And InStrRev(field, """") = Len(field) Then
        ' 囲み文字で囲まれている場合は取り除く
        excludeEncloseChar = Mid$(field, 1, Len(field) - 2)
    Else
        ' 囲み文字の場合はそのまま
        excludeEncloseChar = field
    End If


End Function


