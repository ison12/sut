VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ファイル出力クラス
'
' 作成者　：Ison
' 履歴　　：2008/04/17　新規作成
'
' 特記事項：
' *********************************************************

Private stream As Object

' ファイルパス
Private filePath As String
' 文字コード
Private charSet As String
' BOM付き
Private isWithBom As Boolean
' 改行コード
Private newline As String

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    destroy
End Sub

' =========================================================
' ▽文字コード取得処理
'
' 概要　　　：
' 引数　　　：c            文字コード
' 　　　　　　outChar      戻り値…変換後の文字コード
' 　　　　　　outIsWithBom 戻り値…文字コードがBOM有りかどうか
'
' 戻り値　　：
'
' =========================================================
Private Function getCharSet(ByVal c As String _
                        , ByRef outChar As String _
                        , ByRef outIsWithBom As Boolean)

    If InStr(UCase$(c), "UTF-8") > 0 Then
        ' UTF-8
        
        outChar = "UTF-8"
        
        If InStr(UCase$(c), "BOM") > 0 Then
            ' BOMあり
            outIsWithBom = True
        Else
            ' BOMなし
            outIsWithBom = False
        End If
        
    Else
        ' 上記以外
        
        outChar = c
        outIsWithBom = False
        
    End If

End Function

' =========================================================
' ▽初期化処理
'
' 概要　　　：初期化処理を実行する。
' 引数　　　：f      ファイルパス
' 　　　　　　c      文字コード
' 　　　　　　n      改行コード
' 　　　　　　append 追加有無
'
' 戻り値　　：
'
' =========================================================
Public Sub init(ByVal f As String _
              , ByVal c As String _
              , ByVal n As String _
              , Optional ByVal append = False)
              
    ' ファイルパスを設定
    filePath = f
    ' 文字コードを設定
    getCharSet c, charSet, isWithBom
    ' 改行コードを設定
    newline = n

    ' ストリームオブジェクトを設定
    Set stream = CreateObject("ADODB.Stream")
    
    ' 文字コードプロパティを指定
    ' システムが認識できる文字セット文字列の一覧については、Windows レジストリの HKEY_CLASSES_ROOT\MIME\Database\Charset のサブキーを参照してください。
    stream.charSet = charSet
    
    ' 改行コードプロパティを指定
    stream.LineSeparator = convertStreamConstNewLine(newline)
    
    ' ストリームを開く
    stream.Open
    
    ' 追加モード
    If append = True Then
    
        ' ファイルが存在している場合
        If VBUtil.isExistFile(filePath) = True Then
            
            ' 指定のファイルを読み込む
            stream.LoadFromFile filePath
            stream.position = stream.size
            
        End If
    End If
    
End Sub

' =========================================================
' ▽破棄処理
'
' 概要　　　：初期化処理を実行する。
' 引数　　　：f      ファイルパス
' 　　　　　　c      文字コード
' 　　　　　　n      改行コード
' 　　　　　　append 追加有無
'
' 戻り値　　：
'
' =========================================================
Public Sub destroy()

    On Error GoTo err

    If stream Is Nothing Then
    
        Exit Sub
    End If
    
    ' ストリームが閉じていない場合
    If stream.state <> 0 Then
    
        If UCase$(charSet) = "UTF-8" And isWithBom = False Then
            ' UTF-8 BOMなし
            
            ' バイト位置を先頭に戻す
            stream.position = 0
            stream.Type = 1 ' adTypeBinary
            
            ' BOMの3Byteをスキップ
            If stream.size >= 3 Then
                stream.position = 3
            End If
            
            ' BOMを抜いたバイトデータを書き込む
            Dim buf As Variant: buf = stream.read
            ' 空ファイルの場合は、nullが返却されるのでチェックする
            If isNull(buf) = False Then
                stream.position = 0
                stream.Write buf
                stream.SetEOS
            End If
            
        End If
    
        ' http://msdn.microsoft.com/ja-jp/library/cc389870.aspx
        ' SaveOptionsEnum
        '
        ' Stream オブジェクトからファイルを保存するときにファイルを作成するか、上書きするかを表します。これらの値は AND 演算子で結合することができます。
        '
        ' 定数  値  説明
        ' adSaveCreateNotExist  1   既定値です。FileName パラメータで指定したファイルがない場合は新しいファイルが作成されます。
        ' adSaveCreateOverWrite 2   FileName パラメータで指定したファイルがある場合は、現在開かれている Stream オブジェクトのデータでファイルが上書きされます。
        stream.SaveToFile filePath, 2
    
        ' ストリームを閉じる
        stream.Close
        Set stream = Nothing
        
    End If
    
    Exit Sub
    
err:

    Dim errT As errInfo
    errT = VBUtil.swapErr
        
    If Not stream Is Nothing Then
        ' ストリームを閉じる
        stream.Close
        Set stream = Nothing
    End If

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Sub

' =========================================================
' ▽書き込み処理
'
' 概要　　　：書き込み処理を実施する
' 引数　　　：buffer        書き込み内容
' 　　　　　　appendNewLine 改行コードの追加有無
'
' 戻り値　　：
'
' =========================================================
Public Sub writeText(ByVal buffer As String _
                   , Optional ByVal appendNewLine As Boolean = False)

    stream.writeText buffer, convertStreamConstAppendNewLine(appendNewLine)

End Sub

' =========================================================
' ▽書き込み処理中断
'
' 概要　　　：書き込み処理を中断する
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub cancel()

    ' ストリームが閉じていない場合
    If stream.state <> 0 Then
    
        ' ストリームを閉じる
        stream.Close
        Set stream = Nothing
    End If

End Sub

Private Function convertStreamConstNewLine(ByVal n As String) As Long

    ' http://msdn.microsoft.com/ja-jp/library/cc389826.aspx
    ' LineSeparatorsEnum
    '
    ' テキスト Stream オブジェクトの行区切りに使われている文字を表します。
    '
    ' 定数   値  説明
    ' adCR   13  改行復帰を示します。
    ' adCRLF -1  既定値です。改行復帰行送りを示します。
    ' adLF   10  行送りを示します。
    If n = vbCr & vbLf Then
    
        ' Windows
        convertStreamConstNewLine = -1
    
    ElseIf n = vbCr Then
    
        ' Mac
        convertStreamConstNewLine = 13
    
    ElseIf n = vbLf Then
    
        ' Unix
        convertStreamConstNewLine = 10
        
    ' 当てはまらない場合
    Else
    
        ' Windows
        convertStreamConstNewLine = -1
    
    End If


End Function

Private Function convertStreamConstAppendNewLine(ByVal appendNewLine As Boolean) As Long

    ' StreamWriteEnum
    '
    ' Stream オブジェクトに書き込む文字列に、行区切り文字を追加するかどうかを表します。
    '
    ' 定数  値  説明
    ' adWriteChar   0   既定値です。Stream オブジェクトに対して、Data パラメータで指定したテキスト文字列を書き込みます。
    ' adWriteLine   1   Stream オブジェクトに、テキスト文字列と行区切り文字を書き込みます。LineSeparator プロパティが定義されていない場合は、実行時エラーを返します。
    If appendNewLine = True Then
    
        convertStreamConstAppendNewLine = 1
    Else
    
        convertStreamConstAppendNewLine = 0
    End If

End Function

