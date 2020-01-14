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
' 作成者　：Hideki Isobe
' 履歴　　：2008/04/17　新規作成
'
' 特記事項：
' *********************************************************

Private stream As Object

' ファイルパス
Private filePath As String
' 文字コード
Private charSet As String
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

    If stream Is Nothing Then
    
        Exit Sub
    End If
    
    ' ストリームが閉じていない場合
    If stream.State <> 0 Then
    
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
        
    End If
End Sub

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
    charSet = c
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
Public Sub Cancel()

    ' ストリームが閉じていない場合
    If stream.State <> 0 Then
    
        ' ストリームを閉じる
        stream.Close
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

