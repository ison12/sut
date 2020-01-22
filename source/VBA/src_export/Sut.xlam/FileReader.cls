VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ファイル読み込みクラス
'
' 作成者　：Ison
' 履歴　　：2008/06/28　新規作成
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
    If stream.state <> 0 Then
    
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
              , ByVal n As String)

    ' ファイルパスを設定
    filePath = f
    ' 文字コードを設定
    charSet = c
    ' 改行コードを設定
    newline = n

End Sub

' =========================================================
' ▽読み込み処理
'
' 概要　　　：読み込み処理を実施する
' 引数　　　：
'
' 戻り値　　：読み込んだファイルの内容
'
' =========================================================
Public Function readText() As String

    ' ストリームオブジェクトを設定
    Set stream = CreateObject("ADODB.Stream")
    
    ' 読み取りで開く
    stream.Mode = 3
    ' テキストで開く
    stream.Type = 2

    ' 文字コードプロパティを指定
    ' システムが認識できる文字セット文字列の一覧については、Windows レジストリの HKEY_CLASSES_ROOT\MIME\Database\Charset のサブキーを参照してください。
    stream.charSet = charSet
    
    ' 改行コードプロパティを指定
    stream.LineSeparator = convertStreamConstNewLine(newline)
    
    ' ストリームを開く
    stream.Open
    stream.LoadFromFile filePath
    
    readText = stream.readText
    
    ' ストリームを閉じる
    stream.Close

End Function

' =========================================================
' ▽読み込み処理中断
'
' 概要　　　：読み込み処理を中断する
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub cancel()

    ' ストリームが閉じていない場合
    If stream.state <> 0 Then
    
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
