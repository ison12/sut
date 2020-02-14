VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' INIファイル操作クラス
'
' 制限事項
' ・セクションの省略不可
' ・コメント設定不可
' ・データ閲覧時、改行コードやイコール記号が特殊な制御コードに変換される
'
' 作成者　：Ison
' 履歴　　：2020/02/13　新規作成
'
' 特記事項：
' *********************************************************
' ファイル名
Private fileName As String
' INIデータ
Private inid As IniData

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Public Sub Class_Initialize()
    
    Set inid = New IniData
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Public Sub Class_Terminate()
    
    destroy
    
End Sub

' =========================================================
' ▽初期化処理。
'
' 概要　　　：
' 引数　　　：fileName_ ファイル名
' 戻り値　　：
'
' =========================================================
Public Sub init(ByVal fileName_ As String)

    fileName = fileName_
    
    readFile

End Sub

' =========================================================
' ▽破棄処理。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub destroy()

    writeFile

End Sub

' =========================================================
' ▽INIファイルの読み込み
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Function readFile()

    Dim contents As String

    createFileIfFileNotExists

    Dim fr As FileReader
    Set fr = New FileReader: fr.init fileName, "utf-8", vbNewLine
    
    contents = fr.readText
    
    inid.datas = inid.readContents(contents)
    
    ' 読み込み直後なので変更済みをオフにする
    inid.isModified = False
    
End Function

' =========================================================
' ▽INIファイルの書き込み
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub writeFile()

    If inid.isModified = False Then
        ' 変更なしなので書き込まない
        Exit Sub
    End If

    Dim contents As String
    contents = inid.createWriteContents(inid.datas)

    Dim fr As FileWriter
    Set fr = New FileWriter: fr.init fileName, "utf-8", vbNewLine
    
    fr.writeText contents
    
    ' 書き込んだので変更済みをオフにする
    inid.isModified = False
    
End Sub

' =========================================================
' ▽INIファイルの作成
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub createFileIfFileNotExists()

    ' ファイル番号
    Dim fileNum As Long
    
    ' ファイルが存在しない場合
    If (dir(fileName, vbNormal) = "") Then
    
        VBUtil.createDir VBUtil.extractDirPathFromFilePath(fileName)
    
        '指定ファイルが無いので作成
        Dim fw As FileWriter
        Set fw = New FileWriter: fw.init fileName, "utf-8", vbNewLine
        fw.destroy
    
    End If
    
End Sub

' ------------------------------------------------------------------
' IniDataをラップして各メソッドを呼び出す
' ------------------------------------------------------------------
Public Function getValue _
            (ByVal SECTION As String, _
             ByVal key As String) As String

    getValue = inid.getValue(SECTION, key)
End Function

Public Sub setValue _
            (ByVal SECTION As String, _
             ByVal key As String, _
             ByVal value As String)

    inid.setValue SECTION, key, value
End Sub

Public Function getValues _
            (ByVal SECTION As String) As ValCollection

    Set getValues = inid.getValues(SECTION)
End Function

Public Sub setValues _
            (ByVal SECTION As String, _
             ByRef values As ValCollection)

    inid.setValues SECTION, values
End Sub

Public Sub delete _
            (ByVal SECTION As String, _
             Optional ByVal key As String = "")

    inid.delete SECTION, key
End Sub


