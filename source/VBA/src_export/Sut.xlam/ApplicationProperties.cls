VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationProperties"
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
' INIデータワークシート
Private iniw As IniWorksheet
' INIデータファイル
Private inif As IniFile

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Public Sub Class_Initialize()
    
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
' 引数　　　：book_      ブック
'     　　　：sheetName_ シート名
'     　　　：fileName_  ファイル名
' 戻り値　　：
'
' =========================================================
Public Sub initWorksheet(ByVal book_ As Workbook, ByVal sheetName_ As String, ByVal fileName_ As String)

    Set iniw = New IniWorksheet
    iniw.init book_, sheetName_, fileName_

End Sub

' =========================================================
' ▽初期化処理。
'
' 概要　　　：
' 引数　　　：fileName_ ファイル名
' 戻り値　　：
'
' =========================================================
Public Sub initFile(ByVal fileName_ As String)

    Set inif = New IniFile
    inif.init fileName_

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

    If Not iniw Is Nothing Then
        iniw.destroy
    End If

    If Not inif Is Nothing Then
        inif.destroy
    End If

End Sub

' =========================================================
' ▽読み込み処理。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub readData()

    If Not iniw Is Nothing Then
        iniw.readSheet
    End If

    If Not inif Is Nothing Then
        inif.readFile
    End If

End Sub

' =========================================================
' ▽書き込み処理。
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub writeData()

    If Not iniw Is Nothing Then
        iniw.writeSheet
    End If

    If Not inif Is Nothing Then
        inif.writeFile
    End If

End Sub

' ------------------------------------------------------------------
' IniDataをラップして各メソッドを呼び出す
' ------------------------------------------------------------------
Public Function getValue _
            (ByVal SECTION As String, _
             ByVal key As String) As String

    If Not iniw Is Nothing Then
        If iniw.isExistsData Then
            getValue = iniw.getValue(SECTION, key)
            Exit Function
        End If
    End If

    If Not inif Is Nothing Then
        getValue = inif.getValue(SECTION, key)
        Exit Function
    End If

End Function

Public Sub setValue _
            (ByVal SECTION As String, _
             ByVal key As String, _
             ByVal value As String)

    If Not iniw Is Nothing Then
        iniw.setValue SECTION, key, value
    End If

    If Not inif Is Nothing Then
        inif.setValue SECTION, key, value
    End If

End Sub

Public Function getValues _
            (ByVal SECTION As String) As ValCollection

    If Not iniw Is Nothing Then
        If iniw.isExistsData Then
            Set getValues = iniw.getValues(SECTION)
            Exit Function
        End If
    End If

    If Not inif Is Nothing Then
        Set getValues = inif.getValues(SECTION)
        Exit Function
    End If
    
    Set getValues = New ValCollection
    
End Function

Public Sub setValues _
            (ByVal SECTION As String, _
             ByRef values As ValCollection)

    If Not iniw Is Nothing Then
        iniw.setValues SECTION, values
    End If

    If Not inif Is Nothing Then
        inif.setValues SECTION, values
    End If

End Sub

Public Sub delete _
            (ByVal SECTION As String, _
             Optional ByVal key As String = "")

    If Not iniw Is Nothing Then
        iniw.delete SECTION, key
    End If

    If Not inif Is Nothing Then
        inif.delete SECTION, key
    End If

End Sub




