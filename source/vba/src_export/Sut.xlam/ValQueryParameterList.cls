VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryParameterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリパラメータリストを表現するクラス
'
' 作成者　：Ison
' 履歴　　：2019/12/08　新規作成
'
' 特記事項：
' *********************************************************

' リスト
Private m_list    As ValCollection
' ブック
Private m_book    As Workbook

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    Set m_list = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    Set m_list = Nothing
End Sub

' =========================================================
' ▽初期化処理メソッド
' =========================================================
Public Sub init(ByVal book As Workbook)
    Set m_book = book
End Sub

' =========================================================
' ▽プロパティ
' =========================================================
Public Property Get list() As ValCollection
    Set list = m_list
End Property

Public Property Let list(ByVal newValue As ValCollection)
    Set m_list = newValue
End Property

' =========================================================
' ▽設定情報の生成
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initWorksheet m_book, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "queryParameter" & ".ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ▽クエリパラメータ情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub readForData()
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties

    ' データを読み込む
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
            
    Dim queryParameterList As New ValCollection
    Dim queryParameter As ValQueryParameter
    
    Dim i As Long: i = 1
    Do While True
    
        val = values.getItem(i & "_" & "name", vbVariant)
        If Not IsArray(val) Then
            Exit Do
        End If
        
        Set queryParameter = New ValQueryParameter
                    
        val = values.getItem(i & "_" & "name", vbVariant): If IsArray(val) Then queryParameter.name = val(2)
        val = values.getItem(i & "_" & "value", vbVariant): If IsArray(val) Then queryParameter.value = val(2)
        
        queryParameterList.setItem queryParameter
    
        i = i + 1
    Loop
    
    list = queryParameterList

End Sub

' =========================================================
' ▽クエリパラメータ情報を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub writeForData()

    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' 書き込みデータ
    Dim val As Variant
    Dim values As New ValCollection
    
    Dim i As Long: i = 1
    For Each val In list.col
    
        values.setItem Array(i & "_" & "name", val.name)
        values.setItem Array(i & "_" & "value", val.value)
    
        i = i + 1
    Next

    ' データを書き込む
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
End Sub
