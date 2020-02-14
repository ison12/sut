VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' INIデータ操作クラス
'
' 制限事項
' ・セクションの省略不可
' ・コメント設定不可
' ・データ閲覧時、改行コードやイコール記号が特殊な制御コードに変換される
'
' 作成者　：Ison
' 履歴　　：2020/02/12　新規作成
'
' 特記事項：
' *********************************************************
' データ
Private datas_ As ValCollection
' 変更済みフラグ
Private isModified_ As Boolean

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Public Sub Class_Initialize()
    
    datas = New ValCollection
    
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
Public Sub init(ByVal contents As String)
    
    isModified = False
    datas = readContents(contents)

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
    
    datas = Nothing

End Sub

' =========================================================
' ▽INIファイルから任意のキー値を取得する
'
' 概要　　　：
' 引数　　　：section セクション
' 　　　　　　key     キー
' 戻り値　　：セクションとキーに紐づく値
'
' =========================================================
Public Function getValue _
            (ByVal SECTION As String, _
             ByVal key As String) As String
             
    ' 戻り値の初期値
    getValue = ""
             
    Dim dataSection As ValCollection
    Dim data As Variant
             
    Set dataSection = datas.getItem(SECTION)
    If Not dataSection Is Nothing Then

        data = dataSection.getItem(key, vbVariant)
        If IsArray(data) Then
            getValue = data(2)
        End If
        
    End If
    
End Function

' =========================================================
' ▽INIファイルに任意のキー値を設定する
'
' 概要　　　：
' 引数　　　：section セクション
' 　　　　　　key     キー
'             value   値
' 戻り値　　：Trueの場合、正常に書き込み完了
'
' =========================================================
Public Sub setValue _
            (ByVal SECTION As String, _
             ByVal key As String, _
             ByVal value As String)
             
    Dim dataSection As ValCollection
             
    If datas.exist(SECTION) = False Then
    
        Set dataSection = New ValCollection
        datas.setItem dataSection, SECTION
        
    Else
        Set dataSection = datas.getItem(SECTION)
        
    End If

    dataSection.setItem Array(SECTION, key, value), key

    ' 変更したので変更済みとする
    isModified = True
    
End Sub

' =========================================================
' ▽INIファイルの任意のセクションに複数のキー値を設定する
'
' 概要　　　：戻り値はValCollectionで要素は配列となる
'             ------------------------------
'             element(0 to 2)の配列
'             ------------------------------
' 　　　　　　element(0): セクション
' 　　　　　　element(1): キー
' 　　　　　　element(2): 値
'
' 引数　　　：section セクション
' 戻り値　　：2次元配列
'
' =========================================================
Public Function getValues _
            (ByVal SECTION As String) As ValCollection
             
    Dim dataSection As ValCollection
             
    If datas.exist(SECTION) Then
        
        Set getValues = datas.getItem(SECTION)
        
    Else
        Set getValues = New ValCollection
        
    End If

End Function

' =========================================================
' ▽INIファイルの任意のセクションに複数のキー値を設定する
'
' 概要　　　：引数valuesはValCollectionであること。
' 　　　　　　要素は配列で2パターン想定。
'
'             パターン1
'             ------------------------------
'             element(0 to 2)の配列
'             ------------------------------
' 　　　　　　element(0): セクション
' 　　　　　　element(1): キー
' 　　　　　　element(2): 値
'
'             パターン2
'             ------------------------------
'             element(0 to 1)の配列
'             ------------------------------
' 　　　　　　element(0): キー
' 　　　　　　element(1): 値

' 引数　　　：section セクション
'             values  複数の値
' 戻り値　　：Trueの場合、正常に書き込み完了
'
' =========================================================
Public Sub setValues _
            (ByVal SECTION As String, _
             ByRef values As ValCollection)
             
    Dim dataSection As ValCollection
    Dim data As Variant
             
    If datas.exist(SECTION) = False Then
        Set dataSection = New ValCollection
        datas.setItem dataSection, SECTION
    Else
        Set dataSection = datas.getItem(SECTION)
    End If

    Dim valuesOfElement As Variant
    
    For Each valuesOfElement In values.col
        
        If UBound(valuesOfElement) - LBound(valuesOfElement) + 1 >= 3 Then
            ' 配列のサイズが3つならセクションあり
            dataSection.remove valuesOfElement(1)
            dataSection.setItem Array(valuesOfElement(0), valuesOfElement(1), valuesOfElement(2), valuesOfElement(1))
            
        Else
            ' 配列のサイズが3つ未満ならセクションなし
            dataSection.remove valuesOfElement(0)
            dataSection.setItem Array(SECTION, valuesOfElement(0), valuesOfElement(1), valuesOfElement(0))
        End If
        
    Next

    ' 変更したので変更済みとする
    isModified = True
    
End Sub

' =========================================================
' ▽INIファイルの任意のセクションまたはキーを削除する
'
' 概要　　　：引数keyを省略した場合、引数section配下の全てのキー値が削除される。
' 　　　　　　引数keyを省略しなかった場合、キー値が削除される。
'
' 引数　　　：section セクション
' 　　　　　　key     キー
' 戻り値　　：Trueの場合、正常に書き込み完了
'
' =========================================================
Public Sub delete _
            (ByVal SECTION As String, _
             Optional ByVal key As String = "")

             
    Dim dataSection As ValCollection
             
    Set dataSection = datas.getItem(SECTION)

    If dataSection Is Nothing Then
        ' 後続の処理でNothingエラーが発生しないように空オブジェクトを設定しておく
        Set dataSection = New ValCollection
    End If
    
    If key = "" Then
        
        ' Sectionの削除
        dataSection.removeAll
        
    Else
        
        ' Keyの削除
        dataSection.remove key
        
    End If
    
    ' 変更したので変更済みとする
    isModified = True
    
End Sub

' =========================================================
' ▽INIコンテンツの読み込み
'
' 概要　　　：
' 引数　　　：contents 内容
' 戻り値　　：コンテンツデータ
'
' =========================================================
Public Function readContents(ByVal contents As String) As ValCollection

    ' 戻り値
    Dim datas As ValCollection
    Set datas = New ValCollection
    
    Dim sectionData As ValCollection
    
    ' 改行コードを統一する
    contents = replace(contents, vbCr & vbLf, vbLf)
    contents = replace(contents, vbCr, vbLf)
    
    Dim key As String
    Dim val As String
    Dim findKeyValueSeparator As Long
    
    Dim currentSection As String
    
    Dim contentsArrayEle As String
    Dim contentsArray As Variant
    
    ' 内容を改行コード区切りにする
    contentsArray = Split(contents, vbLf)
    
    Dim i As Long
    For i = LBound(contentsArray) To UBound(contentsArray)
        ' 一行ずつ走査する
    
        contentsArrayEle = contentsArray(i)
        
        If contentsArrayEle = Empty Then
            ' 空文字列の場合はスキップする
            GoTo loop_end
        End If
        
        If Mid(contentsArrayEle, 1, 1) = "[" And _
           Mid(contentsArrayEle, Len(contentsArrayEle), 1) = "]" Then
           ' セクションの場合
            
            currentSection = decodeSpecialChars(Mid(contentsArrayEle, 2, Len(contentsArrayEle) - 2))
            
            Set sectionData = New ValCollection
            datas.setItem sectionData, currentSection
            
            GoTo loop_end
        End If
        
        If sectionData Is Nothing Then
            ' セクションがない場合は、スキップ（セクションが登場するまでデータを読み込まない）
            GoTo loop_end
        End If
        
        ' キーと値を取得する
        findKeyValueSeparator = InStr(contentsArrayEle, "=")
        If findKeyValueSeparator > 0 Then
            ' key=value
            key = decodeSpecialChars(Mid(contentsArrayEle, 1, findKeyValueSeparator - 1))
            val = decodeSpecialChars(Mid(contentsArrayEle, findKeyValueSeparator + 1))
        Else
            ' key
            key = decodeSpecialChars(contentsArrayEle)
            val = ""
        End If
        
        sectionData.setItem Array(currentSection, key, val), key
    
loop_end:

    Next
    
    Set readContents = datas
    
End Function

' =========================================================
' ▽INIファイルの書き込みデータの生成
'
' 概要　　　：
' 引数　　　：datas    データ
' 戻り値　　：書き込みデータ
'
' =========================================================
Public Function createWriteContents(ByVal datas As ValCollection) As String

    Dim ret As New StringBuilder

    Dim sectionName As String
    Dim sectionData As ValCollection
    Dim keyValue As Variant
    
    For Each sectionData In datas.col
    
        If sectionData.count > 0 Then
            
            sectionName = sectionData.getItemByIndex(1, vbVariant)(0)
        
            ret.append "[" & encodeSpecialChars(sectionName) & "]" & vbNewLine
            
            For Each keyValue In sectionData.col
                ret.append encodeSpecialChars(keyValue(1)) & "=" & encodeSpecialChars(keyValue(2)) & vbNewLine
            Next
            
        End If
            
    Next
    
    createWriteContents = ret.str
    
End Function

' =========================================================
' ▽特殊な文字をエンコードする
'
' 概要　　　：
' 引数　　　：val 対象文字列
' 戻り値　　：変換後の文字列
'
' =========================================================
Private Function encodeSpecialChars(ByVal val As String) As String
    
    encodeSpecialChars = val
    
    ' ASCIIコードの17~19 装置制御コードを代用する
    encodeSpecialChars = replace(encodeSpecialChars, vbCr, Chr(17))
    encodeSpecialChars = replace(encodeSpecialChars, vbLf, Chr(18))
    encodeSpecialChars = replace(encodeSpecialChars, "=", Chr(19))
    
End Function

' =========================================================
' ▽特殊な文字をデコードする
'
' 概要　　　：
' 引数　　　：val 対象文字列
' 戻り値　　：変換後の文字列
'
' =========================================================
Private Function decodeSpecialChars(ByVal val As String) As String
    
    decodeSpecialChars = val
    
    ' ASCIIコードの17~19 装置制御コードを代用する
    decodeSpecialChars = replace(decodeSpecialChars, Chr(17), vbCr)
    decodeSpecialChars = replace(decodeSpecialChars, Chr(18), vbLf)
    decodeSpecialChars = replace(decodeSpecialChars, Chr(19), "=")
    
End Function

' =========================================================
' ▽INIデータの取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：データ
'
' =========================================================
Public Property Get datas() As ValCollection

    Set datas = datas_

End Property

' =========================================================
' ▽INIデータの設定
'
' 概要　　　：
' 引数　　　：vNewValue データ
' 戻り値　　：
'
' =========================================================
Public Property Let datas(ByVal vNewValue As ValCollection)

    Set datas_ = vNewValue

End Property

' =========================================================
' ▽変更済みフラグの取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：変更済みフラグ
'
' =========================================================
Public Property Get isModified() As Boolean
    
    isModified = isModified_

End Property

' =========================================================
' ▽変更済みフラグの設定
'
' 概要　　　：
' 引数　　　：vNewValue 変更済みフラグ
' 戻り値　　：
'
' =========================================================
Public Property Let isModified(ByVal vNewValue As Boolean)
    
    isModified_ = vNewValue

End Property
