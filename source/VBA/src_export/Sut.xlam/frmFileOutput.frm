VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileOutput 
   Caption         =   "ファイル出力"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7815
   OleObjectBlob   =   "frmFileOutput.frx":0000
End
Attribute VB_Name = "frmFileOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' ファイル出力を行うフォーム
'
' 作成者　：Ison
' 履歴　　：2008/09/06　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽OKボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event ok(ByVal filePath As String _
              , ByVal characterCode As String _
              , ByVal newline As String)

' =========================================================
' ▽キャンセルボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' 文字コードリスト
Private charcterList As CntListBox

' デフォルトファイル名
Private defaultFileName As String

' 対象ブック
Private targetBook As Workbook
' 対象ブックを取得する
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal  モーダルまたはモードレス表示指定
' 　　　　　　header ヘッダテキスト
' 　　　　　　defFileName デフォルトファイル名
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByVal header As String _
                 , ByVal defFileName As String)

    lblHeader.Caption = header
    defaultFileName = defFileName

    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal

End Sub

' =========================================================
' ▽フォーム非表示
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub HideExt()

    deactivate
    
    Main.storeFormPosition Me.name, Me
    Me.Hide
End Sub

' =========================================================
' ▽フォーム初期化時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
    ' ロード時点のアクティブブックを保持しておく
    Set targetBook = ExcelUtil.getActiveWorkbook
    ' 初期化処理を実行する
    initial

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォーム破棄時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Terminate()

    On Error GoTo err
    
    ' 破棄処理を実行する
    unInitial
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォームアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Activate()

End Sub

' =========================================================
' ▽フォームの閉じる時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' 本処理では処理自体をキャンセルする
        cancel = True
        ' 以下のイベント経由で閉じる
        cmdCancel_Click
    End If
    
End Sub

' =========================================================
' ▽ファイル選択ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
'
' =========================================================
Private Sub btnFileSelect_Click()

    Dim selectFile As String
    
    selectFile = saveFileDialog
    
    If selectFile <> "" Then
        ' ファイルを開くダイアログをオープンしてユーザにファイルを選択させる
        txtFilePath.text = selectFile
    End If
    
End Sub

' =========================================================
' ▽ファイルを開くダイアログオープン
'
' 概要　　　：ファイルを開くダイアログをオープンする
'
' =========================================================
Private Function saveFileDialog() As String

    On Error GoTo err
        
    ' 選択ファイル
    Dim selectFile As String
    
    ' 開くダイアログを選択する
    selectFile = VBUtil.openFileSaveDialog("保存ファイルを選択してください。" _
                                         , "SQLファイル (*.sql),*.sql,すべてのファイル (*.*),*.*" _
                                         , VBUtil.extractFileName(txtFilePath.value))

    ' ファイルパスを設定する
    saveFileDialog = selectFile
    
    Exit Function
    
err:

End Function

' =========================================================
' ▽OKボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err
    
    ' ファイルパス
    Dim filePath As String
    ' ディレクトリパス
    Dim dirPath As String
    ' 文字コード
    Dim characterCode As String
    ' 改行コード
    Dim newline As String
    ' フォルダ作成の成功有無
    Dim isSuccessCreateDir As Boolean

    ' ファイルパスを取得
    filePath = txtFilePath.text
    ' 文字コードを取得
    characterCode = cboChoiceCharacterCode.text
    ' 改行コードを取得
    newline = cboChoiceNewLine.text
    
    If VBUtil.isExistDirectory(filePath) Then
        ' ファイルパスがディレクトリの場合はエラーとする
        VBUtil.showMessageBoxForWarning "フォルダが指定されています。ファイルパスを指定してください。" _
                                      , ConstantsCommon.APPLICATION_NAME _
                                      , Nothing

        Exit Sub
    End If
    
    ' ファイルパスの親ディレクトリを取得する
    dirPath = VBUtil.extractDirPathFromFilePath(filePath)

    ' --------------------------------------
    ' 親フォルダを作成する
    On Error Resume Next
    
    isSuccessCreateDir = False
    
    VBUtil.createDir dirPath
    If err.Number = 0 Then
        ' 作成に成功
        isSuccessCreateDir = True
    End If
    
    On Error GoTo err
    ' --------------------------------------

    ' フォルダへのテスト出力に失敗した場合
    If isSuccessCreateDir = False Or VBUtil.touch(dirPath) = False Then
    
        VBUtil.showMessageBoxForWarning "指定されたファイルパスにファイルが出力できません。" & vbNewLine & "未入力、不正なパス、または権限が不足している可能性があります。" _
                                      , ConstantsCommon.APPLICATION_NAME _
                                      , Nothing
        
        Exit Sub
    End If
    
    ' フォームを閉じる
    HideExt
    
    ' OKイベントを送信する
    RaiseEvent ok(filePath, characterCode, VBUtil.convertNewLineStrToNewLineCode(cboChoiceNewLine.text))
    
    ' ファイル出力オプションを書き込む
    storeFileOutputOption

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub


' =========================================================
' ▽キャンセルボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdCancel_Click()

    On Error GoTo err
    
    ' フォームを閉じる
    HideExt
    
    ' キャンセルイベントを送信する
    RaiseEvent cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽初期化処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub initial()

    ' 文字コードリストのコントロールオブジェクトを初期化する
    Set charcterList = New CntListBox
    
    charcterList.init cboChoiceCharacterCode
    charcterList.addAll VBUtil.getEncodeList
    
    ' 改行コードリストに改行コードを追加する
    Dim newLineList As ValCollection
    Set newLineList = VBUtil.getNewlineList
    
    Dim var As Variant
    
    For Each var In newLineList.col
    
        cboChoiceNewLine.addItem var
    Next
    
    cboChoiceCharacterCode.value = "Shift_JIS"
    cboChoiceNewLine.ListIndex = 0
    
End Sub

' =========================================================
' ▽後始末処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub unInitial()

End Sub

' =========================================================
' ▽アクティブ時の処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()

    ' ファイル出力オプションを読み込む
    restoreFileOutputOption
    
    ' ファイルパスにデフォルトのファイル名を設定する
    If txtFilePath.value = "" Then
        txtFilePath.value = VBUtil.concatFilePath( _
                                        VBUtil.extractDirPathFromFilePath(targetBook.path) _
                                      , defaultFileName)
    Else
        txtFilePath.value = VBUtil.concatFilePath( _
                                        VBUtil.extractDirPathFromFilePath(txtFilePath.value) _
                                      , defaultFileName)
    End If
    
End Sub

' =========================================================
' ▽ノンアクティブ時の処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub deactivate()

End Sub

' =========================================================
' ▽設定情報の生成
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & Me.name & ".ini"
    appProp.initWorksheet targetBook, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & Me.name & ".ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ▽ファイルオプションを保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeFileOutputOption()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' 書き込みデータ
    Dim values As New ValCollection
    
    values.setItem Array(txtFilePath.name, VBUtil.extractDirPathFromFilePath(txtFilePath.value))
    values.setItem Array(cboChoiceCharacterCode.name, cboChoiceCharacterCode.value)
    values.setItem Array(cboChoiceNewLine.name, cboChoiceNewLine.value)

    ' データを書き込む
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽ファイルオプションを読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreFileOutputOption()

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties

    ' データを読み込む
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
            
    val = values.getItem(txtFilePath.name, vbVariant): If IsArray(val) Then txtFilePath.value = val(2)
    val = values.getItem(cboChoiceCharacterCode.name, vbVariant): If IsArray(val) Then cboChoiceCharacterCode.value = val(2)
    val = values.getItem(cboChoiceNewLine.name, vbVariant): If IsArray(val) Then cboChoiceNewLine.value = val(2)
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽テキストボックスチェック成功時のコントロール変更処理
'
' 概要　　　：
' 引数　　　：cnt コントロール
' 戻り値　　：
'
' =========================================================
Private Sub changeControlPropertyByValidTrue(ByRef cnt As MSForms.control)

    With cnt
        .BackColor = &H80000005
        .ForeColor = &H80000012
    
    End With

End Sub

' =========================================================
' ▽テキストボックスチェック失敗時のコントロール変更処理
'
' 概要　　　：
' 引数　　　：cnt コントロール
' 戻り値　　：
'
' =========================================================
Private Sub changeControlPropertyByValidFalse(ByRef cnt As MSForms.control)

    With cnt
        ' テキスト全体を選択する
        .SelStart = 0
        .SelLength = Len(.text)
        
        .BackColor = RGB(&HFF, &HFF, &HCC)
        .ForeColor = reverseRGB(&HFF, &HFF, &HCC)
        
    End With

End Sub

