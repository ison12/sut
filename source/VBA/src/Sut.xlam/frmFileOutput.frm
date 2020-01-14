VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileOutput 
   Caption         =   "ファイル出力"
   ClientHeight    =   3495
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
' 作成者　：Hideki Isobe
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
Public Event Cancel()

Private Const NEW_LINE_STR_CRLF As String = "CRLF"
Private Const NEW_LINE_STR_CR As String = "CR"
Private Const NEW_LINE_STR_LF As String = "LF"

Private Const REG_SUB_KEY_FILE_OUTPUT_OPTION As String = "file_output_option"

' レジストリパス - 文字コード一覧
Private Const REG_PATH_CHARACTER_CODE_LIST As String = "MIME\Database\Charset"
' レジストリキー - 文字コードの別名
Private Const REG_KEY_ALIAS_CHARSET As String = "AliasForCharset"

' 文字コードリスト
Private charcterList As CntListBox

' デフォルトファイル名
Private defaultFileName As String

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
' ▽文字コードリスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboChoiceCharacterCode_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' コレクション
    Dim col As ValCollection
    ' コントロールからコレクションを取得する
    Set col = charcterList.collection

    ' リストに現在入力されているテキストの要素が存在しない場合
    If col.exist(cboChoiceCharacterCode.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NO_LIST_ITEM
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse cboChoiceCharacterCode
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidTrue cboChoiceCharacterCode
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

    
End Sub

' =========================================================
' ▽改行コードリスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboChoiceNewLine_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' コレクション
    Dim col As New ValCollection

    col.setItem NEW_LINE_STR_CRLF, NEW_LINE_STR_CRLF
    col.setItem NEW_LINE_STR_CR, NEW_LINE_STR_CR
    col.setItem NEW_LINE_STR_LF, NEW_LINE_STR_LF
    
    ' リストに現在入力されているテキストの要素が存在しない場合
    If col.exist(cboChoiceNewLine.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NO_LIST_ITEM
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse cboChoiceNewLine
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidTrue cboChoiceNewLine
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

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

    ' 最前面表示にする
    ExcelUtil.setUserFormTopMost Me

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

    ' ファイルパスを取得
    filePath = txtFilePath.text
    ' 文字コードを取得
    characterCode = cboChoiceCharacterCode.text
    ' 改行コードを取得
    newline = cboChoiceNewLine.text
    
    ' ファイルパスの親ディレクトリを取得する
    dirPath = VBUtil.extractDirPathFromFilePath(filePath)
    ' 親フォルダを作成する
    VBUtil.createDir dirPath
    
    ' ファイルパスのディレクトリが存在するかを確認する
    If VBUtil.isExistDirectory(dirPath) = False Then
    
        VBUtil.showMessageBoxForWarning "指定されたファイルパスのフォルダが見つかりません。" _
                                      , ConstantsCommon.APPLICATION_NAME _
                                      , Nothing
        
        Exit Sub
    End If
    
    ' フォームを閉じる
    HideExt
    
    ' OKイベントを送信する
    RaiseEvent ok(filePath, characterCode, convertNewLineStrToNewLineCode(cboChoiceNewLine.text))
    
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
    RaiseEvent Cancel

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
    
    cboChoiceCharacterCode.value = "shift_jis"
    cboChoiceNewLine.ListIndex = 0

    ' ファイル出力オプションを読み込む
    restoreFileOutputOption
    
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
    
    ' ファイルパスにデフォルトのファイル名を設定する
    txtFilePath.value = VBUtil.concatFilePath( _
                                    VBUtil.extractDirPathFromFilePath(txtFilePath.value) _
                                  , defaultFileName)
    
    
    ' エラーメッセージをクリアする
    lblErrorMessage.Caption = ""
    
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

Private Function convertNewLineStrToNewLineCode(ByVal newLineStr As String) As String

    If newLineStr = NEW_LINE_STR_CRLF Then
    
        ' Windows
        convertNewLineStrToNewLineCode = vbCr & vbLf
    
    ElseIf newLineStr = NEW_LINE_STR_CR Then
    
        ' Mac
        convertNewLineStrToNewLineCode = vbCr
    
    ElseIf newLineStr = NEW_LINE_STR_LF Then
    
        ' Unix
        convertNewLineStrToNewLineCode = vbLf
        
    ' 当てはまらない場合
    Else
    
        ' Windows
        convertNewLineStrToNewLineCode = vbCr & vbLf
    
    End If

End Function

' =========================================================
' ▽テキストボックスチェック成功時のコントロール変更処理
'
' 概要　　　：
' 引数　　　：cnt コントロール
' 戻り値　　：
'
' =========================================================
Public Sub changeControlPropertyByValidTrue(ByRef cnt As MSForms.control)

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
Public Sub changeControlPropertyByValidFalse(ByRef cnt As MSForms.control)

    With cnt
        ' テキスト全体を選択する
        .SelStart = 0
        .SelLength = Len(.text)
        
        .BackColor = RGB(&HFF, &HFF, &HCC)
        .ForeColor = reverseRGB(&HFF, &HFF, &HCC)
        
    End With

End Sub


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
    
    Dim j As Long
    
    Dim fileOutputOption(0 To 2 _
                       , 0 To 1) As Variant
    
    
    fileOutputOption(j, 0) = txtFilePath.name
    fileOutputOption(j, 1) = VBUtil.extractDirPathFromFilePath(txtFilePath.value): j = j + 1
    
    fileOutputOption(j, 0) = cboChoiceCharacterCode.name
    fileOutputOption(j, 1) = cboChoiceCharacterCode.value: j = j + 1

    fileOutputOption(j, 0) = cboChoiceNewLine.name
    fileOutputOption(j, 1) = cboChoiceNewLine.value: j = j + 1
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_FILE_OUTPUT_OPTION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' レジストリに情報を設定する
    registry.setValues fileOutputOption
    
    Set registry = Nothing
        
    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet
    
    bookProp.setValue ConstantsBookProperties.TABLE_FILE_OUTPUT_DIALOG, txtFilePath.name, VBUtil.extractDirPathFromFilePath(txtFilePath.value)
    bookProp.setValue ConstantsBookProperties.TABLE_FILE_OUTPUT_DIALOG, cboChoiceCharacterCode.name, cboChoiceCharacterCode.value
    bookProp.setValue ConstantsBookProperties.TABLE_FILE_OUTPUT_DIALOG, cboChoiceNewLine.name, cboChoiceNewLine.value
    ' ----------------------------------------------

    Exit Sub
    
err:
    
    Set registry = Nothing

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
    
    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' 設定情報シートが存在する
        
        Set bookPropVal = bookProp.getValues(ConstantsBookProperties.TABLE_FILE_OUTPUT_DIALOG)
        If bookPropVal.count > 0 Then
            ' 設定情報が存在するので、フォームに反映する
            
            txtFilePath.value = bookPropVal.getItem(txtFilePath.name, vbString)
            cboChoiceCharacterCode.value = bookPropVal.getItem(cboChoiceCharacterCode.name, vbString)
            cboChoiceNewLine.value = bookPropVal.getItem(cboChoiceNewLine.name, vbString)

            Exit Sub
        End If
    End If
    ' ----------------------------------------------

    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_FILE_OUTPUT_OPTION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Dim retFilepath As String
    Dim retChar     As String
    Dim retNewLine  As String
    
    registry.getValue txtFilePath.name, retFilepath
    registry.getValue cboChoiceCharacterCode.name, retChar
    registry.getValue cboChoiceNewLine.name, retNewLine
    
    txtFilePath.value = retFilepath
    cboChoiceCharacterCode.value = retChar
    cboChoiceNewLine.value = retNewLine
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub
