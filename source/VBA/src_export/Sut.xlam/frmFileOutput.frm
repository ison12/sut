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

Private Const REG_SUB_KEY_FILE_OUTPUT_OPTION As String = "file_output_option"

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
    
        VBUtil.showMessageBoxForWarning "指定されたファイルパスにファイルが出力できません。" & vbNewLine & "不正なパス、または権限が不足している可能性があります。" _
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
    
    cboChoiceCharacterCode.value = "shift_jis"
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
    txtFilePath.value = VBUtil.concatFilePath( _
                                    VBUtil.extractDirPathFromFilePath(txtFilePath.value) _
                                  , defaultFileName)
    
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
    If txtFilePath.value = "" Then txtFilePath.value = ThisWorkbook.path
    
    cboChoiceCharacterCode.value = retChar
    If cboChoiceCharacterCode.ListIndex = -1 Then cboChoiceCharacterCode.ListIndex = 0
    
    cboChoiceNewLine.value = retNewLine
    If cboChoiceNewLine.ListIndex = -1 Then cboChoiceNewLine.ListIndex = 0
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub
