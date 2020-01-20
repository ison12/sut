VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBQueryBatch 
   Caption         =   "クエリ一括実行"
   ClientHeight    =   9795.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13905
   OleObjectBlob   =   "frmDBQueryBatch.frx":0000
End
Attribute VB_Name = "frmDBQueryBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' クエリ一括実行フォーム
'
' 作成者　：Hideki Isobe
' 履歴　　：2020/01/18　新規作成
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
Public Event ok(ByVal dbQueryBatchMode As DB_QUERY_BATCH_MODE _
              , ByVal filePath As String _
              , ByVal characterCode As String _
              , ByVal newline As String _
              , ByVal tableWorksheets As ValCollection)

' =========================================================
' ▽キャンセルボタン押下時に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' DBクエリバッチモード
Public Enum DB_QUERY_BATCH_MODE

    ' ファイル出力
    FileOutput
    ' クエリ実行
    QueryExecute

End Enum

Private Const REG_SUB_KEY_DB_QUERY_BATCH_OPTION As String = "db_query_batch"

' DBクエリバッチのクエリ種類の一件毎の編集（子画面）
Private WithEvents frmDBQueryBatchTypeSettingVar As frmDBQueryBatchTypeSetting
Attribute frmDBQueryBatchTypeSettingVar.VB_VarHelpID = -1

' テーブルリストでの選択項目インデックス
Private tableSheetSelectedIndex As Long
' テーブルリストでの選択項目オブジェクト
Private tableSheetSelectedItem As ValDbQueryBatchTableWorksheet

' DBクエリバッチモード
Private dbQueryBatchMode As DB_QUERY_BATCH_MODE
' DBクエリバッチ種類
Private dbQueryBatchType As DB_QUERY_BATCH_TYPE
' 処理対象ワークブック
Private book As Workbook

' 文字コードリスト
Private charcterList As CntListBox
' DBクエリバッチ種類変更コンボボックスリスト
Private dbQueryBatchTypeChangeAll As CntListBox
' DBクエリバッチ種類変更コンボボックスの処理中
Private inProcessDbQueryBatchTypeChangeAll As Boolean

' テーブルリスト
Private tableSheetList  As CntListBox

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal  モーダルまたはモードレス表示指定
' 　　　　　　mode   モード
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByVal dbQueryBatchMode_ As DB_QUERY_BATCH_MODE _
                 , ByVal dbQueryBatchType_ As DB_QUERY_BATCH_TYPE _
                 , ByRef book_ As Workbook)

    dbQueryBatchMode = dbQueryBatchMode_
    dbQueryBatchType = dbQueryBatchType_
    Set book = book_

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
' ▽全ての選択肢を選択済みにするボタンのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdSelectedAll_Click()

    tableSheetList.setSelectedAll True

End Sub

' =========================================================
' ▽全ての選択肢を選択解除にするボタンのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdUnselectedAll_Click()

    tableSheetList.setSelectedAll False

End Sub

' =========================================================
' ▽全てのDBクエリバッチ種類を変更するコンボボックスリストのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboDbQueryBatchTypeChangeAll_Change()

    On Error GoTo err

    If inProcessDbQueryBatchTypeChangeAll = True Then
        ' 既に処理中の場合は処理を終了する
        Exit Sub
    End If

    inProcessDbQueryBatchTypeChangeAll = True

    Dim i As Long
    Dim var As ValDbQueryBatchTableWorksheet
    
    Dim selectedDbQueryBatchType As ValDbQueryBatchType
    
    i = 0
    For Each var In tableSheetList.collection.col
    
        Set selectedDbQueryBatchType = dbQueryBatchTypeChangeAll.getSelectedItem
        var.dbQueryBatchType = selectedDbQueryBatchType.dbQueryBatchType
        
        setTableSheet i, var
        
        i = i + 1
    
    Next
    
    ' 処理の最後に未選択状態に戻す
    dbQueryBatchTypeChangeAll.setSelectedIndex 0

    inProcessDbQueryBatchTypeChangeAll = False
    
    Exit Sub
err:

    inProcessDbQueryBatchTypeChangeAll = False
    
End Sub

' =========================================================
' ▽DBクエリバッチ種類を変更するボタンのイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmbDbQueryBatchTypeChange_Click()

    ' 現在選択されているインデックスを取得
    tableSheetSelectedIndex = tableSheetList.getSelectedIndex

    ' 未選択の場合
    If tableSheetSelectedIndex = -1 Then
    
        ' 終了する
        Exit Sub
    End If

    ' 現在選択されている項目を取得
    Set tableSheetSelectedItem = tableSheetList.getSelectedItem

    Load frmDBQueryBatchTypeSetting
    Set frmDBQueryBatchTypeSettingVar = frmDBQueryBatchTypeSetting
    
    frmDBQueryBatchTypeSettingVar.ShowExt vbModal _
                        , tableSheetSelectedItem.sheetNameOrSheetTableName _
                        , tableSheetSelectedItem.dbQueryBatchType _
                        , dbQueryBatchTypeChangeAll.collection
    
    Set frmDBQueryBatchTypeSettingVar = Nothing

End Sub

' =========================================================
' ▽DBクエリバッチ種類を変更の確定時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmDBQueryBatchTypeSettingVar_ok(ByVal dbQueryBatchType As DB_QUERY_BATCH_TYPE)

    tableSheetSelectedItem.dbQueryBatchType = dbQueryBatchType
    
    setTableSheet tableSheetSelectedIndex, tableSheetSelectedItem
    
End Sub

' =========================================================
' ▽DBクエリバッチ種類を変更のキャンセル時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmDBQueryBatchTypeSettingVar_cancel()

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
    
    selectFile = openFolderDialog
    
    If selectFile <> "" Then
        ' ファイルを開くダイアログをオープンしてユーザにファイルを選択させる
        txtFilePath.text = selectFile
    End If
    
End Sub

' =========================================================
' ▽フォルダを開くダイアログオープン
'
' 概要　　　：フォルダを開くダイアログをオープンする
'
' =========================================================
Private Function openFolderDialog() As String

    On Error GoTo err
            ' 選択ファイル
    Dim selectFile As String
    
    ' 開くダイアログを選択する
    selectFile = VBUtil.openFolderDialog("ファイル出力先フォルダを選択してください。" _
                                         , txtFilePath.value)

    ' ファイルパスを設定する
    openFolderDialog = selectFile
    
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

    ' ファイル出力時のみの処理
    If dbQueryBatchMode = FileOutput Then
        ' ファイルパスを取得
        filePath = txtFilePath.text
        ' 文字コードを取得
        characterCode = cboChoiceCharacterCode.text
        ' 改行コードを取得
        newline = cboChoiceNewLine.text
        
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
        If isSuccessCreateDir = False Or VBUtil.touch(filePath) = False Then
        
            VBUtil.showMessageBoxForWarning "指定されたフォルダパスにファイルが出力できません。" & vbNewLine & "不正なパス、または権限が不足している可能性があります。" _
                                          , ConstantsCommon.APPLICATION_NAME _
                                          , Nothing
            
            Exit Sub
        End If
        
    End If
    
    ' フォームを閉じる
    HideExt
    
    ' OKイベントを送信する
    RaiseEvent ok(dbQueryBatchMode, filePath, characterCode, VBUtil.convertNewLineStrToNewLineCode(cboChoiceNewLine.text), tableSheetList.selectedList)
    
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

    ' コントロールの状態を制御する
    If dbQueryBatchMode = FileOutput Then
        ' ファイル出力時
        lblFilePath.visible = True
        txtFilePath.visible = True
        lblChoiceCharacterCode.visible = True
        cboChoiceCharacterCode.visible = True
        lblChoiceNewLine.visible = True
        cboChoiceNewLine.visible = True
        btnFileSelect.visible = True
    Else
        ' DB実行時
        lblFilePath.visible = False
        txtFilePath.visible = False
        lblChoiceCharacterCode.visible = False
        cboChoiceCharacterCode.visible = False
        lblChoiceNewLine.visible = False
        cboChoiceNewLine.visible = False
        btnFileSelect.visible = False
    End If
    
    ' DBバッチクエリ種類リストに選択肢を追加する
    Set dbQueryBatchTypeChangeAll = New CntListBox
    dbQueryBatchTypeChangeAll.init cboDbQueryBatchTypeChangeAll
    
    Dim dbBatchQueryTypeRawList As New ValCollection
    Dim dbBatchQueryType As ValDbQueryBatchType
    
    If dbQueryBatchMode = FileOutput Then
        ' ファイル出力
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.none: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.insert: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.update: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.deleteOnSheet: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
    
    Else
        ' クエリ実行
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.none: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.insertUpdate: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.insert: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.update: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.deleteOnSheet: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.deleteAll: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.selectAll: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.selectCondition: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
        Set dbBatchQueryType = New ValDbQueryBatchType: dbBatchQueryType.dbQueryBatchType = DB_QUERY_BATCH_TYPE.selectReExec: dbBatchQueryType.dbQueryBatchTypeName = ConstantsEnum.getDbQueryBatchTypeName(dbBatchQueryType.dbQueryBatchType): dbBatchQueryTypeRawList.setItem dbBatchQueryType
    End If
    
    dbQueryBatchTypeChangeAll.addAll dbBatchQueryTypeRawList, "dbQueryBatchTypeName"
    dbQueryBatchTypeChangeAll.setSelectedIndex 0
    
    ' ファイル出力オプションを読み込む
    restoreFileOutputOption
    
    ' テーブルシートを読み込む
    readTableSheet
    
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
    
    If dbQueryBatchMode <> FileOutput Then
        ' ファイル出力モードではない場合
        Exit Sub
    End If
    
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
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_QUERY_BATCH_OPTION) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' レジストリに情報を設定する
    registry.setValues fileOutputOption
    
    Set registry = Nothing
        
    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet
    
    bookProp.setValue ConstantsBookProperties.TABLE_DB_QUERY_BATCH_DIALOG, txtFilePath.name, VBUtil.extractDirPathFromFilePath(txtFilePath.value)
    bookProp.setValue ConstantsBookProperties.TABLE_DB_QUERY_BATCH_DIALOG, cboChoiceCharacterCode.name, cboChoiceCharacterCode.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_QUERY_BATCH_DIALOG, cboChoiceNewLine.name, cboChoiceNewLine.value
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
        
    If dbQueryBatchMode <> FileOutput Then
        ' ファイル出力モードではない場合
        Exit Sub
    End If

    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' 設定情報シートが存在する
        
        Set bookPropVal = bookProp.getValues(ConstantsBookProperties.TABLE_DB_QUERY_BATCH_DIALOG)
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
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_QUERY_BATCH_OPTION) _
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

' =========================================================
' ▽テーブルシートを読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub readTableSheet()

    ' テーブルリスト
    Dim tableList As ValCollection
    Dim tableWorksheet As ValTableWorksheet
    
    Dim dbQueryBatchTableWorksheet As ValDbQueryBatchTableWorksheet
    
    ' テーブルシート読込オブジェクト
    Dim tableSheetReader As ExeTableSheetReader
    Set tableSheetReader = New ExeTableSheetReader
        
    ' シート
    Dim sheet As Worksheet
    
    ' テーブルリストを初期化する
    Set tableList = New ValCollection
    
    ' ブックに含まれているシートを1件ずつ処理する
    For Each sheet In book.Worksheets
    
        Set tableSheetReader.sheet = sheet
        
        ' 対象シートがテーブルシートの場合
        If tableSheetReader.isTableSheet = True Then
        
            ' テーブルシートを読み込んでリストに設定する（テーブル情報のみ取得する）
            Set tableWorksheet = tableSheetReader.readTableInfo(True)
            
            Set dbQueryBatchTableWorksheet = New ValDbQueryBatchTableWorksheet
            dbQueryBatchTableWorksheet.dbQueryBatchType = dbQueryBatchTypeChangeAll.getItem(1).dbQueryBatchType
            Set dbQueryBatchTableWorksheet.tableWorksheet = tableWorksheet
            
            tableList.setItem dbQueryBatchTableWorksheet
        End If
    
    Next
    
    ' リストコントロールにテーブルシート情報を追加する
    Set tableSheetList = New CntListBox: tableSheetList.init lstTableSheet
    tableSheetList.removeAll
    addTableSheetList tableList
    
End Sub

' =========================================================
' ▽テーブルシートリストを追加
'
' 概要　　　：
' 引数　　　：valTableSheetList テーブルシートリスト
'     　　　  isAppend              追加有無フラグ
' 戻り値　　：
'
' =========================================================
Private Sub addTableSheetList(ByVal valTableSheetList As ValCollection, Optional ByVal isAppend As Boolean = True)
    
    tableSheetList.addAll valTableSheetList _
                       , "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ▽テーブルシートを追加
'
' 概要　　　：
' 引数　　　：tableSheet テーブルシート
' 戻り値　　：
'
' =========================================================
Private Sub addTableSheet(ByVal tableSheet As ValDbQueryBatchTableWorksheet)
    
    tableSheetList.addItemByProp tableSheet, "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName"
    
End Sub

' =========================================================
' ▽テーブルシートを変更
'
' 概要　　　：
' 引数　　　：index インデックス
'     　　　  rec   テーブルシート
' 戻り値　　：
'
' =========================================================
Private Sub setTableSheet(ByVal index As Long, ByVal rec As ValDbQueryBatchTableWorksheet)
    
    tableSheetList.setItem index, rec, "sheetNameOrSheetTableName", "tableComment", "dbQueryBatchTypeName"
    
End Sub
