VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableSheetCreator 
   Caption         =   "テーブルシートの作成"
   ClientHeight    =   8790.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535.001
   OleObjectBlob   =   "frmTableSheetCreator.frx":0000
End
Attribute VB_Name = "frmTableSheetCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' テーブルシート作成フォーム
'
' 作成者　：Ison
' 履歴　　：2009/01/25　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽処理が完了した場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event complete(ByRef createTargetTable As ValCollection)

' =========================================================
' ▽処理がキャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

Private Const MULTIPAGE_MIN_PAGE As Long = 0
Private Const MULTIPAGE_MAX_PAGE As Long = 4

Private Const MULTIPAGE_WELCOME            As Long = 0
Private Const MULTIPAGE_CHOICE_SCHEMA      As Long = 1
Private Const MULTIPAGE_CHOICE_TABLE       As Long = 2
Private Const MULTIPAGE_SETTING_ROW_FORMAT As Long = 3
Private Const MULTIPAGE_COMPLETE           As Long = 4

Private Const ROW_FORMAT_STR_TO_UNDER As String = "↓"
Private Const ROW_FORMAT_STR_TO_RIGHT As String = "→"

' アプリケーション設定情報
Private applicationSetting As ValApplicationSetting

' DBコネクションオブジェクト
Private dbConn As Object

' -------------------------------------------------------------
' スキーマリスト
Private schemaInfoList As CntListBox
' テーブルリスト
Private tableInfoList  As CntListBox
' テーブルリスト（行フォーマット）
Private tableInfoListRowFormat As CntListBox

' 選択されたテーブルリスト
Private selectedTableList As ValCollection
' -------------------------------------------------------------

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
' 引数　　　：modal モーダルまたはモードレス表示指定
' 　　　　　　conn  DBコネクション
' 　　　　　　aps   アプリケーション設定情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef aps As ValApplicationSetting, ByRef conn As Object)

    ' アプリケーション情報を設定する
    Set applicationSetting = aps
    ' DBコネクションを設定する
    Set dbConn = conn
    ' アクティブ処理
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

    ' ディアクティブ処理
    deactivate

    Main.storeFormPosition Me.name, Me
    Me.Hide
End Sub

' =========================================================
' ▽アクティブ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()

    ' スキーマリストオブジェクトを初期化する
    Set schemaInfoList = New CntListBox: schemaInfoList.init lstSchemaList
    ' テーブルリストオブジェクトを初期化する
    Set tableInfoList = New CntListBox: tableInfoList.init lstTableList1
    ' テーブルリスト（行フォーマット）オブジェクトを初期化する
    Set tableInfoListRowFormat = New CntListBox: tableInfoListRowFormat.init lstTableListRowFormat
    
    ' 複数のスキーマを利用しない（単体のスキーマのみ参照）
    If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
    
        ' スキーマを一つしか選択できないようにする
        lstSchemaList.multiSelect = fmMultiSelectSingle
        
    ' 複数のスキーマを利用する
    Else
        
        ' スキーマを複数選択できるようにする
        lstSchemaList.multiSelect = fmMultiSelectMulti
    End If

    ' マルチページのページ番号を一番始めのページに設定する
    multiPage.value = MULTIPAGE_MIN_PAGE
    
    ' ウィザード形式のウィンドウを操作するための各ボタンのenableプロパティを設定する
    ' 1ページ目なので戻れない
    btnBack.Enabled = False
    ' 1ページ目なので戻れる
    btnNext.Enabled = True
    ' キャンセルは押下可能
    btnCancel.Enabled = True
    ' 完了は押下不可
    btnFinish.Enabled = False
    
    ' ページ毎の初期化を行う
    ' ウェルカムページ
    initPageWelcome
    ' スキーマ選択ページ
    initPageChoiceSchema
    ' テーブル選択ページ
    initPageChoiceTable
    ' 完了ページ
    initPageComplete
    
End Sub

' =========================================================
' ▽ディアクティブ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub deactivate()

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
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォームクローズ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' 本処理では処理自体をキャンセルする
        Cancel = True
        ' 以下のイベント経由で閉じる
        btnCancel_Click
    End If
    
End Sub

' =========================================================
' ▽戻るボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnBack_Click()

    On Error GoTo err

    Dim page As Long

    ' 1ページ目より前に戻ろうとした場合
    If multiPage.value - 1 <= MULTIPAGE_MIN_PAGE Then
    
        ' マルチページを1ページ目に設定
        page = MULTIPAGE_MIN_PAGE
        
        ' ページ切り替え処理
        changePage page
        
        ' 各ボタンのenableプロパティを1ページの状態に設定
        btnBack.Enabled = False
        btnNext.Enabled = True
        btnCancel.Enabled = True
        btnFinish.Enabled = False
        
    ' 1ページ以外
    Else
    
        ' マルチページを現在のページから1ページ前に設定
        page = multiPage.value - 1
        
        ' ページ切り替え処理
        changePage page
        
        ' 各ボタンのenableプロパティを1ページ以外の状態に設定
        btnBack.Enabled = True
        btnNext.Enabled = True
        btnCancel.Enabled = True
        btnFinish.Enabled = False
        
        ' マルチページを現在のページから1ページ前に設定
        page = multiPage.value - 1

    End If

    ' ページを切り替える
    multiPage.value = page

    Exit Sub
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽次へボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnNext_Click()

    On Error GoTo err

    Dim page As Long

    ' ページ切り替え前のチェック処理
    changePageBefore multiPage.value
    
    ' 最終ページ目以降に進もうとした場合
    If multiPage.value + 1 >= MULTIPAGE_MAX_PAGE Then
    
        ' マルチページを最終ページ目に設定
        page = MULTIPAGE_MAX_PAGE
        
        ' ページ切り替え処理
        changePage page
        
        ' 各ボタンのenableプロパティを最終ページの状態に設定
        btnBack.Enabled = True
        btnNext.Enabled = False
        btnCancel.Enabled = True
        btnFinish.Enabled = True
        
    ' 最終ページ以外
    Else
    
        ' マルチページを現在のページから1ページ後に設定
        page = multiPage.value + 1
        
        ' ページ切り替え処理
        changePage page
        
        ' 各ボタンのenableプロパティを最終ページ以外の状態に設定
        btnBack.Enabled = True
        btnNext.Enabled = True
        btnCancel.Enabled = True
        btnFinish.Enabled = False
        
    End If
    
    ' ページを切り替える
    multiPage.value = page

    Exit Sub
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽キャンセルボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnCancel_Click()

    On Error GoTo err
    
    ' キャンセル判定
    If checkCancel = True Then
    
        ' フォームを非表示にする
        HideExt
    
        ' イベントを発行する
        RaiseEvent Cancel
    End If
    
    
    Exit Sub
err:

    ' エラーメッセージを表示する
    Main.ShowErrorMessage
    
    ' フォームを非表示にする
    HideExt
    
End Sub

' =========================================================
' ▽キャンセル判定
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：True キャンセルする場合
'
' =========================================================
Private Function checkCancel() As Boolean

    ' メッセージボックスの戻り値
    Dim result As Long
    
    ' キャンセル確認用のメッセージボックスを表示する
    result = VBUtil.showMessageBoxForYesNo("終了してもよろしいですか？", ConstantsCommon.APPLICATION_NAME)

    If result = WinAPI_User.IDYES Then
    
        checkCancel = True
    Else
    
        checkCancel = False
    End If
    
End Function

' =========================================================
' ▽完了ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnFinish_Click()

    On Error GoTo err
    
    fixPageComplete
    
    ' 後続のイベント発行時に、新しいフォームを開いてしまうのでここで先にフォームを閉じておく
    Me.Hide
    ' イベントを発行する
    RaiseEvent complete(selectedTableList)
        
    HideExt
        
    Exit Sub
err:

    Main.ShowErrorMessage
        
    HideExt
    
End Sub

' =========================================================
' ▽マルチページのページ移動時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub multiPage_Change()

    On Error GoTo err

    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽テーブルリストのチェック状態を全てONにするボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnTableAll_Click()

    Dim i As Long
    
    ' 全て未選択にする
    For i = 0 To lstTableList1.ListCount - 1
    
        lstTableList1.selected(i) = True
    Next
    
End Sub

' =========================================================
' ▽テーブルリストのチェック状態を全てOFFにするボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnTableNon_Click()

    Dim i As Long
    
    ' 全て未選択にする
    For i = 0 To lstTableList1.ListCount - 1
    
        lstTableList1.selected(i) = False
    Next
    
End Sub

' =========================================================
' ▽行フォーマットリスト選択時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstTableListRowFormat_Change()
    
    ' テーブルリストのインデックス
    Dim i    As Long
    ' テーブルリストのサイズ
    Dim size As Long
    
    ' テーブルリストのサイズを取得する
    size = lstTableListRowFormat.ListCount
        
    ' リスト上で選択されている要素を取得する
    i = lstTableListRowFormat.ListIndex

    ' リストコントロールにて選択されているかを判定する
    If lstTableListRowFormat.selected(i) = True Then
    
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_RIGHT
    Else
    
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_UNDER
    End If
    
End Sub

' =========================================================
' ▽行フォーマットリストの全要素を↓に設定するボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnRowFormatToUnder_Click()

    ' テーブルリストのインデックス
    Dim i    As Long
    ' テーブルリストのサイズ
    Dim size As Long
    
    ' テーブルリストのサイズを取得する
    size = lstTableListRowFormat.ListCount
        
    ' リストコントロールをループさせる
    For i = 0 To size - 1
    
        lstTableListRowFormat.selected(i) = False
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_UNDER
    Next
    
End Sub

' =========================================================
' ▽行フォーマットリストの全要素を→に設定するボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub btnRowFormatToRight_Click()

    ' テーブルリストのインデックス
    Dim i    As Long
    ' テーブルリストのサイズ
    Dim size As Long
    
    ' テーブルリストのサイズを取得する
    size = lstTableListRowFormat.ListCount
        
    ' リストコントロールをループさせる
    For i = 0 To size - 1
    
        lstTableListRowFormat.selected(i) = True
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_RIGHT
    Next
    
End Sub


' =========================================================
' ▽マルチページを切り替える前に呼び出す処理
'
' 概要　　　：
' 引数　　　：pageIndex ページ番号
' 戻り値　　：
'
' =========================================================
Private Sub changePageBefore(ByVal pageIndex As Long)

    If pageIndex = MULTIPAGE_WELCOME Then
    
        ' 検証処理を行う
        validPageWelcome
    
    ElseIf pageIndex = MULTIPAGE_CHOICE_SCHEMA Then
    
        ' 検証処理を行う
        validPageChoiceSchema
        
    ElseIf pageIndex = MULTIPAGE_CHOICE_TABLE Then
    
        ' 検証処理を行う
        validPageChoiceTable
        
    ElseIf pageIndex = MULTIPAGE_SETTING_ROW_FORMAT Then
    
        ' 検証処理を行う
        validPageSettingRowFormat
    
    ElseIf pageIndex = MULTIPAGE_COMPLETE Then
    
        ' 検証処理を行う
        validPageComplete
        
    End If

End Sub

' =========================================================
' ▽マルチページを切り替える処理
'
' 概要　　　：
' 引数　　　：pageIndex ページ番号
' 戻り値　　：
'
' =========================================================
Private Sub changePage(ByVal pageIndex As Long)

    If pageIndex = MULTIPAGE_WELCOME Then
    
        ' 表示処理を行う
        activatePageWelcome
        
    ElseIf pageIndex = MULTIPAGE_CHOICE_SCHEMA Then
    
        ' 表示処理を行う
        activatePageChoiceSchema
        ' 完了処理を行う
        fixPageWelcome
        
    ElseIf pageIndex = MULTIPAGE_CHOICE_TABLE Then
    
        ' 完了処理を行う
        fixPageChoiceSchema
        ' 表示処理を行う
        activatePageChoiceTable
        
    ElseIf pageIndex = MULTIPAGE_SETTING_ROW_FORMAT Then
    
        ' 完了処理を行う
        fixPageChoiceTable
        ' 表示処理を行う
        activatePageSettingRowFormat
    
    ElseIf pageIndex = MULTIPAGE_COMPLETE Then
    
        ' 完了処理を行う
        fixPageSettingRowFormat
        ' 表示処理を行う
        activatePageComplete
        
    End If

End Sub

Private Sub initPageWelcome()

End Sub

Private Sub initPageChoiceSchema()

    On Error GoTo err

    ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
    Dim cursorWait As New ExcelCursorWait: cursorWait.init

    ' 一時変数
    Dim var As ValCollection
    
    Dim dbObjFactory As New DbObjectFactory
    
    Dim dbInfo As IDbMetaInfoGetter
    Set dbInfo = dbObjFactory.createMetaInfoGetterObject(dbConn)
    
    ' 長時間の処理が終了したのでマウスカーソルを元に戻す
    cursorWait.destroy
    
    Set var = dbInfo.getSchemaList
    
    ' スキーマリストボックスにリストを追加する
    schemaInfoList.addAll var, "SchemaName", "SchemaComment"
        
    Exit Sub
    
err:

    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

Private Sub initPageChoiceTable()

End Sub

Private Sub initPageComplete()

End Sub

Private Sub validPageWelcome()

End Sub

Private Sub validPageChoiceSchema()

    Dim cnt As Long
    cnt = schemaInfoList.getSelectedList().count
    
    ' スキーマリストでの選択件数を確認する
    If cnt <= 0 Then
    
        ' 0件の場合エラーを発行する
        err.Raise ERR_NUMBER_NOT_SELECTED_SCHEMA _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_SCHEMA _
                , err.HelpFile _
                , err.HelpContext
    End If

End Sub

Private Sub validPageChoiceTable()

    Dim cnt As Long
    cnt = tableInfoList.getSelectedList().count
    
    ' テーブルリストでの選択件数を確認する
    If cnt <= 0 Then
    
        ' 0件の場合エラーを発行する
        err.Raise ERR_NUMBER_NOT_SELECTED_TABLE _
                , err.Source _
                , ERR_DESC_NOT_SELECTED_TABLE _
                , err.HelpFile _
                , err.HelpContext
    End If

End Sub

Private Sub validPageSettingRowFormat()

End Sub

Private Sub validPageComplete()

End Sub

Private Sub activatePageWelcome()

End Sub

Private Sub activatePageChoiceSchema()

End Sub

Private Sub activatePageChoiceTable()

    On Error GoTo err

    ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
    Dim cursorWait As New ExcelCursorWait: cursorWait.init

    ' 一時変数
    Dim var  As ValCollection
    
    Dim selectedSchema As ValCollection
    Set selectedSchema = schemaInfoList.getSelectedList(vbObject)
    
    Dim dbObjFactory As New DbObjectFactory
    
    Dim dbInfo As IDbMetaInfoGetter
    Set dbInfo = dbObjFactory.createMetaInfoGetterObject(dbConn)
    
    Set var = dbInfo.getTableList(selectedSchema)
    
    ' テーブルリストボックスにリストを追加する
    If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
        tableInfoList.addAll var, "TableName", "TableComment"
    Else
        tableInfoList.addAll var, "SchemaTableName", "TableComment"
    End If
    
    ' 長時間の処理が終了したのでマウスカーソルを元に戻す
    cursorWait.destroy
    
    Exit Sub
    
err:

    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

Private Sub activatePageSettingRowFormat()

    ' 選択済みテーブルリスト
    Dim selectedTableList As ValCollection
    ' 選択済みテーブル
    Dim selectedTable     As ValDbDefineTable
    
    ' 選択済みテーブルリストを取得する
    Set selectedTableList = tableInfoList.getSelectedList
    
    ' テーブルリストボックスにリストを追加する
    If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
        tableInfoListRowFormat.addAll selectedTableList _
                                    , "TableName"
    Else
        tableInfoListRowFormat.addAll selectedTableList _
                                    , "SchemaTableName"
    End If
    
    ' テーブル定義の行フォーマットの状態をコントロールに反映する
    ' ↓の場合、リストを未選択
    ' →の場合、リストを選択
    
    ' テーブルリストのインデックス
    Dim i    As Long
    ' テーブルリストのサイズ
    Dim size As Long
    
    ' テーブルリストのサイズを取得する
    size = lstTableListRowFormat.ListCount
        
    ' リストコントロールをループさせる
    For i = 0 To size - 1
    
        lstTableListRowFormat.list(i, 1) = ROW_FORMAT_STR_TO_UNDER
        lstTableListRowFormat.selected(i) = False
        
    Next
    

End Sub

Private Sub activatePageComplete()

End Sub

Private Sub fixPageWelcome()

End Sub

Private Sub fixPageChoiceSchema()

End Sub

Private Sub fixPageChoiceTable()

End Sub

Private Sub fixPageSettingRowFormat()

End Sub

Private Sub fixPageComplete()

    On Error GoTo err

    ' テーブルリストのインデックス
    Dim i    As Long
    
    ' 一時変数
    Dim var    As ValCollection
    Dim varObj As ValDbDefineTable
    
    Dim tableSheetList As New ValCollection
    Dim tableSheet     As ValTableWorksheet
    
    Set var = tableInfoListRowFormat.collection
    
    ' Tableに設定されているスキーマ名をクリアする
    For Each varObj In var.col
    
        Set tableSheet = New ValTableWorksheet
        Set tableSheet.table = varObj
        tableSheet.recFormat = recFormatToUnder
        
        ' 複数のスキーマを利用しない（単体のスキーマのみ参照）
        If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
        
            tableSheet.omitsSchema = True
        Else
            tableSheet.omitsSchema = False
        End If
        
        ' リストコントロールにて選択されているかを判定する
        If lstTableListRowFormat.selected(i) = True Then
        
            tableSheet.recFormat = REC_FORMAT.recFormatToRight
        Else
        
            tableSheet.recFormat = REC_FORMAT.recFormatToUnder
        End If
        
        tableSheetList.setItem tableSheet, varObj.schemaTableName
        
        i = i + 1
    Next
    
    Set selectedTableList = tableSheetList
    
    Exit Sub
    
err:

    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub
