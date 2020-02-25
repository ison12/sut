VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOption 
   Caption         =   "オプション"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9030.001
   OleObjectBlob   =   "frmOption.frx":0000
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' オプション設定を行うフォーム
'
' 作成者　：Ison
' 履歴　　：2009/03/14　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽決定した際に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：applicationSetting アプリケーション設定情報
'
' =========================================================
Public Event ok(ByRef applicationSetting As ValApplicationSetting)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' カラム書式設定
Private WithEvents frmDBColumnFormatVar As frmDBColumnFormat
Attribute frmDBColumnFormatVar.VB_VarHelpID = -1

' アプリケーション設定情報
Private applicationSetting As ValApplicationSetting
' アプリケーション設定情報（カラム書式）
Private applicationSettingColFmt As ValApplicationSettingColFormat

' フォントリスト コントロール
Private fontList As CntListBox
' フォントサイズリスト コントロール
Private fontSizeList As CntListBox

' カラム書式を設定中のDB
Private settingColFormatDb As DbmsType

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
' 　　　　　　var   アプリケーション設定情報
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByRef var As ValApplicationSetting _
                 , ByRef var2 As ValApplicationSettingColFormat)

    ' メンバ変数にアプリケーション設定情報を設定する
    Set applicationSetting = var
    Set applicationSettingColFmt = var2
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
' ▽フォームアクティブ時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()

    ' 前回最後に設定した情報をフォーム上の各コントロールに復元させる
    ' 読み込みを行う
    applicationSetting.readForData
    restoreOptionInfo applicationSetting
    
    ' エラーメッセージをクリアする
    lblErrorMessage.Caption = ""
    ' ブックタイトルを設定する
    lblBookTitle.Caption = replace(lblBookTitle.Tag, "${book}", targetBook.name)

End Sub

' =========================================================
' ▽フォームディアクティブ時のイベントプロシージャ
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
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' 本処理では処理自体をキャンセルする
        Cancel = True
        ' 以下のイベント経由で閉じる
        cmdCancel_Click
    End If
    
End Sub

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
    
    ' 情報を記録する
    storeOptionInfo
    
    ' フォームを閉じる
    HideExt
    
    ' OKイベントを送信する
    RaiseEvent ok(applicationSetting)
    
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
' ▽デフォルトボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDefault_Click()

    On Error GoTo err
    
    Dim applicationSetting As New ValApplicationSetting
    ' デフォルト値を反映させたアプリケーションデータでコントロールに反映する
    restoreOptionInfo applicationSetting

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

    ' カラム書式設定を初期化する
    If VBUtil.unloadFormIfChangeActiveBook(frmDBColumnFormat) Then Unload frmDBColumnFormat
    Load frmDBColumnFormat
    Set frmDBColumnFormatVar = frmDBColumnFormat
    
    ' フォントリストを初期化する
    Set fontList = New CntListBox: fontList.init cboFontList
    ' フォントサイズリストを初期化する
    Set fontSizeList = New CntListBox: fontSizeList.init cboFontSizeList

    ' フォントリストにExcelで利用可能なフォントを格納する
    fontList.addAll WinAPI_GDI.getFontNameList
    ' フォントサイズリストにExcelのフォントサイズの規定値を格納する
    fontSizeList.addAll ExcelUtil.getFontSizeList
    
    MultiPageGlobalOrBook.value = 0
    MultiPageGlobalOrBook.Pages.item("PageGlobalSetting").MultiPageAll.value = 0
    MultiPageGlobalOrBook.Pages.item("PageBookSetting").MultiPageBook.value = 0

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

    ' カラム書式設定を破棄する
    Set frmDBColumnFormatVar = Nothing
    
End Sub

' =========================================================
' ▽オプション情報を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeOptionInfo()

    With applicationSetting
    
        ' レコード処理単位
        If optRecProcessCountAll.value = True Then
            .recProcessCount = .REC_PROCESS_COUNT_ALL
            .recProcessCountCustom = txtRecProcessCountCustom.value
        Else
            .recProcessCount = .REC_PROCESS_COUNT_COSTOM
            .recProcessCountCustom = txtRecProcessCountCustom.value
        End If
        
        ' コミット確認
        If optCommitConfirmNo.value = True Then
            .commitConfirm = .COMMIT_CONFIRM_NO
        Else
            .commitConfirm = .COMMIT_CONFIRM_YES
        End If
        
        ' SQLエラー時の挙動
        If optSqlErrorHandlingSuspend.value = True Then
            .sqlErrorHandling = .SQL_ERROR_HANDLING_SUSPEND
        Else
            .sqlErrorHandling = .SQL_ERROR_HANDLING_RESUME
        End If
        
        ' スキーマの指定
        If optSchemaUseOne.value = True Then
            .schemaUse = .SCHEMA_USE_ONE
        Else
            .schemaUse = .SCHEMA_USE_MULTIPLE
        End If
        
        ' 正常時のクエリ結果表示有無
        If optQueryResultShowWhenNormalNo.value = True Then
            .queryResultShowWhenNormal = False
        Else
            .queryResultShowWhenNormal = True
        End If
        
        ' フォント名
        .cellFontName = cboFontList.value
        ' フォントサイズ
        .cellFontSize = cboFontSizeList.value
        ' 折り返し有無
        If optWordWrapYes.value = True Then
            .cellWordwrap = True
        Else
            .cellWordwrap = False
        End If
        ' セル幅
        .cellWidth = CDec(txtCellWidth.value)
        ' セル高さ
        .cellHeight = CDec(txtCellHeight.value)
        ' 行高の自動調整
        If optLineHeightAutoAdjustNo.value = True Then
            .lineHeightAutoAdjust = False
        Else
            .lineHeightAutoAdjust = True
        End If
        
        ' 空白セル読み取り方式
        If optEmptyCellReadingDel.value = True Then
            .emptyCellReading = .EMPTY_CELL_READING_DEL
        ElseIf optEmptyCellReadingNonDel.value = True Then
            .emptyCellReading = .EMPTY_CELL_READING_NON_DEL
        Else
            .emptyCellReading = .EMPTY_CELL_READING_NON_DEL_STR_EMPTY
        End If
        
        ' 直接入力文字
        If optDirectInputCharDisable.value = True Then
            .directInputChar = .DIRECT_INPUT_CHAR_DISABLE
        Else
            .directInputChar = .DIRECT_INPUT_CHAR_ENABLE_CUSTOM
        End If
        .directInputCharCustomPrefix = txtDirectInputCharEnableCustomPrefix.value
        .directInputCharCustomSuffix = txtDirectInputCharEnableCustomSuffix.value
        
        ' クエリパラメータの囲み文字
        .queryParameterEncloseCustomPrefix = txtQueryParameterEncloseEnableCustomPrefix.value
        .queryParameterEncloseCustomSuffix = txtQueryParameterEncloseEnableCustomSuffix.value
        
        ' NULL入力文字
        If optNullInputCharDisable.value = True Then
            .nullInputChar = .NULL_INPUT_CHAR_DISABLE
        Else
            .nullInputChar = .NULL_INPUT_CHAR_ENABLE_CUSTOM
        End If
        .nullInputCharCustom = txtNullInputCharEnableCustom.text
        
        ' SELECT時のセルの最大文字数チェック
        If optSelectCheckCellMaxLengthDisable.value = True Then
            .selectCheckCellMaxLength = False
        Else
            .selectCheckCellMaxLength = True
        End If
        
        ' テーブル・カラム名エスケープ
        .tableColumnEscapeOracle = chkTableColumnEscapeOracle.value
        .tableColumnEscapeMysql = chkTableColumnEscapeMysql.value
        .tableColumnEscapePostgresql = chkTableColumnEscapePostgresql.value
        .tableColumnEscapeSqlserver = chkTableColumnEscapeSqlserver.value
        .tableColumnEscapeAccess = chkTableColumnEscapeAccess.value
        .tableColumnEscapeSymfoware = chkTableColumnEscapeSymfoware.value
        
        ' 書き込みを行う
        .writeForData
    
    End With
End Sub

' =========================================================
' ▽オプション情報を読み込む
'
' 概要　　　：
' 引数　　　：applicationsetting アプリケーションデータ
' 戻り値　　：
'
' =========================================================
Private Sub restoreOptionInfo(ByRef applicationSetting As ValApplicationSetting)

    ' アプリケーション設定オブジェクト
    With applicationSetting
        
        ' レコード処理単位
        If .recProcessCount = .REC_PROCESS_COUNT_ALL Then
            optRecProcessCountAll.value = True
        Else
            optRecProcessCountCustom.value = True
        End If
        txtRecProcessCountCustom.value = .recProcessCountCustom
        
        ' コミット確認
        If .commitConfirm = .COMMIT_CONFIRM_NO Then
            optCommitConfirmNo.value = True
        Else
            optCommitConfirmYes.value = True
        End If
        
        ' SQLエラー時の挙動
        If .sqlErrorHandling = .SQL_ERROR_HANDLING_SUSPEND Then
            optSqlErrorHandlingSuspend.value = True
        Else
            optSqlErrorHandlingResume.value = True
        End If
        
        ' スキーマ
        If .schemaUse = .SCHEMA_USE_ONE Then
            optSchemaUseOne.value = True
        Else
            optSchemaUseMultiple.value = True
        End If
        
        ' 正常時のクエリ結果表示有無
        If .queryResultShowWhenNormal = True Then
            optQueryResultShowWhenNormalYes.value = True
        Else
            optQueryResultShowWhenNormalNo.value = True
        End If

        ' フォント名
        cboFontList.value = .cellFontName
        ' フォントサイズ
        cboFontSizeList.value = .cellFontSize
        ' 折り返し有無
        If .cellWordwrap = True Then
            optWordWrapYes.value = True
        Else
            optWordWrapNo.value = True
        End If
        ' セル幅
        txtCellWidth.value = .cellWidth
        ' セル高さ
        txtCellHeight.value = .cellHeight
        ' 行高の自動調整
        If .lineHeightAutoAdjust = True Then
            optLineHeightAutoAdjustYes.value = True
        Else
            optLineHeightAutoAdjustNo.value = True
        End If
        
        ' 空白セル読み取り方式
        If .emptyCellReading = .EMPTY_CELL_READING_DEL Then
            optEmptyCellReadingDel.value = True
        ElseIf .emptyCellReading = .EMPTY_CELL_READING_NON_DEL Then
            optEmptyCellReadingNonDel.value = True
        Else
            optEmptyCellReadingNonDelStrEmpty.value = True
        End If
        
        ' 直接入力文字
        If .directInputChar = .DIRECT_INPUT_CHAR_DISABLE Then
            optDirectInputCharDisable.value = True
        Else
            optDirectInputCharEnableCustom = True
        End If
        
        txtDirectInputCharEnableCustomPrefix = .directInputCharCustomPrefix
        txtDirectInputCharEnableCustomSuffix = .directInputCharCustomSuffix
        
        ' クエリパラメータの囲み文字
        txtQueryParameterEncloseEnableCustomPrefix = .queryParameterEncloseCustomPrefix
        txtQueryParameterEncloseEnableCustomSuffix = .queryParameterEncloseCustomSuffix
        
        ' NULL入力文字
        If .nullInputChar = .NULL_INPUT_CHAR_DISABLE Then
            optNullInputCharDisable.value = True
        Else
            optNullInputCharEnableCustom = True
        End If
        txtNullInputCharEnableCustom.text = .nullInputCharCustom
        
        ' SELECT時のセルの最大文字数チェック
        If .selectCheckCellMaxLength = False Then
            optSelectCheckCellMaxLengthDisable.value = True
        Else
            optSelectCheckCellMaxLengthEnable.value = True
        End If
        
        ' テーブル・カラム名エスケープ
        chkTableColumnEscapeOracle.value = .tableColumnEscapeOracle
        chkTableColumnEscapeMysql.value = .tableColumnEscapeMysql
        chkTableColumnEscapePostgresql.value = .tableColumnEscapePostgresql
        chkTableColumnEscapeSqlserver.value = .tableColumnEscapeSqlserver
        chkTableColumnEscapeAccess.value = .tableColumnEscapeAccess
        chkTableColumnEscapeSymfoware.value = .tableColumnEscapeSymfoware
        
    End With
    
    applicationSettingColFmt.readForData
    
End Sub

' =========================================================
' ▽処理単位テキスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtRecProcessCountCustom_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' 整数かをチェックする
    If validInteger(txtRecProcessCountCustom.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtRecProcessCountCustom
    
    ' 数値範囲チェック
    ElseIf CDec(txtRecProcessCountCustom.text) < 1 Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtRecProcessCountCustom

    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtRecProcessCountCustom
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽セル幅テキスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtCellWidth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' 数値チェック
    If validUnsignedNumeric(txtCellWidth.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NUMERIC
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtCellWidth
    
    ' 数値範囲チェック
    ElseIf CDec(txtCellWidth.text) < applicationSetting.CELL_WIDTH_DEFAULT Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", applicationSetting.CELL_WIDTH_DEFAULT)
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtCellWidth
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtCellWidth
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽セル高さテキスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtCellHeight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' 数値チェック
    If validUnsignedNumeric(txtCellHeight.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NUMERIC
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtCellHeight
    
    ' 数値範囲チェック
    ElseIf CDec(txtCellHeight.text) < applicationSetting.CELL_HEIGHT_DEFAULT Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", applicationSetting.CELL_HEIGHT_DEFAULT)
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtCellHeight
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtCellHeight
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽フォントリスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboFontList_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' リストに現在入力されているテキストの要素が存在しない場合
    If fontList.exist(cboFontList.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NO_LIST_ITEM
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse cboFontList
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue cboFontList
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

    
End Sub

' =========================================================
' ▽フォントサイズリスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboFontSizeList_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:
    
    ' フォントサイズに数値が入力されていない場合
    If validUnsignedNumeric(cboFontSizeList.value) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_NUMERIC
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse cboFontSizeList
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue cboFontSizeList
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

    
End Sub

' =========================================================
' ▽直接入力文字接頭辞　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtDirectInputCharEnableCustomPrefix_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' テキストボックスに入力がない場合
    If txtDirectInputCharEnableCustomPrefix.text = "" Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtDirectInputCharEnableCustomPrefix
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtDirectInputCharEnableCustomPrefix
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽クエリパラメータ囲み文字接頭辞　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtQueryParameterEncloseEnableCustomPrefix_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' テキストボックスに入力がない場合
    If txtQueryParameterEncloseEnableCustomPrefix.text = "" Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtQueryParameterEncloseEnableCustomPrefix
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtQueryParameterEncloseEnableCustomPrefix
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽クエリパラメータ囲み文字接尾辞　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtQueryParameterEncloseEnableCustomSuffix_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' テキストボックスに入力がない場合
    If txtQueryParameterEncloseEnableCustomSuffix.text = "" Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtQueryParameterEncloseEnableCustomSuffix
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtQueryParameterEncloseEnableCustomSuffix
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽NULL直接入力文字　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtNullInputCharEnableCustom_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' テキストボックスに入力がない場合
    If txtNullInputCharEnableCustom.text = "" Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtNullInputCharEnableCustom
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtNullInputCharEnableCustom
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽カラム書式設定（Oracle）ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdColumnTypeFormatOracle_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.Oracle)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ▽カラム書式設定（MySQL）ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdColumnTypeFormatMySQL_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.MySQL)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ▽カラム書式設定（PostgreSQL）ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdColumnTypeFormatPostgreSQL_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.PostgreSQL)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ▽カラム書式設定（SQLServer）ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdColumnTypeFormatSQLServer_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.MicrosoftSqlServer)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ▽カラム書式設定（Access）ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdColumnTypeFormatAccess_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.MicrosoftAccess)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ▽カラム書式設定（Symfoware）ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdColumnTypeFormatSymfoware_Click()

    Dim dbColInfo As ValDbColumnFormatInfo
    Set dbColInfo = applicationSettingColFmt.getDbColFormatInfo(DbmsType.Symfoware)
    
    settingColFormatDb = dbColInfo.dbName
    
    frmDBColumnFormatVar.ShowExt vbModal, dbColInfo
End Sub

' =========================================================
' ▽カラム書式設定ウィンドウのOKボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：dbColumnFormatInfo カラム書式設定ウィンドウで設定された情報
' 戻り値　　：
'
' =========================================================
Private Sub frmDBColumnFormatVar_ok(ByVal dbColumnFormatInfo As ValDbColumnFormatInfo)

    ' アプリケーション設定情報にロードされた情報を設定する
    applicationSettingColFmt.setDbColFormatInfo dbColumnFormatInfo
    
    ' 情報を書き込む
    applicationSettingColFmt.writeForDataDbInfo dbColumnFormatInfo

End Sub

' =========================================================
' ▽カラム書式設定ウィンドウのキャンセルボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub frmDBColumnFormatVar_cancel()

End Sub
