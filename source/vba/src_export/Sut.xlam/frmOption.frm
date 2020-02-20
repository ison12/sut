VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOption 
   Caption         =   "オプション"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8655.001
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
    restoreOptionInfo
    
    ' エラーメッセージをクリアする
    lblErrorMessage.Caption = ""

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
    
        ' レコード処理単位の指定をコントロールに反映する
        If optRecProcessCountAll.value = True Then
        
            .recProcessCount = .REC_PROCESS_COUNT_ALL
            .recProcessCountCustom = txtRecProcessCountUserInput.value
        Else
        
            .recProcessCount = .REC_PROCESS_COUNT_COSTOM
            .recProcessCountCustom = txtRecProcessCountUserInput.value
        
        End If
        
        ' コミット確認の指定をコントロールに反映する
        If optCommitConfirmNo.value = True Then
        
            .commitConfirm = .COMMIT_CONFIRM_NO
        Else
            
            .commitConfirm = .COMMIT_CONFIRM_YES
        End If
        
        ' SQLエラー時の挙動の指定をコントロールに反映する
        If optSqlErrorHandlingSuspend.value = True Then
        
            .sqlErrorHandling = .SQL_ERROR_HANDLING_SUSPEND
        Else
        
            .sqlErrorHandling = .SQL_ERROR_HANDLING_RESUME
        End If
        
        ' 空白セル読み取り方式の指定をコントロールに反映する
        If optEmptyCellReadingDel.value = True Then
        
            .emptyCellReading = .EMPTY_CELL_READING_DEL
        
        Else
        
            .emptyCellReading = .EMPTY_CELL_READING_NON_DEL
            
        End If
        
        ' 直接入力文字の指定をコントロールに反映する
        If optDirectInputCharDisable.value = True Then
        
            .directInputChar = .DIRECT_INPUT_CHAR_DISABLE
            .directInputCharCustom = txtDirectInputCharEnableCustom.value
        Else
        
            .directInputChar = .DIRECT_INPUT_CHAR_ENABLE_CUSTOM
            .directInputCharCustom = txtDirectInputCharEnableCustom.value
        
        End If
        
        ' 正常時のクエリ結果表示有無
        If optQueryResultShowWhenNormalNo.value = True Then
        
            .queryResultShowWhenNormal = False
        Else
        
            .queryResultShowWhenNormal = True
        End If
        
        ' スキーマの指定をコントロールに反映する
        If optSchemaUseOne.value = True Then
        
            .schemaUse = .SCHEMA_USE_ONE
        Else
        
            .schemaUse = .SCHEMA_USE_MULTIPLE
        End If
        
        ' テーブル・カラム名エスケープ
        .tableColumnEscapeOracle = chkTableColumnEscapeOracle.value
        .tableColumnEscapeMysql = chkTableColumnEscapeMysql.value
        .tableColumnEscapePostgresql = chkTableColumnEscapePostgresql.value
        .tableColumnEscapeSqlserver = chkTableColumnEscapeSqlserver.value
        .tableColumnEscapeAccess = chkTableColumnEscapeAccess.value
        .tableColumnEscapeSymfoware = chkTableColumnEscapeSymfoware.value
        
        ' フォント名を反映する
        .cellFontName = cboFontList.value
        
        ' フォントサイズを反映する
        .cellFontSize = cboFontSizeList.value
        
        ' 折り返し有無を反映する
        If optWordWrapYes.value = True Then
        
            .cellWordwrap = True
        Else
        
            .cellWordwrap = False
        End If
        
        ' セル幅を反映する
        .cellWidth = CDbl(txtCellWidth.value)
        
        ' セル高さを反映する
        .cellHeight = CDbl(txtCellHeight.value)
        
        ' 行高の自動調整
        If optLineHeightAutoAdjustNo.value = True Then
        
            .lineHeightAutoAdjust = False
        Else
        
            .lineHeightAutoAdjust = True
        End If
        
        ' レジストリに書き込みを行う
        .writeForData
    
    End With
End Sub

' =========================================================
' ▽オプション情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreOptionInfo()

    ' アプリケーション設定オブジェクト
    With applicationSetting
        
        ' レジストリから読み込みを行う
        .readForData
        
        ' レコード処理単位をコントロールに反映する
        If .recProcessCount = .REC_PROCESS_COUNT_ALL Then
            
            optRecProcessCountAll.value = True
        
        Else
        
            optRecProcessCountCustom.value = True
        End If
        
        txtRecProcessCountUserInput.value = .recProcessCountCustom
        
        ' コミット確認をコントロールに反映する
        If .commitConfirm = .COMMIT_CONFIRM_NO Then
        
            optCommitConfirmNo.value = True
        Else
        
            optCommitConfirmYes.value = True
        End If
        
        ' SQLエラー時の挙動をコントロールに反映する
        If .sqlErrorHandling = .SQL_ERROR_HANDLING_SUSPEND Then
        
            optSqlErrorHandlingSuspend.value = True
        Else
        
            optSqlErrorHandlingResume.value = True
        End If
        
        ' 空白セル読み取り方式をコントロールに反映する
        If .emptyCellReading = .EMPTY_CELL_READING_DEL Then
            
            optEmptyCellReadingDel.value = True
        Else
        
            optEmptyCellReadingNonDel.value = True
        End If
        
        ' 直接入力文字指定をコントロールに反映する
        If .directInputChar = .DIRECT_INPUT_CHAR_DISABLE Then
            
            optDirectInputCharDisable.value = True
        
        Else
        
            optDirectInputCharEnableCustom = True
        End If
        
        txtDirectInputCharEnableCustom = .directInputCharCustom
        
        ' 正常時のクエリ結果表示有無を反映する
        If .queryResultShowWhenNormal = True Then
        
            optQueryResultShowWhenNormalYes.value = True
        Else
        
            optQueryResultShowWhenNormalNo.value = True
        End If
        
        ' スキーマをコントロールに反映する
        If .schemaUse = .SCHEMA_USE_ONE Then
        
            optSchemaUseOne.value = True
        Else
        
            optSchemaUseMultiple.value = True
        End If
        
        ' テーブル・カラム名エスケープをコントロールに反映する
        chkTableColumnEscapeOracle.value = .tableColumnEscapeOracle
        chkTableColumnEscapeMysql.value = .tableColumnEscapeMysql
        chkTableColumnEscapePostgresql.value = .tableColumnEscapePostgresql
        chkTableColumnEscapeSqlserver.value = .tableColumnEscapeSqlserver
        chkTableColumnEscapeAccess.value = .tableColumnEscapeAccess
        chkTableColumnEscapeSymfoware.value = .tableColumnEscapeSymfoware
        
        ' フォント名を反映する
        cboFontList.value = .cellFontName
        ' フォントサイズを反映する
        cboFontSizeList.value = .cellFontSize
                
        ' 折り返し有無を反映する
        If .cellWordwrap = True Then
        
            optWordWrapYes.value = True
        Else
        
            optWordWrapNo.value = True
        End If
        
        ' セル幅を反映する
        txtCellWidth.value = .cellWidth
        
        ' セル高さを反映する
        txtCellHeight.value = .cellHeight
        
        ' 行高の自動調整を反映する
        If .lineHeightAutoAdjust = True Then
        
            optLineHeightAutoAdjustYes.value = True
        Else
        
            optLineHeightAutoAdjustNo.value = True
        End If
        
    End With
End Sub

' =========================================================
' ▽処理単位テキスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtRecProcessCountUserInput_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' 整数かをチェックする
    If validInteger(txtRecProcessCountUserInput.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtRecProcessCountUserInput
    
    ' 数値範囲チェック
    ElseIf CDbl(txtRecProcessCountUserInput.text) < 1 Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtRecProcessCountUserInput

    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtRecProcessCountUserInput
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽直接入力文字テキスト　更新時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtDirectInputCharEnableCustom_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' テキストボックスに入力がない場合
    If txtDirectInputCharEnableCustom.text = "" Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_REQUIRED
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidFalse txtDirectInputCharEnableCustom
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        ' コントロールのプロパティを変更する
        VBUtil.changeControlPropertyByValidTrue txtDirectInputCharEnableCustom
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
    ElseIf CDbl(txtCellWidth.text) < applicationSetting.CELL_WIDTH_DEFAULT Then
    
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
    ElseIf CDbl(txtCellHeight.text) < applicationSetting.CELL_HEIGHT_DEFAULT Then
    
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
