VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectConditionCreator 
   Caption         =   "SELECT"
   ClientHeight    =   8805.001
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7935
   OleObjectBlob   =   "frmSelectConditionCreator.frx":0000
End
Attribute VB_Name = "frmSelectConditionCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' SELECT条件生成フォーム
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/04/03　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽処理が完了した場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：sql SELECT SQL
'
' =========================================================
Public Event ok(ByVal sql As String, ByVal append As Boolean)

' =========================================================
' ▽処理がキャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' ---------------------------------------------------------
' INIファイル関連定数
' ---------------------------------------------------------
' ▼最後に操作された情報
Private Const REG_SUB_KEY_SELECT_CONDITION As String = "select_condition"

' 簡易設定ページ
Private Const PAGE_SIMPLE_SETTING As Long = 0
' 詳細設定ページ
Private Const PAGE_DETAIL_SETTING As Long = 1

' 条件指定数の最小値
Private Const COLUMN_COND_MIN As Long = 1
' 条件指定数の最大値
Private Const COLUMN_COND_MAX As Long = 10

' 順序値 昇順
Private Const ORDER_BY_VALUE_ASC  As Variant = True
' 順序値 降順
Private Const ORDER_BY_VALUE_DESC As Variant = False
' 順序値 指定なし
Private Const ORDER_BY_VALUE_NONE As Variant = Null

' アプリケーション設定
Private applicationSetting As ValApplicationSetting
' アプリケーション設定（カラム書式情報）
Private applicationSettingColFmt As ValApplicationSettingColFormat

' DBコネクション
Private dbConn As Object
' DBMS種類
Private dbms   As DbmsType
' テーブル定義
Private tableSheet As ValTableWorksheet

' 検索条件　配列ントロール　カラム
Private columnCondList()   As CntListBox
' 検索条件　配列ントロール　値
Private valueCondList()    As control
' 検索条件　配列ントロール　順序
Private orderCondList()    As control

' SQL編集フラグ
Private editedSql As Boolean

' =========================================================
' ▽フォーム表示（拡張処理）
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByRef apps As ValApplicationSetting _
                 , ByRef appsColFmt As ValApplicationSettingColFormat _
                 , ByRef conn As Object)

    ' エラーメッセージをクリアする
    lblErrorMessage.Caption = ""

    ' アプリケーション設定情報を設定
    Set applicationSetting = apps
    Set applicationSettingColFmt = appsColFmt
    ' DBコネクションを設定
    Set dbConn = conn
    ' DBMS種類を取得する
    dbms = ADOUtil.getDBMSType(dbConn)
    
    ' アクティブ時の処理
    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal

End Sub

' =========================================================
' ▽フォーム非表示（拡張処理）
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub HideExt()

    ' ディアクティブ時の処理
    deactivate
    
    Main.storeFormPosition Me.name, Me
    Me.Hide

End Sub

' =========================================================
' ▽テストボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdTest_Click()

    On Error GoTo err

    SutWhite.showHourglassWindowOnCenterPt Me
    
    Dim resultSet   As Object
    Dim resultCount As Long

    Set resultSet = ADOUtil.querySelect(dbConn, txtSqlEditor.value, resultCount, adOpenStatic)
    resultCount = resultSet.recordCount
    
    ADOUtil.closeRecordSet resultSet

    lblResultCount.Caption = resultCount & "件"

    SutWhite.closeHourglassWindow
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

    SutWhite.closeHourglassWindow
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
' ▽初期化処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub initial()

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

    ' テーブルシート読込オブジェクト
    Dim tableSheetReader As ExeTableSheetReader
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.sheet = ActiveSheet
    Set tableSheetReader.conn = dbConn
            
    ' テーブル定義を読み込む
    Set tableSheet = tableSheetReader.readTableInfo

    Dim table As ValDbDefineTable
    Set table = tableSheet.table

    Dim i As Long
    
    ' -----------------------------------------------
    ' カラム名
    ' -----------------------------------------------
    ' コントロール配列を解放する
    Erase columnCondList
    ' コントロール配列を確保する
    ReDim columnCondList(COLUMN_COND_MIN To COLUMN_COND_MAX)
    
    i = COLUMN_COND_MIN
    
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond1: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond2: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond3: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond4: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond5: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond6: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond7: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond8: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond9: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond10: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
        
    ' -----------------------------------------------
    ' 値
    ' -----------------------------------------------
    ' コントロール配列を解放する
    Erase valueCondList
    ' コントロール配列を確保する
    ReDim valueCondList(COLUMN_COND_MIN To COLUMN_COND_MAX)
    
    i = COLUMN_COND_MIN
        
    ' TextBoxのオブジェクトをそのまま代入しようとすると何故かString型に変換されるので
    ' Controlsリストから間接的に取得して代入する
    Set valueCondList(i) = Controls.item(txtCondValue1.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue2.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue3.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue4.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue5.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue6.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue7.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue8.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue9.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue10.name): i = i + 1
    
    ' -----------------------------------------------
    ' 順序
    ' -----------------------------------------------
    ' コントロール配列を解放する
    Erase orderCondList
    ' コントロール配列を確保する
    ReDim orderCondList(COLUMN_COND_MIN To COLUMN_COND_MAX)
    
    i = COLUMN_COND_MIN
        
    Set orderCondList(i) = tglOrderCond1: i = i + 1
    Set orderCondList(i) = tglOrderCond2: i = i + 1
    Set orderCondList(i) = tglOrderCond3: i = i + 1
    Set orderCondList(i) = tglOrderCond4: i = i + 1
    Set orderCondList(i) = tglOrderCond5: i = i + 1
    Set orderCondList(i) = tglOrderCond6: i = i + 1
    Set orderCondList(i) = tglOrderCond7: i = i + 1
    Set orderCondList(i) = tglOrderCond8: i = i + 1
    Set orderCondList(i) = tglOrderCond9: i = i + 1
    Set orderCondList(i) = tglOrderCond10: i = i + 1
    
    
    ' ファイルから各コントロールの情報を読み込む
    restoreSelectCondition
    
    ' ページを切り替え処理
    ' SQLエディタにテキストが設定されていない場合
    If txtSqlEditor.value = "" Then
    
        ' 簡易ページへ
        mpageCondition.value = PAGE_SIMPLE_SETTING
        
    ' SQLエディタにテキストが設定されている場合
    Else
    
        ' 詳細ページへ
        mpageCondition.value = PAGE_DETAIL_SETTING
        
        ' 編集フラグをtrueに設定しておく
        editedSql = True
    End If
    
    ' 結果件数表示ラベルを初期化する
    lblResultCount.Caption = ""

    
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
' ▽順序指定トグルボタン変更時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub tglOrderCond1_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond1
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond2_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond2
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond3_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond3
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond4_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond4
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond5_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond5
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond6_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond6
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond7_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond7
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond8_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond8
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond9_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond9
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond10_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond10
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽順序指定トグルボタンのラベル変更
'
' 概要　　　：順序指定トグルボタンの状態に応じてラベルを変更するための処理
' 引数　　　：control トグルボタン
' 戻り値　　：
'
' =========================================================
Private Sub changeLabelOrderControl(ByRef control As ToggleButton)

    If control.value = ORDER_BY_VALUE_ASC Then
    
        control.Caption = "昇順"
    
    ElseIf control.value = ORDER_BY_VALUE_DESC Then
    
        control.Caption = "降順"
    Else
    
        control.Caption = "なし"
    End If

End Sub

' =========================================================
' ▽PKで条件指定クリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdPkCondition_Click()

    On Error GoTo err:

    ' 一度リセットする
    resetWhereOrderby

    Dim i As Long: i = COLUMN_COND_MIN
    
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    ' カラム
    Dim column     As ValDbDefineColumn
    ' カラムリスト
    Dim columnList As ValCollection
    
    ' テーブル制約情報(PK)
    Dim tableConstPk    As New ValDbDefineTableConstraints
    ' PKカラムであるかをあらわすフラグ
    Dim isColumnPk      As Boolean
    
    Dim tableConstTmp   As ValDbDefineTableConstraints
    ' テーブル制約リストからPK制約を取得する
    For Each tableConstTmp In table.constraintsList
    
        If tableConstTmp.constraintType = sutredlib.tableConstPk Then
        
            Set tableConstPk = tableConstTmp
            Exit For
        End If
    Next
    
    ' カラムリストを取得する
    Set columnList = table.columnList
    
    ' カラムリストを1件ずつ処理する
    For Each column In columnList.col
            
        ' PK制約であるかどうかを判定する
        If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
        
            isColumnPk = False
        Else
        
            isColumnPk = True
        End If
        
        ' カラムがPKである場合
        If isColumnPk = True Then
        
            ' PKの数がコントロールの数を上回っているのでメッセージを表示して終了する
            If i > COLUMN_COND_MAX Then
            
                err.Raise ConstantsError.ERR_NUMBER_OVER_SELECT_COND_CONTROL _
                        , _
                        , ConstantsError.ERR_DESC_OVER_SELECT_COND_CONTROL
                Exit Sub
            End If
            
            ' カラム名にPK列名を設定する
            columnCondList(i).control.value = column.columnName
            ' 順序に昇順を設定する
            orderCondList(i).value = ORDER_BY_VALUE_ASC
            i = i + 1
        End If
        
                
    Next

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽レコード取得範囲　開始テキストボックスのチェック
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtRecRangeStart_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' 空の場合、正常
    If txtRecRangeStart.text = "" Then
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecRangeStart

    ' テキストボックスの値が整数かをチェックする
    ElseIf validInteger(txtRecRangeStart.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        ' アラートを表示する
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
        
        changeControlPropertyByValidFalse txtRecRangeStart
    
    ' 数値範囲チェック
    ElseIf CDbl(txtRecRangeStart.text) < 1 Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse txtRecRangeStart
    
    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecRangeStart
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽レコード取得範囲　終了テキストボックスのチェック
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtRecRangeEnd_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' 空の場合、正常
    If txtRecRangeEnd.text = "" Then
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecRangeEnd

    ' テキストボックスの値が整数かをチェックする
    ElseIf validInteger(txtRecRangeEnd.text) = False Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        ' アラートを表示する
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
    
        changeControlPropertyByValidFalse txtRecRangeEnd
        
    ' 数値範囲チェック
    ElseIf CDbl(txtRecRangeEnd.text) < 1 Then
    
        ' 更新をキャンセルする
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' コントロールのプロパティを変更する
        changeControlPropertyByValidFalse txtRecRangeEnd

    ' 正常な場合
    Else
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecRangeEnd
    End If
    
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
' ▽先頭100件を取得するボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdLimit100_Click()

    On Error GoTo err:

    ' レコード範囲 開始を設定する
    txtRecRangeStart.value = 1
    ' レコード範囲 終了を設定する
    txtRecRangeEnd.value = 100
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽次へボタン押下時のイベントプロシージャ
'
' 概要　　　：ページを切り替える
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdNext_Click()

    On Error GoTo err:

    ' SQLを生成し、SQL編集テキストボックスに内容を表示
    ' ページを切り替える前に変更を行う
    txtSqlEditor.value = createSql

    ' ページを切り替える
    mpageCondition.value = PAGE_DETAIL_SETTING
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽戻るボタン押下時のイベントプロシージャ
'
' 概要　　　：ページを切り替える
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdReturn_Click()

    On Error GoTo err:

    ' SQL編集フラグの確認
    ' 内容が編集されている場合
    If editedSql = True Then
    
        ' メッセージボックスの戻り値
        Dim ret As Long
    
        ' 編集後に戻る場合は、警告を表示する
        ret = VBUtil.showMessageBoxForYesNo("戻ると編集内容が消えてしまいますが、よろしいですか？", ConstantsCommon.APPLICATION_NAME)
        
        ' はいを選択した場合
        If ret = WinAPI_User.IDYES Then
        
            ' ページを切り替える
            mpageCondition.value = PAGE_SIMPLE_SETTING
            txtSqlEditor.value = ""
            editedSql = False
        End If
        
    ' 内容が編集されていない場合
    Else
    
        ' ページを切り替える
        mpageCondition.value = PAGE_SIMPLE_SETTING
        txtSqlEditor.value = ""
    
    End If

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽リセットクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdReset_Click()

    On Error GoTo err:

    ' 条件・並び順をリセット
    resetWhereOrderby
    ' レコード範囲指定をリセット
    resetRecRange

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽条件・並び順の設定をリセット
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub resetWhereOrderby()

    Dim i As Long
    
    ' コントロール配列を1件ずつ処理する
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        ' カラム名を空に設定
        columnCondList(i).control.value = ""
        ' 値を空に設定
        valueCondList(i).value = ""
        ' 順序をなしに設定
        orderCondList(i).value = ORDER_BY_VALUE_NONE
    Next
    
End Sub

' =========================================================
' ▽レコード取得の範囲指定をリセット
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub resetRecRange()

    txtRecRangeStart.value = ""
    txtRecRangeEnd.value = ""
    
End Sub

' =========================================================
' ▽SQLエディタ 変更イベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub txtSqlEditor_Change()

    ' 詳細ページで、Changeイベントが発生した場合、編集フラグをONにする
    If mpageCondition.value = PAGE_DETAIL_SETTING Then
    
        editedSql = True
    End If
End Sub


' =========================================================
' ▽OKボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err:
    
    ' SQL
    Dim sql As String
    
    ' ページが簡易設定の場合
    If mpageCondition.value = PAGE_SIMPLE_SETTING Then
    
        ' SQLをコントロールから生成する
        sql = createSql
    
    ' ページが詳細設定の場合
    Else
    
        ' SQLをエディタから取得する
        sql = txtSqlEditor.value
    End If
    
    ' SELECT条件を保存する
    storeSelectCondition
    
    Me.HideExt
    
    RaiseEvent ok(sql, cbxTableSheetAppend.value)

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
Private Sub cmdCancel_Click()

    On Error GoTo err:
    
    Me.HideExt
    RaiseEvent Cancel

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽SQLを生成する
'
' 概要　　　：SQLを生成する。
' 引数　　　：
' 戻り値　　：SELECTクエリー
'
' =========================================================
Private Function createSql() As String

    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    ' SELECT条件
    Dim condition As ValSelectCondition
    ' フォームから条件を生成する
    Set condition = createCondition

    Dim dbObjFactory As New DbObjectFactory
    Dim queryCreator        As IDbQueryCreator

    Set queryCreator = dbObjFactory.createQueryCreator(dbConn _
                                                            , applicationSetting.emptyCellReading _
                                                            , applicationSetting.getDirectInputChar _
                                                            , applicationSettingColFmt.getDbColFormatListByDbConn(dbConn) _
                                                            , applicationSetting.schemaUse _
                                                            , applicationSetting.getTableColumnEscapeByDbConn(dbConn))

    ' SELECT文を生成する
    createSql = queryCreator.createSelect(table, condition)
    
End Function

' =========================================================
' ▽条件生成
'
' 概要　　　：コントロールから条件を生成する。
' 引数　　　：
' 戻り値　　：SELECT条件オブジェクト
'
' =========================================================
Private Function createCondition() As ValSelectCondition

    ' 戻り値
    Dim ret As ValSelectCondition
    ' 戻り値を初期化する
    Set ret = New ValSelectCondition

    ' カラム名
    Dim columnName  As String
    ' 値
    Dim value       As String
    ' 順序
    Dim orderby     As Variant
    ' 順序 (Long型)
    Dim orderByLong As Long
    
    Dim i As Long
    
    ' コントロール配列を1件ずつ処理する
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        ' カラム名を取得
        columnName = columnCondList(i).control.value
        ' 値を取得
        value = valueCondList(i).value
        ' 順序を取得
        orderby = orderCondList(i).value
    
        ' カラム名が設定されている場合のみ、条件として設定する
        If columnName <> "" Then
        
            ' コントロールの値を ValSelectCondition の定数に変換する
            ' 昇順
            If orderby = ORDER_BY_VALUE_ASC Then
            
                orderByLong = ret.ORDER_ASC
                
            ' 降順
            ElseIf orderby = ORDER_BY_VALUE_DESC Then
            
                orderByLong = ret.ORDER_DESC
                
            ' 無し
            Else
            
                orderByLong = ret.ORDER_NONE
            End If
            
            ' 条件を設定する
            ret.setCondition columnName, value, orderByLong
            
        End If
        
    Next
    
    ' レコード範囲指定 開始が設定されている場合、条件として設定
    If txtRecRangeStart.value <> "" Then
    
        ret.recRangeStart = txtRecRangeStart.value
    End If
    
    ' レコード範囲指定 終了が設定されている場合、条件として設定
    If txtRecRangeEnd.value <> "" Then
    
        ret.recRangeEnd = txtRecRangeEnd.value
    End If

    ' 戻り値に設定
    Set createCondition = ret

End Function

' =========================================================
' ▽SELECT条件を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeSelectCondition()

    On Error GoTo err
    
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    ' SELECT条件を格納するための配列変数
    ' 条件指定数×(カラム・値・並び順)＋レコード範囲指定（開始・終了）＋SQLエディタ ＝ 10×3＋2＋1
    Dim selectCondition(COLUMN_COND_MIN _
                    To (COLUMN_COND_MAX * 3 + 2 + 1), 0 To 1) As Variant
    
    
    Dim i As Long
    Dim j As Long
    
    j = COLUMN_COND_MIN
    
    ' コントロール配列を1件ずつ処理する
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        selectCondition(j, 0) = columnCondList(i).control.name
        selectCondition(j, 1) = columnCondList(i).control.value: j = j + 1
        selectCondition(j, 0) = valueCondList(i).name
        selectCondition(j, 1) = valueCondList(i).value: j = j + 1
        selectCondition(j, 0) = orderCondList(i).name
        ' 順序コントロール（トグルボタン）は未選択の場合にNULLを返すので空文字列に変換する
        selectCondition(j, 1) = VBUtil.convertNullToEmptyStr(orderCondList(i).value): j = j + 1
    
    Next

    selectCondition(j, 0) = txtRecRangeStart.name
    selectCondition(j, 1) = txtRecRangeStart.value: j = j + 1

    selectCondition(j, 0) = txtRecRangeEnd.name
    selectCondition(j, 1) = txtRecRangeEnd.value: j = j + 1
    
    selectCondition(j, 0) = txtSqlEditor.name
    selectCondition(j, 1) = txtSqlEditor.value: j = j + 1
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SELECT_CONDITION & "\" & table.schemaTableName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' レジストリに情報を設定する
    registry.setValues selectCondition
    
    Set registry = Nothing
    
    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = tableSheet.sheet
    bookProp.removeValueByPrefixName ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_"
    
    ' コントロール配列を1件ずつ処理する
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & columnCondList(i).control.name, columnCondList(i).control.value
        bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & valueCondList(i).name, valueCondList(i).value
        bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & orderCondList(i).name, VBUtil.convertNullToEmptyStr(orderCondList(i).value)
    
    Next
    
    bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & txtRecRangeStart.name, txtRecRangeStart.value
    bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & txtRecRangeEnd.name, txtRecRangeEnd.value
    bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & txtSqlEditor.name, txtSqlEditor.value
    ' ----------------------------------------------
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽SELECT条件を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreSelectCondition()

    On Error GoTo err
    
    Dim i As Long

    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = tableSheet.sheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' 設定情報シートが存在する
        
        Set bookPropVal = bookProp.getValuesByPrefixName(ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName)
        If bookPropVal.count > 0 Then
        
            ' 設定情報が存在するので、フォームに反映する
            ' コントロール配列を1件ずつ処理する
            For i = COLUMN_COND_MIN To COLUMN_COND_MAX
            
                columnCondList(i).control.value = bookPropVal.getItem(tableSheet.sheetName & "_" & columnCondList(i).control.name, vbString)
                valueCondList(i).value = bookPropVal.getItem(tableSheet.sheetName & "_" & valueCondList(i).name, vbString)
                orderCondList(i).value = bookPropVal.getItem(tableSheet.sheetName & "_" & orderCondList(i).name, vbString)
                
            Next
        
            txtRecRangeStart.value = bookPropVal.getItem(tableSheet.sheetName & "_" & txtRecRangeStart.name, vbString)
            txtRecRangeEnd.value = bookPropVal.getItem(tableSheet.sheetName & "_" & txtRecRangeEnd.name, vbString)
            
            txtSqlEditor.value = bookPropVal.getItem(tableSheet.sheetName & "_" & txtSqlEditor.name, vbString)

            Exit Sub
        End If
    End If
    ' ----------------------------------------------
    
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SELECT_CONDITION & "\" & table.schemaTableName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
        
    Dim retColumn As String
    Dim retValue  As String
    Dim retOrder  As String
    
    Dim retRecRangeStart As String
    Dim retRecRangeEnd   As String
    
    Dim retSqlEdit As String
    
    ' コントロール配列を1件ずつ処理する
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        retColumn = ""
        retValue = ""
        retOrder = ""
        
        registry.getValue columnCondList(i).control.name, retColumn
        registry.getValue valueCondList(i).name, retValue
        registry.getValue orderCondList(i).name, retOrder
    
        columnCondList(i).control.value = retColumn
        valueCondList(i).value = retValue
        orderCondList(i).value = retOrder
        
    Next

    registry.getValue txtRecRangeStart.name, retRecRangeStart
    registry.getValue txtRecRangeEnd.name, retRecRangeEnd
    
    txtRecRangeStart.value = retRecRangeStart
    txtRecRangeEnd.value = retRecRangeEnd
    
    registry.getValue txtSqlEditor.name, retSqlEdit
    
    txtSqlEditor.value = retSqlEdit
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub

