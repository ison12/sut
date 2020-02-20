VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSnapshot 
   Caption         =   "スナップショット取得"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600.001
   OleObjectBlob   =   "frmSnapshot.frx":0000
End
Attribute VB_Name = "frmSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' スナップショット取得フォーム
'
' 作成者　：Ison
' 履歴　　：2008/09/06　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽スナップショット取得実行イベント
'
' 概要　　　：
' 引数　　　：sheet ワークシート
'
' =========================================================
Public Event execSnapshot(ByRef sheet As Worksheet)

' =========================================================
' ▽キャンセルイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' =========================================================
' ▽DB変更イベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event changeDb()

' =========================================================
' ▽SQL変更イベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event changeSql(ByRef sheet As Worksheet)

' =========================================================
' ▽スナップショットクリアイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event clearSnapshot(ByRef sheet As Worksheet)

' =========================================================
' ▽実行イベント
'
' 概要　　　：
' 引数　　　：sheet    シート
'             srcIndex 比較元インデックス
'             desIndex 比較先インデックス
'
' =========================================================
Public Event execDiff(ByRef sheet As Worksheet, ByVal srcIndex As Long, ByVal desIndex As Long)

' アプリケーション設定情報
Private applicationSetting As ValApplicationSetting

' DBコネクションオブジェクト
Private dbConn As Object
' DB接続文字列
Private dbConnStr As String

' 実行SQLリスト
Private executeSqltList  As CntListBox
' スナップショットリスト
Private snapShotList     As CntListBox
' 比較元スナップショットリスト
Private srcSnapshotList  As CntListBox
' 比較先スナップショットリスト
Private desSnapshotList  As CntListBox

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
' 引数　　　：modal    モーダルまたはモードレス表示指定
' 　　　　　　aps      アプリケーション設定情報
' 　　　　　　conn     DBコネクション
' 　　　　　　connStr  DB接続文字列
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef aps As ValApplicationSetting, ByRef conn As Object, ByVal connStr As String)

    ' アプリケーション情報を設定する
    Set applicationSetting = aps
    ' DBコネクションを設定する
    Set dbConn = conn
    dbConnStr = connStr
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

    Main.storeFormPosition Me.name, Me
    Me.Hide
    
    ' 非アクティブ処理
    deactivate
    
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
        cmdClose_Click
    End If
    
End Sub

' =========================================================
' ▽DB変更処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDBConnectedChange_Click()

    On Error GoTo err
    
    RaiseEvent changeDb
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽実行SQL更新処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdExecuteSqlUpdate_Click()

    On Error GoTo err
    
    Dim sheet As Worksheet
    
    Dim ExeSnapSqlDefineSheetReader As ExeSnapSqlDefineSheetReader
    
    ' リストオブジェクトを初期化する
    executeSqltList.removeAll
    executeSqltList.init cboExecuteSql
    
    ' 全シートを対象にする
    For Each sheet In targetBook.Sheets
    
        Set ExeSnapSqlDefineSheetReader = New ExeSnapSqlDefineSheetReader
        Set ExeSnapSqlDefineSheetReader.sheet = sheet
                
        If ExeSnapSqlDefineSheetReader.isSqlDefineSheet = True Then
            ' SQL定義シートの場合、リストに追加
            executeSqltList.addItem sheet.name, sheet
        
        End If
    
    Next
    
    ' 実行SQ選択コンボボックスに追加されたものがあれば
    ' 先頭をデフォルト選択する
    If cboExecuteSql.ListCount >= 1 Then
        cboExecuteSql.ListIndex = 0
    End If
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽実行SQL変更処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboExecuteSql_Change()

    On Error GoTo err
    
    ' 実行SQ選択コンボボックスが未選択の場合
    If cboExecuteSql.ListIndex = -1 Then
        clearSnapshot
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = executeSqltList.getItem(cboExecuteSql.ListIndex)
    
    RaiseEvent changeSql(sheet)
    
    sheet.activate
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽スナップショット一覧クリア処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdSnapshotClear_Click()

    On Error GoTo err
    
    ' 実行SQ選択コンボボックスが未選択の場合
    If cboExecuteSql.ListIndex = -1 Then
        clearSnapshot
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = executeSqltList.getItem(cboExecuteSql.ListIndex)
    
    RaiseEvent clearSnapshot(sheet)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽ページ変更処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub multiPage_Change()

End Sub

' =========================================================
' ▽スナップショット取得処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdSnapshotGet_Click()

    On Error GoTo err
    
    ' 実行SQ選択コンボボックスが未選択の場合
    If cboExecuteSql.ListIndex = -1 Then
    
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = executeSqltList.getItem(cboExecuteSql.ListIndex)
    
    RaiseEvent execSnapshot(sheet)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽スナップショットリスト比較元変更イベント
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lstSnapshotSrc_Change()

    On Error GoTo err
    
    refreshLstSnapshotDes lstSnapshotSrc.ListIndex

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽比較結果出力イベント
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdResultOut_Click()

    On Error GoTo err
    
    ' 実行SQ選択コンボボックスが未選択の場合
    If cboExecuteSql.ListIndex = -1 Then
    
        Exit Sub
    End If

    Dim srcIndex As Long
    Dim desIndex As Long

    srcIndex = lstSnapshotSrc.ListIndex
    desIndex = lstSnapshotDes.ListIndex

    If srcIndex = desIndex Then
        ' 同じにならないはず
        Exit Sub
    End If

    If srcIndex = -1 Or desIndex = -1 Then
        ' 未選択状態
        Exit Sub
    End If
    
    If srcIndex <= desIndex Then
        ' 同じにならないはず
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = executeSqltList.getItem(cboExecuteSql.ListIndex)
    
    RaiseEvent execDiff(sheet, srcIndex, desIndex)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽閉じる処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdClose_Click()

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

    Set executeSqltList = New CntListBox
    Set snapShotList = New CntListBox
    snapShotList.init lstSnapshot
    Set srcSnapshotList = New CntListBox
    srcSnapshotList.init lstSnapshotSrc
    Set desSnapshotList = New CntListBox
    desSnapshotList.init lstSnapshotDes

    ' 閉じるボタンを非表示にする
    cmdClose.Width = 0

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

    Set executeSqltList = Nothing
    Set snapShotList = Nothing
End Sub

' =========================================================
' ▽アクティブ処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub activate()

    cmdExecuteSqlUpdate_Click
    
    txtDBConnected.text = dbConnStr
    multiPage.value = 0 ' 最初のページを表示
    
End Sub

' =========================================================
' ▽非アクティブ処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub deactivate()

End Sub

' =========================================================
' ▽DBコネクションの更新
'
' 概要　　　：conn    DBコネクション
'             connStr DB接続文字列
'
' =========================================================
Public Sub updateDbConn(ByRef conn As Object, ByVal connStr As String)

    On Error GoTo err
    
    Set dbConn = conn
    dbConnStr = connStr

    txtDBConnected.text = dbConnStr

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽スナップショットの削除
'
' 概要　　　：
'
' =========================================================
Public Sub clearSnapshot()

    On Error GoTo err
    
    snapShotList.removeAll
    snapShotList.init lstSnapshot
    
    srcSnapshotList.removeAll
    srcSnapshotList.init lstSnapshotSrc
    
    desSnapshotList.removeAll
    desSnapshotList.init lstSnapshotDes
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub



' =========================================================
' ▽スナップショットの追加
'
' 概要　　　：label ラベル
'             value 値
'
' =========================================================
Public Sub addSnapshot(ByRef label As String, ByRef value As String)

    On Error GoTo err
    
    snapShotList.addItem label, value
    srcSnapshotList.addItem label, value
    
    ' 末尾を選択
    srcSnapshotList.setSelectedIndex srcSnapshotList.count - 1
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽スナップショットリスト比較先更新
'
' 概要　　　：
' 引数　　　：lstSnapshotSrcListIndex スナップショットリスト比較元インデックス
' 戻り値　　：
'
' =========================================================
Private Sub refreshLstSnapshotDes(ByVal lstSnapshotSrcListIndex As Long)

    Dim snapshot As ValSnapRecordsSet

    If desSnapshotList Is Nothing Then
        Set desSnapshotList = New CntListBox
        desSnapshotList.init lstSnapshotDes
    Else
        desSnapshotList.removeAll
        desSnapshotList.init lstSnapshotDes
    End If
    
    Dim i As Long
    
    i = 0
    For i = 0 To snapShotList.count - 1
    
        If i < lstSnapshotSrcListIndex Then
            desSnapshotList.addItem snapShotList.control.list(i), Empty
        End If
    
    Next
    
    If lstSnapshotDes.ListCount > 0 Then
        lstSnapshotDes.ListIndex = lstSnapshotDes.ListCount - 1
    End If

End Sub

