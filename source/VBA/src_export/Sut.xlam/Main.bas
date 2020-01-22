Attribute VB_Name = "Main"
Option Explicit

' ________________________________________________________
' メンバ変数
' ________________________________________________________

Private menuDB    As UIMenuDB
Private menuTable As UIMenuTable
Private menuData  As UIMenuData
Private menuDiff  As UIMenuDiff
Private menuFile  As UIMenuFile
Private menuTool  As UIMenuTool
Private menuHelp  As UIMenuHelp

' ■DBコネクション
Public dbConn As Object
' ■DB接続文字列
Public dbConnStr As String
' ■DB接続文字列（単純な接続文字列）
Public dbConnSimpleStr As String

' ■アプリケーション設定情報
Private applicationSetting As ValApplicationSetting

' ■アプリケーション設定情報（ショートカット）
Private applicationSettingShortcut As ValApplicationSettingShortcut

' ■アプリケーション設定情報（カラム書式情報）
Private applicationSettingColFormat As ValApplicationSettingColFormat

' アドインのファイルクローズ時の処理
Public Sub Auto_Close()

    On Error GoTo err

    #If (DEBUG_MODE = 1) Then
    
        Debug.Print "Auto_Close"
    #End If
    
    'VBUtil.showMessageBoxForInformation "Auto_Close", "テスト"
    
    SutDestroy
    
    Exit Sub
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Sub

' アドインのアンインストール時の対策
Public Function Auto_Remove()

    '何もしません
    
End Function

' =========================================================
' ▽Mainモジュールで管理しているDBコネクションを更新する。
'
' 概要　　　：
' 特記事項　：
'
' =========================================================
Public Function SutUpdateDbConn(ByRef dbConn_ As Object, ByRef dbConnStr_ As String, ByRef dbConnSimpleStr_ As String)

    If Not dbConn_ Is Nothing Then
    
        ADOUtil.closeDB dbConn
    End If
    
    Set dbConn = dbConn_
    dbConnStr = dbConnStr_
    dbConnSimpleStr = dbConnSimpleStr_
    
    If Not menuTable Is Nothing Then
        menuTable.updateDbConn dbConn_
    End If
    
    If Not menuDiff Is Nothing Then
        menuDiff.updateDbConn dbConn_, dbConnSimpleStr_
    End If
    
    If ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        changeDbConnectStatus False
    Else
        changeDbConnectStatus True
    End If

End Function

' =========================================================
' ▽Mainモジュールのメンバを初期化する
'
' 概要　　　：
' 特記事項　：ツールバーの初期化後に呼び出しを行うこと
'
' =========================================================
Public Function SutInit()
    
    ' 各種メンバのGetメソッドをコールすることでメンバを初期化する
    getApplicationSetting
    getApplicationSettingShortcut
    getApplicationSettingColFormat
    
    initUIObject
    
End Function

' =========================================================
' ▽Mainモジュールのメンバを解放する
'
' 概要　　　：
' 特記事項　：ツールバーの削除前に呼び出しを行うこと
'
' =========================================================
Public Function SutRelease()

    ADOUtil.closeDB dbConn
    Set dbConn = Nothing
    
    dbConnStr = Empty
    dbConnSimpleStr = Empty
    
    Set applicationSetting = Nothing
    Set applicationSettingShortcut = Nothing
    
    Set menuDB = Nothing
    Set menuTable = Nothing
    Set menuData = Nothing
    Set menuDiff = Nothing
    Set menuFile = Nothing
    Set menuTool = Nothing
    Set menuHelp = Nothing
    
End Function

' =========================================================
' ▽UIオブジェクトの初期化
'
' 概要　　　：
'
' =========================================================
Private Sub initUIObject()

    If menuDB Is Nothing Then
        
        Set menuDB = New UIMenuDB
    End If
    
    If menuTable Is Nothing Then
    
        Set menuTable = New UIMenuTable
    End If
    
    If menuData Is Nothing Then
    
        Set menuData = New UIMenuData
    End If
    
    If menuDiff Is Nothing Then
    
        Set menuDiff = New UIMenuDiff
    End If
    
    If menuFile Is Nothing Then
    
        Set menuFile = New UIMenuFile
    End If
    
    If menuTool Is Nothing Then
    
        Set menuTool = New UIMenuTool
    End If
    
    If menuHelp Is Nothing Then
        
        Set menuHelp = New UIMenuHelp
    End If

End Sub

' =========================================================
' ▽アドインをロードする。
'
' 概要　　　：
'
' =========================================================
Public Function SutPreload()

    On Error GoTo err

    initLoadingToolbar
    
    Exit Function
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Function

' =========================================================
' ▽Sutを完全に破棄する
'
' 概要　　　：
'
' =========================================================
Public Function SutDestroy()

    On Error GoTo err

    ' ツールバーを削除する前に呼び出す
    ' グローバル領域のデータを解放する
    Main.SutRelease
    
    ' ツールバーを削除する
    deleteToolbar
    
    Exit Function
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Function

' =========================================================
' ▽アドインをロードする。
'
' 概要　　　：
'
' =========================================================
Public Function SutLoad()

    On Error GoTo err
    
    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init

    ' カレントドライブとカレントディレクトリを切り替える
    ChDrive SutWorkbook.path
    ChDir SutWorkbook.path
    
    ' ツールバーを初期化する
    initToolbar
    
    ' ツールバーの初期化後に呼び出す
    ' グローバル領域のデータを初期化する
    Main.SutInit
    
    Exit Function
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Function

' =========================================================
' ▽アドインをアンロードする。
'
' 概要　　　：
'
' =========================================================
Public Function SutUnload()

    On Error GoTo err

    ' Excel.Applicationプロパティを操作するオブジェクト
    ' 関数を抜けると自動でApplicationプロパティが復元される
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' ツールバーを削除する前に呼び出す
    ' グローバル領域のデータを解放する
    Main.SutRelease
    
    ' ツールバーの一部を削除する
    deleteToolbarExcludeSomeItems
    
    Exit Function
    
err:

    ' エラー発生
    Main.ShowErrorMessage

End Function

' =========================================================
' ▽DB接続設定フォーム表示
'
' 概要　　　：
'
' =========================================================
Public Function SutConnectDB()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    menuDB.init
    menuDB.connectDb
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽DB接続切断
'
' 概要　　　：
'
' =========================================================
Public Function SutDisConnectDB()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    menuDB.init
    menuDB.disconnectDb
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽DBエクスプローラ表示
'
' 概要　　　：
'
' =========================================================
Public Function SutShowDbExplorer()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim conn As Object: Set conn = getDBConnection
    
    menuTable.init appSetting, conn
    menuTable.showDbExplorer
    
    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽テーブルシート一覧表示
'
' 概要　　　：
'
' =========================================================
Public Function SutShowTableSheetList()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting

    ' オートメーションエラーが発生してしまうためダミーのオブジェクトを作っておく
    ' （原因は不明）
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    menuTable.init appSetting, conn
    menuTable.showTableSheetList
    
    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽テーブルシート作成
'
' 概要　　　：
'
' =========================================================
Public Function SutCreateTableSheet()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim conn As Object: Set conn = getDBConnection

    menuTable.init appSetting, conn
    menuTable.createTableSheet
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽テーブルシート更新
'
' 概要　　　：
'
' =========================================================
Public Function SutUpdateTableSheet()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim conn As Object: Set conn = getDBConnection

    menuTable.init appSetting, conn
    menuTable.updateTableSheet
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽INSERT実行
'
' 概要　　　：
'
' =========================================================
Public Function SutInsertUpdateAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.insertUpdateAll
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽INSERT実行（選択領域）
'
' 概要　　　：
'
' =========================================================
Public Function SutInsertUpdateSelection()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.insertUpdateSelection
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function
' =========================================================
' ▽INSERT実行
'
' 概要　　　：
'
' =========================================================
Public Function SutInsertAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.insertAll
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽INSERT実行（選択領域）
'
' 概要　　　：
'
' =========================================================
Public Function SutInsertSelection()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.insertSelection
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽UPDATE実行
'
' 概要　　　：
'
' =========================================================
Public Function SutUpdateAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.updateAll
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽UPDATE実行（選択領域）
'
' 概要　　　：
'
' =========================================================
Public Function SutUpdateSelection()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.updateSelection
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽DELETE実行
'
' 概要　　　：
'
' =========================================================
Public Function SutDeleteAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.deleteAll
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽DELETE実行（選択領域）
'
' 概要　　　：
'
' =========================================================
Public Function SutDeleteSelection()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.deleteSelection
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽DELETE実行（テーブル上の全レコード）
'
' 概要　　　：
'
' =========================================================
Public Function SutDeleteAllOfTable()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.deleteAllOfTable
    
    doAfterProcess

    menuData.showQueryResultWhenSettingResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽SELECT実行
'
' 概要　　　：
'
' =========================================================
Public Function SutSelectAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.selectAll
    
    doAfterProcess
    
    menuData.showQueryResultWhenSettingResult

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽SELECT実行（条件指定）
'
' 概要　　　：
'
' =========================================================
Public Function SutSelectCondition()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.selectCondition
    
    doAfterProcess
    
    menuData.showQueryResultWhenSettingResult

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽SELECT実行（再実行）
'
' 概要　　　：
'
' =========================================================
Public Function SutSelectReExecute()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.selectReExecute
    
    doAfterProcess
    
    menuData.showQueryResultWhenSettingResult

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽行の追加
'
' 概要　　　：
'
' =========================================================
Public Function SutRecordAdd()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    
    menuData.init appSetting, appSettingColFmt, Nothing
    menuData.recordAdd
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽クエリエディタ（ !!! 未実装 !!! ）
'
' 概要　　　：
'
' =========================================================
Public Function SutQueryEditor()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.queryEditor
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽一括クエリ
'
' 概要　　　：
'
' =========================================================
Public Function SutQueryBatch()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuData.init appSetting, appSettingColFmt, conn
    menuData.queryBatch
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽クエリ結果表示
'
' 概要　　　：
'
' =========================================================
Public Function SutShowQueryResult()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    
    menuData.init appSetting, appSettingColFmt, Nothing, False ' クエリ結果を消去しない
    menuData.showQueryResult
    
    Exit Function
err:
    
    Main.ShowErrorMessage
    
End Function

' =========================================================
' ▽スナップショットSQL定義シート作成
'
' 概要　　　：
'
' =========================================================
Public Function SutCreateNewSheetSnapSqlDefine()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    
    menuDiff.init appSetting, appSettingColFmt, Nothing
    menuDiff.createNewSheetSnapSqlDefine
    
    doAfterProcess
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽スナップショット実行フォーム呼び出し
'
' 概要　　　：
'
' =========================================================
Public Function SutShowSnapshot()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuDiff.init appSetting, appSettingColFmt, conn
    menuDiff.showSnapshot
    
    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽オプション設定
'
' 概要　　　：
'
' =========================================================
Public Function SutSettingQueryParameter()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    menuTool.settingQueryParameter
    
    'doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽オプション設定
'
' 概要　　　：
'
' =========================================================
Public Function SutSettingOption()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    menuTool.settingOption
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽右クリックメニューの設定
'
' 概要　　　：
'
' =========================================================
Public Function SutSettingRClickMenu()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    menuTool.settingRClickMenu
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ショートカットキーの設定
'
' 概要　　　：
'
' =========================================================
Public Function SutSettingShortCutKey()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    menuTool.settingShortCutKey
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽バージョンの表示
'
' 概要　　　：
'
' =========================================================
Public Function SutSettingPopupMenu()

    On Error GoTo err
    
    ' UIオブジェクトの初期化
    initUIObject
    
    menuTool.settingPopupMenu

    doAfterProcess

    Exit Function
    
err:
    
    Main.ShowErrorMessage

End Function

' =========================================================
' ▽ファイル出力 - INSERT + UPDATE（全て）
'
' 概要　　　：
'
' =========================================================
Public Function SutFileInsertUpdateAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.insertUpdateAll
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ファイル出力 - INSERT + UPDATE（選択範囲）
'
' 概要　　　：
'
' =========================================================
Public Function SutFileInsertUpdateSelection()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.insertUpdateSelection
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function
' =========================================================
' ▽ファイル出力 - INSERT（全て）
'
' 概要　　　：
'
' =========================================================
Public Function SutFileInsertAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.insertAll
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ファイル出力 - INSERT（選択範囲）
'
' 概要　　　：
'
' =========================================================
Public Function SutFileInsertSelection()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.insertSelection
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ファイル出力 - UPDATE（全て）
'
' 概要　　　：
'
' =========================================================
Public Function SutFileUpdateAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.updateAll
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ファイル出力 - UPDATE（選択範囲）
'
' 概要　　　：
'
' =========================================================
Public Function SutFileUpdateSelection()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.updateSelection
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ファイル出力 - DELETE（全て）
'
' 概要　　　：
'
' =========================================================
Public Function SutFileDeleteAll()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.deleteAll
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ファイル出力 - DELETE（選択範囲）
'
' 概要　　　：
'
' =========================================================
Public Function SutFileDeleteSelection()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.deleteSelection
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ファイル出力 - 一括出力
'
' 概要　　　：
'
' =========================================================
Public Function SutFileBatch()

    On Error GoTo err
    
    ' ブックのチェックを行う
    validWorkbook
    
    ' UIオブジェクトの初期化
    initUIObject
    
    Dim appSetting As Object: Set appSetting = getApplicationSetting
    Dim appSettingColFmt As Object: Set appSettingColFmt = getApplicationSettingColFormat
    Dim conn As Object: Set conn = getDBConnection
    
    menuFile.init appSetting, appSettingColFmt, conn
    menuFile.batchFile
    
    doAfterProcess

    Exit Function
err:
    
    Main.ShowErrorMessage
        
End Function

' =========================================================
' ▽ヘルプファイルの表示
'
' 概要　　　：
'
' =========================================================
Public Function SutShowHelpFile()

    On Error GoTo err
    
    ' 戻り値
    Dim ret As Long
    
    ' ヘルプファイルを表示する
    ret = WinAPI_Shell.ShellExecute(0 _
                           , "open" _
                           , VBUtil.concatFilePath(ThisWorkbook.path _
                                                 , ConstantsCommon.HELP_FILE) _
                           , "" _
                           , ThisWorkbook.path _
                           , 1)
    
    ' 戻り値が32以下の場合エラー
    If ret <= 32 Then
    
        VBUtil.showMessageBoxForInformation "ヘルプファイルを開くことができませんでした。", ConstantsCommon.APPLICATION_NAME
    
    End If

    Exit Function
    
err:
    
    Main.ShowErrorMessage

End Function

' =========================================================
' ▽バージョンの表示
'
' 概要　　　：
'
' =========================================================
Public Function SutShowVersion()

    On Error GoTo err
    
    frmSplash.Show vbModal

    Exit Function
    
err:
    
    Main.ShowErrorMessage

End Function

Private Function SutShowPopupCommon(ByVal index As Long)

    On Error GoTo err
    
    Dim appSetting As ValApplicationSettingShortcut
    Set appSetting = Main.getApplicationSettingShortcut
    
    Dim popupMenu As ValPopupmenu
    Set popupMenu = appSetting.popupMenuList.getItemByIndex(index)
    
    If Not popupMenu Is Nothing Then
    
        ' ポップアップコントロールを取得する
        Dim popup As CommandBar
        Set popup = popupMenu.commandBarPopup
        
        If Not popup Is Nothing Then
        
            ' 表示する
            popup.ShowPopup
        End If
    
        
    End If
    
    Exit Function
    
err:
    
    Main.ShowErrorMessage

End Function

Public Function SutShowPopup1()

    SutShowPopupCommon 1
End Function

Public Function SutShowPopup2()

    SutShowPopupCommon 2
End Function

Public Function SutShowPopup3()
    
    SutShowPopupCommon 3
End Function

Public Function SutShowPopup4()
    
    SutShowPopupCommon 4
End Function

Public Function SutShowPopup5()
    
    SutShowPopupCommon 5
End Function

Public Function SutShowPopup6()
    
    SutShowPopupCommon 6
End Function

Public Function SutShowPopup7()
    
    SutShowPopupCommon 7
End Function

Public Function SutShowPopup8()
    
    SutShowPopupCommon 8
End Function

Public Function SutShowPopup9()
    
    SutShowPopupCommon 9
End Function

Public Function SutShowPopup10()
    
    SutShowPopupCommon 10
End Function

' =========================================================
' ▽DBコネクション取得
'
' 概要　　　：DBコネクションを取得する
'
' =========================================================
Public Function getDBConnection() As Object

    ' DBコネクションが初期化されている場合
    If Not dbConn Is Nothing Then
    
        #If DEBUG_MODE = 1 Then

            Debug.Print "Connection Ver. " & dbConn.version
        #End If
    
        ' 接続されているか確認する
        If ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        
            ' エラーを投げる
            err.Raise ConstantsError.ERR_NUMBER_DISCONNECT_DB _
                    , _
                    , ConstantsError.ERR_DESC_DISCONNECT_DB
            
        End If
    
    ' DBコネクションが初期化されていない場合
    Else
        ' DB接続フォームを表示する
        SutConnectDB
        
        ' DBに接続されていない場合
        If dbConn Is Nothing Then
        
            ' エラーを投げる
            err.Raise ConstantsError.ERR_NUMBER_DISCONNECT_DB _
                    , _
                    , ConstantsError.ERR_DESC_DISCONNECT_DB
        End If
        
    End If

    ' 戻り値を設定する
    Set getDBConnection = dbConn

End Function

' =========================================================
' ▽アプリケーション設定情報取得
'
' 概要　　　：アプリケーション設定情報を取得する
'
' =========================================================
Public Function getApplicationSetting() As Object

    ' 初期化されている場合
    If Not applicationSetting Is Nothing Then
    
    
    ' 初期化されていない場合
    Else
    
        Set applicationSetting = New ValApplicationSetting
        applicationSetting.readForRegistry
        
    End If

    ' 戻り値を設定する
    Set getApplicationSetting = applicationSetting

End Function

' =========================================================
' ▽アプリケーション設定情報取得
'
' 概要　　　：アプリケーション設定情報を取得する
'
' =========================================================
Public Function getApplicationSettingShortcut() As Object

    ' 初期化されている場合
    If Not applicationSettingShortcut Is Nothing Then
    
    
    ' 初期化されていない場合
    Else
    
        Set applicationSettingShortcut = New ValApplicationSettingShortcut
        applicationSettingShortcut.init
        
    End If

    ' 戻り値を設定する
    Set getApplicationSettingShortcut = applicationSettingShortcut

End Function

' =========================================================
' ▽アプリケーション設定情報取得（カラム書式情報）
'
' 概要　　　：アプリケーション設定情報（カラム書式情報）を取得する
'
' =========================================================
Public Function getApplicationSettingColFormat() As Object

    ' 初期化されている場合
    If Not applicationSettingColFormat Is Nothing Then
    
    
    ' 初期化されていない場合
    Else
    
        Set applicationSettingColFormat = New ValApplicationSettingColFormat
        applicationSettingColFormat.init
        
    End If

    ' 戻り値を設定する
    Set getApplicationSettingColFormat = applicationSettingColFormat

End Function

' =========================================================
' ▽ワークブックのチェックを行う
'
' 概要　　　：
'
' =========================================================
Public Function validWorkbook()

    ' ブックオブジェクト
    Dim book As Workbook
    
    ' ブックオブジェクトを取得する
    Set book = ActiveWorkbook
    
    ' ブックオブジェクトのチェック
    If book Is Nothing Then
    
        err.Raise ERR_NUMBER_NON_ACTIVE_BOOK _
                , _
                , ERR_DESC_NON_ACTIVE_BOOK
            
    End If

End Function

' =========================================================
' ▽エラーメッセージを表示する
'
' 概要　　　：アプリケーションエラーかどうかを判定して
' 　　　　　　適切なＩＦでエラーメッセージを表示する。
'
' =========================================================
Public Function ShowErrorMessage()

    If ConstantsError.isApplicationError(err.Number) = True Then
    
        ' アプリケーションエラーが発生した場合、vbObjectErrorと固定数[512]を引いて、本来のエラー番号を算出する
        err.Number = err.Number - vbObjectError - 512
        ' エラー情報を表示する
        VBUtil.showMessageBoxForWarning "", ConstantsCommon.APPLICATION_NAME, err
    Else
    
        VBUtil.showMessageBoxForError ConstantsError.ERR_MSG_ERROR_LEVEL, ConstantsCommon.APPLICATION_NAME, err
    End If
    
End Function

' =========================================================
' ▽フォームポジションを復元する
'
' 概要　　　：formName フォームの識別子
' 　　　　　　formObj  フォームオブジェクト
'
' =========================================================
Public Function restoreFormPosition(ByVal formName As String _
                                  , ByRef formObj As Object)
    
    Dim formRect As New ValRectPt
    formRect.Left = formObj.Left
    formRect.Top = formObj.Top
    formRect.Width = formObj.Width
    formRect.Height = formObj.Height
    
    Dim formPosition As New ValFormPosition: formPosition.init formName
    Call formPosition.readForRegistry(formRect)

    formObj.Top = formRect.Top
    formObj.Left = formRect.Left

End Function

' =========================================================
' ▽フォームポジションを保存する
'
' 概要　　　：formName フォームの識別子
' 　　　　　　formObj  フォームオブジェクト
'
' =========================================================
Public Function storeFormPosition(ByVal formName As String _
                                , ByRef formObj As Object)

    Dim formRect As New ValRectPt
    formRect.Left = formObj.Left
    formRect.Top = formObj.Top
    formRect.Width = formObj.Width
    formRect.Height = formObj.Height
    
    Dim formPosition As New ValFormPosition: formPosition.init formName
    Call formPosition.writeForRegistry(formRect)

End Function

' =========================================================
' ▽ツールバーの初期化処理
'
' 概要　　　：
'
' =========================================================
Private Function initLoadingToolbar()

    On Error Resume Next
    
    ' カレントドライブとカレントディレクトリを切り替える
    ChDrive SutWorkbook.path
    ChDir SutWorkbook.path

    ' エクセルのバージョン
    Dim excelVer As ExcelVersion: excelVer = ExcelUtil.getExcelVersion
    
    ' コマンドバー
    Dim cb   As CommandBar
    
    Set cb = Application.CommandBars.Add( _
                            name:=ConstantsCommon.COMMANDBAR_MENU_NAME _
                          , Temporary:=True _
                          , position:=msoBarFloating)
        
    ' 既に追加されている場合は、変数cbがnothingになる
    ' 変数cbがnothingの場合は、処理を中断する
    If cb Is Nothing Then
    
        Exit Function
        
    End If
    
    ' -----------------------------------------------------------------------
    ' アプリケーションアイコンを設定する
    ' -----------------------------------------------------------------------
    ' アプリケーションアイコンボタン
    Dim appIcon As CommandBarButton
    
    ' Excel2002以降のプロパティ
    If excelVer >= Ver2002 Then
        
        Set appIcon = cb.Controls.Add(Type:=msoControlButton)
        
        With appIcon
        
            .Style = msoButtonIcon
            .OnAction = "Main.SutShowVersion"
            ' 削除対象から除外
            .Tag = ConstantsCommon.COMMANDBAR_DONT_DELETE_TARGET
            
            setCommandBarControlIcon appIcon _
                                   , RESOURCE_ICON.database
            
            ' ※DescriptionTextプロパティに明示的に空文字列を設定する
            ' 　ショートカットキーの機能リストに本コントロールは追加しない
            .DescriptionText = ""
            
        
        End With

    End If
    
    ' -----------------------------------------------------------------------
    ' 機能別にコマンドバーにコントロールを追加する
    ' -----------------------------------------------------------------------
    
    ' ***************************************************************
    ' DB接続
    ' ***************************************************************
    ' ファイルポップアップ
    Dim popFile                   As commandBarPopup
    ' ロードボタン
    Dim btnLoad                   As CommandBarButton
    ' アンロードボタン
    Dim btnUnload                 As CommandBarButton
    
    ' ファイルポップアップを追加する
    Set popFile = cb.Controls.Add(Type:=msoControlPopup)
    
    With popFile
        ' 削除対象から除外
        .Tag = ConstantsCommon.COMMANDBAR_DONT_DELETE_TARGET
        .Caption = "ファイル"
    End With
        
    ' ロードボタンをコマンドバーにボタンを追加する
    Set btnLoad = popFile.Controls.Add(Type:=msoControlButton)
    
    ' ロードボタンのプロパティを設定する
    With btnLoad
    
        .Style = msoButtonIconAndCaption
        .Caption = "アプリケーションの開始"
        .OnAction = "Main.SutLoad"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutLoad"
        
        ' ※DescriptionTextプロパティに明示的に空文字列を設定する
        ' 　ショートカットキーの機能リストに本コントロールは追加しない
        .DescriptionText = ""
        
    End With
        
    ' ロードボタンをコマンドバーにボタンを追加する
    Set btnUnload = popFile.Controls.Add(Type:=msoControlButton)
    
    ' ロードボタンのプロパティを設定する
    With btnUnload
    
        .Style = msoButtonIconAndCaption
        .Caption = "アプリケーションの終了"
        .OnAction = "Main.SutUnload"
        .enabled = False
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutUnload"
        
        ' ※DescriptionTextプロパティに明示的に空文字列を設定する
        ' 　ショートカットキーの機能リストに本コントロールは追加しない
        .DescriptionText = ""
        
    End With
    
    ' ***************************************************************
    
    cb.visible = True

    On Error GoTo 0
    
End Function


' =========================================================
' ▽ツールバーの初期化処理
'
' 概要　　　：
'
' =========================================================
Private Function initToolbar()

    On Error Resume Next
    
    ' ディレクトリを一時的に変更する
    ' アイコン設定のために SutYellow.dll を呼び出すために必要な処置
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ' エクセルのバージョン
    Dim excelVer As ExcelVersion: excelVer = ExcelUtil.getExcelVersion
    
    ' コマンドバー
    Dim cb   As CommandBar
    
    Set cb = Application.CommandBars(ConstantsCommon.COMMANDBAR_MENU_NAME)
    
    ' 取得に失敗した場合、変数cbがnothingになる
    ' initToolbar呼び出しの前提として、既にメニューが追加されている必要がある。
    ' 変数cbがnothingの場合は、処理を中断する
    If cb Is Nothing Then
    
        Exit Function
        
    End If

    ' -----------------------------------------------------------------------
    ' 機能別にコマンドバーにコントロールを追加する
    ' -----------------------------------------------------------------------
    
    ' ***************************************************************
    ' DB接続
    ' ***************************************************************
    ' DBポップアップ
    Dim popDB                     As commandBarPopup
    ' DB接続ボタン
    Dim btnDBConnect              As CommandBarButton
    ' DB切断ボタン
    Dim btnDBDisConnect           As CommandBarButton
    ' DB接続状態
    Dim btnDBInfo                 As CommandBarButton
    
    ' DBポップアップを追加する
    Set popDB = cb.Controls.Add(Type:=msoControlPopup)
    
    With popDB
    
        .Caption = "DB"
    End With
        
    ' DB接続ボタンをコマンドバーにボタンを追加する
    Set btnDBConnect = popDB.Controls.Add(Type:=msoControlButton)
    
    ' DB接続ボタンのプロパティを設定する
    With btnDBConnect
    
        .Style = msoButtonIconAndCaption
        .Caption = "接続"
        .DescriptionText = "DB接続"
        .OnAction = "Main.SutConnectDB"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutConnectDB"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDBConnect _
                                   , RESOURCE_ICON.databaseSetting
        End If
        
    End With
        
    ' DB切断ボタンをコマンドバーにボタンを追加する
    Set btnDBDisConnect = popDB.Controls.Add(Type:=msoControlButton)
    
    ' DB切断ボタンのプロパティを設定する
    With btnDBDisConnect
    
        .Style = msoButtonIconAndCaption
        .Caption = "切断"
        .DescriptionText = "DB切断"
        .OnAction = "Main.SutDisconnectDB"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDisconnectDB"
        .state = msoButtonDown ' DBが切断されていることが分かるように初期状態はボタン押下状態にする
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDBDisConnect _
                                   , RESOURCE_ICON.deleteDatabase
        End If
        
    End With
    
    ' ***************************************************************
    
        
    ' ***************************************************************
    ' テーブル
    ' ***************************************************************
    ' テーブルポップアップ
    Dim popTable                  As commandBarPopup
    ' DBエクスプローラ
    Dim btnDbExplorer             As CommandBarButton
    ' テーブル一覧ボタン
    Dim btnTableList              As CommandBarButton
    ' テーブル生成ウィザードボタン
    Dim btnTableCreateSheetWizard As CommandBarButton
    ' テーブル更新ボタン
    Dim btnTableUpdateSheetWizard As CommandBarButton
    
    ' テーブルポップアップを追加する
    Set popTable = cb.Controls.Add(Type:=msoControlPopup)
    
    With popTable
    
        .Caption = "テーブル"
    End With
    
    ' DBエクスプローラボタンをコマンドバーにボタンを追加する
    Set btnDbExplorer = popTable.Controls.Add(Type:=msoControlButton)
    
    ' DBエクスプローラボタンのプロパティを設定する
    With btnDbExplorer
    
        .Style = msoButtonIconAndCaption
        .Caption = "DBエクスプローラ"
        .DescriptionText = "DBエクスプローラ"
        .OnAction = "Main.SutShowDbExplorer"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowDbExplorer"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDbExplorer _
                                   , RESOURCE_ICON.databaseSearch
        End If
        
    End With
    
    ' テーブル一覧ボタンをコマンドバーにボタンを追加する
    Set btnTableList = popTable.Controls.Add(Type:=msoControlButton)
    
    ' テーブル一覧ボタンのプロパティを設定する
    With btnTableList
    
        .Style = msoButtonIconAndCaption
        .Caption = "テーブルシート一覧"
        .DescriptionText = "テーブルシート一覧"
        .OnAction = "Main.SutShowTableSheetList"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowTableSheetList"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnTableList _
                                   , RESOURCE_ICON.searchWindow
        End If
        
    End With
    
    ' テーブル生成ウィザードボタンをコマンドバーにボタンを追加する
    Set btnTableCreateSheetWizard = popTable.Controls.Add(Type:=msoControlButton)
    
    ' テーブル生成ウィザードボタンのプロパティを設定する
    With btnTableCreateSheetWizard
    
        .Style = msoButtonIconAndCaption
        .Caption = "テーブルシート作成"
        .DescriptionText = "テーブルシート作成"
        .OnAction = "Main.SutCreateTableSheet"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutCreateTableSheet"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnTableCreateSheetWizard _
                                   , RESOURCE_ICON.addFolder
        End If
        
    End With
    
    ' テーブル更新ボタンをコマンドバーにボタンを追加する
    Set btnTableUpdateSheetWizard = popTable.Controls.Add(Type:=msoControlButton)
    
    ' テーブル更新ボタンのプロパティを設定する
    With btnTableUpdateSheetWizard
    
        .Style = msoButtonIconAndCaption
        .Caption = "テーブルシート更新"
        .DescriptionText = "テーブルシート更新"
        .OnAction = "Main.SutUpdateTableSheet"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutUpdateTableSheet"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnTableUpdateSheetWizard _
                                   , RESOURCE_ICON.windowImport
        End If
    End With

    ' ***************************************************************
    
    
    ' ***************************************************************
    ' データ
    ' ***************************************************************
    ' データポップアップ
    Dim popData                   As commandBarPopup
    
    ' データポップアップを追加する
    Set popData = cb.Controls.Add(Type:=msoControlPopup)
    
    With popData
    
        .Caption = "データ"
    End With
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' INSERT + UPDATE
    ' ***************************************************************
    ' INSERT + UPDATEポップアップ
    Dim popInsertUpdate                 As commandBarPopup
    ' INSERT + UPDATEボタン
    Dim btnInsertUpdate                 As CommandBarButton
    ' INSERT + UPDATE（範囲選択）ボタン
    Dim btnInsertUpdateSelected         As CommandBarButton
    
    ' INSERT + UPDATEポップアップを追加する
    Set popInsertUpdate = popData.Controls.Add(Type:=msoControlPopup)
    
    With popInsertUpdate
        
        .Caption = "INSERT + UPDATE"
    End With
    
    ' INSERT + UPDATEボタンをコマンドバーにボタンを追加する
    Set btnInsertUpdate = popInsertUpdate.Controls.Add(Type:=msoControlButton)
    
    ' INSERT + UPDATEボタンのプロパティを設定する
    With btnInsertUpdate
    
        .Style = msoButtonIconAndCaption
        .Caption = "全て"
        .DescriptionText = "INSERT + UPDATE - 全て"
        .OnAction = "Main.SutInsertUpdateAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutInsertUpdateAll"

        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnInsertUpdate _
                                   , RESOURCE_ICON.Add
        End If

    End With
    
    ' INSERT + UPDATE（範囲選択）ボタンをコマンドバーにボタンを追加する
    Set btnInsertUpdateSelected = popInsertUpdate.Controls.Add(Type:=msoControlButton)
    
    ' INSERT + UPDATE（範囲選択）ボタンのプロパティを設定する
    With btnInsertUpdateSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "範囲選択"
        .DescriptionText = "INSERT + UPDATE - 範囲選択"
        .OnAction = "Main.SutInsertUpdateSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutInsertUpdateSelection"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnInsertUpdateSelected _
                                   , RESOURCE_ICON.areaAdd
        End If
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' INSERT
    ' ***************************************************************
    ' INSERTポップアップ
    Dim popInsert                 As commandBarPopup
    ' INSERTボタン
    Dim btnInsert                 As CommandBarButton
    ' INSERT（範囲選択）ボタン
    Dim btnInsertSelected         As CommandBarButton
    
    ' INSERTポップアップを追加する
    Set popInsert = popData.Controls.Add(Type:=msoControlPopup)
    
    With popInsert
        
        .Caption = "INSERT"
    End With
    
    ' INSERTボタンをコマンドバーにボタンを追加する
    Set btnInsert = popInsert.Controls.Add(Type:=msoControlButton)
    
    ' INSERTボタンのプロパティを設定する
    With btnInsert
    
        .Style = msoButtonIconAndCaption
        .Caption = "全て"
        .DescriptionText = "INSERT - 全て"
        .OnAction = "Main.SutInsertAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutInsertAll"

        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnInsert _
                                   , RESOURCE_ICON.Add
        End If

    End With
    
    ' INSERT（範囲選択）ボタンをコマンドバーにボタンを追加する
    Set btnInsertSelected = popInsert.Controls.Add(Type:=msoControlButton)
    
    ' INSERT（範囲選択）ボタンのプロパティを設定する
    With btnInsertSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "範囲選択"
        .DescriptionText = "INSERT - 範囲選択"
        .OnAction = "Main.SutInsertSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutInsertSelection"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnInsertSelected _
                                   , RESOURCE_ICON.areaAdd
        End If
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' UPDATE
    ' ***************************************************************
    ' UPDATEポップアップ
    Dim popUpdate                 As commandBarPopup
    ' UPDATEボタン
    Dim btnupdate                 As CommandBarButton
    ' UPDATE（範囲選択）ボタン
    Dim btnUpdateSelected         As CommandBarButton
    
    ' テーブルポップアップを追加する
    Set popUpdate = popData.Controls.Add(Type:=msoControlPopup)
    
    With popUpdate
    
        .Caption = "UPDATE"
    End With
    
    ' UPDATEボタンをコマンドバーにボタンを追加する
    Set btnupdate = popUpdate.Controls.Add(Type:=msoControlButton)
    
    ' UPDATEボタンのプロパティを設定する
    With btnupdate
    
        .Style = msoButtonIconAndCaption
        .Caption = "全て"
        .DescriptionText = "UPDATE - 全て"
        .OnAction = "Main.SutUpdateAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutUpdateAll"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnupdate _
                                   , RESOURCE_ICON.Edit
        End If
    End With
    
    ' UPDATE（範囲選択）ボタンをコマンドバーにボタンを追加する
    Set btnUpdateSelected = popUpdate.Controls.Add(Type:=msoControlButton)
    
    ' UPDATE（範囲選択）ボタンのプロパティを設定する
    With btnUpdateSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "範囲選択"
        .DescriptionText = "UPDATE - 範囲選択"
        .OnAction = "Main.SutUpdateSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutUpdateSelection"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnUpdateSelected _
                                   , RESOURCE_ICON.areaEdit
        End If
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' DELETE
    ' ***************************************************************
    ' DELETEポップアップ
    Dim popDelete                 As commandBarPopup
    ' DELETEボタン
    Dim btnDelete                 As CommandBarButton
    ' DELETE（範囲選択）ボタン
    Dim btnDeleteSelected         As CommandBarButton
    ' DELETE（テーブル上の全レコード）ボタン
    Dim btnDeleteAllOfTable       As CommandBarButton
    
    ' テーブルポップアップを追加する
    Set popDelete = popData.Controls.Add(Type:=msoControlPopup)
    
    With popDelete
    
        .Caption = "DELETE"
    End With
    
    ' DELETEボタンをコマンドバーにボタンを追加する
    Set btnDelete = popDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETEボタンのプロパティを設定する
    With btnDelete
    
        .Style = msoButtonIconAndCaption
        .Caption = "全て"
        .DescriptionText = "DELETE - 全て"
        .OnAction = "Main.SutDeleteAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDeleteAll"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDelete _
                                   , RESOURCE_ICON.remove
        End If
    End With
    
    ' DELETE（範囲選択）ボタンをコマンドバーにボタンを追加する
    Set btnDeleteSelected = popDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETE（範囲選択）ボタンのプロパティを設定する
    With btnDeleteSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "範囲選択"
        .DescriptionText = "DELETE - 範囲選択"
        .OnAction = "Main.SutDeleteSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDeleteSelection"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDeleteSelected _
                                   , RESOURCE_ICON.areaRemove
        End If
    End With
    
    ' DELETE（テーブル上の全レコード）ボタンをコマンドバーにボタンを追加する
    Set btnDeleteAllOfTable = popDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETE（テーブル上の全レコード）ボタンのプロパティを設定する
    With btnDeleteAllOfTable
    
        .Style = msoButtonIconAndCaption
        .Caption = "テーブル上の全レコード"
        .DescriptionText = "DELETE - テーブル上の全レコード"
        .OnAction = "Main.SutDeleteAllOfTable"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDeleteAllOfTable"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnDeleteAllOfTable _
                                   , RESOURCE_ICON.bug
        End If
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' SELECT
    ' ***************************************************************
    ' SELECTポップアップ
    Dim popSelect                 As commandBarPopup
    ' SELECTボタン
    Dim btnSelect                 As CommandBarButton
    ' SELECT（条件指定）ボタン
    Dim btnSelectSelected         As CommandBarButton
    ' SELECT（前回の条件で実行）ボタン
    Dim btnSelectReExecute        As CommandBarButton
    
    ' テーブルポップアップを追加する
    Set popSelect = popData.Controls.Add(Type:=msoControlPopup)
    
    With popSelect
    
        .Caption = "SELECT"
    End With
    
    ' SELECTボタンをコマンドバーにボタンを追加する
    Set btnSelect = popSelect.Controls.Add(Type:=msoControlButton)
    
    ' SELECTボタンのプロパティを設定する
    With btnSelect
    
        .Style = msoButtonIconAndCaption
        .Caption = "全て"
        .DescriptionText = "SELECT - 全て"
        .OnAction = "Main.SutSelectAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSelectAll"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnSelect _
                                   , RESOURCE_ICON.Search
        End If
    End With
    
    ' SELECT（条件指定）ボタンをコマンドバーにボタンを追加する
    Set btnSelectSelected = popSelect.Controls.Add(Type:=msoControlButton)
    
    ' SELECT（条件指定）ボタンのプロパティを設定する
    With btnSelectSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "条件指定"
        .DescriptionText = "SELECT - 条件指定"
        .OnAction = "Main.SutSelectCondition"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSelectCondition"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnSelectSelected _
                                   , RESOURCE_ICON.areaSearch
        End If
    End With
    
    ' SELECT（再実行）ボタンをコマンドバーにボタンを追加する
    Set btnSelectReExecute = popSelect.Controls.Add(Type:=msoControlButton)
    
    ' SELECT（再実行）ボタンのプロパティを設定する
    With btnSelectReExecute
    
        .Style = msoButtonIconAndCaption
        .Caption = "再実行"
        .DescriptionText = "SELECT - 再実行"
        .OnAction = "Main.SutSelectReExecute"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSelectReExecute"
        
    End With
    
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' クエリエディタ（ !!! 未実装 !!! ）
    ' ***************************************************************
'    ' クエリエディタの追加
'    Dim btnQueryEditor As CommandBarButton
'
'    ' クエリエディタボタンをコマンドバーにボタンを追加する
'    Set btnQueryEditor = popData.Controls.Add(Type:=msoControlButton)
'
'    ' クエリエディタボタンのプロパティを設定する
'    With btnQueryEditor
'
'        .Style = msoButtonIconAndCaption
'        .Caption = "クエリエディタ"
'        .DescriptionText = "クエリエディタ"
'        .OnAction = "Main.SutQueryEditor"
'        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutQueryEditor"
'
'        ' Excel2002以降のプロパティ
'        If excelVer >= Ver2002 Then
'            setCommandBarControlIcon btnQueryEditor _
'                                   , RESOURCE_ICON.EDIT _
'                                   , RESOURCE_ICON.EDIT_MASK
'        End If
'
'    End With
    
    ' ***************************************************************
    ' 一括クエリ
    ' ***************************************************************
    ' 一括クエリの追加
    Dim btnQueryBatch As CommandBarButton
    
    ' 行の追加ボタンをコマンドバーにボタンを追加する
    Set btnQueryBatch = popData.Controls.Add(Type:=msoControlButton)
    
    ' SELECTボタンのプロパティを設定する
    With btnQueryBatch
    
        .Style = msoButtonIconAndCaption
        .Caption = "クエリ一括実行"
        .DescriptionText = "クエリ一括実行"
        .OnAction = "Main.SutQueryBatch"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutQueryBatch"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnQueryBatch _
                                   , RESOURCE_ICON.Forward
        End If
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' クエリ結果
    ' ***************************************************************
    ' クエリ結果
    Dim btnQueryResult             As CommandBarButton
    
    ' クエリ結果ボタンをコマンドバーにボタンを追加する
    Set btnQueryResult = popData.Controls.Add(Type:=msoControlButton)
    
    ' クエリ結果ボタンのプロパティを設定する
    With btnQueryResult
    
        .Style = msoButtonIconAndCaption
        .Caption = "クエリ結果"
        .DescriptionText = "クエリ結果"
        .OnAction = "Main.SutShowQueryResult"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowQueryResult"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnQueryResult _
                                   , RESOURCE_ICON.Paste
        End If
        
    End With
    
    ' ***************************************************************
    ' 行の追加・削除
    ' ***************************************************************
    ' 行の追加
    Dim btnRecAdd As CommandBarButton
    
    ' 行の追加ボタンをコマンドバーにボタンを追加する
    Set btnRecAdd = popData.Controls.Add(Type:=msoControlButton)
    
    ' SELECTボタンのプロパティを設定する
    With btnRecAdd
    
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "行の追加・削除"
        .DescriptionText = "行の追加・削除"
        .OnAction = "Main.SutRecordAdd"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutRecordAdd"
        
    End With
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' Diff
    ' ***************************************************************
    ' Diffポップアップ
    Dim popDiff                   As commandBarPopup
    
    ' Diffポップアップを追加する
    Set popDiff = cb.Controls.Add(Type:=msoControlPopup)
    
    With popDiff
    
        .Caption = "Diff"
    End With
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' DBスナップショット取得フォーム呼び出し
    ' ***************************************************************
    ' スナップショット取得
    Dim btnShowDBSnapshot As CommandBarButton
    
    ' スナップショット取得ボタンをコマンドバーにボタンを追加する
    Set btnShowDBSnapshot = popDiff.Controls.Add(Type:=msoControlButton)
    
    ' スナップショット取得ボタンのプロパティを設定する
    With btnShowDBSnapshot
    
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "スナップショット取得・比較"
        .DescriptionText = "スナップショット取得・比較"
        .OnAction = "Main.SutShowSnapshot"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowSnapshot"
        
    End With
    
    ' ***************************************************************
    ' DBスナップショットSQL定義シート追加
    ' ***************************************************************
    ' スナップショットSQLシート追加
    Dim btnNewSheetDataSnapshotSqlDefine As CommandBarButton
    
    ' スナップショットSQLシート追加ボタンをコマンドバーにボタンを追加する
    Set btnNewSheetDataSnapshotSqlDefine = popDiff.Controls.Add(Type:=msoControlButton)
    
    ' スナップショットSQLシート追加ボタンのプロパティを設定する
    With btnNewSheetDataSnapshotSqlDefine
    
        .BeginGroup = False
        .Style = msoButtonIconAndCaption
        .Caption = "スナップショットSQLシート追加"
        .DescriptionText = "スナップショットSQLシート追加"
        .OnAction = "Main.SutCreateNewSheetSnapSqlDefine"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutCreateNewSheetSnapSqlDefine"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' ファイル
    ' ***************************************************************
    ' ファイルポップアップ
    Dim popFile                   As commandBarPopup
    
    ' ファイルポップアップを追加する
    Set popFile = cb.Controls.Add(Type:=msoControlPopup)
    
    With popFile
    
        .Caption = "ファイル"
    End With
    ' ***************************************************************
    
    ' ***************************************************************
    ' INSERT出力
    ' ***************************************************************
    ' INSERTポップアップ
    Dim popFileInsert                 As commandBarPopup
    ' INSERTボタン
    Dim btnFileInsert                 As CommandBarButton
    ' INSERT（範囲選択）ボタン
    Dim btnFileInsertSelected         As CommandBarButton
    
    ' INSERTポップアップを追加する
    Set popFileInsert = popFile.Controls.Add(Type:=msoControlPopup)
    
    With popFileInsert
        
        .Caption = "INSERT SQL"
    End With
    
    ' INSERTボタンをコマンドバーにボタンを追加する
    Set btnFileInsert = popFileInsert.Controls.Add(Type:=msoControlButton)
    
    ' INSERTボタンのプロパティを設定する
    With btnFileInsert
    
        .Style = msoButtonIconAndCaption
        .Caption = "全て"
        .DescriptionText = "ファイル出力 INSERT SQL - 全て"
        .OnAction = "Main.SutFileInsertAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileInsertAll"
        
    End With
    
    ' INSERT（範囲選択）ボタンをコマンドバーにボタンを追加する
    Set btnFileInsertSelected = popFileInsert.Controls.Add(Type:=msoControlButton)
    
    ' INSERT（範囲選択）ボタンのプロパティを設定する
    With btnFileInsertSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "範囲選択"
        .DescriptionText = "ファイル出力 INSERT SQL - 範囲選択"
        .OnAction = "Main.SutFileInsertSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileInsertSelection"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' UPDATE出力
    ' ***************************************************************
    ' UPDATEポップアップ
    Dim popFileUpdate                 As commandBarPopup
    ' UPDATEボタン
    Dim btnFileUpdate                 As CommandBarButton
    ' UPDATE（範囲選択）ボタン
    Dim btnFileUpdateSelected         As CommandBarButton
    
    ' UPDATEポップアップを追加する
    Set popFileUpdate = popFile.Controls.Add(Type:=msoControlPopup)
    
    With popFileUpdate
        
        .Caption = "UPDATE SQL"
    End With
    
    ' UPDATEボタンをコマンドバーにボタンを追加する
    Set btnFileUpdate = popFileUpdate.Controls.Add(Type:=msoControlButton)
    
    ' UPDATEボタンのプロパティを設定する
    With btnFileUpdate
    
        .Style = msoButtonIconAndCaption
        .Caption = "全て"
        .DescriptionText = "ファイル出力 UPDATE SQL - 全て"
        .OnAction = "Main.SutFileUpdateAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileUpdateAll"
        
    End With
    
    ' UPDATE（範囲選択）ボタンをコマンドバーにボタンを追加する
    Set btnFileUpdateSelected = popFileUpdate.Controls.Add(Type:=msoControlButton)
    
    ' UPDATE（範囲選択）ボタンのプロパティを設定する
    With btnFileUpdateSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "範囲選択"
        .DescriptionText = "ファイル出力 UPDATE SQL - 範囲選択"
        .OnAction = "Main.SutFileUpdateSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileUpdateSelection"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' DELETE出力
    ' ***************************************************************
    ' DELETEポップアップ
    Dim popFileDelete                 As commandBarPopup
    ' DELETEボタン
    Dim btnFileDelete                 As CommandBarButton
    ' DELETE（範囲選択）ボタン
    Dim btnFileDeleteSelected         As CommandBarButton
    
    ' DELETEポップアップを追加する
    Set popFileDelete = popFile.Controls.Add(Type:=msoControlPopup)
    
    With popFileDelete
        
        .Caption = "DELETE SQL"
    End With
    
    ' DELETEボタンをコマンドバーにボタンを追加する
    Set btnFileDelete = popFileDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETEボタンのプロパティを設定する
    With btnFileDelete
    
        .Style = msoButtonIconAndCaption
        .Caption = "全て"
        .DescriptionText = "ファイル出力 DELETE SQL - 全て"
        .OnAction = "Main.SutFileDeleteAll"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileDeleteAll"
        
    End With
    
    ' DELETE（範囲選択）ボタンをコマンドバーにボタンを追加する
    Set btnFileDeleteSelected = popFileDelete.Controls.Add(Type:=msoControlButton)
    
    ' DELETE（範囲選択）ボタンのプロパティを設定する
    With btnFileDeleteSelected
    
        .Style = msoButtonIconAndCaption
        .Caption = "範囲選択"
        .DescriptionText = "ファイル出力 DELETE SQL - 範囲選択"
        .OnAction = "Main.SutFileDeleteSelection"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileDeleteSelection"
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' 一括ファイル出力
    ' ***************************************************************
    ' DELETE（範囲選択）ボタン
    Dim btnFileBatch         As CommandBarButton
    
    ' DELETEボタンをコマンドバーにボタンを追加する
    Set btnFileBatch = popFile.Controls.Add(Type:=msoControlButton)
    
    ' DELETEボタンのプロパティを設定する
    With btnFileBatch
    
        .Style = msoButtonIconAndCaption
        .Caption = "一括出力"
        .DescriptionText = "ファイル一括出力"
        .OnAction = "Main.SutFileBatch"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutFileBatch"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnFileBatch _
                                   , RESOURCE_ICON.Forward
        End If
        
    End With
    
    ' ***************************************************************
    
    ' ***************************************************************
    ' ツール
    ' ***************************************************************
    ' ツールポップアップ
    Dim popTool             As commandBarPopup
    ' クエリパラメーボタン
    Dim btnQueryParameter As CommandBarButton
    ' オプションボタン
    Dim btnOption           As CommandBarButton
    ' 右クリックメニューのカスタマイズボタン
    Dim btnRClickMenuCustom As CommandBarButton
    ' ショートカットキーの割り当てボタン
    Dim btnShortCutKey      As CommandBarButton
    ' ポップアップメニューのカスタマイズボタン
    Dim btnPopupKey         As CommandBarButton
    
    ' ツールポップアップを追加する
    Set popTool = cb.Controls.Add(Type:=msoControlPopup)
    
    With popTool
    
        .Caption = "ツール"
    End With
    
    ' クエリパラメーボタンをコマンドバーにボタンを追加する
    Set btnQueryParameter = popTool.Controls.Add(Type:=msoControlButton)
    
    ' クエリパラメーボタンのプロパティを設定する
    With btnQueryParameter
    
        .Style = msoButtonIconAndCaption
        .Caption = "クエリパラメータ"
        .DescriptionText = "クエリパラメータ"
        .OnAction = "Main.SutSettingQueryParameter"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingQueryParameter"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            'setCommandBarControlIcon btnOption _
            '                       , RESOURCE_ICON.SETTINGS _
            '                       , RESOURCE_ICON.SETTINGS_MASK
        End If
    End With

    ' オプションボタンをコマンドバーにボタンを追加する
    Set btnOption = popTool.Controls.Add(Type:=msoControlButton)
    
    ' オプションボタンのプロパティを設定する
    With btnOption
    
        .Style = msoButtonIconAndCaption
        .Caption = "オプション"
        .DescriptionText = "オプション"
        .OnAction = "Main.SutSettingOption"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingOption"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnOption _
                                   , RESOURCE_ICON.settings
        End If
    End With
    
    ' 右クリックメニューのカスタマイズボタンをコマンドバーにボタンを追加する
    Set btnRClickMenuCustom = popTool.Controls.Add(Type:=msoControlButton)
    
    ' 右クリックメニューのカスタマイズボタンのプロパティを設定する
    With btnRClickMenuCustom
    
        .BeginGroup = True
        .Style = msoButtonIconAndCaption
        .Caption = "右クリックメニューの設定"
        .DescriptionText = "右クリックメニューの設定"
        .OnAction = "Main.SutSettingRClickMenu"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingRClickMenu"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnRClickMenuCustom _
                                   , RESOURCE_ICON.flagRed
        End If
    End With
    
    ' ショートカットキーの割り当てボタンをコマンドバーにボタンを追加する
    Set btnShortCutKey = popTool.Controls.Add(Type:=msoControlButton)
    
    ' ショートカットキーの割り当てボタンのプロパティを設定する
    With btnShortCutKey
    
        .Style = msoButtonIconAndCaption
        .Caption = "ショートカットキーの設定"
        .DescriptionText = "ショートカットキーの設定"
        .OnAction = "Main.SutSettingShortCutKey"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingShortCutKey"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnShortCutKey _
                                   , RESOURCE_ICON.flagGreen
        End If
    End With
    
    ' ショートカットキーの割り当てボタンをコマンドバーにボタンを追加する
    Set btnPopupKey = popTool.Controls.Add(Type:=msoControlButton)
    
    ' ショートカットキーの割り当てボタンのプロパティを設定する
    With btnPopupKey
    
        .Style = msoButtonIconAndCaption
        .Caption = "ポップアップメニューの設定"
        .DescriptionText = "ポップアップメニューの設定"
        .OnAction = "Main.SutSettingPopupMenu"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutSettingPopupMenu"
        
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnPopupKey _
                                   , RESOURCE_ICON.flagBlue
        End If
    End With
    
    
    ' ***************************************************************
    
    
    ' ***************************************************************
    ' ヘルプ
    ' ***************************************************************
    ' ヘルプポップアップ
    Dim popHelp           As commandBarPopup
    ' ヘルプ
    Dim btnHelp           As CommandBarButton
    ' ライセンス
    Dim btnLicence        As CommandBarButton
    ' バージョン
    Dim btnVersion        As CommandBarButton
    
    ' ツールポップアップを追加する
    Set popHelp = cb.Controls.Add(Type:=msoControlPopup)
    
    With popHelp
    
        .Caption = "ヘルプ"
    End With
    
    ' ヘルプボタンをコマンドバーにボタンを追加する
    Set btnHelp = popHelp.Controls.Add(Type:=msoControlButton)
    
    ' ヘルプボタンのプロパティを設定する
    With btnHelp
    
        .Style = msoButtonIconAndCaption
        .Caption = "Sutヘルプ"
        .DescriptionText = "Sutヘルプ"
        .OnAction = "Main.SutShowHelpFile"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowHelpFile"
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnHelp _
                                   , RESOURCE_ICON.book
        End If
        
    End With
    
    ' バージョン情報をコマンドバーにボタンを追加する
    Set btnVersion = popHelp.Controls.Add(Type:=msoControlButton)
    
    ' バージョン情報ボタンのプロパティを設定する
    With btnVersion
    
        .Style = msoButtonIconAndCaption
        .Caption = "バージョン情報"
        .DescriptionText = "バージョン情報"
        .OnAction = "Main.SutShowVersion"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutShowVersion"
        ' Excel2002以降のプロパティ
        If excelVer >= Ver2002 Then
            setCommandBarControlIcon btnVersion _
                                   , RESOURCE_ICON.alertMessage
        End If
        
    End With
    
    ' ***************************************************************
    
    ' ロードボタンを押下不可にする
    Dim btnLoad As CommandBarButton
    Set btnLoad = cb.FindControl(Tag:=COMMANDBAR_CONTROL_BASE_ID & "Main.SutLoad", recursive:=True)
    
    If Not btnLoad Is Nothing Then
    
        btnLoad.enabled = False
    End If

    ' アンロードボタンを押下可能にする
    Dim btnUnload As CommandBarButton
    Set btnUnload = cb.FindControl(Tag:=COMMANDBAR_CONTROL_BASE_ID & "Main.SutUnload", recursive:=True)
    
    If Not btnUnload Is Nothing Then
    
        btnUnload.enabled = True
    End If
    
    cb.visible = True

    On Error GoTo 0
    
End Function

' =========================================================
' ▽ツールバーの削除処理
'
' 概要　　　：
'
' =========================================================
Private Function deleteToolbar()

    On Error Resume Next
    
    ' コマンドバー
    Dim cb   As CommandBar
    
    Set cb = Application.CommandBars.item(ConstantsCommon.COMMANDBAR_MENU_NAME)
        
    ' 取得に失敗した場合、変数cbがnothingになる
    ' 変数cbがnothingの場合は、処理を中断する
    If cb Is Nothing Then
    
        Exit Function
        
    End If
    
    cb.delete
    
    On Error GoTo 0
    
End Function

' =========================================================
' ▽ツールバーの削除処理（特定のメニューは残す）
'
' 概要　　　：
'
' =========================================================
Private Function deleteToolbarExcludeSomeItems()

    On Error Resume Next
    
    ' コマンドバー
    Dim cb   As CommandBar
    
    Set cb = Application.CommandBars.item(ConstantsCommon.COMMANDBAR_MENU_NAME)
        
    ' 取得に失敗した場合、変数cbがnothingになる
    ' 変数cbがnothingの場合は、処理を中断する
    If cb Is Nothing Then
    
        Exit Function
        
    End If
    
    Dim ctl As commandBarControl
    
    For Each ctl In cb.Controls
    
        If ctl.Tag <> ConstantsCommon.COMMANDBAR_DONT_DELETE_TARGET Then
        
            ' コントロールを削除する
            ctl.delete
        End If
    Next
    
    ' ロードボタンを押下可能にする
    Dim btnLoad As CommandBarButton
    Set btnLoad = cb.FindControl(Tag:=COMMANDBAR_CONTROL_BASE_ID & "Main.SutLoad", recursive:=True)
    
    If Not btnLoad Is Nothing Then
    
        btnLoad.enabled = True
    End If

    ' アンロードボタンを押下不可にする
    Dim btnUnload As CommandBarButton
    Set btnUnload = cb.FindControl(Tag:=COMMANDBAR_CONTROL_BASE_ID & "Main.SutUnload", recursive:=True)
    
    If Not btnUnload Is Nothing Then
    
        btnUnload.enabled = False
    End If
    
    On Error GoTo 0
    
End Function

Private Function setCommandBarControlIcon(ByVal control As Object _
                                        , ByVal icon As RESOURCE_ICON)
                                   
    'control.Picture = LoadPicture("")
    'control.Mask = LoadPicture("")

End Function

Private Function doAfterProcess()

    On Error Resume Next
    
    ' 処理終了後に、Excelウィンドウがアクティブにならずに、他のウィンドウがアクティブになる事象を確認
    ' これを受けて、以下のように、現在のアクティブブックをアクティブにするように明示的に指定する
    Application.ActiveWindow.activate

    On Error GoTo 0

End Function

Private Sub changeDbConnectStatus(ByVal conn As Boolean)

    ' コマンドバー
    Dim cb   As CommandBar
    Set cb = Application.CommandBars.item(ConstantsCommon.COMMANDBAR_MENU_NAME)
        
    ' 既に追加されている場合は、変数cbがnothingになる
    ' 変数cbがnothingの場合は、処理を中断する
    If cb Is Nothing Then
        Exit Sub
    End If
    
    Dim btnConn    As CommandBarButton
    Dim btnDisconn As CommandBarButton
    
    Set btnConn = cb.FindControl(Tag:=ConstantsCommon.COMMANDBAR_CONTROL_BASE_ID & "Main.SutConnectDB", recursive:=True)
    Set btnDisconn = cb.FindControl(Tag:=ConstantsCommon.COMMANDBAR_CONTROL_BASE_ID & "Main.SutDisconnectDB", recursive:=True)
    
    If _
        Not btnConn Is Nothing And _
        Not btnDisconn Is Nothing Then
    
        If conn = True Then
        
            btnConn.state = msoButtonDown
            btnDisconn.state = msoButtonUp
        Else
        
            btnConn.state = msoButtonUp
            btnDisconn.state = msoButtonDown
        End If
    End If
    
End Sub

' =========================================================
' ▽バージョン情報を取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：バージョン情報
' 特記事項　：
'
' =========================================================
Public Function getVersionInfo() As String
    
    Dim version     As String
    Dim machineName As String
    
    version = ConstantsCommon.version
    
    #If VBA7 And Win64 Then
        machineName = "64bit"
    #Else
        machineName = "32bit"
    #End If
    
    #If DEBUG_MODE = "1" Then
        machineName = machineName & " !!! IS DEBUG MODE"
    #End If
    
    getVersionInfo = machineName & " - ver " & version
    
End Function
