VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnect 
   Caption         =   "DB接続"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   OleObjectBlob   =   "frmDBConnect.frx":0000
End
Attribute VB_Name = "frmDBConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB接続を行うフォーム
'
' 作成者　：Ison
' 履歴　　：2008/09/06　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽接続するDBが決定した際に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event ok(ByVal connStr As String, ByVal connSimpleStr As String, ByVal connectInfo As ValDBConnectInfo)

' =========================================================
' ▽DBの接続がキャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' 接続文字列 配列インデックス最小値
Private Const CONNECT_STR_MIN As Long = 1
' 接続文字列 配列インデックス最大値
Private Const CONNECT_STR_MAX As Long = 5

' コントロール有効フラグ インデックス データソース
Private Const CONTROL_ENABLE_IDX_DATASOURCE As Long = 1
' コントロール有効フラグ インデックス ホスト
Private Const CONTROL_ENABLE_IDX_HOST       As Long = 2
' コントロール有効フラグ インデックス DB
Private Const CONTROL_ENABLE_IDX_DB         As Long = 3
' コントロール有効フラグ インデックス ポート
Private Const CONTROL_ENABLE_IDX_PORT       As Long = 4
' コントロール有効フラグ インデックス ユーザ
Private Const CONTROL_ENABLE_IDX_USER       As Long = 5
' コントロール有効フラグ インデックス パスワード
Private Const CONTROL_ENABLE_IDX_PASSWORD   As Long = 6
' コントロール有効フラグ インデックス ファイル選択ボタン
Private Const CONTROL_ENABLE_IDX_FILE_SELECT   As Long = 7

' 接続文字列
Private connectStr(1 To 5) As String
' プロバイダラベル
Private providerLabel(1 To 5) As String
' デフォルトポート番号
Private defaultPort(1 To 5) As String
' コントロール有効フラグ
Private controlEnable(1 To 5, 1 To 7) As Boolean

' ---------------------------------------------------------
' レジストリファイルキー
' ---------------------------------------------------------
Private Const REG_SUB_KEY_DB_CONNECT As String = "db_connect"

Private WithEvents history  As frmDBConnectHistory
Attribute history.VB_VarHelpID = -1
Private WithEvents favorite As frmDBConnectFavorite
Attribute favorite.VB_VarHelpID = -1

Private dbConnectListener As IDbConnectListener

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal                    モーダルまたはモードレス表示指定
'     　　　：dbConnectInfo            DB接続情報
'     　　　：dbConnectListener        DB接続リスナー
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, _
                    Optional ByVal dbConnectInfo As ValDBConnectInfo = Nothing, _
                    Optional ByVal dbConnectListener_ As IDbConnectListener = Nothing)
    
    Set dbConnectListener = dbConnectListener_

    If Not dbConnectListener_ Is Nothing Then
        cmdHistoryChoice.visible = False
        cmdFavoriteSave.visible = False
        cmdFavoriteLoad.visible = False
    Else
        cmdHistoryChoice.visible = True
        cmdFavoriteSave.visible = True
        cmdFavoriteLoad.visible = True
    End If
    
    ' DB接続情報の初期値を設定する
    setDbConnectInfo dbConnectInfo

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
End Sub

Private Function getDbConnectInfo() As ValDBConnectInfo

    ' DB接続情報を生成しコントロールから情報を集め設定する
    Dim connectInfo As New ValDBConnectInfo
    connectInfo.type_ = cboDBType.value
    connectInfo.dsn = cboDataSourceName.value
    connectInfo.host = txtHost.value
    connectInfo.port = txtPort.value
    connectInfo.db = txtDB.value
    connectInfo.user = txtUser.value
    connectInfo.password = txtPassword.value
    connectInfo.option_ = txtOption.value
    
    Set getDbConnectInfo = connectInfo

End Function

Private Sub setDbConnectInfo(ByRef connectInfo As ValDBConnectInfo)

    On Error Resume Next
    
    cboDBType.value = connectInfo.type_
    cboDataSourceName.value = connectInfo.dsn
    txtHost.value = connectInfo.host
    txtPort.value = connectInfo.port
    txtDB.value = connectInfo.db
    txtUser.value = connectInfo.user
    txtPassword.value = connectInfo.password
    txtOption.value = connectInfo.option_
    
    On Error GoTo 0
End Sub

Private Sub favorite_ok(ByVal connectInfo As ValDBConnectInfo)

    setDbConnectInfo connectInfo
End Sub

Private Sub history_ok(ByVal connectInfo As ValDBConnectInfo)

    setDbConnectInfo connectInfo
End Sub

Private Sub cmdHistoryChoice_Click()

    ' --------------------------------------
    ' 履歴情報ウィンドウを表示する
    Load frmDBConnectHistory
    Set history = frmDBConnectHistory
    
    history.ShowExt vbModal
    
    Set history = Nothing
    ' --------------------------------------

End Sub

Private Sub cmdFavoriteLoad_Click()

    ' お気に入りフォームではfrmDBConnectフォームを編集用に開く必要がある。
    ' その際に、すでに開かれたfrmDBConnectフォームが存在しているとVBAの仕様上エラーになるため、一旦自フォームを閉じるようにする
    
    ' 自身のフォームを閉じる
    HideExt

    ' --------------------------------------
    ' お気に入り情報ウィンドウを表示する
    Load frmDBConnectFavorite
    Set favorite = frmDBConnectFavorite
    
    favorite.ShowExt vbModal
    
    Set favorite = Nothing
    ' --------------------------------------
    
    ' 自身のフォームを再度開く
    ShowExt vbModal

End Sub

Private Sub cmdFavoriteSave_Click()

    ' DB接続情報を生成しコントロールから情報を集め設定する
    Dim connectInfo As ValDBConnectInfo
    Set connectInfo = getDbConnectInfo
    
    ' DbConnectInfo.Nameプロパティのデフォルト値
    Dim defaultName As String
    If cboDBType.value = "汎用ODBC" Then
    
        defaultName = cboDataSourceName.value
    ElseIf cboDBType.value = "Oracle Provider for OLE DB" Then
    
        defaultName = txtHost.value & " " & txtDB.value
        
    ElseIf cboDBType.value = "Microsoft OLE DB for SQL Server" Then
    
        defaultName = txtHost.value & " " & txtDB.value
    End If
    
    ' DbConnectInfo.Nameプロパティの入力を行うプロンプトを表示する
    Dim inputedName As String
    inputedName = InputBox("現在の入力内容でDB接続情報を保存します。名前を入力してください。", "DB接続の設定保存", defaultName)
    
    If StrPtr(inputedName) = 0 Then
        ' キャンセルボタンが押下された場合
        Exit Sub
    End If
    
    connectInfo.name = inputedName
    
    ' DB接続情報をレジストリに登録する
    frmDBConnectFavorite.registDbConnectInfo connectInfo
    
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
' ▽DB種類コンボボックス変更時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cboDBType_Change()

    On Error GoTo err
    
    ' DB種類のインデックス
    Dim index As Long
    ' ポート番号
    Dim port As Long
    
    ' コンボボックスで選択されているインデックスを取得する
    index = cboDBType.ListIndex + 1
    
    ' インデックスが範囲外の場合
    If index < CONNECT_STR_MIN Or index > CONNECT_STR_MAX Then
    
        ' 終了
        Exit Sub
    End If
    
    ' 各コントロールの設定値をリセットする
    txtHost.text = ""
    txtDB.text = ""
    txtPort.text = ""
    txtUser.text = ""
    txtPassword.text = ""
    txtOption.text = ""

    ' 各コントロールの有効・無効を設定する
    changeControlByEnableStatus cboDataSourceName, controlEnable(index, CONTROL_ENABLE_IDX_DATASOURCE)
    changeControlByEnableStatus txtHost, controlEnable(index, CONTROL_ENABLE_IDX_HOST)
    changeControlByEnableStatus txtDB, controlEnable(index, CONTROL_ENABLE_IDX_DB)
    changeControlByEnableStatus txtPort, controlEnable(index, CONTROL_ENABLE_IDX_PORT)
    changeControlByEnableStatus txtUser, controlEnable(index, CONTROL_ENABLE_IDX_USER)
    changeControlByEnableStatus txtPassword, controlEnable(index, CONTROL_ENABLE_IDX_PASSWORD)
    changeControlByVisibleStatus cmdFileSelection, controlEnable(index, CONTROL_ENABLE_IDX_FILE_SELECT)

    ' デフォルトポート番号を取得する
    txtPort.text = defaultPort(index)

    ' ○データソースリストを更新する
    updateDataSourceList
    
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

Private Sub changeControlByEnableStatus(ByRef c As control, ByVal enable As Boolean)

    If enable = True Then
    
        c.enabled = True
        c.BackColor = &H80000005
    Else
        c.enabled = False
        c.BackColor = &H8000000F
    
    End If

End Sub

Private Sub changeControlByVisibleStatus(ByRef c As control, ByVal visible As Boolean)

    c.visible = visible
End Sub

' =========================================================
' ▽ODBC設定ラベルクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub lblODBCSetting_Click()
    
    On Error GoTo err
    
    ' 戻り値格納用変数
    Dim ret        As Long
    
    ' システムルート環境変数
    Dim systemRoot As String
    
    ' システムルート環境変数を取得
    systemRoot = WinAPI_Shell.getEnvironmentVariable("SystemRoot")
    
    ' ODBC管理コンソールを起動する
    ret = WinAPI_Shell.ShellExecute(0 _
                           , "open" _
                           , systemRoot & "\system32\odbcad32.exe" _
                           , "" _
                           , systemRoot & "\system32" _
                           , 1)
                           
    ' 戻り値が32以下の場合エラー
    If ret <= 32 Then
    
        VBUtil.showMessageBoxForWarning "ODBC管理コンソールを開くことができませんでした。", ConstantsCommon.APPLICATION_NAME, Nothing
    
    End If
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽DSN更新ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdDSNUpdate_Click()

    On Error GoTo err
    
    ' ○データソースリストを更新する
    updateDataSourceList

    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽接続テストクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdConnectTest_Click()

    On Error GoTo err
    
    ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
    Dim cursorWait As New ExcelCursorWait: cursorWait.init
    
    ' 接続テスト処理を実施する
    connectDBTest
    
    ' 長時間の処理が終了したのでマウスカーソルを元に戻す
    cursorWait.destroy
    
    ' 成功した場合
    VBUtil.showMessageBoxForInformation "DBの接続に成功しました。", ConstantsCommon.APPLICATION_NAME
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽ファイル選択ボタンクリック時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdFileSelection_Click()

    ' ファイルを選択する
    Dim filePath As String
    filePath = VBUtil.openFileDialog("Accessファイルを選択してください", "")

    ' ファイルが選択されたかどうかの判定
    If filePath <> "" Then
    
        ' DBテキストにファイルパスを設定する
        txtDB.text = filePath
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
    
    Dim connStr As String
    
    ' 接続テスト実施結果が失敗だった場合に
    ' 再度設定を行うかをユーザに選択させる
    
    On Error Resume Next
    
    ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
    Dim cursorWait As New ExcelCursorWait: cursorWait.init
    
    ' DBに接続する
    connStr = connectDBTest
    
    ' 長時間の処理が実行されるのでマウスカーソルを砂時計にする
    cursorWait.destroy

    If err.Number <> 0 Then
        
        showMessageBoxForError "エラーが発生しました。", ConstantsCommon.APPLICATION_NAME, err

        ' 設定を終了するかどうかを選択させる
        If VBUtil.showMessageBoxForYesNo("再度設定しますか？" _
                , ConstantsCommon.APPLICATION_NAME) = WinAPI_User.IDYES Then
        
            ' 処理を中断する
            Exit Sub
            
        Else
            ' キャンセルボタン押下時と同じ処理を行い処理を中断する
            cmdCancel_Click
        
            Exit Sub
        End If
        
    End If
    
    On Error GoTo err
    
    ' 通常時の処理（リスナー未設定時=通常の接続、リスナー設定時＝DB接続お気に入りフォームなどからの呼び出し）
    If dbConnectListener Is Nothing Then
        ' DB接続情報を記録する
        storeDbConnectInfo
    End If
    
    ' フォームを閉じる
    HideExt
    
        
    ' 接続文字列
    Dim connSimpleStr As String
    
    ' 接続文字列を生成する
    connSimpleStr = createConnectSimpleString(cboDBType.text _
                                , cboDataSourceName.text _
                                , txtHost.text _
                                , txtPort.text _
                                , txtDB.text _
                                , txtUser.text _
                                , txtPassword.text _
                                , txtOption.text)
                                
    Dim connectInfo As ValDBConnectInfo
    Set connectInfo = New ValDBConnectInfo
    connectInfo.name = ""
    connectInfo.type_ = cboDBType.text
    connectInfo.name = cboDataSourceName.text
    connectInfo.host = txtHost.text
    connectInfo.port = txtPort.text
    connectInfo.db = txtDB.text
    connectInfo.user = txtUser.text
    connectInfo.password = txtPassword.text
    connectInfo.option_ = txtOption.text
    
    ' DB接続OKイベントを送信する
    RaiseEvent ok(connStr, connSimpleStr, connectInfo)
    ' リスナーにもイベントを通知する
    If Not dbConnectListener Is Nothing Then
        dbConnectListener.connect connectInfo
    End If
    
    ' 通常時の処理（リスナー未設定時=通常の接続、リスナー設定時＝DB接続お気に入りフォームなどからの呼び出し）
    If dbConnectListener Is Nothing Then
        ' --------------------------------------
        Load frmDBConnectHistory
        Set history = frmDBConnectHistory
        
        history.registDbConnectInfo getDbConnectInfo
        
        Set history = Nothing
        ' --------------------------------------
    End If
    
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
    
    ' DB接続キャンセルイベントを送信する
    RaiseEvent cancel
    ' リスナーにもイベントを通知する
    If Not dbConnectListener Is Nothing Then
        dbConnectListener.cancel
    End If

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

    Dim i As Long
    Dim j As Long
    
    ' 以下の配列変数は、同一インデックスによって対応している。
    ' ・接続文字列
    ' ・プロバイダラベル
    ' ・デフォルトポート番号
    ' ・コントロール有効フラグ
    
    ' ----------------------------------------------
    ' 接続文字列　初期化
    ' ----------------------------------------------
    i = CONNECT_STR_MIN
    
    ' ODBC
    ' ※MSDASQL.1は、マイクロソフト製のODBC用OLE DBプロバイダ
    connectStr(i) = "Provider=MSDASQL.1;" & _
                    "Data Source=${dsn};" & _
                    "User ID=${user};" & _
                    "Password=${password};"
                    
    i = i + 1
    
'    ' PostgreSQL（OLEDB）
'    connectStr(i) = "Provider=PostgreSQL OLE DB Provider;" & _
'                                                 "Data Source=${host};" & _
'                                                 "Location=${db};" & _
'                                                 "User ID=${user};" & _
'                                                 "Password=${password};"
'
'    i = i + 1
'
'    ' MySQL（ODBC）
'    connectStr(i) = "Driver={MySQL ODBC 3.51 Driver};" & _
'                                                 "Server=${host};" & _
'                                                 "Port=${port};" & _
'                                                 "Database=${db};" & _
'                                                 "User=${user};" & _
'                                                 "Password=${password};" & _
'                                                 "Option=3;"
'
'    i = i + 1
    
    ' Oracle（OLEDB Oracle）
    connectStr(i) = "Provider=OraOLEDB.Oracle;" & _
                                                 "Data Source=${db};" & _
                                                 "User Id=${user};" & _
                                                 "Password=${password};"
                                                 
    i = i + 1
    
'    ' Oracle（OLEDB Microsoft）
'    connectStr(i) = "Provider=msdaora;" & _
'                                                 "Data Source=${db};" & _
'                                                 "User Id=${user};" & _
'                                                 "Password=${password};"
                                                 
    ' Microsoft SQL Server（OLEDB）
    connectStr(i) = "Provider=SQLOLEDB;" & _
                                                 "Data Source=${host};" & _
                                                 "Initial Catalog=${db};" & _
                                                 "User Id=${user};" & _
                                                 "Password=${password};"
                                                 
    i = i + 1
    
    ' Microsoft Access
    connectStr(i) = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                                 "Data Source=${db};" & _
                                                 "User Id=${user};" & _
                                                 "Password=${password};"
                                                 
    i = i + 1
    
    ' Microsoft Access for 2007
    connectStr(i) = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                                 "Data Source=${db};" & _
                                                 "User Id=${user};" & _
                                                 "Password=${password};"
                                                 
    i = i + 1
                                                 
    ' ----------------------------------------------
    ' プロバイダラベル　初期化
    ' ----------------------------------------------
    i = CONNECT_STR_MIN

    providerLabel(i) = "汎用ODBC": i = i + 1
'    providerLabel(i) = "PostgreSQL (OLE DB)": i = i + 1
'    providerLabel(i) = "MySQL (MyODBC 3.51)": i = i + 1
    providerLabel(i) = "Oracle Provider for OLE DB": i = i + 1
'    providerLabel(i) = "Oracle Provider for OLE DB (Microsoft)": i = i + 1
    providerLabel(i) = "Microsoft OLE DB for SQL Server": i = i + 1
    providerLabel(i) = "Microsoft Access (Jet Provider)": i = i + 1
    providerLabel(i) = "Microsoft Access (Ace Provider)": i = i + 1

    ' ----------------------------------------------
    ' デフォルトポート番号　初期化
    ' ----------------------------------------------
    i = CONNECT_STR_MIN

    defaultPort(i) = "": i = i + 1
'    defaultPort(i) = "5432": i = i + 1
'    defaultPort(i) = "3306": i = i + 1
    defaultPort(i) = "": i = i + 1
'    defaultPort(i) = "": i = i + 1
    defaultPort(i) = "": i = i + 1
    defaultPort(i) = "": i = i + 1
    defaultPort(i) = "": i = i + 1
    
    ' ----------------------------------------------
    ' コントロール有効フラグ　初期化
    ' ※プロバイダが変更された場合に対応するコントロールの有効・無効を決定する値
    ' ----------------------------------------------
    i = CONNECT_STR_MIN
    j = CONTROL_ENABLE_IDX_DATASOURCE
    
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = False: j = j + 1
    controlEnable(i, j) = False: j = j + 1
    controlEnable(i, j) = False: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = False: j = j + 1

    i = i + 1
    j = CONTROL_ENABLE_IDX_DATASOURCE
    
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'
'    i = i + 1
'    j = CONTROL_ENABLE_IDX_DATASOURCE
'
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'
'    i = i + 1
'    j = CONTROL_ENABLE_IDX_DATASOURCE
    
    controlEnable(i, j) = False: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = True: j = j + 1
    controlEnable(i, j) = False: j = j + 1

    i = i + 1
    j = CONTROL_ENABLE_IDX_DATASOURCE
    
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = False: j = j + 1
'    controlEnable(i, j) = True: j = j + 1
'    controlEnable(i, j) = True: j = j + 1

    controlEnable(i, CONTROL_ENABLE_IDX_DATASOURCE) = False
    controlEnable(i, CONTROL_ENABLE_IDX_HOST) = True
    controlEnable(i, CONTROL_ENABLE_IDX_DB) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PORT) = False
    controlEnable(i, CONTROL_ENABLE_IDX_USER) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PASSWORD) = True
    controlEnable(i, CONTROL_ENABLE_IDX_FILE_SELECT) = False

    i = i + 1
    j = CONTROL_ENABLE_IDX_DATASOURCE

    controlEnable(i, CONTROL_ENABLE_IDX_DATASOURCE) = False
    controlEnable(i, CONTROL_ENABLE_IDX_HOST) = False
    controlEnable(i, CONTROL_ENABLE_IDX_DB) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PORT) = False
    controlEnable(i, CONTROL_ENABLE_IDX_USER) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PASSWORD) = True
    controlEnable(i, CONTROL_ENABLE_IDX_FILE_SELECT) = True

    i = i + 1
    j = CONTROL_ENABLE_IDX_DATASOURCE

    controlEnable(i, CONTROL_ENABLE_IDX_DATASOURCE) = False
    controlEnable(i, CONTROL_ENABLE_IDX_HOST) = False
    controlEnable(i, CONTROL_ENABLE_IDX_DB) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PORT) = False
    controlEnable(i, CONTROL_ENABLE_IDX_USER) = True
    controlEnable(i, CONTROL_ENABLE_IDX_PASSWORD) = True
    controlEnable(i, CONTROL_ENABLE_IDX_FILE_SELECT) = True

    ' ○DB種類コンボボックスにリストを追加する
    cboDBType.list = providerLabel

    ' 通常時の処理（リスナー未設定時=通常の接続、リスナー設定時＝DB接続お気に入りフォームなどからの呼び出し）
    If dbConnectListener Is Nothing Then
        ' ○前回最後に接続した情報をフォーム上の各コントロールに復元させる
        restoreDbConnectInfo
    End If
    
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

    Set history = Nothing
    Set favorite = Nothing

End Sub

' =========================================================
' ▽データソースリストの更新処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub updateDataSourceList()

    Dim dataSourceList As ValCollection
    Dim dataSource     As ValCollection
    
    Set dataSourceList = WinAPI_ODBC.getDataSourceList
    
    cboDataSourceName.clear
    
    For Each dataSource In dataSourceList.col
    
        cboDataSourceName.addItem dataSource.getItemByIndex(1, vbVariant)
        
    Next
End Sub

' =========================================================
' ▽接続テスト処理
'
' 概要　　　：DBへの接続を行う
' 引数　　　：
' 戻り値　　：DB接続文字列
'
' =========================================================
Private Function connectDBTest() As String

    On Error GoTo err
    
    ' コネクションオブジェクト
    Dim conn As Object
    
    ' 接続文字列
    Dim connStr As String
    
    ' 接続文字列を生成する
    connStr = createConnectString(cboDBType.text _
                                , cboDataSourceName.text _
                                , txtHost.text _
                                , txtPort.text _
                                , txtDB.text _
                                , txtUser.text _
                                , txtPassword.text _
                                , txtOption.text)
                                      
    
    ' DBに接続する
    Set conn = ADOUtil.connectDb(connStr)
    
    ' DBに接続している場合、DBの接続を切断する
    If Not conn Is Nothing Then
    
        ADOUtil.closeDB conn
        Set conn = Nothing
        
    End If
    
    connectDBTest = connStr
    
    Exit Function

err:

    ' DBに接続している場合、DBの接続を切断する
    If Not conn Is Nothing Then
    
        ADOUtil.closeDB (conn)
        Set conn = Nothing
        
    End If
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Function

' =========================================================
' ▽DB接続文字列生成処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Function createConnectString(ByVal dbType As String _
                                   , ByVal dsn As String _
                                   , ByVal host As String _
                                   , ByVal port As String _
                                   , ByVal db As String _
                                   , ByVal user As String _
                                   , ByVal password As String _
                                   , ByVal options As String _
                                   ) As String

    ' 接続文字列
    Dim connStr As String
    
    ' DB種類のインデックス
    Dim index As Long
    
    ' コンボボックスで選択されているインデックスを取得する
    index = cboDBType.ListIndex + 1
    
    ' インデックスが範囲外の場合
    If index < CONNECT_STR_MIN Or index > CONNECT_STR_MAX Then
    
        ' 終了
        Exit Function
    End If
    
    connStr = connectStr(index)

    ' Oracleの場合
    If dbType = "Oracle Provider for OLE DB" Then
    
        Dim dbVar As String
        dbVar = db
        If Trim$(host) <> "" And Trim$(port) <> "" Then
            dbVar = host & ":" & port & "/" & dbVar
        ElseIf Trim$(host) <> "" And Trim$(port) = "" Then
            dbVar = host & "/" & dbVar
        End If
        
        connStr = replace(connStr, "${db}", dbVar)
        connStr = replace(connStr, "${user}", user)
        connStr = replace(connStr, "${password}", password)
        connStr = connStr & options
            
    Else
    
        connStr = replace(connStr, "${dsn}", dsn)
        connStr = replace(connStr, "${host}", host)
        connStr = replace(connStr, "${port}", port)
        connStr = replace(connStr, "${db}", db)
        connStr = replace(connStr, "${user}", user)
        connStr = replace(connStr, "${password}", password)
        connStr = connStr & options
        
    End If
        
    createConnectString = connStr
    
End Function

' =========================================================
' ▽DB接続文字列（単純）生成処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Function createConnectSimpleString(ByVal dbType As String _
                                   , ByVal dsn As String _
                                   , ByVal host As String _
                                   , ByVal port As String _
                                   , ByVal db As String _
                                   , ByVal user As String _
                                   , ByVal password As String _
                                   , ByVal options As String _
                                   ) As String

    ' 接続文字列
    Dim connStr As String
    Dim joinStr As String
    
    If dsn <> "" Then
        connStr = connStr & joinStr & "DSN=" & dsn: joinStr = ", "
    End If
    
    If host <> "" Then
        connStr = connStr & joinStr & "ホスト=" & host: joinStr = ", "
    End If
    
    If port <> "" Then
        connStr = connStr & joinStr & "ポート=" & port: joinStr = ", "
    End If
    
    If db <> "" Then
        connStr = connStr & joinStr & "DB=" & db: joinStr = ", "
    End If
        
    createConnectSimpleString = connStr
    
End Function

' =========================================================
' ▽DBの接続情報を保存する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub storeDbConnectInfo()

    On Error GoTo err
    
    ' DB接続情報を格納する配列
    Dim dbConnectInfoArray(0 To 7 _
                         , 0 To 1) As Variant
    
    
    Dim i As Long
    
    i = 0
    
    dbConnectInfoArray(i, 0) = cboDBType.name
    dbConnectInfoArray(i, 1) = cboDBType.value: i = i + 1
    dbConnectInfoArray(i, 0) = cboDataSourceName.name
    dbConnectInfoArray(i, 1) = cboDataSourceName.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtHost.name
    dbConnectInfoArray(i, 1) = txtHost.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtPort.name
    dbConnectInfoArray(i, 1) = txtPort.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtDB.name
    dbConnectInfoArray(i, 1) = txtDB.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtUser.name
    dbConnectInfoArray(i, 1) = txtUser.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtPassword.name
    dbConnectInfoArray(i, 1) = txtPassword.value: i = i + 1
    dbConnectInfoArray(i, 0) = txtOption.name
    dbConnectInfoArray(i, 1) = txtOption.value: i = i + 1
    
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' レジストリに情報を設定する
    registry.setValues dbConnectInfoArray
    
    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet
    
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, cboDBType.name, cboDBType.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, cboDataSourceName.name, cboDataSourceName.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtHost.name, txtHost.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtPort.name, txtPort.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtDB.name, txtDB.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtUser.name, txtUser.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtPassword.name, txtPassword.value
    bookProp.setValue ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG, txtOption.name, txtOption.value
    ' ----------------------------------------------
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽DBの接続情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub restoreDbConnectInfo()

    On Error GoTo err
    
    ' ----------------------------------------------
    ' ブック設定情報
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' 設定情報シートが存在する
        
        Set bookPropVal = bookProp.getValues(ConstantsBookProperties.TABLE_DB_CONNECT_DIALOG)
        If bookPropVal.count > 0 Then
            ' 設定情報が存在するので、フォームに反映する
            cboDBType.value = bookPropVal.getItem(cboDBType.name, vbString)
            cboDataSourceName.value = bookPropVal.getItem(cboDataSourceName.name, vbString)
            txtHost.value = bookPropVal.getItem(txtHost.name, vbString)
            txtPort.value = bookPropVal.getItem(txtPort.name, vbString)
            txtDB.value = bookPropVal.getItem(txtDB.name, vbString)
            txtUser.value = bookPropVal.getItem(txtUser.name, vbString)
            txtPassword.value = bookPropVal.getItem(txtPassword.name, vbString)
            txtOption.value = bookPropVal.getItem(txtOption.name, vbString)
        
            Exit Sub
        End If
    End If
    ' ----------------------------------------------
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_DB_CONNECT) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
    
    Dim retDbType         As String
    Dim retDataSourceName As String
    Dim retHost           As String
    Dim retPort           As String
    Dim retDB             As String
    Dim retUser           As String
    Dim retPassword       As String
    Dim retOption         As String
    
    ' レジストリから情報を取得する
    registry.getValue cboDBType.name, retDbType
    registry.getValue cboDataSourceName.name, retDataSourceName
    registry.getValue txtHost.name, retHost
    registry.getValue txtPort.name, retPort
    registry.getValue txtDB.name, retDB
    registry.getValue txtUser.name, retUser
    registry.getValue txtPassword.name, retPassword
    registry.getValue txtOption.name, retOption

    cboDBType.value = retDbType
    cboDataSourceName.value = retDataSourceName
    txtHost.value = retHost
    txtPort.value = retPort
    txtDB.value = retDB
    txtUser.value = retUser
    txtPassword.value = retPassword
    txtOption.value = retOption
    
    Exit Sub

err:

    Main.ShowErrorMessage


End Sub
