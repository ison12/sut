VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenuSetting 
   Caption         =   "メニュー設定"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7440
   OleObjectBlob   =   "frmMenuSetting.frx":0000
End
Attribute VB_Name = "frmMenuSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' 右クリックメニューの設定
'
' 作成者　：Ison
' 履歴　　：2009/06/02　新規作成
'
' 特記事項：
' 　フォームの上にフォームを重ねて表示すると
'   以後､Excel本体のIMEモードが無効になり操作不能になってしまうという現象に遭遇
'   これを防ぐために､一旦親フォームを隠して､本フォームを閉じるときに再表示することで､この現象を防ぐ
' 　そのために、ShowExtメソッドに親フォームを渡すよう引数に追加している
'
' *********************************************************

' ▽イベント
' =========================================================
' ▽決定した際に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：appSettingShortcut アプリケーション設定ショートカット
' 　　　　　　selectedItemList 選択済み項目リスト
' 　　　　　　menuName 新しいメニュー名
'
' =========================================================
Public Event ok(ByRef appSettingShortcut As ValApplicationSettingShortcut _
              , ByRef selectedItemList As ValCollection _
              , ByVal menuName As String)

' =========================================================
' ▽キャンセルされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' =========================================================
' ▽リセットされた場合に呼び出されるイベント
'
' 概要　　　：
' 引数　　　：appSettingShortcut アプリケーション設定ショートカット
' 　　　　　　cancel キャンセルフラグ
'
' =========================================================
Public Event reset(ByRef appSettingShortcut As ValApplicationSettingShortcut _
                 , ByRef Cancel As Boolean)

' アイコン画像
Private iconImage As IPictureDisp

' アプリケーション設定情報（ショートカットキー）
Private applicationSetting As ValApplicationSettingShortcut

' 選択済み項目リスト
' 右側のリストボックスに設定する項目を格納しているリスト
' 以下のキー値を元に、機能リストから選択されている項目を抽出する
' [ Key ] : CommandBarControl.Tag
' [ Val ] : CommandBarControl.Tag
Private selectedItemList As ValCollection

' メニューリスト コントロール
Private menuList As CntListBox
' 機能リスト コントロール
Private appMenuList As CntListBox

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
' 引数　　　：icon             アイコン
' 　　　　　　modal            モーダルまたはモードレス表示指定
' 　　　　　　var              アプリケーション設定情報
'             var2             選択済み項目リスト
' 　　　　　　title            フォームタイトル
' 　　　　　　message          フォームのメッセージ
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByRef icon As Object _
                 , ByVal modal As FormShowConstants _
                 , ByRef var As ValApplicationSettingShortcut _
                 , ByRef var2 As ValCollection _
                 , ByVal title As String _
                 , ByVal message As String _
                 , ByVal menuName As String _
                 , Optional ByVal menuNameDisable As Boolean = False)

    If Not icon Is Nothing Then
        ' 自身のアイコンを退避させる
        'Set iconImage = Me.imgIcon.Picture
        ' アイコンを親フォームの画像で置き換える
        'Me.imgIcon.Picture = icon
    End If
    
    ' メンバ変数にアプリケーション設定情報を設定する
    Set applicationSetting = var
    ' 選択済み項目リストを設定する
    Set selectedItemList = var2
    ' タイトルを設定する
    Me.Caption = title
    ' メッセージを設定する
    lblMessage.Caption = message
    ' メニュー名を設定する
    txtMenuName.value = menuName
    If menuNameDisable = True Then
    
        txtMenuName.Enabled = False
        lstAppList.SetFocus
    Else
    
        txtMenuName.Enabled = True
        txtMenuName.SetFocus
    End If
    
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
    
    If Not iconImage Is Nothing Then
        ' アイコン画像を設定する
        Me.imgIcon.Picture = iconImage
    End If

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

    initListControl
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
    
    ' 登録する情報
    Dim storedList As New ValCollection

    ' コントロールオブジェクト
    Dim control As commandBarControl
    
    ' リストに存在する項目を右クリックメニューとして追加する
    For Each control In menuList.collection.col
    
        storedList.setItem control.Tag, control.Tag
    Next
    
    ' フォームを閉じる
    HideExt
    
    ' OKイベントを送信する
    RaiseEvent ok(applicationSetting _
                , storedList _
                , txtMenuName.value)
    
    
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
' ▽追加ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdMenuAdd_Click()

    On Error GoTo err
    
    ' インデックス
    Dim i As Long
    ' 個数
    Dim cnt As Long
    
    ' リストボックスの要素
    Dim appMenuItem As commandBarControl
    
    ' 削除する要素リスト
    Dim removeItem As New ValCollection
    
    ' リストボックスの個数を取得する
    cnt = lstAppList.ListCount
    
    ' 機能リストにてチェックされている要素を
    ' 右クリックメニューリストに移し変える
    For i = 0 To cnt - 1
    
        ' 選択されているかチェック
        If lstAppList.selected(i) = True Then
        
            ' 削除要素リストに削除すべきインデックスを追加
            removeItem.setItem i
            
            ' 機能を取得
            Set appMenuItem = appMenuList.getItem(i)
            ' 右クリックメニューリストに機能を追加する
            menuList.addItem appMenuItem.DescriptionText, appMenuItem
            
        End If
        
    Next
    
    ' 削除処理の実行
    ' 最後尾から最初に向かってリストをループさせるのは
    ' 要素の削除によってインデックスにずれが発生するのを防ぐため
    cnt = removeItem.count
    For i = cnt - 1 To 0 Step -1
    
        appMenuList.removeItem removeItem.getItemByIndex(i + 1, vbLong)
    
    Next
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽削除ボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdMenuRemove_Click()

    On Error GoTo err
    
    ' 削除した要素
    Dim removedItem As commandBarControl
    
    ' 選択されている項目のインデックス
    Dim selectedIndex As Long
    
    ' 現在リストで選択されているインデックスを取得する
    selectedIndex = lstMenu.ListIndex
    
    ' 未選択の場合
    If selectedIndex = -1 Then
    
        Exit Sub
    End If
    
    ' 右クリックメニューリストから項目を取得し削除する
    Set removedItem = menuList.getItem(selectedIndex)
    menuList.removeItem selectedIndex
    
    ' 機能リストに項目を追加する
    appMenuList.addItem removedItem.DescriptionText, removedItem
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽下へボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdMenuDown_Click()

    On Error GoTo err
    
    ' 選択済みインデックス
    Dim selectedIndex As Long
    
    ' 現在リストで選択されているインデックスを取得する
    selectedIndex = lstMenu.ListIndex
    
    If selectedIndex < lstMenu.ListCount - 1 Then
    
        menuList.swapItem selectedIndex _
                        , selectedIndex + 1
                              
        lstMenu.selected(selectedIndex + 1) = True
            
    End If
    
    lstMenu.SetFocus
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
        
End Sub

' =========================================================
' ▽上へボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdMenuUp_Click()

    On Error GoTo err
    
    ' 選択済みインデックス
    Dim selectedIndex As Long
    
    ' 現在リストで選択されているインデックスを取得する
    selectedIndex = lstMenu.ListIndex
    
    If selectedIndex > 0 Then
    
        menuList.swapItem selectedIndex _
                        , selectedIndex - 1
                              
        lstMenu.selected(selectedIndex - 1) = True
            
    End If
    
    lstMenu.SetFocus
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽リセットボタン押下時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdReset_Click()

    On Error GoTo err
    
    Dim isCancel As Boolean: isCancel = False
    
    ' リセットイベントを発行する
    RaiseEvent reset(applicationSetting, isCancel)
    
    ' キャンセルされた場合
    If isCancel = True Then
    
        Exit Sub
    End If
    
    ' 選択済み項目を初期化し、サイズを0にする
    Set selectedItemList = New ValCollection
    
    ' リセットイベントを受信した側で、リセットが行われているため
    ' リストコントロールを初期化する
    initListControl
    
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

    Set iconImage = Nothing
End Sub

' =========================================================
' ▽ショートカット情報を読み込む
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub initListControl()

    ' 右クリックメニューリストと機能リストをリセットする
    lstMenu.clear
    lstAppList.clear
    
    ' 右クリックメニューリストの初期化
    Set menuList = New CntListBox: menuList.init lstMenu
    ' 機能リストの初期化
    Set appMenuList = New CntListBox: appMenuList.init lstAppList
    
    ' Sutメニューの要素
    Dim sutMenuItem As commandBarControl
    
    Dim shortcutInfo As ValShortcutKey
    
    ' 機能のショートカットリストを取得する
    Dim shortCutList As ValCollection
    Set shortCutList = applicationSetting.shortcutAppList
    
    ' ---------------------------------------------------------
    ' 機能リストの初期化
    ' ---------------------------------------------------------
    For Each shortcutInfo In shortCutList.col
    
        ' メニューの要素を取得する
        Set sutMenuItem = shortcutInfo.commandBarControl
        
        ' 保存されていない場合は、機能リストに追加
        If selectedItemList.exist(sutMenuItem.Tag) = False Then
        
            ' 左側のメニューに項目を追加する
            appMenuList.addItem sutMenuItem.DescriptionText, sutMenuItem
        
        Else
        
        End If
    
    Next

    ' ---------------------------------------------------------
    ' メニューリストの初期化
    ' ※メニューリストの順序性を維持するために
    ' 　menuListから要素を順番に取り出してリストに格納する
    ' ---------------------------------------------------------
    Dim menuId As Variant
    
    For Each menuId In selectedItemList.col
    
        ' ショートカット情報を取得する
        Set shortcutInfo = shortCutList.getItem(menuId)
        
        If Not shortcutInfo Is Nothing Then
        
            ' メニューの要素を取得する
            Set sutMenuItem = shortcutInfo.commandBarControl
        
            ' 右側のメニューに項目を追加する
            menuList.addItem sutMenuItem.DescriptionText, sutMenuItem
        
        End If
        
    Next

End Sub
