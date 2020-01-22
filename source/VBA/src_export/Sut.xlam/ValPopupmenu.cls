VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValPopupmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ポップアップメニューを表現するクラス
'
' 作成者　：Ison
' 履歴　　：2009/06/07　新規作成
'
' 特記事項：
' *********************************************************

' アプリケーションコマンドバーID
Private m_appBarId As String

' ポップアップメニューの名前
Private m_popupMenuName    As String
' ショートカットキーラベル
Private m_shortcutKeyLabel As String
' ショートカットキーコード Application#OnKeyに設定するコード
Private m_shortcutKeyCode  As String

' ポップアップメニューの項目リスト
Private m_itemList As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set m_itemList = New ValCollection
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set m_itemList = Nothing
End Sub

Public Sub init(ByVal appBarId As String)

    m_appBarId = appBarId
End Sub

' =========================================================
' ▽プロパティ
' =========================================================
Public Property Get popupMenuName() As String

    popupMenuName = m_popupMenuName
End Property

Public Property Let popupMenuName(ByVal vNewValue As String)

    m_popupMenuName = vNewValue
End Property

Public Property Get shortcutKeyCode() As String

    shortcutKeyCode = m_shortcutKeyCode
End Property

Public Property Let shortcutKeyCode(ByVal vNewValue As String)

    m_shortcutKeyCode = vNewValue
End Property

Public Property Get shortcutKeyLabel() As String

    shortcutKeyLabel = m_shortcutKeyLabel
End Property

Public Property Let shortcutKeyLabel(ByVal vNewValue As String)

    m_shortcutKeyLabel = vNewValue
End Property

Public Property Get commandBarPopup() As CommandBar

    ' Application.CommandBars.itemメソッドは項目が見つからない場合、エラーを投げるため
    ' これを防ぐためエラーを無視する以下の宣言を行う
    On Error Resume Next

    Set commandBarPopup = Application.CommandBars.item(m_appBarId & "_" & m_popupMenuName)

    If err.Number <> 0 Then
    
        Set commandBarPopup = Nothing
    End If
    
    On Error GoTo 0
    
End Property

Public Property Get itemList() As ValCollection

    Set itemList = m_itemList
End Property

Public Property Let itemList(ByVal vNewValue As ValCollection)

    Set m_itemList = vNewValue
End Property


' =========================================================
' ▽ポップアップの生成
'
' 概要　　　：項目リストからポップアップメニューを生成する。
' 引数　　　：
' 戻り値　　：True 生成に成功
'
' =========================================================
Public Function createPopup() As Boolean

    On Error GoTo err

    ' 戻り値を初期化
    createPopup = False
    
    ' アプリケーションのコマンドバー
    Dim appBar As CommandBar
    
    Set appBar = Application.CommandBars(m_appBarId)

    ' ポップアップ
    Dim popup As CommandBar
    
    ' -----------------------------------------------------------------
    ' Application.CommandBars.itemメソッドは項目が見つからない場合、エラーを投げるため
    ' これを防ぐためエラーを無視する以下の宣言を行う
    On Error Resume Next
    
    ' ポップアップコントロールを取得する
    Set popup = Application.CommandBars.item(m_appBarId & "_" & m_popupMenuName)
    
    On Error GoTo err
    ' -----------------------------------------------------------------
    
    ' ポップアップコントロールが既に存在する場合は、一旦削除する
    If Not popup Is Nothing Then
    
        popup.delete
        
    End If
    
    ' 新規で追加する
    Set popup = Application.CommandBars.Add( _
                            name:=m_appBarId & "_" & m_popupMenuName _
                          , Temporary:=True _
                          , position:=msoBarPopup)
    
    ' 元になるコントロール
    Dim srcControl As commandBarControl
    ' ポップアップに追加するコントロール（元になるコントロールの内容をコピーする）
    Dim desControl As commandBarControl
    
    ' ポップアップ項目
    Dim popupMenuId As Variant
    
    ' 項目リストを1つずる処理する
    For Each popupMenuId In m_itemList.col
    
        ' 機能IDを条件にアプリケーションのコマンドバーからコントロールを検索する
        Set srcControl = appBar.FindControl(Tag:=popupMenuId, recursive:=True)
        
        ' 存在している場合
        If Not srcControl Is Nothing Then
        
            ' ポップアップにコントロールを追加
            Set desControl = popup.Controls.Add(Type:=msoControlButton)
            ' ポップアップに追加したコントロール（真っ白な状態）をアプリケーションバーから取得したコントロールと同じ内容に変更する
            ExcelUtil.copyCommandBarControl srcControl, desControl
        
        End If
    
    Next
    
    
    ' 作成に成功
    createPopup = True
    
    Exit Function
    
err:

    If Not popup Is Nothing Then
    
        popup.delete
    End If

    ' 作成に成功
    createPopup = False
    
End Function

' =========================================================
' ▽ポップアップメニューのクローンを生成する。
'
' 概要　　　：オブジェクトのクローンを生成する。
' 引数　　　：
' 戻り値　　：ValPopupMenu
'
' =========================================================
Public Function Clone() As ValPopupmenu

    Dim ret As ValPopupmenu
    
    Set ret = New ValPopupmenu:
    
    ret.init m_appBarId
    ret.popupMenuName = m_popupMenuName
    ret.shortcutKeyCode = m_shortcutKeyCode
    ret.shortcutKeyLabel = m_shortcutKeyLabel
    
    Dim newItemList As New ValCollection
    
    Dim var As Variant
    For Each var In m_itemList.col
    
        newItemList.setItem var, var
    Next
    
    ret.itemList = newItemList
    
    Set Clone = ret

End Function
