VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelCursorWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' Excelマクロでカーソルのみ待機状態にするためのクラス
'
' 作成者　：Ison
' 履歴　　：2009/04/12　新規作成

' 使用方法：
' 　本クラスのオブジェクトを New する。
' 　関数を抜けると自動で Class_Terminate で後始末が行われる。
'
'   Dim var As New ExcelCursorWait: var.init
'
' *********************************************************

' 破棄後も継続するフラグ
Private isContinueWhenDestroyed As Boolean
' 長時間処理クラス
Private excelLongTimeProcess As ExcelLongTimeProcessing

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    On Error Resume Next
    
    Set excelLongTimeProcess = New ExcelLongTimeProcessing
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set excelLongTimeProcess = Nothing
    
End Sub

' =========================================================
' ▽初期化処理
'
' 概要　　　：isContinueWhenDestroyed_ 破棄後も継続するフラグ
'
' =========================================================
Public Sub init(Optional ByVal isContinueWhenDestroyed_ As Boolean = False)
              
    ' 破棄後も継続するフラグ
    isContinueWhenDestroyed = isContinueWhenDestroyed_
    
    If isContinueWhenDestroyed = True Then
        ' 破棄後も継続するので強制的にカーソルを待機状態にする
        Application.cursor = xlWait
        
        Exit Sub
    End If
              
    If Application.cursor = xlWait Then
        ' 既にカーソルが待機状態なので、何も有効にしない
    Else
        ' カーソルのみ有効にする
        excelLongTimeProcess.init False, False, False, True, False, False, False
    End If
              

End Sub

' =========================================================
' ▽破棄処理
'
' 概要　　　：
'
' =========================================================
Public Sub destroy()

    excelLongTimeProcess.destroy

End Sub

' =========================================================
' ▽強制的にカーソルを元の状態に戻す処理
'
' 概要　　　：
'
' =========================================================
Public Sub forceRestore()

    Application.cursor = xlDefault

End Sub

