VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのトランザクションを管理するオブジェクト
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/04/12　新規作成

' 使用方法：
' 　本クラスのオブジェクトを New する。
' 　関数を抜けると自動で Class_Terminate が呼ばれロールバックされる。
' 　コミットする場合は、commitメソッドを明示的にコールする。
'
'   Dim var As New DBTransaction: var.init
'
'
' *********************************************************

' コネクションオブジェクト
Private m_conn As Object
' コミットされているかどうかを表すフラグ
Private commited As Boolean

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' コミットフラグがオフの場合
    If commited = False Then
        
        ' ロールバックを実行する
        m_conn.RollbackTrans
        
        #If (DEBUG_MODE <> 1) Then
            Debug.Print "DBTransaction Rollback"
        #End If

    End If
End Sub

' =========================================================
' ▽初期化処理
'
' 概要　　　：トランザクションを開始する
' 引数　　　：conn コネクションオブジェクト
'
' =========================================================
Public Sub init(ByRef conn As Object)

    ' コミットフラグを初期化する
    commited = False
    
    ' コネクションオブジェクトを設定する
    Set m_conn = conn
    ' トランザクションを開始する
    m_conn.begintrans
End Sub

' =========================================================
' ▽コミット処理
'
' 概要　　　：トランザクションをコミットする
' 引数　　　：
'
' =========================================================
Public Sub commit()

    ' トランザクションをコミットする
    m_conn.CommitTrans
    ' コミットフラグをオンにする
    commited = True
End Sub
