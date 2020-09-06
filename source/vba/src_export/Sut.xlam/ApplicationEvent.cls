VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' アプリケーションパスを一時的に変更するクラス
'
' 作成者　：Ison
' 履歴　　：2013/07/16　新規作成
'
' 特記事項：
' *********************************************************

' アプリケーションオブジェクト
Public WithEvents myEvent As Application
Attribute myEvent.VB_VarHelpID = -1

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ▽初期化メソッド
' =========================================================
Public Sub init(ByRef app As Application)

    Set myEvent = app
    
End Sub

' =========================================================
' ▽ブックの新規作成ｖ
'
' 概要　　　：
'
' 引数　　　：Wb ワークブック
' 戻り値　　：無し
' =========================================================
Private Sub myEvent_NewWorkbook(ByVal wb As Workbook)

    Debug.Print wb.name
End Sub

' =========================================================
' ▽ブックの開くイベント
'
' 概要　　　：
'
' 引数　　　：Wb ワークブック
' 戻り値　　：無し
' =========================================================
Private Sub myEvent_WorkbookOpen(ByVal wb As Workbook)

    Debug.Print wb.name
End Sub

' =========================================================
' ▽ブックのクローズ時イベント
'
' 概要　　　：
'
' 引数　　　：Wb ワークブック
' 戻り値　　：無し
' =========================================================
Private Sub myEvent_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)

    Debug.Print wb.name
    
    If Not SutWorkbook Is wb Then
    
        Main.SutRelease
    End If
    
End Sub

