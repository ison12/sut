VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationCurDirChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' アプリケーションパスを一時的に変更するクラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2013/07/16　新規作成
'
' 特記事項：
' *********************************************************

' アプリケーションパス
Private applicationPath As String

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    ' カレントパスを一時退避状態にする
    applicationPath = CurDir$
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' パスを元に戻す
    ChDrive applicationPath
    ChDir applicationPath
    
End Sub

' =========================================================
' ▽初期化メソッド
' =========================================================
Public Sub init(Optional ByVal path As String)
    
    ' 現在のパスを任意のパスに設定する
    ChDrive path
    ChDir path

End Sub

' =========================================================
' ▽初期化メソッド
' =========================================================
Public Sub initByThisWorkbook()
    
    ' 現在のパスをワークブックのパスに設定する
    ChDrive ThisWorkbook.path
    ChDir ThisWorkbook.path

End Sub

