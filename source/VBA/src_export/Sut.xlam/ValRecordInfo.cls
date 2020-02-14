VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValRecordInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' レコードを表現するクラス
'
' 作成者　：Ison
' 履歴　　：2008/02/01　新規作成
'
' 特記事項：
' *********************************************************

Public columnInfoList  As collection  ' ValColumnInfoクラスを要素に持つCollectionオブジェクト

Public tableName       As String      ' テーブル名

Public targetSheetName As String      ' 対象シート名
Public targetRow       As Long        ' 対象行

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    Set columnInfoList = New collection
    
End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    Set columnInfoList = Nothing
    
End Sub

' =========================================================
' ▽カラム情報追加
'
' 概要　　　：プロパティにカラム情報を追加する
' 引数　　　：valColumnInfo カラム情報
'             key           キー
'
' =========================================================
Public Sub addColumnInfo(ByRef valColumnInfo, Optional ByVal key As Variant = "")

    'VBUtil.setItem columnInfoList, valColumnInfo, key
End Sub

' =========================================================
' ▽カラム情報取得
'
' 概要　　　：カラム情報を取得する
' 引数　　　：key キー
'
' =========================================================
Public Function getColumnInfo(ByVal key As Variant)

    'VBUtil.getItem columnInfoList, key
End Function
