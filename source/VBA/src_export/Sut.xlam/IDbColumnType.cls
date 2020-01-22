VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbColumnType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのデフォルトのカラム書式情報を表現するクラス
' 対応するDB毎に本インターフェースをImplementsする。
'
' 作成者　：Ison
' 履歴　　：2009/08/14　新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' ▽DBコネクション設定
'
' 概要　　　：DBコネクションを設定する。
' 引数　　　：vNewValue DBコネクション
'
' 特記事項　：
'
'
' =========================================================
Public Property Let dbConn(ByVal vNewValue As Variant)

End Property

' =========================================================
' ▽DBのデフォルトのカラム書式情報を取得する
'
' 概要　　　：
' 引数　　　：
'
' 戻り値　　：DBカラム書式情報
'
' 特記事項　：
'
'
' =========================================================
Public Function getDefaultColumnFormat() As ValCollection

End Function
