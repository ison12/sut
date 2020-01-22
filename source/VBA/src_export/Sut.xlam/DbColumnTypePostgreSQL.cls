VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbColumnTypePostgreSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDbColumnType

' *********************************************************
' DBのカラム書式情報を表現するクラス
'
' 作成者　：Ison
' 履歴　　：2009/08/14　新規作成
'
' 特記事項：
'
' *********************************************************

Private m_dbConn As Object

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
Private Property Let IDbColumnType_dbConn(ByVal RHS As Variant)

    Set m_dbConn = RHS

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
Private Function IDbColumnType_getDefaultColumnFormat() As ValCollection

    ' データ型読み込みオブジェクトを生成する
    Dim dataTypeReader As New ExeDataTypeReader
    Set dataTypeReader.sheet = SutWorkbook.Worksheets("data_type_postgresql")
    
    ' 戻り値を設定する
    Set IDbColumnType_getDefaultColumnFormat = dataTypeReader.execute

End Function

