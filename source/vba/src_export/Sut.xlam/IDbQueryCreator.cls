VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbQueryCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBのクエリを生成するインターフェース。
' 対応するDB毎に本インターフェースをImplementsする。
'
' 作成者　：Ison
' 履歴　　：2009/03/28　新規作成
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
' ▽空白文字読み取り方式設定
'
' 概要　　　：空白文字読み取り方式を設定する。
' 引数　　　：vNewValue 空白文字読み取り方式
'
' 特記事項　：
'
'
' =========================================================
Public Property Let emptyValueReading(ByVal vNewValue As Long)

End Property

' =========================================================
' ▽直接入力文字設定
'
' 概要　　　：直接入力文字を設定する。
' 引数　　　：vNewValue 直接入力文字
'
' 特記事項　：
'
'
' =========================================================
Public Property Let directInputChar(ByVal vNewValue As String)

End Property

' =========================================================
' ▽カラム書式情報の設定
'
' 概要　　　：
' 引数　　　：RHS カラム書式情報オブジェクト
'
' 特記事項　：
'
'
' =========================================================
Public Property Let columnFormatList(ByVal RHS As Variant)

End Property

' =========================================================
' ▽スキーマ使用の設定
'
' 概要　　　：
' 引数　　　：vNewValue スキーマ使用有無
'
' 特記事項　：
'
'
' =========================================================
Public Property Let schemaUse(ByVal vNewValue As Variant)

End Property

' =========================================================
' ▽テーブル・カラム名エスケープの設定
'
' 概要　　　：
' 引数　　　：vNewValue テーブル・カラム名エスケープ
'
' 特記事項　：
'
'
' =========================================================
Public Property Let tableColumnEscape(ByVal vNewValue As Variant)

End Property

' =========================================================
' ▽INSERT文生成
'
' 概要　　　：INSERT SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　record レコード
'
' 戻り値　　：INSERT文
'
' 特記事項　：
'
'
' =========================================================
Public Function createInsert(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

End Function

' =========================================================
' ▽UPDATE文生成
'
' 概要　　　：UPDATE SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　record レコード
'
' 戻り値　　：UPDATE文
'
' 特記事項　：
'
'
' =========================================================
Public Function createUpdate(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

End Function

' =========================================================
' ▽DELETE文生成
'
' 概要　　　：DELETE SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　record レコード
'
' 戻り値　　：DELETE文
'
' 特記事項　：
'
'
' =========================================================
Public Function createDelete(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

End Function

' =========================================================
' ▽DELETE文生成（テーブル上の全レコードを削除）
'
' 概要　　　：DELETE SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　record レコード
'
' 戻り値　　：DELETE文
'
' 特記事項　：
'
'
' =========================================================
Public Function createDeleteOfAllTable(ByRef table As ValDbDefineTable) As String

End Function

' =========================================================
' ▽SELECT文生成
'
' 概要　　　：INSERT SQLを生成する
' 引数　　　：table  テーブル定義
' 　　　　　　condition 条件句
'
' 戻り値　　：SELECT文
'
' 特記事項　：
'
'
' =========================================================
Public Function createSelect(ByRef table As ValDbDefineTable, ByRef condition As ValSelectCondition) As String

End Function
