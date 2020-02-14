VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbMetaInfoGetterAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DBメタ情報取得クラス
' Access用
'
' IDbMetaInfoGetterを実装したクラス
'
' 作成者　：Ison
' 履歴　　：2019/12/15　新規作成
'
' 特記事項：
' *********************************************************

Implements IDbMetaInfoGetter

' DBコネクション
Private m_dbConn As Object
' DBMS種類
Private m_dbms   As DbmsType

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
' ▽DBコネクション設定
'
' 概要　　　：DBコネクションを設定する。
' 引数　　　：vNewValue DBコネクション
'
' 特記事項　：
'
'
' =========================================================
Private Property Let IDbMetaInfoGetter_dbConn(ByVal RHS As Variant)

    ' DBコネクションを設定
    Set m_dbConn = RHS
    ' DB種類を設定
    m_dbms = ADOUtil.getDBMSType(m_dbConn)
    
End Property

' =========================================================
' ▽スキーマ情報リスト取得
'
' 概要　　　：接続しているDBの全スキーマ情報を取得する
' 引数　　　：
'
' 戻り値　　：スキーマリスト。
'
' 特記事項　：
'
'
' =========================================================
Private Function IDbMetaInfoGetter_getSchemaList() As ValCollection

    On Error GoTo err
    
    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
        
End Function

' =========================================================
' ▽テーブル情報取得（＋カラム情報）
'
' 概要　　　：任意のテーブル情報を全て取得する。
' 引数　　　：schemaList スキーマリスト
'
' 戻り値　　：テーブル定義情報。
'
' 特記事項　：
'
'
' =========================================================
Private Function IDbMetaInfoGetter_getTableList(ByRef schemaList As ValCollection) As ValCollection

    On Error GoTo err
    
    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
        
End Function

' =========================================================
' ▽テーブルリスト取得（＋カラム情報）
'
' 概要　　　：任意のスキーマ・テーブル名をキーにしてテーブル情報を取得する。
' 引数　　　：schemaName スキーマ名
' 　　　　　　tableName  テーブル名
'
' 戻り値　　：テーブル定義リスト。
'
' 特記事項　：
'
'
' =========================================================
Private Function IDbMetaInfoGetter_getTable(schemaName As String, tableName As String) As ValDbDefineTable

    On Error GoTo err
    
    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext

End Function

