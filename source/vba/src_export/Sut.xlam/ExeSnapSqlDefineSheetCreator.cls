VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSnapSqlDefineSheetCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スナップショットSQL定義シートを生成するクラス
'
' 作成者　：Ison
' 履歴　　：2019/01/03 新規作成
'
' 特記事項：
' *********************************************************

' ________________________________________________________
' イベント
' ________________________________________________________

' =========================================================
' ▽スナップショットSQL定義シート生成処理開始前イベント
'
' 概要　　　：スナップショットSQL定義シートの生成前に呼び出されるイベント
' 引数　　　：
'
' =========================================================
Public Event createSnapshotSqlDefineSheetBegin()

' =========================================================
' ▽スナップショットSQL定義シート生成イベント
'
' 概要　　　：スナップショットSQL定義シートの生成完了時に呼び出されるイベント
' 引数　　　：
'
' =========================================================
Public Event createSnapshotSqlDefineSheet()

' =========================================================
' プロパティ
' =========================================================
' アプリケーション設定情報
Private m_applicationSetting As ValApplicationSetting
' 処理対象ワークブック
Private m_book               As Workbook
' キャンセルフラグ
Private m_canceled           As Boolean

Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
End Property

Public Property Let book(ByVal vNewValue As Variant)

    Set m_book = vNewValue
    
End Property

Public Property Let canceled(ByVal vNewValue As Variant)

    m_canceled = vNewValue
    
End Property

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
' ▽シートを生成する
'
' 概要　　　：指定されたテーブルのシートを生成する
' 引数　　　：
' 戻り値    ：シート
'
' =========================================================
Public Function createSheet() As Worksheet

    ' シート名
    Dim sheetName     As String

    ' ◇スナップショットSQL定義シート生成イベント
    RaiseEvent createSnapshotSqlDefineSheetBegin

    sheetName = "DiffSQL_" & Format(Now, "yyyy-mm-dd_hhmmss")

    ' シートの削除（存在していたら）と挿入を実行する
    deleteAndInsertSheet sheetName
    
    ' 新たにコピーされたシートオブジェクトを取得する
    Set createSheet = m_book.Worksheets(sheetName)
    
    ' ▽新シートにテーブル定義情報を埋め込む
    With createSheet

        ' シートを可視にする
        .visible = xlSheetVisible

    End With
    
    ' ◇スナップショットSQL定義シート生成完了イベント
    RaiseEvent createSnapshotSqlDefineSheet

End Function


' =========================================================
' ▽シートの削除・挿入
'
' 概要　　　：シートが既に存在する場合は削除し、挿入を行う。
' 引数　　　：SnapshotSqlDefineSheet      テーブル情報
' 　　　　　　targetSheetName 対象シート名
'
' =========================================================
Private Sub deleteAndInsertSheet(ByRef targetSheetName As String)

    ' コピーシート
    Dim copySheetName As String
    copySheetName = ConstantsSnapSqlDefine.SHEET_NAME_TEMPLATE

    ExcelUtil.copySheetAppend SutWorkbook _
                            , copySheetName _
                            , m_book _
                            , targetSheetName

End Sub


