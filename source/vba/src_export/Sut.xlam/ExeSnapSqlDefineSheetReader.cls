VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSnapSqlDefineSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' シートからスナップショットSQL定義情報を読み取る処理を定義したクラス
'
' 作成者　：Ison
' 履歴　　：2019/01/04　新規作成
'     　　：2020/02/20　パラメータセルを削除
'
' 特記事項：
' *********************************************************

' ワークシート
Public sheet     As Worksheet
' DBコネクション
Public conn      As Object
' キャンセルフラグ
Public canceled  As Boolean

' =========================================================
' ▽SQL定義読込開始前イベント
'
' 概要　　　：SQL定義の読込を開始する前に呼び出されるイベント
' 引数　　　：
'
' =========================================================
Public Event readSqlDefineInfoBegin()

' =========================================================
' ▽SQL定義読込イベント
'
' 概要　　　：SQL定義の読込時に呼び出されるイベント
' 引数　　　：sqlDefineSheet SQL定義シート
'
' =========================================================
Public Event readSqlDefineInfo(ByRef sqlDefineSheet As ValSnapSqlDefineWorkSheet)

' =========================================================
' ▽SQLリスト読込開始前イベント
'
' 概要　　　：SQLリストの読込を開始する前に呼び出されるイベント
' 引数　　　：cnt 件数
'
' =========================================================
Public Event readSqlListBegin(ByVal cnt As Long)

' =========================================================
' ▽SQL情報読込イベント
'
' 概要　　　：SQL情報読込時に呼び出されるイベント
' 引数　　　：sql SQL情報
'
' =========================================================
Public Event readSql(ByRef sql As ValSnapSqlDefineSql)

' =========================================================
' ▽SQLリスト読込イベント
'
' 概要　　　：SQLリスト読込時に呼び出されるイベント
' 引数　　　：
'
' =========================================================
Public Event readSqlList()

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
' ▽SQL定義シート判定
'
' 概要　　　：対象シート（メンバのsheet）が
' 　　　　　　SQL定義シートであるかを判定する。
'
' 引数　　　：
' 戻り値　　：True SQL定義シートの場合
' =========================================================
Public Function isSqlDefineSheet() As Boolean

    ' テーブルシート判定用のオートシェイプ
    Dim s As Shape

    ' 戻り値を初期化する
    isSqlDefineSheet = False
    
    ' ------------------------------------------------
    ' エラーが発生しても継続する
    On Error Resume Next

    Set s = sheet.Shapes(ConstantsSnapSqlDefine.SHEET_CHECK_IMAGE)
    
    ' 取得できた場合
    If Not s Is Nothing Then
    
        isSqlDefineSheet = True
        Exit Function
    End If
        
    On Error GoTo 0
    ' ------------------------------------------------

End Function

' =========================================================
' ▽テーブルシート判定
'
' 概要　　　：対象シート（メンバのsheet）SQL定義シートであるかを判定する。
' 　　　　　　SQL定義シートではない場合、エラーを発行する。
'
' 引数　　　：
' 戻り値　　：
' =========================================================
Public Sub validSqlDefineSheet()

    ' テーブルシートではない場合
    If isSqlDefineSheet = False Then
    
        err.Raise ConstantsError.ERR_NUMBER_IS_NOT_SQL_DEFINE_SHEET _
            , _
            , ConstantsError.ERR_DESC_IS_NOT_SQL_DEFINE_SHEET
    End If

End Sub

' =========================================================
' ▽SQL定義読み込み
'
' 概要　　　：シート上のSQL定義を読み込んで
' 　　　　　　SQL定義シートオブジェクトを生成する。
'
' 引数　　　：SqlDefineInfoOnly テーブル情報のみ取得（カラム情報は取得しない）
' 戻り値　　：テーブルシート
'
' =========================================================
Public Function readSqlDefineInfo() As ValSnapSqlDefineWorkSheet

    ' SQL定義
    Dim sqlDefine As ValSnapSqlDefine
    ' SQL定義情報を生成する
    Set sqlDefine = New ValSnapSqlDefine
    ' SQL定義シート
    Dim sqlDefineSheet As ValSnapSqlDefineWorkSheet
    ' SQL定義シート情報を生成する
    Set sqlDefineSheet = New ValSnapSqlDefineWorkSheet
    Set sqlDefineSheet.snapSqlDefine = sqlDefine
    
    ' 名前
    Dim name As String
    
    ' テーブル生成開始イベント
    RaiseEvent readSqlDefineInfoBegin
    
    With sheet

        ' テーブル名を取得する
        name = sheet.name
        sqlDefine.name = name
        
        ' テーブル情報取得元シートを設定する
        Set sqlDefineSheet.sheet = sheet
        sqlDefineSheet.sheetName = sheet.name
        
        readSqlDefineSqlsInfo sqlDefineSheet
        
    End With ' sheet

    ' SQL定義生成完了イベント
    RaiseEvent readSqlDefineInfo(sqlDefineSheet)

    ' 戻り値を設定する
    Set readSqlDefineInfo = sqlDefineSheet

End Function

' =========================================================
' ▽SQL定義のSQLリスト読み込み
'
' 概要　　　：シート上のSQL定義のSQLリストを読み込む。
'
' 引数　　　：sqlDefineSheet SQL定義シート
' 戻り値　　：
'
' =========================================================
Private Function readSqlDefineSqlsInfo(ByRef sqlDefineSheet As ValSnapSqlDefineWorkSheet)

    Dim i As Long
    
    ' 改行コード（セルの改行コードはLF）
    Dim newline As String: newline = vbLf

    ' SQL定義
    Dim sqlDefine      As ValSnapSqlDefine
    Set sqlDefine = sqlDefineSheet.snapSqlDefine
    
    ' SQL
    Dim sql   As ValSnapSqlDefineSql
    Dim sqlPkObj As ValSnapSqlDefineSqlPk
    
    ' ループ継続フラグ
    Dim continue   As Variant
    
    ' 現在行
    Dim currentRow As Long
    
    ' SQL情報
    Dim sqlText    As String    ' SQL
    Dim sqlPk      As String    ' Primary Key
    Dim sqlPkArr() As String
    Dim sqlMemo    As String    ' 備考
    
    ' 行の最初の位置を取得
    currentRow = ConstantsSnapSqlDefine.SQL_DEFINE_START_ROW
    
    ' SQLサイズ
    Dim sqlSize As Long
    ' SQLサイズを取得する
    sqlSize = getSqlRowSize(currentRow)
    
    ' SQLサイズが0の場合
    If sqlSize <= 0 Then
    
        Exit Function
    End If

    ' SQL生成開始イベント
    RaiseEvent readSqlListBegin(sqlSize)
    
    ' ---------------------------------------------------------------------
    With sheet
    
        ' ループを繰り返し行うかどうかを判定するための文字列を取得する
        continue = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_ROW_COL)
        
        ' ループ継続条件（continueが空ではない場合継続）
        Do While continue <> Empty
        
            ' レコードにカラムを一つ追加する
            Set sql = New ValSnapSqlDefineSql
            
            ' パラメータ名を取得
            sqlText = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_SQL_COL)
            sql.sql = sqlText
            
            ' パラメータ値を取得
            sqlPk = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_PRIMARY_KEY_COL)
            sqlPk = replace(sqlPk, vbNewLine, vbLf)
            sqlPk = replace(sqlPk, vbCr, vbLf)
            sqlPkArr = Split(sqlPk, newline)
            
            For i = LBound(sqlPkArr) To UBound(sqlPkArr)
                Set sqlPkObj = New ValSnapSqlDefineSqlPk
                sqlPkObj.pk = sqlPkArr(i)
                sql.addPk sqlPkObj
            Next i
            
            ' 備考を取得
            sqlMemo = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_MEMO_COL)
            sql.memo = sqlMemo
            
            If sqlText <> "" And sqlPk <> "" Then
            
                sqlDefine.addSql sql
                
                ' SQL生成完了イベント
                RaiseEvent readSql(sql)
                
            End If
        
            ' 行を一つ進める
            currentRow = currentRow + 1
            
            ' ループを繰り返し行うかどうかを判定するための文字列を取得する
            continue = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_ROW_COL)
        
        Loop
        ' ---------------------------------------------------------------------
    
        ' SQLリスト生成完了イベント
        RaiseEvent readSqlList
    
    End With ' sheet

End Function

' =========================================================
' ▽シートのSQLリストサイズ取得
'
' 概要　　　：SQLリストサイズの取得を行う
'
' 引数　　　：
' 戻り値　　：SQLリストサイズ
'
' =========================================================
Public Function getSqlRowSize(ByVal startRow As Long) As Long

    ' 戻り値変数
    Dim ret As Long
    ret = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsSnapSqlDefine.SQL_DEFINE_ROW_COL) - startRow + 1
    
    ' 戻り値を設定する
    getSqlRowSize = ret
    
End Function

