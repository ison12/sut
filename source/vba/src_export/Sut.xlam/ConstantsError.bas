Attribute VB_Name = "ConstantsError"
Option Explicit

' *********************************************************
' エラーに関連した定数モジュール
'
' 作成者　：Ison
' 履歴　　：2009/03/31　新規作成
'
' 特記事項：
'
' *********************************************************

Public Const ERR_NUMBER_PROC_CANCEL            As Long = 1 + vbObjectError + 512
Public Const ERR_NUMBER_SQL_EXECUTE_FAILED     As Long = 2 + vbObjectError + 512
Public Const ERR_NUMBER_OUT_OF_RANGE_SHEET     As Long = 3 + vbObjectError + 512
Public Const ERR_NUMBER_OUT_OF_RANGE_SELECTION As Long = 4 + vbObjectError + 512
Public Const ERR_NUMBER_DISCONNECT_DB          As Long = 5 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_SELECTED_SCHEMA    As Long = 6 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_SELECTED_TABLE     As Long = 7 + vbObjectError + 512
Public Const ERR_NUMBER_DUPLICATE_SELECTION_CELL As Long = 8 + vbObjectError + 512
Public Const ERR_NUMBER_OVER_SELECT_COND_CONTROL As Long = 9 + vbObjectError + 512
Public Const ERR_NUMBER_IS_NOT_TABLE_SHEET      As Long = 10 + vbObjectError + 512
Public Const ERR_NUMBER_UNSUPPORT_DB           As Long = 11 + vbObjectError + 512
Public Const ERR_NUMBER_NON_ACTIVE_BOOK        As Long = 12 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_EXIST_TABLE_INFO   As Long = 13 + vbObjectError + 512
Public Const ERR_NUMBER_DLL_FUNCTION_WARNING   As Long = 14 + vbObjectError + 512
Public Const ERR_NUMBER_SHORTCUT_SETTING_FAILED As Long = 15 + vbObjectError + 512
Public Const ERR_NUMBER_POPUP_SETTING_FAILED As Long = 16 + vbObjectError + 512
Public Const ERR_NUMBER_RCLICKMENU_SETTING_FAILED As Long = 17 + vbObjectError + 512
Public Const ERR_NUMBER_FILE_OUTPUT_FAILED As Long = 18 + vbObjectError + 512
Public Const ERR_NUMBER_SQL_EMPTY            As Long = 19 + vbObjectError + 512
Public Const ERR_NUMBER_IS_NOT_SQL_DEFINE_SHEET      As Long = 20 + vbObjectError + 512
Public Const ERR_NUMBER_PK_COLUMN_NOT_FOUND   As Long = 21 + vbObjectError + 512
Public Const ERR_NUMBER_SNAP_DIFF__EXEC_ERROR   As Long = 22 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_SELECTED_DB_CONNECT   As Long = 23 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_SELECTED_TABLE_SHEET  As Long = 25 + vbObjectError + 512
Public Const ERR_NUMBER_CREATE_WORKSHEET_FAILED  As Long = 26 + vbObjectError + 512
Public Const ERR_NUMBER_ACTIVE_ADDIN_BOOK        As Long = 27 + vbObjectError + 512
Public Const ERR_NUMBER_SHEET_MISSING        As Long = 28 + vbObjectError + 512
Public Const ERR_NUMBER_CELL_MAX_LENGTH_OVER As Long = 29 + vbObjectError + 512
Public Const ERR_NUMBER_CELL_MAX_LENGTH_OVER_REFLECT As Long = 30 + vbObjectError + 512

Public Const ERR_NUMBER_REG_EXP_NOT_CREATED   As Long = 997 + vbObjectError + 512
Public Const ERR_NUMBER_REGISTRY_ACCESS_FAILED   As Long = 998 + vbObjectError + 512
Public Const ERR_NUMBER_DLL_FUNCTION_FAILED      As Long = 999 + vbObjectError + 512

Public Const ERR_DESC_PROC_CANCEL              As String = "処理がキャンセルされました。"
Public Const ERR_DESC_SQL_EXECUTE_FAILED       As String = "SQL実行時にエラーが発生しました。"
Public Const ERR_DESC_OUT_OF_RANGE_SHEET       As String = "レコード数が多いため、全てのレコードをシートに取り込むことができませんでした。"
Public Const ERR_DESC_OUT_OF_RANGE_SELECTION   As String = "セルの選択領域が入力範囲外にあります。"
Public Const ERR_DESC_DISCONNECT_DB            As String = "データベースに接続されていません。"
Public Const ERR_DESC_NOT_SELECTED_SCHEMA      As String = "スキーマを1つ以上選択してください。"
Public Const ERR_DESC_NOT_SELECTED_TABLE       As String = "テーブルを1つ以上選択してください。"
Public Const ERR_DESC_DUPLICATE_SELECTION_CELL As String = "選択したセルが重複しています。"
Public Const ERR_DESC_OVER_SELECT_COND_CONTROL As String = "プライマリキーがコントロールより多いため正しく設定できませんでした。"
Public Const ERR_DESC_IS_NOT_TABLE_SHEET       As String = "テーブルシートではないため実行できません。"
Public Const ERR_DESC_UNSUPPORT_DB             As String = "未対応のDBに接続されています。"
Public Const ERR_DESC_NON_ACTIVE_BOOK          As String = "ワークブックがアクティブになっていないため実行できません。"
Public Const ERR_DESC_NOT_EXIST_TABLE_INFO     As String = "テーブル情報が取得できませんでした。" & vbNewLine & _
                                                           "接続中のDBに対象テーブルが存在しているか確認してください。"
Public Const ERR_DESC_REG_EXP_NOT_CREATED   As String = "正規表現オブジェクトの生成に失敗しました。PCにIE5.0以上がインストールされている必要があります。"
Public Const ERR_DESC_DLL_FUNCTION_WARNING     As String = "DLLの呼び出しに失敗しました。"
Public Const ERR_DESC_SHORTCUT_SETTING_FAILED  As String = "ショートカットキーの設定に失敗しました。"
Public Const ERR_DESC_POPUP_SETTING_FAILED     As String = "ポップアップメニューの設定に失敗しました。"
Public Const ERR_DESC_RCLICKMENU_SETTING_FAILED As String = "右クリックメニューの設定に失敗しました。"
Public Const ERR_DESC_FILE_OUTPUT_FAILED As String = "ファイル出力に失敗しました。"
Public Const ERR_DESC_SQL_EMPTY                As String = "SQLが未入力です。"
Public Const ERR_DESC_IS_NOT_SQL_DEFINE_SHEET  As String = "SQL定義シートではないため実行できません。"
Public Const ERR_DESC_PK_COLUMN_NOT_FOUND      As String = "PKカラムが見つかりませんでした。"
Public Const ERR_DESC_SNAP_DIFF__EXEC_ERROR    As String = "スナップショット比較実行時にエラーが発生しました。"

Public Const ERR_DESC_NOT_SELECTED_DB_CONNECT  As String = "接続情報を選択してください。"
Public Const ERR_DESC_NOT_SELECTED_TABLE_SHEET As String = "シートを1つ以上選択してください。"
Public Const ERR_DESC_CREATE_WORKSHEET_FAILED  As String = "ワークブックの作成に失敗しました。"
Public Const ERR_DESC_ACTIVE_ADDIN_BOOK        As String = "アドインブックがアクティブになっています。" & vbNewLine & "他のブックをアクティブにして再度実行してください。"
Public Const ERR_DESC_SHEET_MISSING            As String = "対象となるシートから情報を読み取れません。" & vbNewLine & "削除された可能性があります。"
Public Const ERR_DESC_CELL_MAX_LENGTH_OVER     As String = "セルへの入力可能な最大文字数（32767文字）を超えて、データが欠落しました。"
Public Const ERR_DESC_CELL_MAX_LENGTH_OVER_REFLECT    As String = "セルへのデータ反映時に最大文字数（32767文字）を超えて、データが欠落しました。"

Public Const ERR_DESC_REGISTRY_ACCESS_FAILED   As String = "レジストリのアクセスに失敗しました。"
Public Const ERR_DESC_DLL_FUNCTION_FAILED      As String = "DLLの呼び出しに失敗しました。"

Public Const ERR_DESC_COLUMN_SIZE_OVER_SHEET_SIZE As String = "カラム数が多いため、全てのカラムをシートに取り込むことができませんでした。"

Public Const ERR_MSG_ERROR_LEVEL               As String = "エラーが発生しました。"

Public Const VALID_ERR_NUMERIC                 As String = "数値を入力してください。"
Public Const VALID_ERR_INTEGER                 As String = "数値を入力してください。(小数部分含まず)"
Public Const VALID_ERR_NO_LIST_ITEM            As String = "リストから項目を選択してください。"
Public Const VALID_ERR_REQUIRED                As String = "必須入力です。"
Public Const VALID_ERR_AND_OVER                As String = "{1}以上の数値を入力してください。"
Public Const VALID_ERR_AND_LESS                As String = "{1}以下の数値を入力してください。"
Public Const VALID_ERR_INVALID                 As String = "入力値が不正です。"
Public Const VALID_ERR_INVALID_SIZE            As String = "入力値のサイズが不正です。"
Public Const VALID_ERR_NOT_ALPHA_NUM_MARK_FULL As String = "英数字または記号( -   のみ_)と全角文字のみを入力してください。"

' =========================================================
' ▽アプリケーションエラーチェック
'
' 概要　　　：本アプリケーションで発生したエラーであるかをチェックする。
' 引数　　　：num エラー番号
' 戻り値　　：True 本アプリケーションで発生したエラー
'
' =========================================================
Public Function isApplicationError(ByVal num As Long) As Boolean

    If 1 + vbObjectError + 512 <= num And num <= 900 + vbObjectError + 512 Then
    
        isApplicationError = True
    Else
    
        isApplicationError = False
    End If
    
End Function
