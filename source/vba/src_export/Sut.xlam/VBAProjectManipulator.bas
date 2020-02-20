Attribute VB_Name = "VBAProjectManipulator"
Option Explicit

' *********************************************************
' アドインブックのモジュールをエクスポート・インポートする機能。
'
' 作成者　：Ison
' 履歴　　：2020/02/17　新規作成
'
' 特記事項：
'
' *********************************************************

' =========================================================
' ▽全ファイルをエクスポート
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
'
' 戻り値　　：
'
' =========================================================
Public Sub exportAll(Optional ByVal filePath As String = "")

    Dim module                  As Object      ' モジュール
    Dim moduleList              As Object      ' VBAプロジェクトの全モジュール
    Dim extension               As String      ' モジュールの拡張子
    
    Dim exportFilePath          As String      ' エクスポートファイルパス
    
    Dim targetBook              As Workbook    ' 処理対象ブックオブジェクト
    
    Set targetBook = ThisWorkbook
    
    If filePath = "" Then
        filePath = ThisWorkbook.path & "\module"
    End If
    
    ' エクスポート先のディレクトリを一旦削除して再度作成する
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) = True Then
        fso.DeleteFolder filePath
        fso.CreateFolder filePath
    Else
        fso.CreateFolder filePath
    End If
    
    ' 処理対象ブックのモジュール一覧を取得
    Set moduleList = targetBook.VBProject.VBComponents
    
    ' VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
    
        ' クラス
        If (module.Type = 2) Then
            extension = "cls"
        ' フォーム
        ElseIf (module.Type = 3) Then
            ' .frxも一緒にエクスポートされる
            extension = "frm"
        ' 標準モジュール
        ElseIf (module.Type = 1) Then
            extension = "bas"
        ' その他
        Else
            ' エクスポート対象外のため次ループへ
            GoTo continue
        End If
        
        ' エクスポート実施
        exportFilePath = filePath & "\" & module.name & "." & extension
        Call module.export(exportFilePath)
        
        ' 出力先確認用ログ出力
        Debug.Print exportFilePath
        
continue:

    Next
    
End Sub

' =========================================================
' ▽全ファイルをインポートまたは削除
'
' 概要　　　：
' 引数　　　：filePath     ファイルパス
'           : isDeleteOnly 削除のみフラグ
'
' 戻り値　　：
'
' =========================================================
Public Sub importOrDeleteAll(Optional ByVal filePath As String = "", Optional ByVal isDeleteOnly As Boolean = False)
    
    On Error Resume Next
    
    ' 対象ブック
    Dim book As Workbook
    Set book = ThisWorkbook
    
    Dim oFso As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    Dim moduleList()     As String       ' モジュールファイル配列
    Dim module           As Variant      ' モジュールファイル
    Dim extension        As String       ' 拡張子
    
    If filePath = "" Then
        filePath = ThisWorkbook.path & "\module"
    End If
    
    ReDim moduleList(0)
    
    ' 全モジュールのファイルパスを取得
    Call searchAllFile(filePath, moduleList)
    
    ' 全モジュールをループ
    For Each module In moduleList
        
        ' 拡張子を小文字で取得
        extension = LCase(oFso.GetExtensionName(module))
        
        ' 拡張子がcls、frm、basのいずれかの場合
        If (extension = "cls" Or extension = "frm" Or extension = "bas") Then
            
            If oFso.getbasename(module) <> "VBAProjectManipulator.bas" Then
            
                ' 同名モジュールを削除
                Call book.VBProject.VBComponents.remove(book.VBProject.VBComponents(oFso.getbasename(module)))
                
                If isDeleteOnly = False Then
                    ' モジュールを追加
                    Call book.VBProject.VBComponents.Import(module)
                End If
                ' 確認用ログ出力
                Debug.Print module
            End If
        
        End If
    Next
    
End Sub

' =========================================================
' ▽任意のディレクトリ内の全ファイルを再帰的に検索する
'
' 概要　　　：
' 引数　　　：dirPath  ディレクトリパス
'     　　　：fileList ファイルリスト
'
' 戻り値　　：
'
' =========================================================
Private Sub searchAllFile(dirPath As String, fileList() As String)
    
    Dim oFso        As Object
    Dim oFolder     As Object
    Dim oSubFolder  As Object
    Dim oFile       As Object
    
    Dim i
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダがない場合
    If (oFso.folderexists(dirPath) = False) Then
        Exit Sub
    End If
    
    Set oFolder = oFso.GetFolder(dirPath)
    
    ' サブフォルダを再帰（サブフォルダを探す必要がない場合はこのFor文を削除してください）
    For Each oSubFolder In oFolder.SubFolders
        Call searchAllFile(oSubFolder.path, fileList)
    Next
    
    i = UBound(fileList)
    
    ' カレントフォルダ内のファイルを取得
    For Each oFile In oFolder.Files
    
        If (i <> 0 Or fileList(i) <> "") Then
            i = i + 1
            ReDim Preserve fileList(i)
        End If
        
        ' ファイルパスを配列に格納
        fileList(i) = oFile.path
    Next
    
End Sub
