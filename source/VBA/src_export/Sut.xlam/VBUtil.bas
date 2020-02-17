Attribute VB_Name = "VBUtil"
Option Explicit

' *********************************************************
' VB関連の共通関数モジュール
'
' 作成者　：Ison
' 履歴　　：2008/08/10　新規作成
'
' 特記事項：
'
' *********************************************************

' エラー情報を格納する構造体
Public Type errInfo

    Source       As Variant
    Number       As Variant
    Description  As Variant
    LastDllError As Variant
    HelpFile     As Variant
    HelpContext  As Variant
    
End Type

Private Const KEY_CODE_CTRL  As String = "^"
Private Const KEY_CODE_SHIFT As String = "+"
Private Const KEY_CODE_ALT   As String = "%"

' Application#OnKeyの文字マップ
Private applicationOnKeyMap1 As ValCollection ' 論理名をキーにしている
Private applicationOnKeyMap2 As ValCollection ' コードをキーにしている

' レジストリパス - 文字コード一覧
Private Const REG_PATH_CHARACTER_CODE_LIST As String = "MIME\Database\Charset"
' レジストリキー - 文字コードの別名
Private Const REG_KEY_ALIAS_CHARSET As String = "AliasForCharset"

Public Const NEW_LINE_STR_CRLF As String = "CRLF"
Public Const NEW_LINE_STR_CR As String = "CR"
Public Const NEW_LINE_STR_LF As String = "LF"


Public Function getAppOnKeyNameOfShiftByCode(ByVal shiftCode As String)

    If KEY_CODE_CTRL = shiftCode Then
    
        getAppOnKeyNameOfShiftByCode = "Ctrl"
    ElseIf KEY_CODE_SHIFT = shiftCode Then
    
        getAppOnKeyNameOfShiftByCode = "Shift"
    ElseIf KEY_CODE_ALT = shiftCode Then
    
        getAppOnKeyNameOfShiftByCode = "Alt"
    Else
    
        getAppOnKeyNameOfShiftByCode = ""
    End If

End Function

' =========================================================
' ▽Application#OnKey関数に適用可能なKeyコードリストを初期化する
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub initializeAppOnKeyMap()

    If applicationOnKeyMap1 Is Nothing Then
    
        Set applicationOnKeyMap1 = New ValCollection
        applicationOnKeyMap1.setItem "1", "1"
        applicationOnKeyMap1.setItem "2", "2"
        applicationOnKeyMap1.setItem "3", "3"
        applicationOnKeyMap1.setItem "4", "4"
        applicationOnKeyMap1.setItem "5", "5"
        applicationOnKeyMap1.setItem "6", "6"
        applicationOnKeyMap1.setItem "7", "7"
        applicationOnKeyMap1.setItem "8", "8"
        applicationOnKeyMap1.setItem "9", "9"
        applicationOnKeyMap1.setItem "0", "0"
        applicationOnKeyMap1.setItem "a", "a"
        applicationOnKeyMap1.setItem "b", "b"
        applicationOnKeyMap1.setItem "c", "c"
        applicationOnKeyMap1.setItem "d", "d"
        applicationOnKeyMap1.setItem "e", "e"
        applicationOnKeyMap1.setItem "f", "f"
        applicationOnKeyMap1.setItem "g", "g"
        applicationOnKeyMap1.setItem "h", "h"
        applicationOnKeyMap1.setItem "i", "i"
        applicationOnKeyMap1.setItem "j", "j"
        applicationOnKeyMap1.setItem "k", "k"
        applicationOnKeyMap1.setItem "l", "l"
        applicationOnKeyMap1.setItem "m", "m"
        applicationOnKeyMap1.setItem "n", "n"
        applicationOnKeyMap1.setItem "o", "o"
        applicationOnKeyMap1.setItem "p", "p"
        applicationOnKeyMap1.setItem "q", "q"
        applicationOnKeyMap1.setItem "r", "r"
        applicationOnKeyMap1.setItem "s", "s"
        applicationOnKeyMap1.setItem "t", "t"
        applicationOnKeyMap1.setItem "u", "u"
        applicationOnKeyMap1.setItem "v", "v"
        applicationOnKeyMap1.setItem "w", "w"
        applicationOnKeyMap1.setItem "x", "x"
        applicationOnKeyMap1.setItem "y", "y"
        applicationOnKeyMap1.setItem "z", "z"
        applicationOnKeyMap1.setItem "-", "-"
        applicationOnKeyMap1.setItem "{^}", "^"
        applicationOnKeyMap1.setItem "\", "\"
        applicationOnKeyMap1.setItem "@", "@"
        applicationOnKeyMap1.setItem "{[}", "["
        applicationOnKeyMap1.setItem ";", ";"
        applicationOnKeyMap1.setItem ":", ":"
        applicationOnKeyMap1.setItem "{]}", "]"
        applicationOnKeyMap1.setItem ",", ","
        applicationOnKeyMap1.setItem ".", "."
        applicationOnKeyMap1.setItem "/", "/"
        applicationOnKeyMap1.setItem "\", "\"
        applicationOnKeyMap1.setItem "{F1}", "F1"
        applicationOnKeyMap1.setItem "{F2}", "F2"
        applicationOnKeyMap1.setItem "{F3}", "F3"
        applicationOnKeyMap1.setItem "{F4}", "F4"
        applicationOnKeyMap1.setItem "{F5}", "F5"
        applicationOnKeyMap1.setItem "{F6}", "F6"
        applicationOnKeyMap1.setItem "{F7}", "F7"
        applicationOnKeyMap1.setItem "{F8}", "F8"
        applicationOnKeyMap1.setItem "{F9}", "F9"
        applicationOnKeyMap1.setItem "{F10}", "F10"
        applicationOnKeyMap1.setItem "{F11}", "F11"
        applicationOnKeyMap1.setItem "{F12}", "F12"
        applicationOnKeyMap1.setItem "{F13}", "F13"
        applicationOnKeyMap1.setItem "{F14}", "F14"
        applicationOnKeyMap1.setItem "{F15}", "F15"
        applicationOnKeyMap1.setItem "{NUMLOCK}", "Num Lock"
        applicationOnKeyMap1.setItem "{96}", "10key(0)"
        applicationOnKeyMap1.setItem "{97}", "10key(1)"
        applicationOnKeyMap1.setItem "{98}", "10key(2)"
        applicationOnKeyMap1.setItem "{99}", "10key(3)"
        applicationOnKeyMap1.setItem "{100}", "10key(4)"
        applicationOnKeyMap1.setItem "{101}", "10key(5)"
        applicationOnKeyMap1.setItem "{102}", "10key(6)"
        applicationOnKeyMap1.setItem "{103}", "10key(7)"
        applicationOnKeyMap1.setItem "{104}", "10key(8)"
        applicationOnKeyMap1.setItem "{105}", "10key(9)"
        applicationOnKeyMap1.setItem "{106}", "10key(*)"
        applicationOnKeyMap1.setItem "{107}", "10key(+)"
        applicationOnKeyMap1.setItem "{109}", "10key(-)"
        applicationOnKeyMap1.setItem "{110}", "10key(.)"
        applicationOnKeyMap1.setItem "{111}", "10key(/)"
        applicationOnKeyMap1.setItem "{ENTER}", "10key(Enter)"
        applicationOnKeyMap1.setItem "{ESC}", "Esc"
        applicationOnKeyMap1.setItem "{TAB}", "Tab"
        applicationOnKeyMap1.setItem "{CAPSLOCK}", "CapsLock"
        applicationOnKeyMap1.setItem "{BS}", "BackSpace"
        applicationOnKeyMap1.setItem "~", "Enter"
        applicationOnKeyMap1.setItem "{RETURN}", "Return"
        applicationOnKeyMap1.setItem "{SCROLLLOCK}", "ScrollLock"
        applicationOnKeyMap1.setItem "{BREAK}", "Break"
        applicationOnKeyMap1.setItem "{CLEAR}", "Clear"
        applicationOnKeyMap1.setItem "{INSERT}", "Ins"
        applicationOnKeyMap1.setItem "{DEL}", "Delete(Del)"
        applicationOnKeyMap1.setItem "{HOME}", "Home"
        applicationOnKeyMap1.setItem "{END}", "End"
        applicationOnKeyMap1.setItem "{PGUP}", "PageUp"
        applicationOnKeyMap1.setItem "{PGDN}", "PageDown"
        applicationOnKeyMap1.setItem "{HELP}", "Help"
        applicationOnKeyMap1.setItem "{UP}", "↑"
        applicationOnKeyMap1.setItem "{DOWN}", "↓"
        applicationOnKeyMap1.setItem "{LEFT}", "←"
        applicationOnKeyMap1.setItem "{RIGHT}", "→"

    End If
    
    If applicationOnKeyMap2 Is Nothing Then
    
        Set applicationOnKeyMap2 = New ValCollection
        applicationOnKeyMap2.setItem "1", "1"
        applicationOnKeyMap2.setItem "2", "2"
        applicationOnKeyMap2.setItem "3", "3"
        applicationOnKeyMap2.setItem "4", "4"
        applicationOnKeyMap2.setItem "5", "5"
        applicationOnKeyMap2.setItem "6", "6"
        applicationOnKeyMap2.setItem "7", "7"
        applicationOnKeyMap2.setItem "8", "8"
        applicationOnKeyMap2.setItem "9", "9"
        applicationOnKeyMap2.setItem "0", "0"
        applicationOnKeyMap2.setItem "a", "a"
        applicationOnKeyMap2.setItem "b", "b"
        applicationOnKeyMap2.setItem "c", "c"
        applicationOnKeyMap2.setItem "d", "d"
        applicationOnKeyMap2.setItem "e", "e"
        applicationOnKeyMap2.setItem "f", "f"
        applicationOnKeyMap2.setItem "g", "g"
        applicationOnKeyMap2.setItem "h", "h"
        applicationOnKeyMap2.setItem "i", "i"
        applicationOnKeyMap2.setItem "j", "j"
        applicationOnKeyMap2.setItem "k", "k"
        applicationOnKeyMap2.setItem "l", "l"
        applicationOnKeyMap2.setItem "m", "m"
        applicationOnKeyMap2.setItem "n", "n"
        applicationOnKeyMap2.setItem "o", "o"
        applicationOnKeyMap2.setItem "p", "p"
        applicationOnKeyMap2.setItem "q", "q"
        applicationOnKeyMap2.setItem "r", "r"
        applicationOnKeyMap2.setItem "s", "s"
        applicationOnKeyMap2.setItem "t", "t"
        applicationOnKeyMap2.setItem "u", "u"
        applicationOnKeyMap2.setItem "v", "v"
        applicationOnKeyMap2.setItem "w", "w"
        applicationOnKeyMap2.setItem "x", "x"
        applicationOnKeyMap2.setItem "y", "y"
        applicationOnKeyMap2.setItem "z", "z"
        applicationOnKeyMap2.setItem "-", "-"
        applicationOnKeyMap2.setItem "^", "{^}"
        applicationOnKeyMap2.setItem "\", "\"
        applicationOnKeyMap2.setItem "@", "@"
        applicationOnKeyMap2.setItem "[", "{[}"
        applicationOnKeyMap2.setItem ";", ";"
        applicationOnKeyMap2.setItem ":", ":"
        applicationOnKeyMap2.setItem "]", "{]}"
        applicationOnKeyMap2.setItem ",", ","
        applicationOnKeyMap2.setItem ".", "."
        applicationOnKeyMap2.setItem "/", "/"
        applicationOnKeyMap2.setItem "\", "\"
        applicationOnKeyMap2.setItem "F1", "{F1}"
        applicationOnKeyMap2.setItem "F2", "{F2}"
        applicationOnKeyMap2.setItem "F3", "{F3}"
        applicationOnKeyMap2.setItem "F4", "{F4}"
        applicationOnKeyMap2.setItem "F5", "{F5}"
        applicationOnKeyMap2.setItem "F6", "{F6}"
        applicationOnKeyMap2.setItem "F7", "{F7}"
        applicationOnKeyMap2.setItem "F8", "{F8}"
        applicationOnKeyMap2.setItem "F9", "{F9}"
        applicationOnKeyMap2.setItem "F10", "{F10}"
        applicationOnKeyMap2.setItem "F11", "{F11}"
        applicationOnKeyMap2.setItem "F12", "{F12}"
        applicationOnKeyMap2.setItem "F13", "{F13}"
        applicationOnKeyMap2.setItem "F14", "{F14}"
        applicationOnKeyMap2.setItem "F15", "{F15}"
        applicationOnKeyMap2.setItem "Num Lock", "{NUMLOCK}"
        applicationOnKeyMap2.setItem "10key(0)", "{96}"
        applicationOnKeyMap2.setItem "10key(1)", "{97}"
        applicationOnKeyMap2.setItem "10key(2)", "{98}"
        applicationOnKeyMap2.setItem "10key(3)", "{99}"
        applicationOnKeyMap2.setItem "10key(4)", "{100}"
        applicationOnKeyMap2.setItem "10key(5)", "{101}"
        applicationOnKeyMap2.setItem "10key(6)", "{102}"
        applicationOnKeyMap2.setItem "10key(7)", "{103}"
        applicationOnKeyMap2.setItem "10key(8)", "{104}"
        applicationOnKeyMap2.setItem "10key(9)", "{105}"
        applicationOnKeyMap2.setItem "10key(*)", "{106}"
        applicationOnKeyMap2.setItem "10key(+)", "{107}"
        applicationOnKeyMap2.setItem "10key(-)", "{109}"
        applicationOnKeyMap2.setItem "10key(.)", "{110}"
        applicationOnKeyMap2.setItem "10key(/)", "{111}"
        applicationOnKeyMap2.setItem "10key(Enter)", "{ENTER}"
        applicationOnKeyMap2.setItem "Esc", "{ESC}"
        applicationOnKeyMap2.setItem "Tab", "{TAB}"
        applicationOnKeyMap2.setItem "CapsLock", "{CAPSLOCK}"
        applicationOnKeyMap2.setItem "BackSpace", "{BS}"
        applicationOnKeyMap2.setItem "Enter", "~"
        applicationOnKeyMap2.setItem "Return", "{RETURN}"
        applicationOnKeyMap2.setItem "ScrollLock", "{SCROLLLOCK}"
        applicationOnKeyMap2.setItem "Break", "{BREAK}"
        applicationOnKeyMap2.setItem "Clear", "{CLEAR}"
        applicationOnKeyMap2.setItem "Ins", "{INSERT}"
        applicationOnKeyMap2.setItem "Delete(Del)", "{DEL}"
        applicationOnKeyMap2.setItem "Home", "{HOME}"
        applicationOnKeyMap2.setItem "End", "{END}"
        applicationOnKeyMap2.setItem "PageUp", "{PGUP}"
        applicationOnKeyMap2.setItem "PageDown", "{PGDN}"
        applicationOnKeyMap2.setItem "Help", "{HELP}"
        applicationOnKeyMap2.setItem "↑", "{UP}"
        applicationOnKeyMap2.setItem "↓", "{DOWN}"
        applicationOnKeyMap2.setItem "←", "{LEFT}"
        applicationOnKeyMap2.setItem "→", "{RIGHT}"
        
    End If

End Sub

' =========================================================
' ▽Application#OnKey関数に適用可能なKeyコードリストを取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：Keyコードリスト
'
' =========================================================
Public Function getAppOnKeyCodeList() As ValCollection

    initializeAppOnKeyMap
    Set getAppOnKeyCodeList = applicationOnKeyMap2
End Function

' =========================================================
' ▽Application#OnKey関数のKeyコードを取得
'
' 概要　　　：論理名をキーにしてKeyコードを取得する。
' 引数　　　：name 論理名
' 戻り値　　：Keyコード
'
' =========================================================
Public Function getAppOnKeyCodeByName(ByVal name As String) As String

    initializeAppOnKeyMap
    getAppOnKeyCodeByName = applicationOnKeyMap1.getItem(name, vbString)
    
End Function

' =========================================================
' ▽Application#OnKey関数のKeyコードに紐づく論理名を取得
'
' 概要　　　：Keyコードに紐づく論理名を取得する。
' 引数　　　：code Keyコード
' 戻り値　　：論理名
'
' =========================================================
Public Function getAppOnKeyNameByCode(ByVal code As String) As String

    initializeAppOnKeyMap
    getAppOnKeyNameByCode = applicationOnKeyMap2.getItem(code, vbString)

End Function

' =========================================================
' ▽Application#OnKey関数のキー値を解析
'
' 概要　　　：Application#OnKey関数のキー値を解析し
' 　　　　　　戻り値用の引数に返す。
' 引数　　　：keyCode    キーコード
' 　　　　　　shiftCtrl  Ctrlキー
' 　　　　　　shiftShift Shiftキー
' 　　　　　　shiftAlt   Altキー
' 　　　　　　keyName    キー値
' 戻り値　　：
'
' =========================================================
Public Function resolveAppOnKey(ByVal KeyCode As String _
                                      , ByRef shiftCtrl As Boolean _
                                      , ByRef shiftShift As Boolean _
                                      , ByRef shiftAlt As Boolean _
                                      , ByRef keyName As String)

    initializeAppOnKeyMap
    
    ' 文字列インデックス
    Dim i      As Long
    ' 文字列長さ
    Dim length As Long
    ' 文字列から抽出した1文字
    Dim char   As String
    
    ' 戻り値用の引数を初期化する
    shiftCtrl = False
    shiftShift = False
    shiftAlt = False
    keyName = ""
    
    ' keyCodeの長さを取得する
    length = Len(KeyCode)
    
    For i = 1 To length
    
        ' 1文字抽出する
        char = Mid$(KeyCode, i, 1)
        
        ' Ctrlキー
        If char = KEY_CODE_CTRL Then
        
            shiftCtrl = True
            
        ' Shiftキー
        ElseIf char = KEY_CODE_SHIFT Then
        
            shiftShift = True
            
        ' Altキー
        ElseIf char = KEY_CODE_ALT Then
        
            shiftAlt = True
            
        ' その他のキー
        Else
        
            keyName = getAppOnKeyNameByCode(Mid$(KeyCode, i, length))
            Exit For
        End If
        
    Next

End Function

' =========================================================
' ▽Application#OnKey関数に与えるキーコードの取得
'
' 概要　　　：幾つかのパラメータからApplication#OnKey関数のキーコードを取得する。
' 引数　　　：shiftCtrl  Ctrlキー
' 　　　　　　shiftShift Shiftキー
' 　　　　　　shiftAlt   Altキー
' 　　　　　　name       キーの論理名
' 戻り値　　：キーコード
'
' =========================================================
Public Function getAppOnKeyCodeBySomeParams(ByVal shiftCtrl As Boolean _
                                                  , ByVal shiftShift As Boolean _
                                                  , ByVal shiftAlt As Boolean _
                                                  , ByVal name As String) As String

    initializeAppOnKeyMap
    
    ' 戻り値
    Dim ret As String
    
    ' Ctrlキー
    If shiftCtrl = True Then
    
        ret = ret & KEY_CODE_CTRL
    End If
        
    ' Shiftキー
    If shiftShift = True Then
    
        ret = ret & KEY_CODE_SHIFT
    End If
        
    ' Altキー
    If shiftAlt = True Then
    
        ret = ret & KEY_CODE_ALT
    End If

    ' キーを取得する
    ret = ret & getAppOnKeyCodeByName(name)

    ' 戻り値を設定する
    getAppOnKeyCodeBySomeParams = ret

End Function

' =========================================================
' ▽Application#OnKey関数に与えるキーコードの取得
'
' 概要　　　：幾つかのパラメータからApplication#OnKey関数のキーコードを取得する。
' 引数　　　：shiftCtrl  Ctrlキー
' 　　　　　　shiftShift Shiftキー
' 　　　　　　shiftAlt   Altキー
' 　　　　　　name       キーの論理名
' 戻り値　　：キーコード
'
' =========================================================
Public Function getAppOnKeyNameBySomeParams(ByVal shiftCtrl As Boolean _
                                                  , ByVal shiftShift As Boolean _
                                                  , ByVal shiftAlt As Boolean _
                                                  , ByVal name As String) As String

    initializeAppOnKeyMap
    
    ' 戻り値
    Dim ret As String
    ' 結合文字列
    Dim juncStr As String
    
    ' Ctrlキー
    If shiftCtrl = True Then
    
        ret = ret & getAppOnKeyNameOfShiftByCode(KEY_CODE_CTRL)
    End If
        
    ' Shiftキー
    If shiftShift = True Then
    
        If ret <> "" Then
            juncStr = "+"
        Else
            juncStr = ""
        End If
        
        ret = ret & juncStr & getAppOnKeyNameOfShiftByCode(KEY_CODE_SHIFT)
    End If
        
    ' Altキー
    If shiftAlt = True Then
    
        If ret <> "" Then
            juncStr = "+"
        Else
            juncStr = ""
        End If
            
        ret = ret & juncStr & getAppOnKeyNameOfShiftByCode(KEY_CODE_ALT)
    End If

    If ret <> "" Then
        juncStr = "+"
    Else
        juncStr = ""
    End If
    
    ' キーを取得する
    ret = ret & juncStr & name

    ' 戻り値を設定する
    getAppOnKeyNameBySomeParams = ret

End Function

Public Function getAppOnKeyNameByMultipleCode(ByVal KeyCode As String) As String

    Dim a As Boolean
    Dim b As Boolean
    Dim c As Boolean
    Dim d As String

    resolveAppOnKey KeyCode _
                            , a _
                            , b _
                            , c _
                            , d
                            
    getAppOnKeyNameByMultipleCode = getAppOnKeyNameBySomeParams(a _
                                                                              , b _
                                                                              , c _
                                                                              , d)
    

End Function

' =========================================================
' ▽Errオブジェクトの情報を構造体に退避
'
' 概要　　　：Errオブジェクトの情報を構造体に設定して返す。
' 引数　　　：
' 戻り値　　：エラー情報
'
' 特記事項　：エラーハンドラで別の関数を呼び出すとErrオブジェクトの情報が消えてしまうことがあり
' 　　　　　　この状態で、Err.Raiseすると正しい情報を上位のモジュールにで伝播できない。
' 　　　　　　正しい情報を伝播する場合には、本関数を利用して、一度エラー情報を退避してからErr.Raiseしてやると良い。
'
' 　　　　　　使用例：
' 　　　　　　　Dim errT As errInfo
' 　　　　　　　errT = VBUtil.swapErr

' 　　　　　　　・・・エラー時の後始末処理など
'
' 　　　　　　　Err.Raise errT.Number, errT.Source・・・
'
' =========================================================
Public Function swapErr() As errInfo

    swapErr.Source = err.Source
    swapErr.Number = err.Number
    swapErr.Description = err.Description
    swapErr.LastDllError = err.LastDllError
    swapErr.HelpFile = err.HelpFile
    swapErr.HelpContext = err.HelpContext

End Function

' =========================================================
' ▽Errオブジェクトに情報を設定する
'
' 概要　　　：
' 引数　　　：errT エラー情報
' 戻り値　　：
'
' 特記事項　：
'
' =========================================================
Public Sub setErr(ByRef errT As errInfo)

    err.Source = errT.Source
    err.Number = errT.Number
    err.Description = errT.Description
    'err.LastDllError = errT.LastDllError
    err.HelpFile = errT.HelpFile
    err.HelpContext = errT.HelpContext

End Sub

' =========================================================
' ▽保存ダイアログ表示
'
' 概要　　　：保存ダイアログを表示する
' 引数　　　：title           ダイアログのタイトル
' 　　　　　　filter          フィルタ
' 　　　　　　initialFileName 初期ファイル名
' 戻り値　　：保存ファイルパス
'
' =========================================================
Public Function openFileSaveDialog(ByVal title As String, ByVal filter As String, ByVal initialFileName As String) As String

    ' アプリケーション
    Dim xlsApp   As Application
    
    ' ファイルパス
    Dim filePath As Variant

    ' Applicationオブジェクト取得
    Set xlsApp = Application
    
    ' ダイアログで選択されたファイル名を格納
    filePath = xlsApp.GetSaveAsFilename(initialFileName:=initialFileName _
                                      , fileFilter:=filter _
                                      , title:=title)
                                      
    ' キャンセルされたかを判定する
    If filePath = False Then
    
        ' キャンセルされた場合 空文字列を返す
        openFileSaveDialog = ""
        
    Else
        ' 保存を選択された場合 ファイル名を返す
        openFileSaveDialog = filePath
    End If

End Function

' =========================================================
' ▽フォルダを開くダイアログ表示
'
' 概要　　　：フォルの開くダイアログを表示する
' 引数　　　：title           ダイアログのタイトル
'     　　　  initialFileName 初期ファイル名
' 戻り値　　：選択したフォルのファイルパス
'
' =========================================================
Public Function openFolderDialog(ByVal title As String, ByVal initialFileName As String) As Variant

    ' アプリケーション
    Dim xlsApp   As Application
    
    ' ファイルパス
    Dim fileDialogObj As FileDialog
    ' ファイルパス
    Dim filePath As Variant

    ' Applicationオブジェクト取得
    Set xlsApp = Application
    
    ' ダイアログで選択されたファイル名を格納
    Set fileDialogObj = xlsApp.FileDialog(msoFileDialogFolderPicker)
    fileDialogObj.title = title
    fileDialogObj.initialFileName = initialFileName
    fileDialogObj.Show

    If fileDialogObj.SelectedItems.count <= 0 Then
    
        ' キャンセルされた場合 空を返す
        openFolderDialog = Empty
    Else
    
        ' 単一選択される場合
        openFolderDialog = fileDialogObj.SelectedItems(1)
    
    End If

End Function

' =========================================================
' ▽開くダイアログ表示
'
' 概要　　　：開くダイアログを表示する
' 引数　　　：title           ダイアログのタイトル
' 　　　　　　filter          フィルタ
' 　　　　　　multiSelect     複数選択
' 戻り値　　：選択したファイルのファイルパス
'
' =========================================================
Public Function openFileDialog(ByVal title As String, ByVal filter As String, Optional ByVal multiSelect As Boolean = False) As Variant

    ' アプリケーション
    Dim xlsApp   As Application
    
    ' ファイルパス
    Dim filePath As Variant

    ' Applicationオブジェクト取得
    Set xlsApp = Application
    
    ' ダイアログで選択されたファイル名を格納
    filePath = xlsApp.GetOpenFilename(fileFilter:=filter _
                                    , title:=title _
                                    , multiSelect:=multiSelect)

    ' 複数選択の場合、戻り値として配列が返されるので配列かどうかを判定する
    If IsArray(filePath) Then
    
        ' 保存を選択された場合 ファイル名を返す
        openFileDialog = filePath
    
    ' 選択がキャンセルされた場合
    ElseIf filePath = False Then
    
        ' キャンセルされた場合 空を返す
        openFileDialog = Empty
        
    Else
        ' 保存を選択された場合 ファイル名を返す
        openFileDialog = filePath
    
    End If

End Function

' =========================================================
' ▽ファイルの拡張子チェック
'
' 概要　　　：ファイルの拡張子をチェックする
' 引数　　　：file      ファイル名
' 　　　　　　extension 拡張子
' 戻り値　　：ファイルの拡張子が指定された引数extensionの場合Trueを返す
'
' =========================================================
Public Function checkFileExtension(ByRef file As String _
                                 , ByRef extension As String) As Boolean

    ' ファイル名から抽出した拡張子
    Dim fileExtension As String
    
    ' インデックス
    Dim index As Long
    
    ' ファイル名と拡張子の区切り文字であるドット(.)を検索する
    index = InStrRev(file, ".")
    
    ' ドット(.)が見つからない場合
    If index <= 0 Then
    
        Exit Function
    End If
    
    ' ファイル名から拡張子を抽出する
    fileExtension = Mid$(file, index + 1, Len(file))

    If fileExtension = extension Then
    
        checkFileExtension = True
    Else
    
        checkFileExtension = False
    End If

End Function

' =========================================================
' ▽ファイルパスからファイル名抽出
'
' 概要　　　：ファイルパスからファイル名を抽出する
' 引数　　　：filePath ファイルパス
' 戻り値　　：ファイル名
'
' =========================================================
Public Function extractFileName(ByRef filePath As String) As String
    
    ' ファイルパス区切り文字
    Const FILE_SEPARATE As String = "\"

    ' ファイルパスの右後方からはじめに出現した区切り文字の文字位置
    Dim index As Long
    
    ' 区切り文字の位置を取得する
    index = InStrRev(filePath, FILE_SEPARATE)

    ' 区切り文字を発見した場合
    If index > 0 Then
    
        extractFileName = Mid$(filePath, index + 1)
    
    ' 区切り文字を発見できなかった場合
    Else
        extractFileName = filePath
    
    End If

End Function

' =========================================================
' ▽インフォメッセージボックスを表示
'
' 概要　　　：インフォメッセージボックスを表示する
' 引数　　　：basePrompt 基本メッセージ
'             title      メッセージボックスのタイトル
' 　　　　　　err        エラーオブジェクト
'
' =========================================================
Public Sub showMessageBoxForInformation(ByRef basePrompt As String _
                                      , ByRef title As String _
                                      , Optional ByRef err As ErrObject = Nothing)

    MsgBox basePrompt, vbOKOnly, title
         
End Sub

' =========================================================
' ▽エラーメッセージボックスを表示
'
' 概要　　　：エラーメッセージボックスを表示する
' 引数　　　：basePrompt 基本メッセージ
'             title      メッセージボックスのタイトル
' 　　　　　　err        エラーオブジェクト
'
' =========================================================
Public Sub showMessageBoxForError(ByRef basePrompt As String _
                                , ByRef title As String _
                                , ByRef err As ErrObject)

    MsgBox basePrompt & vbNewLine & vbNewLine & _
           err.Description & vbNewLine & _
           "Error no [" & err.Number & "]" & vbNewLine & _
           "Source [" & err.Source & "]" _
           , vbOKOnly + vbCritical _
           , title

End Sub

' =========================================================
' ▽警告メッセージボックスを表示
'
' 概要　　　：警告メッセージボックスを表示する
' 引数　　　：basePrompt 基本メッセージ
'             title      メッセージボックスのタイトル
' 　　　　　　err        エラーオブジェクト
'
' =========================================================
Public Sub showMessageBoxForWarning(ByVal basePrompt As String _
                                  , ByVal title As String _
                                  , ByRef err As ErrObject)

    If err Is Nothing Then
    
        MsgBox basePrompt _
               , vbOKOnly + vbExclamation _
               , title
    
    ElseIf err.Number = 0 Then
    
        MsgBox basePrompt _
               , vbOKOnly + vbExclamation _
               , title
               
    Else
    
        If basePrompt <> "" Then
        
            basePrompt = basePrompt & vbNewLine & vbNewLine
        End If
        
        MsgBox basePrompt & _
               err.Description & vbNewLine & _
               "Error no [" & err.Number & "]" _
               , vbOKOnly + vbExclamation _
               , title
    End If
         
End Sub

' =========================================================
' ▽はい・いいえ・キャンセルメッセージボックスを表示
'
' 概要　　　：はい・いいえ・キャンセルメッセージボックスを表示する
' 引数　　　：basePrompt 基本メッセージ
'             title      メッセージボックスのタイトル
'
' =========================================================
Public Function showMessageBoxForYesNoCancel(ByRef basePrompt As String _
                                , ByRef title As String) As Long
    
    showMessageBoxForYesNoCancel = MsgBox(basePrompt _
           , vbYesNoCancel + vbDefaultButton2 _
           , title)

End Function

' =========================================================
' ▽はい・いいえメッセージボックスを表示
'
' 概要　　　：はい・いいえメッセージボックスを表示する
' 引数　　　：basePrompt 基本メッセージ
'             title      メッセージボックスのタイトル
'
' =========================================================
Public Function showMessageBoxForYesNo(ByRef basePrompt As String _
                                , ByRef title As String) As Long
    
    showMessageBoxForYesNo = MsgBox(basePrompt _
           , vbYesNo + vbDefaultButton2 _
           , title)

End Function

' =========================================================
' ▽INIファイルパス取得
'
' 概要　　　：アプリケーションのINIファイルパスを取得する
' 引数　　　：fileName ファイル名
' 戻り値　　：INIファイルパス
'
' =========================================================
Public Function getApplicationIniFilePath(Optional ByVal fileName As String = "") As String

    ' iniファイルのパスを取得する
    getApplicationIniFilePath = ThisWorkbook.path & "\resource\config\" & fileName
    
End Function

' =========================================================
' ▽レジストリパス取得
'
' 概要　　　：アプリケーションのレジストリパスを取得する
' 　　　　　　ルートキーは、HKEY_CURRENT_USER
' 引数　　　：companyName 会社名
' 　　　　　　appName     アプリケーション名
' 　　　　　　suffix      レジストリパスの接尾辞
' 戻り値　　：INIファイルパス
'
' =========================================================
Public Function getApplicationRegistryPath(ByVal companyName As String _
                                         , Optional ByVal suffix As String = "" _
                                         , Optional ByVal appName As String = "") As String

    ' アプリケーション名が設定されていない場合
    If appName = "" Then
    
        ' プロジェクト名を設定する
        appName = ConstantsCommon.APPLICATION_NAME
    End If

    ' iniファイルのパスを取得する
    ' 本ブックのパス＋プロジェクト名＋".ini"
    getApplicationRegistryPath = "Software\" & companyName & "\" & appName
    
    If suffix <> "" Then
    
        getApplicationRegistryPath = getApplicationRegistryPath & "\" & suffix
    End If
    
End Function

' =========================================================
' ▽配列サイズ取得
'
' 概要　　　：配列のサイズを取得する
' 引数　　　：var       配列
' 　　　　　　dimension 次元
'
' =========================================================
Public Function arraySize(ByRef var As Variant, Optional ByVal dimension As Long = 1) As Long

    If IsArray(var) = True Then
    
        arraySize = UBound(var, dimension) - LBound(var, dimension) + 1
        
    Else
        arraySize = 0
    
    End If
    

End Function

' =========================================================
' ▽2次元配列の任意の行を1次元配列として返す
'
' 概要　　　：
' 引数　　　：val 配列
'             i   配列のインデックス
'
' =========================================================
Public Function convert2to1Array(ByRef val As Variant, ByVal i As Long) As Variant

    ' 戻り値
    Dim ret() As Variant

    Dim j As Long
    
    ReDim ret(LBound(val, 2) To UBound(val, 2))
    
    For j = LBound(ret) To UBound(ret)
    
        ret(j) = val(i, j)
    
    Next
    
    convert2to1Array = ret

End Function

' =========================================================
' ▽2次元配列をデバッグウィンドウに出力する
'
' 概要　　　：
' 引数　　　：val 配列
'
' =========================================================
Public Function debugPrintArray(ByRef val As Variant)

    ' 配列のインデックス
    Dim i As Long
    Dim j As Long
    
    ' デバッグウィンドウに出力する文字列
    Dim str As String
    
    str = "Output Array" & vbNewLine
    
    ' -------------------------------------------------
    ' 配列として初期化されている場合に出力を実施する
    ' -------------------------------------------------
    If VarType(val) = (vbArray + vbVariant) Then
    
        ' ループ処理
        For i = LBound(val, 1) To UBound(val, 1)
        
            str = str & "+   [" & i & "] - {"
        
            For j = LBound(val, 2) To UBound(val, 2)
            
                str = str & val(i, j) & ", "
            Next
            
            str = str & "}" & vbNewLine
            
        Next
        
    Else
        str = str & "   ... Empty"
        
    End If
    ' -------------------------------------------------
    
    Debug.Print str
    
End Function

' =========================================================
' ▽2次元配列の要素入れ替え
'
' 概要　　　：2次元配列の要素を(x,y)から(y,x)に設定しなおす。
' 引数　　　：v 2次元配列
'
' 戻り値　　：2次元配列
'
' =========================================================
Public Function transposeDim(ByRef v As Variant) As Variant
    
    Dim X As Long
    Dim Y As Long
    
    Dim Xlower As Long
    Dim Xupper As Long
    
    Dim Ylower As Long
    Dim Yupper As Long
    
    Dim tempArray As Variant
    
    Xlower = LBound(v, 2)
    Xupper = UBound(v, 2)
    Ylower = LBound(v, 1)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xlower To Xupper, Ylower To Yupper)
    
    For X = Xlower To Xupper
        For Y = Ylower To Yupper
        
            tempArray(X, Y) = v(Y, X)
        
        Next Y
    Next X
    
    transposeDim = tempArray

End Function

' =========================================================
' ▽整数チェック
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 整数
'
' =========================================================
Public Function validInteger(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False

    ' チェック対象が数値で且つ、小数点を含まない場合、OKとする
    If _
            IsNumeric(value) = True _
        And InStr(value, ".") = 0 Then
    
        ret = True
    
    End If

    ' 戻り値を返す
    validInteger = ret

End Function

' =========================================================
' ▽整数チェック（負数は含まない）
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 整数
'
' =========================================================
Public Function validUnsignedInteger(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False

    ' チェック対象が数値で且つ、マイナス記号を含まず小数点を含まない場合、OKとする
    If _
            IsNumeric(value) = True _
        And InStr(value, ".") = 0 _
        And InStr(value, "-") = 0 _
    Then
    
        ret = True
    
    End If

    ' 戻り値を返す
    validUnsignedInteger = ret

End Function

' =========================================================
' ▽16進数チェック
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 16進数
'
' =========================================================
Public Function validHex(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = True

    ' インデックス
    Dim i    As Long
    ' 文字のサイズ
    Dim size As Long
    
    ' 文字列の1文字分
    Dim one    As String
    ' 1文字分のASCIIコード
    Dim oneAsc As Long
    
    ' 文字のサイズを取得する
    size = Len(value)
    
    ' 文字列から1文字ずつ取り出しループを実行する
    For i = 1 To size
    
        ' 1文字取り出す
        one = Mid$(value, i, 1)
        ' 取り出した文字のASCIIコードを調べる
        oneAsc = Asc(one)
        
        ' 文字列が以下の範囲内であるかを確認する
        ' 0-9 a-f A-F
        If _
             (65 <= oneAsc And oneAsc <= 70) _
          Or (97 <= oneAsc And oneAsc <= 102) _
          Or (48 <= oneAsc And oneAsc <= 57) Then
        
            ' 正常
            
        Else
        
            ' エラー時
            ret = False
            Exit For
        
        End If
        
    Next

    ' 戻り値を返す
    validHex = ret

End Function

' =========================================================
' ▽数値であるかをチェックする
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 整数
'
' =========================================================
Public Function validNumeric(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False

    ' チェック対象が数値の場合、OKとする
    If _
            IsNumeric(value) = True Then
    
        ret = True
    
    End If

    ' 戻り値を返す
    validNumeric = ret

End Function

' =========================================================
' ▽数値であるかをチェックする（負数は含まない）
'
' 概要　　　：
' 引数　　　：value チェック文字列
' 戻り値　　：True 整数
'
' =========================================================
Public Function validUnsignedNumeric(ByVal value As String) As Boolean

    ' 戻り値
    Dim ret As Boolean: ret = False

    ' チェック対象が数値で且つマイナス記号を含まない場合、OKとする
    If _
            IsNumeric(value) = True _
        And InStr(value, "-") = 0 _
    Then
    
        ret = True
    
    End If

    ' 戻り値を返す
    validUnsignedNumeric = ret

End Function

' =========================================================
' ▽コード値チェック
'
' 概要　　　：引数で与えられたコードリストに一致するものがあるかをチェックする。
' 引数　　　：value    チェック文字列
' 　　　　　　codeList コードリスト
' 戻り値　　：True コードリストに一致する値がある
'
' =========================================================
Public Function validCode(ByVal value As String, ParamArray codeList() As Variant) As Boolean

    ' チェック対象が空の場合、OKとする
    Dim i As Long
    
    ' valueがenumsの何れかの値と一致しているかどうかを確認する
    For i = LBound(codeList) To UBound(codeList)
    
        ' 一致している場合
        If value = CStr(codeList(i)) Then
        
            ' Trueを返す
            validCode = True
            
            Exit Function
        End If
    
    Next
    
    ' 一致するものがなかったので、Falseを返す
    validCode = False

End Function

' =========================================================
' ▽RGB反転
'
' 概要　　　：RGBを反転させる。
' 引数　　　：r 赤
' 　　　　　　g 緑
' 　　　　　　b 青
' 戻り値　　：反転色
'
' =========================================================
Public Function reverseRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long

    reverseRGB = (Not RGB(r, g, b)) And &HFFFFFF

End Function

' =========================================================
' ▽NULL→空文字列変換
'
' 概要　　　：Nullを空文字列に変換する。
' 引数　　　：value VARIANTデータ
' 戻り値　　：空文字列
' 特記事項　：Null 値は、データ アイテム に有効なデータが
' 　　　　　　格納されていないことを示すのに使用されるバリアント型 (Variant) の内部処理形式です。
'
' =========================================================
Public Function convertNullToEmptyStr(ByRef value As Variant) As String

    ' NULLの場合
    If isNull(value) = True Then
    
        ' 空文字列に変換
        convertNullToEmptyStr = ""
        
    ' 配列の場合
    ElseIf IsArray(value) Then
    
        ' 空文字列に変換
        convertNullToEmptyStr = ""
        
    ' その他
    Else
    
        ' 文字列に変換して格納する
        convertNullToEmptyStr = CStr(value)
    End If
    
End Function

' =========================================================
' ▽クイックソート
'
' 概要　　　：クイックソートを行う。配列変数の要素はLong型を前提とする。
' 引数　　　：a 配列
' 戻り値　　：
'
' =========================================================
Public Sub quickSort(ByRef a As Variant)

    quickSortSub a, LBound(a), UBound(a)
    
End Sub

' =========================================================
' ▽クイックソート
'
' 概要　　　：クイックソートを行う。配列変数の要素はLong型を前提とする。
' 引数　　　：a     配列
' 　　　　　　left  左位置
' 　　　　　　right 右位置
' 戻り値　　：
'
' =========================================================
Private Sub quickSortSub(ByRef a As Variant _
                       , ByVal Left As Long _
                       , ByVal right As Long)

    ' スタックオブジェクト
    Dim stack As New ValStack
    
    ' スタックに格納する値
    ' （配列を走査する情報、左端と右端のインデックスを格納する）
    Dim stackVal As Variant
    ' 配列変数を生成する
    ReDim stackVal(1 To 2)
    
    
    ' ベースとなる値
    Dim base As Long
    ' 一時変数
    Dim temp As Long
    
    ' 中心のインデックス
    Dim center As Long
    
    Dim i      As Long
    Dim j      As Long

    ' スタックに最初に設定する変数を設定
    stackVal(1) = Left
    stackVal(2) = right
    ' スタックにプッシュする
    stack.push stackVal

    ' スタックの中身がなくなるまで実行
    Do While stack.count > 0
        
        ' スタックから値を取り出す
        stackVal = stack.pop
        
        ' 左端を取得
        Left = stackVal(1)
        ' 右端を取得
        right = stackVal(2)
        
        ' ここからクイックソートのアルゴリズム（教科書どおり）
        If Left < right Then
        
            center = Int((Left + right) / 2)
            
            base = a(center)
            
            i = Left
            j = right
            
            Do While i <= j
            
                ' ※ここが大小の比較部分
                Do While a(i) < base
                
                    i = i + 1
                Loop
            
                ' ※ここが大小の比較部分
                Do While a(j) > base
                
                    j = j - 1
                Loop
            
                If i <= j Then
                
                    temp = a(i)
                    a(i) = a(j)
                    a(j) = temp
                    
                    i = i + 1
                    j = j - 1
                End If
            
            Loop
            
            ' 再帰呼び出しではなく、スタックに情報を詰める
            ' 新たな走査部分をスタックに詰める
            
            ' 右部分の情報
            stackVal(1) = i
            stackVal(2) = right
            stack.push stackVal
            
            ' 左部分の情報
            stackVal(1) = Left
            stackVal(2) = j
            stack.push stackVal
            
        End If
    
    Loop
End Sub

' =========================================================
' ▽ファイル出力できるかを確認する
'
' 概要　　　：
' 引数　　　：folderPath ファイルパス
' 戻り値　　：True ファイル出力可能、False ファイル出力不可
'
' =========================================================
Public Function touch(ByVal folderPath As String) As Boolean

    On Error GoTo err

    ' 出力ファイルパス
    Dim touchedFilePath As String
    
    Dim i As Long
    Dim fw As FileWriter
    
    ' 重複したファイルが存在することも考慮してループを回してカウンタ変数をファイルパスの一部に使用する
    ' （おそらく100回以内には、ユニークなファイル名になるはず）
    For i = 0 To 100
    
        touchedFilePath = VBUtil.concatFilePath(folderPath, "sut_touch________" & (i + 1))
        
        If Not VBUtil.isExistFile(touchedFilePath) Then
        
            Set fw = New FileWriter
            fw.init touchedFilePath, "Shift_JIS", vbNewLine
            fw.writeText "touch"
            fw.destroy
            
            ' ここまで来たら正常にファイル出力されたとみなす
            Exit For
            
        End If
    Next
    
    ' 最後にファイルを削除する
    Kill touchedFilePath

    touch = True

    Exit Function
    
err:

    touch = False

End Function

' =========================================================
' ▽ファイルが存在するかをチェックする
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：True ファイルが存在する場合
'
' =========================================================
Public Function isExistFile(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        ' ファイルが存在する場合
        isExistFile = True
    Else
        ' ファイルが存在しない場合
        isExistFile = False
    End If

End Function

' =========================================================
' ▽ファイルが存在するかをチェックする
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：True ファイルが存在する場合
'
' =========================================================
Public Function isExistDirectory(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) Then
        ' ファイルが存在する場合
        isExistDirectory = True
    Else
        ' ファイルが存在しない場合
        isExistDirectory = False
    End If

End Function

' =========================================================
' ▽ファイルパスからディレクトリパスを抽出する
'   ディレクトリの場合、引数を返却
'       ファイルの場合、ディレクトリパスを抽出
'   何も存在しない場合、ディレクトリパスを抽出
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：ディレクトリパス
'
' =========================================================
Public Function extractDirPathFromFilePath(filePath As String) As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) Then
        ' ファイルではない（＝ディレクトリなどの）場合
        extractDirPathFromFilePath = filePath
        Exit Function
    End If
    
    ' 戻り値
    Dim ret As String
    
    ' ディレクトリ位置
    Dim dirPoint As Long

    ' 文字列の右端から"\"を検索し、左端からの位置を取得する
    dirPoint = InStrRev(filePath, "\")
    
    ' "\"が見つからない場合
    If dirPoint <> 0 Then
    
        ' ディレクトリパスの取得
        ret = Left$(filePath, dirPoint - 1)
        
        extractDirPathFromFilePath = ret
    
    Else
        extractDirPathFromFilePath = ""
    
    End If
    
End Function

' =========================================================
' ▽ディレクトリパスとファイルパスを連結する
'
' 概要　　　：
' 引数　　　：dir      ディレクトリパス
' 　　　　　　filePath ファイルパス
' 戻り値　　：連結後の文字列
'
' =========================================================
Public Function concatFilePath(ByVal dir As String, ByVal fileName As String) As String

    ' 文字列の最後尾に "\" が付いているかを確認する
    If InStrRev(dir, "\") = Len(dir) Then
    
        concatFilePath = dir & fileName
    Else
    
        concatFilePath = dir & "\" & fileName
    End If
    
End Function

' =========================================================
' ▽ディレクトリを作成する
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：True ディレクトリ作成時はTrueを返却
'
' =========================================================
Public Function createDir(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) = False And _
       fso.FileExists(filePath) = False Then
        fso.CreateFolder filePath
        createDir = True
    End If

    createDir = False
        
End Function

' =========================================================
' ▽ディレクトリを削除する
'
' 概要　　　：
' 引数　　　：filePath ファイルパス
' 戻り値　　：True ディレクトリ削除時はTrueを返却
'
' =========================================================
Public Function deleteDir(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) = True Then
        fso.DeleteFolder filePath
        deleteDir = True
    End If

    deleteDir = False
        
End Function

Public Function convertKeyCodeToKeyAscii(ByVal KeyCode As Long) As String

    If vbKey0 = KeyCode Then
        convertKeyCodeToKeyAscii = "0"
    ElseIf vbKey1 = KeyCode Then convertKeyCodeToKeyAscii = "1"
    ElseIf vbKey2 = KeyCode Then convertKeyCodeToKeyAscii = "2"
    ElseIf vbKey3 = KeyCode Then convertKeyCodeToKeyAscii = "3"
    ElseIf vbKey4 = KeyCode Then convertKeyCodeToKeyAscii = "4"
    ElseIf vbKey5 = KeyCode Then convertKeyCodeToKeyAscii = "5"
    ElseIf vbKey6 = KeyCode Then convertKeyCodeToKeyAscii = "6"
    ElseIf vbKey7 = KeyCode Then convertKeyCodeToKeyAscii = "7"
    ElseIf vbKey8 = KeyCode Then convertKeyCodeToKeyAscii = "8"
    ElseIf vbKey9 = KeyCode Then convertKeyCodeToKeyAscii = "9"
    ElseIf vbKeyA = KeyCode Then convertKeyCodeToKeyAscii = "A"
    ElseIf vbKeyB = KeyCode Then convertKeyCodeToKeyAscii = "B"
    ElseIf vbKeyC = KeyCode Then convertKeyCodeToKeyAscii = "C"
    ElseIf vbKeyD = KeyCode Then convertKeyCodeToKeyAscii = "D"
    ElseIf vbKeyE = KeyCode Then convertKeyCodeToKeyAscii = "E"
    ElseIf vbKeyF = KeyCode Then convertKeyCodeToKeyAscii = "F"
    ElseIf vbKeyG = KeyCode Then convertKeyCodeToKeyAscii = "G"
    ElseIf vbKeyH = KeyCode Then convertKeyCodeToKeyAscii = "H"
    ElseIf vbKeyI = KeyCode Then convertKeyCodeToKeyAscii = "I"
    ElseIf vbKeyJ = KeyCode Then convertKeyCodeToKeyAscii = "J"
    ElseIf vbKeyK = KeyCode Then convertKeyCodeToKeyAscii = "K"
    ElseIf vbKeyL = KeyCode Then convertKeyCodeToKeyAscii = "L"
    ElseIf vbKeyM = KeyCode Then convertKeyCodeToKeyAscii = "M"
    ElseIf vbKeyN = KeyCode Then convertKeyCodeToKeyAscii = "N"
    ElseIf vbKeyO = KeyCode Then convertKeyCodeToKeyAscii = "O"
    ElseIf vbKeyP = KeyCode Then convertKeyCodeToKeyAscii = "P"
    ElseIf vbKeyQ = KeyCode Then convertKeyCodeToKeyAscii = "Q"
    ElseIf vbKeyR = KeyCode Then convertKeyCodeToKeyAscii = "R"
    ElseIf vbKeyS = KeyCode Then convertKeyCodeToKeyAscii = "S"
    ElseIf vbKeyT = KeyCode Then convertKeyCodeToKeyAscii = "T"
    ElseIf vbKeyU = KeyCode Then convertKeyCodeToKeyAscii = "U"
    ElseIf vbKeyV = KeyCode Then convertKeyCodeToKeyAscii = "V"
    ElseIf vbKeyW = KeyCode Then convertKeyCodeToKeyAscii = "W"
    ElseIf vbKeyX = KeyCode Then convertKeyCodeToKeyAscii = "X"
    ElseIf vbKeyY = KeyCode Then convertKeyCodeToKeyAscii = "Y"
    ElseIf vbKeyZ = KeyCode Then convertKeyCodeToKeyAscii = "Z"
    End If

End Function

' =========================================================
' ▽ポイントからピクセルに単位を変換する
'
' 概要　　　：
' 引数　　　：d     DPI
' 　　　　　　pixel ピクセル
' 戻り値　　：ポイント
'
' =========================================================
Public Function convertPixelToPoint(ByVal d As Long, ByVal pixel As Long) As Single

    convertPixelToPoint = CSng(pixel) / d * 72

End Function

' =========================================================
' ▽ピクセルからポイントに単位を変換する
'
' 概要　　　：
' 引数　　　：d     DPI
' 　　　　　　pixel ピクセル
' 戻り値　　：ポイント
'
' =========================================================
Public Function convertPointToPixel(ByVal d As Long, ByVal Point As Single) As Long

    convertPointToPixel = Point * d / 72
    
End Function

' =========================================================
' ▽真偽値文字列を真偽データに変換する。
'
' 概要　　　：
' 引数　　　：str 文字列
' 戻り値　　：変換後の真偽データ
'
' =========================================================
Public Function convertBoolStrToBool(ByVal str As String) As Boolean

    If str = Empty Then
        ' 未入力時
        convertBoolStrToBool = False
    Else
        ' 入力時
    
        If LCase$(str) = "true" Then
            ' 真
            convertBoolStrToBool = True
        Else
            ' 偽
            convertBoolStrToBool = False
        End If
    
    End If
    
End Function

' =========================================================
' ▽中心座標を計算する
'
' 概要　　　：計算後の座標が、dx・dyに格納される
' 引数　　　：sx 基準となる矩形 座標X
' 　　　　　　sy 基準となる矩形 座標Y
' 　　　　　　sw 基準となる矩形 幅
' 　　　　　　sh 基準となる矩形 高さ
' 　　　　　　dx 比較する矩形 座標X
' 　　　　　　dy 比較する矩形 座標Y
' 　　　　　　dw 比較する矩形 幅
' 　　　　　　dh 比較する矩形 高さ
'
' =========================================================
Public Sub calcCenterPoint( _
                           ByVal sx As Single _
                         , ByVal sy As Single _
                         , ByVal sw As Single _
                         , ByVal sh As Single _
                         , ByRef dx As Single _
                         , ByRef dy As Single _
                         , ByVal dw As Single _
                         , ByVal dh As Single)

    ' 中心を計算する
    Dim newX As Single
    Dim newY As Single
    
    newX = sw / 2 - dw / 2 + sx
    newY = sh / 2 - dh / 2 + sy

    ' 中心を設定する
    dx = newX
    dy = newY

End Sub

' =========================================================
' ▽矩形AとBを比較しAがB内に収まっているかを確認する
'
' 概要　　　：
' 引数　　　：sx 基準となる矩形 座標X
' 　　　　　　sy 基準となる矩形 座標Y
' 　　　　　　sw 基準となる矩形 幅
' 　　　　　　sh 基準となる矩形 高さ
' 　　　　　　dx 比較する矩形 座標X
' 　　　　　　dy 比較する矩形 座標Y
' 　　　　　　dw 比較する矩形 幅
' 　　　　　　dh 比較する矩形 高さ
' 戻り値　　：True 矩形A内に収まっている場合
'
' =========================================================
Public Function isInnerScreen( _
                           ByVal sx As Single _
                         , ByVal sy As Single _
                         , ByVal sw As Single _
                         , ByVal sh As Single _
                         , ByRef dx As Single _
                         , ByRef dy As Single _
                         , ByRef dw As Single _
                         , ByRef dh As Single) As Boolean

    isInnerScreen = True

    ' 枠をはみ出していないかを確認する
    If sx > dx Then
    
        isInnerScreen = False
        
    ElseIf sy > dy Then
    
        isInnerScreen = False
        
    ElseIf (sx + sw) < (dx + dw) Then
    
        isInnerScreen = False
        
    ElseIf (sy + sh) < (dy + dh) Then
    
        isInnerScreen = False
        
    End If

End Function

' =========================================================
' ▽パディング関数
'
' 概要　　　：文字列の左側に特定の文字を任意の桁数になるように詰める
' 引数　　　：value  値
' 　　　　　　length 桁数
' 　　　　　　char   文字
' 戻り値　　：パディング結果
'
' =========================================================
Public Function padLeft(ByVal value As String _
                      , ByVal length As Long _
                      , Optional ByVal char As String = "0") As String

    ' パディングする桁数
    Dim padLen As Long
    padLen = length - Len(value)
    
    If padLen < 1 Then
    
        padLeft = value
        Exit Function
    End If

    padLeft = String(length - Len(value), char) & value

End Function

' =========================================================
' ▽パディング関数
'
' 概要　　　：文字列の右側に特定の文字を任意の桁数になるように詰める
' 引数　　　：value  値
' 　　　　　　length 桁数
' 　　　　　　char   文字
' 戻り値　　：パディング結果
'
' =========================================================
Public Function padRight(ByVal value As String _
                       , ByVal length As Long _
                       , Optional ByVal char As String = "0") As String

    ' パディングする桁数
    Dim padLen As Long
    padLen = length - Len(value)
    
    If padLen < 1 Then
    
        padRight = value
        Exit Function
    End If

    padRight = value & String(length - Len(value), char)

End Function

' =========================================================
' ▽エンコードリスト取得関数
'
' 概要　　　：エンコードリストを取得する
' 引数　　　：
' 戻り値　　：エンコードリスト
'
' =========================================================
Public Function getEncodeList() As ValCollection

    ' 対象とする文字コードリスト
    Dim includeChars As New ValCollection
    includeChars.setItem "Shift_JIS", "Shift_JIS"
    includeChars.setItem "EUC-JP", "EUC-JP"
    includeChars.setItem "UTF-8", "UTF-8"
    includeChars.setItem "UTF-8 (with bom)", "UTF-8 (with bom)"
    includeChars.setItem "UNICODE", "UNICODE"
    
    Set getEncodeList = includeChars
    
End Function


' =========================================================
' ▽エンコードリスト取得関数（レジストリから取得）
'
' 概要　　　：エンコードリストを取得する
' 引数　　　：
' 戻り値　　：エンコードリスト
'
' =========================================================
Public Function getEncodeListFromRegistry() As ValCollection

    ' 文字コードリスト取得用レジストリオブジェクト
    Dim regChar As New RegistryManipulator
    ' 文字コードリスト取得用のレジストリオブジェクトを初期化する
    regChar.init RegKeyConstants.HKEY_CLASS_ROOT _
               , REG_PATH_CHARACTER_CODE_LIST _
               , RegAccessConstants.KEY_READ _
               , False
               
    ' エイリアス確認用レジストリオブジェクト
    Dim regCharAlias As RegistryManipulator
    
    ' 文字コード一覧
    Dim charList As ValCollection
    ' 文字コードリストを取得する
    Set charList = regChar.getKeyList
    ' 文字コード一覧（エイリアスを除外）
    Dim charListRemovalAlias As New ValCollection

    ' 文字コード
    Dim char As Variant
    ' 文字コード エイリアス
    Dim charAlias As String
    
    For Each char In charList.col
    
        ' エイリアス確認用レジストリオブジェクト初期化
        Set regCharAlias = New RegistryManipulator
        
        regCharAlias.init RegKeyConstants.HKEY_CLASS_ROOT _
                        , REG_PATH_CHARACTER_CODE_LIST & "\" & char _
                        , RegAccessConstants.KEY_READ _
                        , False
                        
        ' 文字コードのエイリアスであるかを判定する
        If regCharAlias.getValue(REG_KEY_ALIAS_CHARSET, charAlias) = False Then
        
            ' エイリアスではない場合、追加する
            charListRemovalAlias.setItem char, char
        End If
    
        ' 破棄する
        Set regCharAlias = Nothing
    Next
    
    Set getEncodeList = charListRemovalAlias
    
End Function

' =========================================================
' ▽改行コードリスト取得関数
'
' 概要　　　：改行コードリスト取得
' 引数　　　：
' 戻り値　　：改行コードリスト
'
' =========================================================
Public Function getNewlineList() As ValCollection

    Set getNewlineList = New ValCollection
    
    getNewlineList.setItem NEW_LINE_STR_CRLF
    getNewlineList.setItem NEW_LINE_STR_CR
    getNewlineList.setItem NEW_LINE_STR_LF
    
End Function

' =========================================================
' ▽改行コード文字を実際の改行コード値に変換する関数
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：実際の改行コード値
'
' =========================================================
Public Function convertNewLineStrToNewLineCode(ByVal newLineStr As String) As String

    If newLineStr = NEW_LINE_STR_CRLF Then
    
        ' Windows
        convertNewLineStrToNewLineCode = vbCr & vbLf
    
    ElseIf newLineStr = NEW_LINE_STR_CR Then
    
        ' Mac
        convertNewLineStrToNewLineCode = vbCr
    
    ElseIf newLineStr = NEW_LINE_STR_LF Then
    
        ' Unix
        convertNewLineStrToNewLineCode = vbLf
        
    ' 当てはまらない場合
    Else
    
        ' Windows
        convertNewLineStrToNewLineCode = vbCr & vbLf
    
    End If

End Function

' =========================================================
' ▽正規表現の検索文字列のエスケープ処理関数
'
' 概要　　　：正規表現の検索文字列のエスケープ処理
' 引数　　　：keyword 検索文字列キーワード
' 戻り値　　：変換後の検索文字列
'
' =========================================================
Public Function escapeRegExpKeyword(ByVal keyword As String) As String

    escapeRegExpKeyword = keyword
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "\", "\\")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "*", "\*")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "+", "\+")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, ".", "\.")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "?", "\?")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "{", "\{")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "}", "\}")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "(", "\(")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, ")", "\)")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "[", "\[")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "]", "\]")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "^", "\^")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "$", "\$")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "-", "\-")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "|", "\|")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "/", "\/")

End Function

' =========================================================
' ▽ワイルドカードフィルタ処理関数
'
' 概要　　　：ワイルドカード文字を加味した検索を実施する。
'
'             filterKeywordに使用できる特殊文字。
'             * → 任意の文字の連続
'             ? → 任意の1文字
'
'             listOfElementPropNameには、ドットを使用することでネストしたプロパティを指定することも可能。
'             例）"table.tableName"のようにすると、まずtableプロパティを取得して、tableプロパティ内のtableNameを取得する
'
' 引数　　　：list                  フィルタ対象リスト
'     　　　  listOfElementPropName フィルタ対象リスト内のプロパティ名
'     　　　  filterKeyword         フィルタに使用するキーワード
' 戻り値　　：検索対象リストをフィルタした結果
'
' =========================================================
Public Function filterWildcard(ByVal list As ValCollection, _
                               ByVal listOfElementPropName As String, _
                               ByVal filterKeyword As String) As ValCollection

    ' ---------------------------------------
    Dim convertedFilterKeyword As String
    convertedFilterKeyword = filterKeyword
    
    ' ワイルドカード文字を"適当な制御コード"（画面で入力・表示できない文字列）に変換する
    ' 制御コードに変換する理由として、後続で正規表現文字列をエスケープする処理があり、そちらでエスケープが実施されないようにするため
    ' ※極端な話、制御コードなら何でもよい（以下を使用）
    ' 装置制御DC1 = Char(17)
    ' 装置制御DC2 = Char(18)
    ' 装置制御DC3 = Char(19)
    ' 装置制御DC4 = Char(20)
    convertedFilterKeyword = replace(convertedFilterKeyword, "~*", Chr(19)) ' チルダ付きなので通常の文字扱いにする
    convertedFilterKeyword = replace(convertedFilterKeyword, "~?", Chr(20)) ' チルダ付きなので通常の文字扱いにする
    convertedFilterKeyword = replace(convertedFilterKeyword, "*", Chr(17))
    convertedFilterKeyword = replace(convertedFilterKeyword, "?", Chr(18))
    
    ' チルダ付きの特殊文字なので通常の文字にしたいので元の値に戻して、後続処理のVBUtil.escapeRegExpKeywordでエスケープして通常の文字として解釈されるようにする
    convertedFilterKeyword = replace(convertedFilterKeyword, Chr(19), "*")
    convertedFilterKeyword = replace(convertedFilterKeyword, Chr(20), "?")
    
    ' キーワードに含まれる正規表現の特殊文字をエスケープする
    convertedFilterKeyword = VBUtil.escapeRegExpKeyword(convertedFilterKeyword)
    
    ' ワイルドカードに対応した制御コードを正規表現に変換する
    convertedFilterKeyword = replace(convertedFilterKeyword, Chr(17), ".*") ' *は任意の文字の0個以上の連続
    convertedFilterKeyword = replace(convertedFilterKeyword, Chr(18), ".") ' ?は任意の1文字
    ' ---------------------------------------
    
    Set filterWildcard = filterRegExp(list, listOfElementPropName, convertedFilterKeyword)

End Function

' =========================================================
' ▽正規表現フィルタ処理関数
'
' 概要　　　：正規表現文字を加味した検索を実施する。
'
'             filterKeywordに使用できる特殊文字はRegExpに準ずる。
'
'             listOfElementPropNameには、ドットを使用することでネストしたプロパティを指定することも可能。
'             例）"table.tableName"のようにすると、まずtableプロパティを取得して、tableプロパティ内のtableNameを取得する
'
' 引数　　　：list                  フィルタ対象リスト
'     　　　  listOfElementPropName フィルタ対象リスト内のプロパティ名
'     　　　  filterKeyword         フィルタに使用するキーワード
' 戻り値　　：検索対象リストをフィルタした結果
'
' =========================================================
Public Function filterRegExp(ByVal list As ValCollection, _
                             ByVal listOfElementPropName As String, _
                             ByVal filterKeyword As String) As ValCollection

    ' 戻り値
    Set filterRegExp = New ValCollection
    
    ' 正規表現オブジェクトを生成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        ' 検索対象文字列
        .Pattern = "^" & filterKeyword & "$"
        ' 大文字小文字無視フラグ
        .IgnoreCase = True
        ' 文字列全体を繰り返し検索するフラグ
        .Global = False
    End With

    ' list変数の要素
    Dim rec As Variant
    ' フィルタ対象
    Dim searchExpression As Variant
    ' リストの要素のプロパティ名の配列
    Dim listOfElementPropNameArray As Variant
    listOfElementPropNameArray = Split(listOfElementPropName, ".")
    
    Dim i As Long

    For Each rec In list.col
    
        Set searchExpression = rec
        For i = LBound(listOfElementPropNameArray) To UBound(listOfElementPropNameArray)
            If i = UBound(listOfElementPropNameArray) Then
                ' 末尾の場合は、プリミティブなデータ
                If Not searchExpression Is Nothing Then
                    searchExpression = CallByName(searchExpression, listOfElementPropNameArray(i), VbGet)
                End If
            Else
                ' 末尾ではない場合は、オブジェクト型
                If Not searchExpression Is Nothing Then
                    Set searchExpression = CallByName(searchExpression, listOfElementPropNameArray(i), VbGet)
                End If
            End If
        Next
    
        If Not IsObject(searchExpression) Then
            ' オブジェクト型ではない（検索なので文字列型を想定）
            If reg.test(CStr(searchExpression)) Then
                filterRegExp.setItem rec
            End If
        End If
    Next

End Function

' =========================================================
' ▽ユーザーフォームをリロードする
'
' 概要　　　：ユーザーフォームは起動時のアクティブなブックを保持する性質がある。
'             ブックが切り替わった場合は一度フォームをアンロードする必要があるため、それらの判定を実施して適切にリロードする。
' 引数　　　：obj ユーザーフォーム
' 戻り値　　：
'
' 特記事項　：
'
' =========================================================
Public Function unloadFormIfChangeActiveBook(ByRef obj As Variant) As Boolean

    If obj Is Nothing Then
    
        unloadFormIfChangeActiveBook = False
        Exit Function
    
    End If
    
    Dim book As Workbook
    Set book = CallByName(obj, "getTargetBook", VbMethod)
    
    If ActiveWorkbook Is book Then
    
        unloadFormIfChangeActiveBook = False
        Exit Function
    
    End If
    
    unloadFormIfChangeActiveBook = True
        
End Function

