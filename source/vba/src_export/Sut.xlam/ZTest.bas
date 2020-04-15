Attribute VB_Name = "ZTest"
Option Explicit

#If DEBUG_MODE = 1 Then

#If VBA7 And Win64 Then
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Private Sub assert(ByVal test As Boolean)

    Debug.Assert test

    If Not test Then
        err.Raise vbObjectError + 513, "Assert", "Assert Error"
    End If

End Sub

Private Sub taestByValByRef()

    Dim timeBegin As Long
    Dim timeEnd   As Long

    Dim obj As Object
    Set obj = SutWorkbook

    Dim str As String
    str = "あいうえおかきくけこあいうえおかきくけこあいうえおかきくけこあいうえおかきくけこあいうえおかきくけこあいうえおかきくけこあいうえおかきくけこあいうえおかきくけこあいうえおかきくけこあいうえおかきくけこ"
    
    Dim i As Long
    Dim count As Long: count = 10000000

    timeBegin = GetTickCount
    For i = 1 To count
        taestByValForObjByRef obj
    Next
    timeEnd = GetTickCount
    Debug.Print "Obj参照渡し：" & (timeEnd - timeBegin) & "ミリ秒"
    
    timeBegin = GetTickCount
    For i = 1 To count
        taestByValForObjByVal obj
    Next
    timeEnd = GetTickCount
    Debug.Print "Obj　値渡し：" & (timeEnd - timeBegin) & "ミリ秒"

    timeBegin = GetTickCount
    For i = 1 To count
        taestByValForStrByRef str
    Next
    timeEnd = GetTickCount
    Debug.Print "Str参照渡し：" & (timeEnd - timeBegin) & "ミリ秒"
    
    timeBegin = GetTickCount
    For i = 1 To count
        taestByValForStrByVal str
    Next
    timeEnd = GetTickCount
    Debug.Print "Str　値渡し：" & (timeEnd - timeBegin) & "ミリ秒"

End Sub

Private Sub taestByValForObjByVal(ByVal a As Object)
End Sub

Private Sub taestByValForObjByRef(ByRef a As Object)
End Sub

Private Sub taestByValForStrByVal(ByVal a As String)
End Sub

Private Sub taestByValForStrByRef(ByRef a As String)
End Sub

Private Sub testAll()

    testIniFile
    testIniFilePerform
    testIniWorksheet
    testIniWorksheetPerform
    
End Sub

Private Sub taest3()

    Dim var(1 To 3) As Variant
    
    var(1) = "aaaa"
    var(2) = 1234
    var(3) = Now
    
    Debug.Print VBUtil.convertNullToEmptyStr(var)
End Sub

Private Sub test4_1()

    Dim a As New ValStack
    
    a.push "aaaa"
    Debug.Print a.pop
    Debug.Print a.pop
    
    a.push "aaaa"
    a.push "bbbb"
    
    Debug.Print a.pop
    Debug.Print a.pop

End Sub

Private Sub test4()

    Dim var(1 To 1048576) As Variant
    
    Randomize    ' 乱数発生ルーチンを初期化します。
    
    Dim i As Long
    
    For i = LBound(var) To UBound(var)
    
        var(i) = Int((10000 * Rnd) + 1)
    Next
    
'    Debug.Print "[before]"
'    For i = LBound(var) To UBound(var)
'
'        Debug.Print i & " " & var(i)
'    Next
    
    quickSort var

'    Debug.Print "[after]"
'    For i = LBound(var) To UBound(var)
'
'        Debug.Print i & " " & var(i)
'    Next

    MsgBox "完了"
End Sub

Private Sub test6()

    Dim a As RegistryManipulator
    Dim b As ValCollection
    
    Set a = New RegistryManipulator: a.init WinAPI_ADVAP.HKEY_CLASS_ROOT, "\MIME\Database\Charset", WinAPI_ADVAP.KEY_READ, False

    Set b = a.getKeyList

    Dim tmp As Variant

    For Each tmp In b.col

        Debug.Print tmp
    Next
    
    Set a = New RegistryManipulator: a.init WinAPI_ADVAP.HKEY_CURRENT_USER, "Software\ison\Sut", WinAPI_ADVAP.KEY_ALL_ACCESS, True

    Dim value As String
    Set b = a.getKeyList
    a.setValue "key1", "あいうえおかきくけこ"
    Debug.Print a.GetValue("key1", value)
    Debug.Print a.GetValue("key2", value)
    
    a.deleteValue "key1"
    
    Set a = New RegistryManipulator: a.init WinAPI_ADVAP.HKEY_CURRENT_USER, "Software\ison", WinAPI_ADVAP.KEY_ALL_ACCESS, True
    a.delete "Sut"
    
End Sub

Private Sub test7()

    frmFileOutput.ShowExt vbModal, "ヘッダです。", "file.sql"

End Sub

Private Sub test8()

    Dim file As New FileWriter: file.init "test.txt", "euc-jp", vbCrLf, True
    
    file.writeText "あいうえお", True
    file.writeText "かきくけこ", True
    file.writeText "さしすせそ", True
    
    file.writeText "たちつてと", True
    
End Sub

Private Sub outProperties(ByRef properties As Object)

    Dim i   As Long
    Dim cnt As Long: cnt = properties.count
    
    Dim propertie As Object
    
    For Each propertie In properties
    
        Debug.Print "[" & i & "]"
        Debug.Print "  Attributes : " & propertie.Attributes
        Debug.Print "  Name       : " & propertie.name
        Debug.Print "  Type       : " & propertie.Type
        Debug.Print "  Value      : " & propertie.value
    
    Next

End Sub

Private Sub outFontInfo()

    ' ループインデックス
    Dim i As Long
    
    ' コンボボックス
    Dim c As CommandBarComboBox
    
    ' フォントリストを取得する
    Set c = Application.CommandBars.FindControl(Id:=1728)
    
    Debug.Print c.BuiltIn
    
    ' リストの内容を全て表示する
    For i = 1 To c.ListCount
    
        ' リストの文字列を出力
        Debug.Print c.list(i)
    Next
    
    ' フォントサイズリストを取得する
    Set c = Application.CommandBars.FindControl(Id:=10000)
    
    Debug.Print c.BuiltIn
    
    ' リストの内容を全て表示する
    For i = 1 To c.ListCount
    
        ' リストの文字列を出力
        Debug.Print c.list(i)
    Next
End Sub

Private Sub test9()

    Dim r As Range
    
    Set r = Workbooks("Book2").Worksheets(1).Range("A1:B3")

    ExcelUtil.changeColWidth r, 100
    ExcelUtil.changeRowHeight r, 100
End Sub

Private Sub test10(ByRef a As String)

    a = "new"
End Sub

Private Sub test11(ByVal a As String)

    a = "new"
End Sub

Private Sub test12()

    Dim aa As String
    
    aa = "original"

    Debug.Print "before : " & aa
    test10 aa
    Debug.Print "after : " & aa

    aa = "original"

    Debug.Print "before : " & aa
    test11 aa
    Debug.Print "after : " & aa

End Sub

Private Sub test13()

    Dim a As ValCollection
    
    Set a = WinAPI_GDI.getFontNameList

End Sub

Private Sub test20()

    Dim aaa As New ValApplicationSettingShortcut
    aaa.init
    
    Dim temp As New ValCollection
    temp.setItem "final_clash"
    temp.setItem "final_atack"

    Set aaa.rclickMenuItemList = temp
    
    aaa.readForDataRClick
    aaa.writeForDataRClick
End Sub


Private Sub test21()

    ' エクセルのバージョン
    Dim excelVer As ExcelVersion: excelVer = ExcelUtil.getExcelVersion
    
    ' コマンドバー
    Dim cb   As CommandBar
    
    On Error Resume Next
    
    Set cb = Application.CommandBars.item("TEstmagicgendesu")
    
    If cb Is Nothing Then
    
        Set cb = Application.CommandBars.Add( _
                                name:="TEstmagicgendesu" _
                              , Temporary:=True _
                              , position:=msoBarPopup)
    End If
    
    On Error GoTo 0
        
    ' DB接続ボタン
    Dim btnDBConnect              As CommandBarButton
    ' DB切断ボタン
    Dim btnDBDisConnect           As CommandBarButton
    
    ' DB接続ボタンをコマンドバーにボタンを追加する
    Set btnDBConnect = cb.Controls.Add(Type:=msoControlButton)
    
    ' DB接続ボタンのプロパティを設定する
    With btnDBConnect
    
        .Style = msoButtonIconAndCaption
        .Caption = "接続"
        .DescriptionText = "DB接続"
        .OnAction = "Main.SutConnectDB"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutConnectDB"
        
    End With
        
    ' DB切断ボタンをコマンドバーにボタンを追加する
    Set btnDBDisConnect = cb.Controls.Add(Type:=msoControlButton)
    
    ' DB切断ボタンのプロパティを設定する
    With btnDBDisConnect
    
        .Style = msoButtonIconAndCaption
        .Caption = "切断"
        .DescriptionText = "DB切断"
        .OnAction = "Main.SutDisconnectDB"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutDisconnectDB"
        
    End With
    
    ' ***************************************************************

    cb.ShowPopup
    
End Sub

Private Function test22()

    Debug.Print VBUtil.getAppOnKeyCodeByName("Ctrl")
    Debug.Print VBUtil.getAppOnKeyCodeByName("Shift")
    Debug.Print VBUtil.getAppOnKeyCodeByName("Alt")
    Debug.Print VBUtil.getAppOnKeyCodeByName("Home")
    Debug.Print VBUtil.getAppOnKeyCodeByName("hogehoge")

    Debug.Print VBUtil.getAppOnKeyNameByCode("^")
    Debug.Print VBUtil.getAppOnKeyNameByCode("+")
    Debug.Print VBUtil.getAppOnKeyNameByCode("%")
    Debug.Print VBUtil.getAppOnKeyNameByCode("{HOME}")
    Debug.Print VBUtil.getAppOnKeyNameByCode("hogehoge")

    Dim a As Boolean
    Dim b As Boolean
    Dim c As Boolean
    
    Dim k As String
    
    VBUtil.resolveAppOnKey "^+%{HOME}", a, b, c, k
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print k
    
    Debug.Print VBUtil.getAppOnKeyCodeBySomeParams(a, b, c, k)


    VBUtil.resolveAppOnKey "^{HOME}", a, b, c, k
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print k
    
    Debug.Print VBUtil.getAppOnKeyCodeBySomeParams(a, b, c, k)

    VBUtil.resolveAppOnKey "+{HOME}", a, b, c, k
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print k
    
    Debug.Print VBUtil.getAppOnKeyCodeBySomeParams(a, b, c, k)

    VBUtil.resolveAppOnKey "%{HOME}", a, b, c, k
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print k
    
    Debug.Print VBUtil.getAppOnKeyCodeBySomeParams(a, b, c, k)

    VBUtil.resolveAppOnKey "%", a, b, c, k
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print k
    
    Debug.Print VBUtil.getAppOnKeyCodeBySomeParams(a, b, c, k)

    VBUtil.resolveAppOnKey "{HOME}", a, b, c, k
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print k
    
    Debug.Print VBUtil.getAppOnKeyCodeBySomeParams(a, b, c, k)

    VBUtil.resolveAppOnKey "%afasfa", a, b, c, k
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print k
    
    Debug.Print VBUtil.getAppOnKeyCodeBySomeParams(a, b, c, k)

    VBUtil.resolveAppOnKey "sdafafasfa", a, b, c, k
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print k
    
    Debug.Print VBUtil.getAppOnKeyCodeBySomeParams(a, b, c, k)

End Function

Private Sub test333()

    Dim a As String
    Dim b As String

    a = "aaaaaaaa "
    b = "cccccccc "
    
    a = b
    
    Debug.Print a
    Debug.Print b

    a = "買えちゃうよ"
    
    Debug.Print a
    Debug.Print b
End Sub

Private Sub test444()

    Dim a As ScreenSizePt
    
    a = WinAPI_User.getScreenSizePt


End Sub

Private Sub test999999()

    Dim i As Long
    
    For i = 0 To 100000
    
        Debug.Print i
    Next

End Sub

Private Sub testValAppSettingColFormatR()


    Dim a As New ValApplicationSettingColFormat

    a.init ActiveWorkbook


End Sub

Private Sub testExeDataTypeReader()

    Dim ct As ValCollection

    Dim impl As IDbColumnType
    Dim fac As New DbObjectFactory
    Set impl = fac.createColumnType(DbmsType.Oracle): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.MySQL): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.PostgreSQL): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.Symfoware): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.MicrosoftAccess): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.MicrosoftSqlServer): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.Other): Set ct = impl.getDefaultColumnFormat
    
'
'    Dim b As New ExeDataTypeReader
'    Dim c As Variant
'
'    Set b.sheet = Worksheets("data_type_mysql")
'
'    c = b.execute
'
End Sub

Private Sub testExeDataTypeReader2()

    Main.getApplicationSettingColFormat ActiveWorkbook

End Sub

Private Sub s()
    
End Sub

Public Sub pallet()

    Application.CommandBars("Fill Color").visible = True
    Application.Dialogs.item(xlDialogColorPalette).Show
End Sub

Public Sub StringBuilderTest()

    Dim str As StringBuilder
    Set str = New StringBuilder
    
    assert str.length = 0
    
    str.clear
    str.append "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf
    assert str.length = 118
    assert str.str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf
    
    str.clear
    str.append "a"
    assert str.length = 1
    assert str.str = "a"
    
    str.clear
    assert str.length = 0
    assert str.str = ""
    
    str.append "ab"
    
    ' 引数不正
    str.remove 0, 0
    str.remove 1, 0
    str.remove 0, 1
    
    assert str.length = 2
    assert str.str = "ab"
    
    str.append "あいうえお"
    assert str.length = 7
    assert str.str = "abあいうえお"
    
    str.append "かきくけこ"
    assert str.length = 12
    assert str.str = "abあいうえおかきくけこ"
    
    str.remove 1, 5
    assert str.length = 7
    assert str.str = "えおかきくけこ"
    
    str.remove 1, 5
    assert str.length = 2
    assert str.str = "けこ"
    
    str.append "あいうえお"
    str.append "かきくけこ"
    str.remove 1, 11
    assert str.length = 1
    assert str.str = "こ"
    
    str.clear
    str.append "あいうえお"
    
    ' 引数不正
    str.insert 0, "かきくけこ"
    assert str.length = 5
    assert str.str = "あいうえお"
    
    ' 引数不正
    str.insert 7, "かきくけこ"
    assert str.length = 5
    assert str.str = "あいうえお"
    
    str.clear
    str.insert 1, "かきくけこ"
    assert str.length = 5
    assert str.str = "かきくけこ"
    
    str.clear
    str.append "あいうえお　かきくけこ"
    str.replace "いう", "置換REP"
    assert str.length = 14
    assert str.str = "あ置換REPえお　かきくけこ"
    
    str.clear
    str.append "あいうえお　かきくけこ"
    str.replace "いうお", "置換REP"
    assert str.length = 11
    assert str.str = "あいうえお　かきくけこ"
    
    str.clear
    str.append "あいうえお　かきくけこ"
    str.replace "あ", "ア"
    assert str.length = 11
    assert str.str = "アいうえお　かきくけこ"
    
    str.clear
    str.append "あいうえお　かきくけこ"
    str.replace "こ", "コ"
    assert str.length = 11
    assert str.str = "あいうえお　かきくけコ"
    
    str.clear
    str.append "あいうえお　かきくけこあ"
    str.replace "あ", "亜"
    assert str.length = 12
    assert str.str = "亜いうえお　かきくけこ亜"
    
    str.clear
    str.append String(1000, "あ")
    assert str.length = 1000
    assert str.str = String(1000, "あ")
    
    str.append String(1000, "あ")
    assert str.length = 2000
    assert str.str = String(2000, "あ")
    
    str.append String(1000, "あ")
    assert str.length = 3000
    assert str.str = String(3000, "あ")
    
    str.append String(1000, "あ")
    assert str.length = 4000
    assert str.str = String(4000, "あ")
    
    str.append String(1000, "あ")
    assert str.length = 5000
    assert str.str = String(5000, "あ")
  
End Sub

Private Sub testCsvParser()

    Dim ret  As ValCollection
    
    Dim var  As Variant
    Dim var2 As Variant
    
    Dim varBuff As New StringBuilder
    
    Dim testStr As String
    
    Dim csvp As New CsvParser: csvp.init
    
    ' ------------------------------------------------------
    testStr = "a,b,c,d"
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret) = testStr
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    testStr = "a,b,c,d" & vbNewLine & "e,f,g,h"
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret) = testStr
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    testStr = "a,b,c,d" & vbCr & "e,f,g,h"
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, ",", vbCr) = testStr
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    testStr = "a,b,c,d" & vbLf & "e,f,g,h"
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, ",", vbLf) = testStr
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    testStr = """a"",b,c,d,""e,e"",""""""e,e"""""""
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret) = "a,b,c,d,e,e,""e,e"""
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    testStr = """a"",b,c,d,""e,e"",""""""e,e""""""" & vbNewLine & "あいうえお,かきくけこ,さしすせそ" & vbNewLine & "あ""いうえお,か""きくけ""こ,さ""しすせ""そ"
    Set ret = csvp.parse(testStr)

    assert testCsvParserToString(ret) = "a,b,c,d,e,e,""e,e""" & vbNewLine & "あいうえお,かきくけこ,さしすせそ" & vbNewLine & "あ""いうえお,か""きくけ""こ,さ""しすせ""そ"
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    testStr = """param 1"","""""
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret) = "param 1,"
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    testStr = """"","""""
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret) = ","
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    csvp.init vbTab
    
    testStr = "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & "orsys" & vbTab
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, vbTab) = testStr
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    csvp.init vbTab
    
    testStr = "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, vbTab) = testStr
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    csvp.init vbTab
    
    testStr = "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab & vbNewLine & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab & vbLf & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab & vbCr & _
              vbCr & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, vbTab, vbLf) = replace(replace(testStr, vbNewLine, vbLf), vbCr, vbLf)
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    csvp.init vbTab
    
    testStr = "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & "あいうえお" & vbTab & vbNewLine & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab & vbLf & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab & vbCr & _
              vbCr & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, vbTab, vbLf) = replace(replace(testStr, vbNewLine, vbLf), vbCr, vbLf)
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    csvp.init vbTab
    
    testStr = "abcあいうえお" & vbTab & """あい" & vbTab & vbNewLine & "うえお""" & vbNewLine & _
              "abcあいうえお" & vbTab & """あい" & vbTab & "うえお"""
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, vbTab, vbNewLine) = "abcあいうえお" & vbTab & "あい" & vbTab & vbNewLine & "うえお" & vbNewLine & _
                                                          "abcあいうえお" & vbTab & "あい" & vbTab & "うえお"
    ' ------------------------------------------------------

End Sub

Private Function testCsvParserToString(ByVal list As ValCollection, Optional ByVal s As String = ",", Optional ByVal n As String = vbNewLine) As String

    Dim buff       As New StringBuilder
    
    Dim recOfList  As ValCollection
    Dim rec        As Variant
    
    Dim i As Long: i = 0
    
    For Each recOfList In list.col
        
        If i > 0 Then
            buff.append n
        End If
    
        For Each rec In recOfList.col
            buff.append rec & s
        Next
        
        If recOfList.count > 0 Then
        
            buff.remove buff.length, 1
        End If
        
        i = i + 1
    Next
    
    testCsvParserToString = buff.str

End Function

Public Sub testValCollection()

    Dim a As New ValCollection
    a.setItem "あいうえお"
    a.setItem "かきくけこ"
    a.setItemByIndexAfter "さしすせそ", 1
    a.setItemByIndexAfter "たちつてと", 1
    a.setItemByIndexBefore "なにぬねの", 1
    a.setItemByIndexBefore "はひふへほ", 1
    
    Debug.Print a.getItemByIndex(1, vbVariant)
    Debug.Print a.getItemByIndex(2, vbVariant)
    Debug.Print a.getItemByIndex(3, vbVariant)
    Debug.Print a.getItemByIndex(4, vbVariant)
    Debug.Print a.getItemByIndex(5, vbVariant)
    Debug.Print a.getItemByIndex(6, vbVariant)

End Sub

Public Sub testFilterWildcard()

    Dim ret As ValCollection

    Dim tableList As New ValCollection
    Dim tableWorksheet As ValTableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "シートA"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "シートB"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "シートC"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|,<.>/?_;+:*]}@`[{"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "*シートE?"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    Set tableWorksheet.table = New ValDbDefineTable
    tableWorksheet.table.tableName = "シートZ"
    tableList.setItem tableWorksheet

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "シートA")
    assert ret.count = 1

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "シートD")
    assert ret.count = 0

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "シート*")
    assert ret.count = 3

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "シ?ト")
    assert ret.count = 0

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "シ*ト?")
    assert ret.count = 3

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "シー*")
    assert ret.count = 3

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "*ート*")
    assert ret.count = 4

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "*A")
    assert ret.count = 1

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|,<.>/?_;+:*]}@`[{")
    assert ret.count = 1

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "~*シートE~?")
    assert ret.count = 1

    Set ret = VBUtil.filterWildcard(tableList, "table.tableName", "シートZ")
    assert ret.count = 1
    
    Debug.Print "testFilterWildcard complete"

End Sub

' =========================================================
' ▽ExcelCursorWaitクラスのテスト
'
' 概要　　　：
'
' =========================================================
Public Sub testExcelLongTimeProcessing()

    Dim tmpCalculation As XlCalculation
    tmpCalculation = Application.calculation

    Dim e As ExcelLongTimeProcessing
    Set e = New ExcelLongTimeProcessing
    
    ' ---------------------------------------------
    ' 初期化～破棄 destroy
    e.init True, True, True, True, True, True, True
    assert Application.calculation = xlCalculationManual
    assert Application.displayAlerts = False
    'assert Application.enableCancelKey = xlDisabled ' Debug時はプロパティの設定が有効にならないのでassertしない
    assert Application.enableEvents = False
    assert Application.cursor = xlWait
    assert Application.screenUpdating = False
    'assert Application.interactive = False ' Debug時はプロパティの設定が有効にならないのでassertしない
    
    e.destroy
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    ' ---------------------------------------------
        
    ' ---------------------------------------------
    ' 初期化～破棄 nothing
    e.init True, True, True, True, True, True, True
    assert Application.calculation = xlCalculationManual
    assert Application.displayAlerts = False
    'assert Application.enableCancelKey = xlDisabled ' Debug時はプロパティの設定が有効にならないのでassertしない
    assert Application.enableEvents = False
    assert Application.cursor = xlWait
    assert Application.screenUpdating = False
    'assert Application.interactive = False ' Debug時はプロパティの設定が有効にならないのでassertしない
    
    Set e = Nothing
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    ' ---------------------------------------------

    Set e = New ExcelLongTimeProcessing
         
    ' ---------------------------------------------
    ' 初期化～破棄 displayAlerts のみ有効
    e.init True, False, False, False, False, False, False
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = False
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    
    e.destroy
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    ' ---------------------------------------------
         
    ' ---------------------------------------------
    ' 初期化～破棄 enableEvents のみ有効
    e.init False, False, True, False, False, False, False
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = False
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    
    e.destroy
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    ' ---------------------------------------------

    ' ---------------------------------------------
    ' 初期化～破棄 cursor のみ有効
    e.init False, False, False, True, False, False, False
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlWait
    assert Application.screenUpdating = True
    assert Application.interactive = True
    
    e.destroy
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    ' ---------------------------------------------

    ' ---------------------------------------------
    ' 初期化～破棄 screenUpdating のみ有効
    e.init False, False, False, False, True, False, False
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = False
    assert Application.interactive = True
    
    e.destroy
    assert Application.calculation = tmpCalculation
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    ' ---------------------------------------------

    ' ---------------------------------------------
    ' 初期化～破棄 calculation のみ有効
    Application.calculation = xlCalculationSemiautomatic
    e.init False, False, False, False, False, True, False
    assert Application.calculation = xlCalculationManual
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    
    e.destroy
    assert Application.calculation = xlCalculationSemiautomatic
    assert Application.displayAlerts = True
    assert Application.enableCancelKey = xlInterrupt
    assert Application.enableEvents = True
    assert Application.cursor = xlDefault
    assert Application.screenUpdating = True
    assert Application.interactive = True
    ' ---------------------------------------------

    assert Application.calculation = tmpCalculation
End Sub

' =========================================================
' ▽ExcelCursorWaitクラスのテスト
'
' 概要　　　：
'
' =========================================================
Public Sub testExcelCursorWait()

    Application.cursor = xlDefault

    Dim w As ExcelCursorWait
    Set w = New ExcelCursorWait
    
    ' ---------------------------------------------
    ' 初期化～破棄 destroy
    w.init
    assert Application.cursor = xlWait
    
    w.destroy
    assert Application.cursor = xlDefault
    ' ---------------------------------------------
    
    ' ---------------------------------------------
    ' 初期化～破棄 forceRestore
    w.init
    assert Application.cursor = xlWait
    
    w.forceRestore
    assert Application.cursor = xlDefault
    ' ---------------------------------------------
    
    ' ---------------------------------------------
    ' 初期化～破棄 nothing
    w.init
    assert Application.cursor = xlWait
    
    Set w = Nothing
    assert Application.cursor = xlDefault
    ' ---------------------------------------------
    
    ' ---------------------------------------------
    ' 初期化～破棄 終了後も継続して待機状態にするのでxlWaitのまま
    Set w = New ExcelCursorWait
    w.init True
    assert Application.cursor = xlWait
    
    Set w = Nothing
    assert Application.cursor = xlWait
    ' ---------------------------------------------

    Application.cursor = xlDefault

End Sub

' =========================================================
' ▽Error発生時のテスト
'
' 概要　　　：
'
' =========================================================
Public Sub testErrorRaise()

    On Error GoTo err
    
    ' 何か適当なオブジェクトを初期化しておく
    Dim obj As ExcelLongTimeProcessing
    Set obj = New ExcelLongTimeProcessing
    obj.init

    err.Raise 1000, "my source", "my description"

    Exit Sub

err:

    ' エラー情報を退避する
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    ' 何か適当なオブジェクトを破棄する（デストラクタ内でErrに関する操作が行われること）
    Set obj = Nothing

    assert err.Number = 0
    assert err.Source = ""
    assert err.Description = ""
    
    ' 退避したエラー情報を設定しなおす
    VBUtil.setErr errT

    assert err.Number = 1000
    assert err.Source = "my source"
    assert err.Description = "my description"
    
    On Error GoTo 0

    assert err.Number = 0
    assert err.Source = ""
    assert err.Description = ""

End Sub

' =========================================================
' ▽Iniファイル操作のテスト
'
' 概要　　　：
'
' =========================================================
Public Sub testIniFile()

    Dim retValue As String
    Dim retValueArray As ValCollection
    
    Dim values As New ValCollection
    

    Dim i As Long
    Dim im As IniFile
    
    Dim testFilePath As String
    testFilePath = VBUtil.getApplicationIniFilePath("test.ini")
    
    ' ファイルが存在する場合は、削除する
    If (dir(testFilePath, vbNormal) <> "") Then
        Kill testFilePath
    End If
    
    ' ---------------------------------------------------------
    ' ファイル書き込み
    Set im = New IniFile
    im.init testFilePath
    
    im.setValue "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
              , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
              , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf
    im.destroy
    
        ' ファイル読み込み
    Set im = New IniFile
    im.init testFilePath
    
    retValue = im.GetValue("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
                         , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf)
    assert retValue = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf

    im.destroy
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' ファイル書き込み
    Set im = New IniFile
    im.init testFilePath
    
    im.setValue "section", "key", "value"
    im.setValue "section", "key2", ""
    im.setValue "セクション", "キー", "値"
    im.setValue "セクション", "キー=" & vbCr & vbLf, "値=" & vbCr & vbLf
    
    values.setItem Array("key", "value")
    values.setItem Array("キー", "値" & ChrW(&H9EB4))
    values.setItem Array("key3", "")
    im.setValues "sectionArray", values
    
    im.destroy
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' ファイル読み込み
    Set im = New IniFile
    im.init testFilePath
    
    retValue = im.GetValue("section", "key")
    assert retValue = "value"
    
    retValue = im.GetValue("section", "key2")
    assert retValue = ""
    
    retValue = im.GetValue("sectionNotExists", "key")
    assert retValue = ""
    
    retValue = im.GetValue("section", "keyNotExists")
    assert retValue = ""
    
    retValue = im.GetValue("セクション", "キー")
    assert retValue = "値"
    
    retValue = im.GetValue("セクション", "キー=" & vbCr & vbLf)
    assert retValue = "値=" & vbCr & vbLf
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.getItemByIndex(1, vbVariant)(1) = "key"
    assert retValueArray.getItemByIndex(1, vbVariant)(2) = "value"
    assert retValueArray.getItemByIndex(2, vbVariant)(1) = "キー"
    assert retValueArray.getItemByIndex(2, vbVariant)(2) = "値" & ChrW(&H9EB4)
    assert retValueArray.getItemByIndex(3, vbVariant)(1) = "key3"
    assert retValueArray.getItemByIndex(3, vbVariant)(2) = ""
    
    im.destroy
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' ファイル書き込みして削除
    Set im = New IniFile
    im.init testFilePath
    
    ' ※一回目
    im.delete "sectionArray"
    im.delete "section", "key"
    
    ' ※二回呼び出してみる
    im.delete "sectionArray"
    im.delete "section", "key"
    
    retValue = im.GetValue("section", "key")
    assert retValue = "" ' 削除したキーなので存在しない
    
    retValue = im.GetValue("セクション", "キー")
    assert retValue = "値" ' 何もしていないので存在する
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.count = 0 ' セクションごと削除
    
    im.destroy
    ' ---------------------------------------------------------
    

End Sub

' =========================================================
' ▽Iniファイル操作のパフォーマンステスト
'
' 概要　　　：
'
' =========================================================
Public Sub testIniFilePerform()

    Dim timeBegin As Long
    Dim timeEnd   As Long
    
    Dim retValue As String
    Dim retValueArray As ValCollection
    

    Dim i As Long
    Dim im As IniFile
    
    Dim testFilePath As String
    testFilePath = VBUtil.getApplicationIniFilePath("test_manydata.ini")
    
    ' ファイルが存在する場合は、削除する
    If (dir(testFilePath, vbNormal) <> "") Then
        Kill testFilePath
    End If
    
    ' ---------------------------------------------------------
    ' ファイル書き込み
    timeBegin = GetTickCount
    
    Set im = New IniFile
    im.init testFilePath
    
    For i = 1 To 10000
        im.setValue "section", "key" & i, "valueを書き込む・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・"
    Next
    
    im.destroy
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini書き込み：" & timeEnd - timeBegin & "ミリ秒経過"
    ' ---------------------------------------------------------
        
    ' ---------------------------------------------------------
    ' ファイル読み込み
    timeBegin = GetTickCount
    
    Set im = New IniFile
    im.init testFilePath
    
    Set retValueArray = im.getValues("section")
    
    im.destroy
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini読み込み：" & timeEnd - timeBegin & "ミリ秒経過"
    ' ---------------------------------------------------------

End Sub

' =========================================================
' ▽Iniワークシート操作のテスト
'
' 概要　　　：
'
' =========================================================
Public Sub testIniWorksheet()

    Dim retValue As String
    Dim retValueArray As ValCollection
    
    Dim values As New ValCollection
    

    Dim i As Long
    Dim im As IniWorksheet
        
    Dim testFileName As String
    testFileName = "test.ini"

    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
        
    ' ---------------------------------------------------------
    ' ファイル書き込み
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    ' 既にブックは存在するがデータはない場合
    assert Not im.isExistsData
    
    im.setValue "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
              , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
              , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf
    im.writeSheet
    
    ' データが存在する場合
    assert im.isExistsData
    
    ' ファイル読み込み
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    retValue = im.GetValue("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
                         , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf)
    assert retValue = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_あいうえおアイウエオｱｲｳｴｵ亜衣兎絵尾" & ChrW(&H9EB4) & vbTab & vbCr & vbLf

    im.writeSheet
    ' ---------------------------------------------------------

    ' ---------------------------------------------------------
    ' ファイル書き込み
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    im.setValue "section", "key", "value"
    im.setValue "section", "key2", ""
    im.setValue "セクション", "キー", "値"
    im.setValue "セクション", "キー=" & vbCr & vbLf, "値=" & vbCr & vbLf
    
    values.setItem Array("key", "value")
    values.setItem Array("キー", "値" & ChrW(&H9EB4))
    values.setItem Array("key3", "")
    im.setValues "sectionArray", values
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' ファイル読み込み
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    retValue = im.GetValue("section", "key")
    assert retValue = "value"
    
    retValue = im.GetValue("section", "key2")
    assert retValue = ""
    
    retValue = im.GetValue("sectionNotExists", "key")
    assert retValue = ""
    
    retValue = im.GetValue("section", "keyNotExists")
    assert retValue = ""
    
    retValue = im.GetValue("セクション", "キー")
    assert retValue = "値"
    
    retValue = im.GetValue("セクション", "キー=" & vbCr & vbLf)
    assert retValue = "値=" & vbCr & vbLf
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.getItemByIndex(1, vbVariant)(1) = "key"
    assert retValueArray.getItemByIndex(1, vbVariant)(2) = "value"
    assert retValueArray.getItemByIndex(2, vbVariant)(1) = "キー"
    assert retValueArray.getItemByIndex(2, vbVariant)(2) = "値" & ChrW(&H9EB4)
    assert retValueArray.getItemByIndex(3, vbVariant)(1) = "key3"
    assert retValueArray.getItemByIndex(3, vbVariant)(2) = ""
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' ファイル書き込みして削除
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    ' ※一回目
    im.delete "sectionArray"
    im.delete "section", "key"
    
    ' ※二回呼び出してみる
    im.delete "sectionArray"
    im.delete "section", "key"
    
    retValue = im.GetValue("section", "key")
    assert retValue = "" ' 削除したキーなので存在しない
    
    retValue = im.GetValue("セクション", "キー")
    assert retValue = "値" ' 何もしていないので存在する
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.count = 0 ' セクションごと削除
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' ファイル再読み込み
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    retValue = im.GetValue("section", "key2")
    assert retValue = ""
    
    retValue = im.GetValue("セクション", "キー")
    assert retValue = "値"
    
    retValue = im.GetValue("セクション", "キー=" & vbCr & vbLf)
    assert retValue = "値=" & vbCr & vbLf
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.count = 0
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' ファイル名を変更
    testFileName = "test2.ini"
    
    ' ---------------------------------------------------------
    ' ファイル書き込み
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    ' 既にブックは存在するがデータはない場合
    assert Not im.isExistsData
    
    im.setValue "section", "key", "value"
    im.setValue "section", "key2", ""
    im.setValue "セクション", "キー", "値"
    im.setValue "セクション", "キー=" & vbCr & vbLf, "値=" & vbCr & vbLf
    
    values.setItem Array("key", "value")
    values.setItem Array("キー", "値" & ChrW(&H9EB4))
    values.setItem Array("key3", "")
    im.setValues "sectionArray", values
    
    im.writeSheet
    
    ' データが存在する
    assert im.isExistsData
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' ファイル読み込み
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    retValue = im.GetValue("section", "key")
    assert retValue = "value"
    
    retValue = im.GetValue("section", "key2")
    assert retValue = ""
    
    retValue = im.GetValue("sectionNotExists", "key")
    assert retValue = ""
    
    retValue = im.GetValue("section", "keyNotExists")
    assert retValue = ""
    
    retValue = im.GetValue("セクション", "キー")
    assert retValue = "値"
    
    retValue = im.GetValue("セクション", "キー=" & vbCr & vbLf)
    assert retValue = "値=" & vbCr & vbLf
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.getItemByIndex(1, vbVariant)(1) = "key"
    assert retValueArray.getItemByIndex(1, vbVariant)(2) = "value"
    assert retValueArray.getItemByIndex(2, vbVariant)(1) = "キー"
    assert retValueArray.getItemByIndex(2, vbVariant)(2) = "値" & ChrW(&H9EB4)
    assert retValueArray.getItemByIndex(3, vbVariant)(1) = "key3"
    assert retValueArray.getItemByIndex(3, vbVariant)(2) = ""
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' ファイル名を変更
    testFileName = "test3.ini"
    
    ' ---------------------------------------------------------
    ' ファイル書き込み
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    im.setValue "section", "key", "value"
    im.setValue "section", "key2", ""
    im.setValue "セクション", "キー", "値"
    im.setValue "セクション", "キー=" & vbCr & vbLf, "値=" & vbCr & vbLf
    
    values.setItem Array("key", "value")
    values.setItem Array("キー", "値" & ChrW(&H9EB4))
    values.setItem Array("key3", "")
    im.setValues "sectionArray", values
    
    im.writeSheet
    ' ---------------------------------------------------------
    
End Sub

' =========================================================
' ▽Iniワークシート操作のパフォーマンステスト
'
' 概要　　　：
'
' =========================================================
Public Sub testIniWorksheetPerform()

    Dim timeBegin As Long
    Dim timeEnd   As Long
    
    Dim retValue As String
    Dim retValueArray As ValCollection
    

    Dim i As Long
    Dim im As IniWorksheet
    
    Dim testFilePath As String
    testFilePath = "test_manydata.ini"

    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    
    ' ---------------------------------------------------------
    ' ファイル書き込み
    timeBegin = GetTickCount
    
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFilePath
    
    For i = 1 To 10000
        im.setValue "section", "key" & i, "valueを書き込む・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・・"
    Next
    
    im.writeSheet
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini書き込み：" & timeEnd - timeBegin & "ミリ秒経過"
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' ファイル読み込み
    timeBegin = GetTickCount
    
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFilePath
    
    Set retValueArray = im.getValues("section")
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini読み込み：" & timeEnd - timeBegin & "ミリ秒経過"
    
    timeBegin = GetTickCount
    
    im.delete "section"
    im.writeSheet
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini消し込み：" & timeEnd - timeBegin & "ミリ秒経過"
    ' ---------------------------------------------------------

End Sub

#End If

