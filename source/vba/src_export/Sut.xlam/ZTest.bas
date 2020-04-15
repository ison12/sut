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
    str = "��������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������"
    
    Dim i As Long
    Dim count As Long: count = 10000000

    timeBegin = GetTickCount
    For i = 1 To count
        taestByValForObjByRef obj
    Next
    timeEnd = GetTickCount
    Debug.Print "Obj�Q�Ɠn���F" & (timeEnd - timeBegin) & "�~���b"
    
    timeBegin = GetTickCount
    For i = 1 To count
        taestByValForObjByVal obj
    Next
    timeEnd = GetTickCount
    Debug.Print "Obj�@�l�n���F" & (timeEnd - timeBegin) & "�~���b"

    timeBegin = GetTickCount
    For i = 1 To count
        taestByValForStrByRef str
    Next
    timeEnd = GetTickCount
    Debug.Print "Str�Q�Ɠn���F" & (timeEnd - timeBegin) & "�~���b"
    
    timeBegin = GetTickCount
    For i = 1 To count
        taestByValForStrByVal str
    Next
    timeEnd = GetTickCount
    Debug.Print "Str�@�l�n���F" & (timeEnd - timeBegin) & "�~���b"

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
    
    Randomize    ' �����������[�`�������������܂��B
    
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

    MsgBox "����"
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
    a.setValue "key1", "��������������������"
    Debug.Print a.GetValue("key1", value)
    Debug.Print a.GetValue("key2", value)
    
    a.deleteValue "key1"
    
    Set a = New RegistryManipulator: a.init WinAPI_ADVAP.HKEY_CURRENT_USER, "Software\ison", WinAPI_ADVAP.KEY_ALL_ACCESS, True
    a.delete "Sut"
    
End Sub

Private Sub test7()

    frmFileOutput.ShowExt vbModal, "�w�b�_�ł��B", "file.sql"

End Sub

Private Sub test8()

    Dim file As New FileWriter: file.init "test.txt", "euc-jp", vbCrLf, True
    
    file.writeText "����������", True
    file.writeText "����������", True
    file.writeText "����������", True
    
    file.writeText "�����Ă�", True
    
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

    ' ���[�v�C���f�b�N�X
    Dim i As Long
    
    ' �R���{�{�b�N�X
    Dim c As CommandBarComboBox
    
    ' �t�H���g���X�g���擾����
    Set c = Application.CommandBars.FindControl(Id:=1728)
    
    Debug.Print c.BuiltIn
    
    ' ���X�g�̓��e��S�ĕ\������
    For i = 1 To c.ListCount
    
        ' ���X�g�̕�������o��
        Debug.Print c.list(i)
    Next
    
    ' �t�H���g�T�C�Y���X�g���擾����
    Set c = Application.CommandBars.FindControl(Id:=10000)
    
    Debug.Print c.BuiltIn
    
    ' ���X�g�̓��e��S�ĕ\������
    For i = 1 To c.ListCount
    
        ' ���X�g�̕�������o��
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

    ' �G�N�Z���̃o�[�W����
    Dim excelVer As ExcelVersion: excelVer = ExcelUtil.getExcelVersion
    
    ' �R�}���h�o�[
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
        
    ' DB�ڑ��{�^��
    Dim btnDBConnect              As CommandBarButton
    ' DB�ؒf�{�^��
    Dim btnDBDisConnect           As CommandBarButton
    
    ' DB�ڑ��{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnDBConnect = cb.Controls.Add(Type:=msoControlButton)
    
    ' DB�ڑ��{�^���̃v���p�e�B��ݒ肷��
    With btnDBConnect
    
        .Style = msoButtonIconAndCaption
        .Caption = "�ڑ�"
        .DescriptionText = "DB�ڑ�"
        .OnAction = "Main.SutConnectDB"
        .Tag = COMMANDBAR_CONTROL_BASE_ID & "Main.SutConnectDB"
        
    End With
        
    ' DB�ؒf�{�^�����R�}���h�o�[�Ƀ{�^����ǉ�����
    Set btnDBDisConnect = cb.Controls.Add(Type:=msoControlButton)
    
    ' DB�ؒf�{�^���̃v���p�e�B��ݒ肷��
    With btnDBDisConnect
    
        .Style = msoButtonIconAndCaption
        .Caption = "�ؒf"
        .DescriptionText = "DB�ؒf"
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

    a = "�������Ⴄ��"
    
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
    str.append "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf
    assert str.length = 118
    assert str.str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf
    
    str.clear
    str.append "a"
    assert str.length = 1
    assert str.str = "a"
    
    str.clear
    assert str.length = 0
    assert str.str = ""
    
    str.append "ab"
    
    ' �����s��
    str.remove 0, 0
    str.remove 1, 0
    str.remove 0, 1
    
    assert str.length = 2
    assert str.str = "ab"
    
    str.append "����������"
    assert str.length = 7
    assert str.str = "ab����������"
    
    str.append "����������"
    assert str.length = 12
    assert str.str = "ab��������������������"
    
    str.remove 1, 5
    assert str.length = 7
    assert str.str = "��������������"
    
    str.remove 1, 5
    assert str.length = 2
    assert str.str = "����"
    
    str.append "����������"
    str.append "����������"
    str.remove 1, 11
    assert str.length = 1
    assert str.str = "��"
    
    str.clear
    str.append "����������"
    
    ' �����s��
    str.insert 0, "����������"
    assert str.length = 5
    assert str.str = "����������"
    
    ' �����s��
    str.insert 7, "����������"
    assert str.length = 5
    assert str.str = "����������"
    
    str.clear
    str.insert 1, "����������"
    assert str.length = 5
    assert str.str = "����������"
    
    str.clear
    str.append "�����������@����������"
    str.replace "����", "�u��REP"
    assert str.length = 14
    assert str.str = "���u��REP�����@����������"
    
    str.clear
    str.append "�����������@����������"
    str.replace "������", "�u��REP"
    assert str.length = 11
    assert str.str = "�����������@����������"
    
    str.clear
    str.append "�����������@����������"
    str.replace "��", "�A"
    assert str.length = 11
    assert str.str = "�A���������@����������"
    
    str.clear
    str.append "�����������@����������"
    str.replace "��", "�R"
    assert str.length = 11
    assert str.str = "�����������@���������R"
    
    str.clear
    str.append "�����������@������������"
    str.replace "��", "��"
    assert str.length = 12
    assert str.str = "�����������@������������"
    
    str.clear
    str.append String(1000, "��")
    assert str.length = 1000
    assert str.str = String(1000, "��")
    
    str.append String(1000, "��")
    assert str.length = 2000
    assert str.str = String(2000, "��")
    
    str.append String(1000, "��")
    assert str.length = 3000
    assert str.str = String(3000, "��")
    
    str.append String(1000, "��")
    assert str.length = 4000
    assert str.str = String(4000, "��")
    
    str.append String(1000, "��")
    assert str.length = 5000
    assert str.str = String(5000, "��")
  
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
    testStr = """a"",b,c,d,""e,e"",""""""e,e""""""" & vbNewLine & "����������,����������,����������" & vbNewLine & "��""��������,��""������""��,��""������""��"
    Set ret = csvp.parse(testStr)

    assert testCsvParserToString(ret) = "a,b,c,d,e,e,""e,e""" & vbNewLine & "����������,����������,����������" & vbNewLine & "��""��������,��""������""��,��""������""��"
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
    
    testStr = "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & "����������" & vbTab & vbNewLine & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab & vbLf & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab & vbCr & _
              vbCr & _
              "a" & vbTab & "Microsoft OLE DB for SQL Server" & vbTab & "" & vbTab & "10.12.3.176" & vbTab & "" & vbTab & "ORSYS_DATA" & vbTab & "sa" & vbTab & vbTab
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, vbTab, vbLf) = replace(replace(testStr, vbNewLine, vbLf), vbCr, vbLf)
    ' ------------------------------------------------------

    ' ------------------------------------------------------
    csvp.init vbTab
    
    testStr = "abc����������" & vbTab & """����" & vbTab & vbNewLine & "������""" & vbNewLine & _
              "abc����������" & vbTab & """����" & vbTab & "������"""
    Set ret = csvp.parse(testStr)
    
    assert testCsvParserToString(ret, vbTab, vbNewLine) = "abc����������" & vbTab & "����" & vbTab & vbNewLine & "������" & vbNewLine & _
                                                          "abc����������" & vbTab & "����" & vbTab & "������"
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
    a.setItem "����������"
    a.setItem "����������"
    a.setItemByIndexAfter "����������", 1
    a.setItemByIndexAfter "�����Ă�", 1
    a.setItemByIndexBefore "�Ȃɂʂ˂�", 1
    a.setItemByIndexBefore "�͂Ђӂւ�", 1
    
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
    tableWorksheet.sheetName = "�V�[�gA"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "�V�[�gB"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "�V�[�gC"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|,<.>/?_;+:*]}@`[{"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    tableWorksheet.sheetName = "*�V�[�gE?"
    tableList.setItem tableWorksheet
    
    Set tableWorksheet = New ValTableWorksheet
    Set tableWorksheet.table = New ValDbDefineTable
    tableWorksheet.table.tableName = "�V�[�gZ"
    tableList.setItem tableWorksheet

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "�V�[�gA")
    assert ret.count = 1

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "�V�[�gD")
    assert ret.count = 0

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "�V�[�g*")
    assert ret.count = 3

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "�V?�g")
    assert ret.count = 0

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "�V*�g?")
    assert ret.count = 3

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "�V�[*")
    assert ret.count = 3

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "*�[�g*")
    assert ret.count = 4

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "*A")
    assert ret.count = 1

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|,<.>/?_;+:*]}@`[{")
    assert ret.count = 1

    Set ret = VBUtil.filterWildcard(tableList, "sheetName", "~*�V�[�gE~?")
    assert ret.count = 1

    Set ret = VBUtil.filterWildcard(tableList, "table.tableName", "�V�[�gZ")
    assert ret.count = 1
    
    Debug.Print "testFilterWildcard complete"

End Sub

' =========================================================
' ��ExcelCursorWait�N���X�̃e�X�g
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub testExcelLongTimeProcessing()

    Dim tmpCalculation As XlCalculation
    tmpCalculation = Application.calculation

    Dim e As ExcelLongTimeProcessing
    Set e = New ExcelLongTimeProcessing
    
    ' ---------------------------------------------
    ' �������`�j�� destroy
    e.init True, True, True, True, True, True, True
    assert Application.calculation = xlCalculationManual
    assert Application.displayAlerts = False
    'assert Application.enableCancelKey = xlDisabled ' Debug���̓v���p�e�B�̐ݒ肪�L���ɂȂ�Ȃ��̂�assert���Ȃ�
    assert Application.enableEvents = False
    assert Application.cursor = xlWait
    assert Application.screenUpdating = False
    'assert Application.interactive = False ' Debug���̓v���p�e�B�̐ݒ肪�L���ɂȂ�Ȃ��̂�assert���Ȃ�
    
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
    ' �������`�j�� nothing
    e.init True, True, True, True, True, True, True
    assert Application.calculation = xlCalculationManual
    assert Application.displayAlerts = False
    'assert Application.enableCancelKey = xlDisabled ' Debug���̓v���p�e�B�̐ݒ肪�L���ɂȂ�Ȃ��̂�assert���Ȃ�
    assert Application.enableEvents = False
    assert Application.cursor = xlWait
    assert Application.screenUpdating = False
    'assert Application.interactive = False ' Debug���̓v���p�e�B�̐ݒ肪�L���ɂȂ�Ȃ��̂�assert���Ȃ�
    
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
    ' �������`�j�� displayAlerts �̂ݗL��
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
    ' �������`�j�� enableEvents �̂ݗL��
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
    ' �������`�j�� cursor �̂ݗL��
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
    ' �������`�j�� screenUpdating �̂ݗL��
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
    ' �������`�j�� calculation �̂ݗL��
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
' ��ExcelCursorWait�N���X�̃e�X�g
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub testExcelCursorWait()

    Application.cursor = xlDefault

    Dim w As ExcelCursorWait
    Set w = New ExcelCursorWait
    
    ' ---------------------------------------------
    ' �������`�j�� destroy
    w.init
    assert Application.cursor = xlWait
    
    w.destroy
    assert Application.cursor = xlDefault
    ' ---------------------------------------------
    
    ' ---------------------------------------------
    ' �������`�j�� forceRestore
    w.init
    assert Application.cursor = xlWait
    
    w.forceRestore
    assert Application.cursor = xlDefault
    ' ---------------------------------------------
    
    ' ---------------------------------------------
    ' �������`�j�� nothing
    w.init
    assert Application.cursor = xlWait
    
    Set w = Nothing
    assert Application.cursor = xlDefault
    ' ---------------------------------------------
    
    ' ---------------------------------------------
    ' �������`�j�� �I������p�����đҋ@��Ԃɂ���̂�xlWait�̂܂�
    Set w = New ExcelCursorWait
    w.init True
    assert Application.cursor = xlWait
    
    Set w = Nothing
    assert Application.cursor = xlWait
    ' ---------------------------------------------

    Application.cursor = xlDefault

End Sub

' =========================================================
' ��Error�������̃e�X�g
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub testErrorRaise()

    On Error GoTo err
    
    ' �����K���ȃI�u�W�F�N�g�����������Ă���
    Dim obj As ExcelLongTimeProcessing
    Set obj = New ExcelLongTimeProcessing
    obj.init

    err.Raise 1000, "my source", "my description"

    Exit Sub

err:

    ' �G���[����ޔ�����
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    ' �����K���ȃI�u�W�F�N�g��j������i�f�X�g���N�^����Err�Ɋւ��鑀�삪�s���邱�Ɓj
    Set obj = Nothing

    assert err.Number = 0
    assert err.Source = ""
    assert err.Description = ""
    
    ' �ޔ������G���[����ݒ肵�Ȃ���
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
' ��Ini�t�@�C������̃e�X�g
'
' �T�v�@�@�@�F
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
    
    ' �t�@�C�������݂���ꍇ�́A�폜����
    If (dir(testFilePath, vbNormal) <> "") Then
        Kill testFilePath
    End If
    
    ' ---------------------------------------------------------
    ' �t�@�C����������
    Set im = New IniFile
    im.init testFilePath
    
    im.setValue "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
              , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
              , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf
    im.destroy
    
        ' �t�@�C���ǂݍ���
    Set im = New IniFile
    im.init testFilePath
    
    retValue = im.GetValue("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
                         , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf)
    assert retValue = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf

    im.destroy
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' �t�@�C����������
    Set im = New IniFile
    im.init testFilePath
    
    im.setValue "section", "key", "value"
    im.setValue "section", "key2", ""
    im.setValue "�Z�N�V����", "�L�[", "�l"
    im.setValue "�Z�N�V����", "�L�[=" & vbCr & vbLf, "�l=" & vbCr & vbLf
    
    values.setItem Array("key", "value")
    values.setItem Array("�L�[", "�l" & ChrW(&H9EB4))
    values.setItem Array("key3", "")
    im.setValues "sectionArray", values
    
    im.destroy
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' �t�@�C���ǂݍ���
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
    
    retValue = im.GetValue("�Z�N�V����", "�L�[")
    assert retValue = "�l"
    
    retValue = im.GetValue("�Z�N�V����", "�L�[=" & vbCr & vbLf)
    assert retValue = "�l=" & vbCr & vbLf
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.getItemByIndex(1, vbVariant)(1) = "key"
    assert retValueArray.getItemByIndex(1, vbVariant)(2) = "value"
    assert retValueArray.getItemByIndex(2, vbVariant)(1) = "�L�["
    assert retValueArray.getItemByIndex(2, vbVariant)(2) = "�l" & ChrW(&H9EB4)
    assert retValueArray.getItemByIndex(3, vbVariant)(1) = "key3"
    assert retValueArray.getItemByIndex(3, vbVariant)(2) = ""
    
    im.destroy
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' �t�@�C���������݂��č폜
    Set im = New IniFile
    im.init testFilePath
    
    ' ������
    im.delete "sectionArray"
    im.delete "section", "key"
    
    ' �����Ăяo���Ă݂�
    im.delete "sectionArray"
    im.delete "section", "key"
    
    retValue = im.GetValue("section", "key")
    assert retValue = "" ' �폜�����L�[�Ȃ̂ő��݂��Ȃ�
    
    retValue = im.GetValue("�Z�N�V����", "�L�[")
    assert retValue = "�l" ' �������Ă��Ȃ��̂ő��݂���
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.count = 0 ' �Z�N�V�������ƍ폜
    
    im.destroy
    ' ---------------------------------------------------------
    

End Sub

' =========================================================
' ��Ini�t�@�C������̃p�t�H�[�}���X�e�X�g
'
' �T�v�@�@�@�F
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
    
    ' �t�@�C�������݂���ꍇ�́A�폜����
    If (dir(testFilePath, vbNormal) <> "") Then
        Kill testFilePath
    End If
    
    ' ---------------------------------------------------------
    ' �t�@�C����������
    timeBegin = GetTickCount
    
    Set im = New IniFile
    im.init testFilePath
    
    For i = 1 To 10000
        im.setValue "section", "key" & i, "value���������ށE�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E"
    Next
    
    im.destroy
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini�������݁F" & timeEnd - timeBegin & "�~���b�o��"
    ' ---------------------------------------------------------
        
    ' ---------------------------------------------------------
    ' �t�@�C���ǂݍ���
    timeBegin = GetTickCount
    
    Set im = New IniFile
    im.init testFilePath
    
    Set retValueArray = im.getValues("section")
    
    im.destroy
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini�ǂݍ��݁F" & timeEnd - timeBegin & "�~���b�o��"
    ' ---------------------------------------------------------

End Sub

' =========================================================
' ��Ini���[�N�V�[�g����̃e�X�g
'
' �T�v�@�@�@�F
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
    ' �t�@�C����������
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    ' ���Ƀu�b�N�͑��݂��邪�f�[�^�͂Ȃ��ꍇ
    assert Not im.isExistsData
    
    im.setValue "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
              , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
              , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf
    im.writeSheet
    
    ' �f�[�^�����݂���ꍇ
    assert im.isExistsData
    
    ' �t�@�C���ǂݍ���
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    retValue = im.GetValue("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf _
                         , "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf)
    assert retValue = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!""#$%&'()-=^~\|@`[{;+:*]},<.>/?_�����������A�C�E�G�I��������ߓe�G��" & ChrW(&H9EB4) & vbTab & vbCr & vbLf

    im.writeSheet
    ' ---------------------------------------------------------

    ' ---------------------------------------------------------
    ' �t�@�C����������
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    im.setValue "section", "key", "value"
    im.setValue "section", "key2", ""
    im.setValue "�Z�N�V����", "�L�[", "�l"
    im.setValue "�Z�N�V����", "�L�[=" & vbCr & vbLf, "�l=" & vbCr & vbLf
    
    values.setItem Array("key", "value")
    values.setItem Array("�L�[", "�l" & ChrW(&H9EB4))
    values.setItem Array("key3", "")
    im.setValues "sectionArray", values
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' �t�@�C���ǂݍ���
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
    
    retValue = im.GetValue("�Z�N�V����", "�L�[")
    assert retValue = "�l"
    
    retValue = im.GetValue("�Z�N�V����", "�L�[=" & vbCr & vbLf)
    assert retValue = "�l=" & vbCr & vbLf
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.getItemByIndex(1, vbVariant)(1) = "key"
    assert retValueArray.getItemByIndex(1, vbVariant)(2) = "value"
    assert retValueArray.getItemByIndex(2, vbVariant)(1) = "�L�["
    assert retValueArray.getItemByIndex(2, vbVariant)(2) = "�l" & ChrW(&H9EB4)
    assert retValueArray.getItemByIndex(3, vbVariant)(1) = "key3"
    assert retValueArray.getItemByIndex(3, vbVariant)(2) = ""
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' �t�@�C���������݂��č폜
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    ' ������
    im.delete "sectionArray"
    im.delete "section", "key"
    
    ' �����Ăяo���Ă݂�
    im.delete "sectionArray"
    im.delete "section", "key"
    
    retValue = im.GetValue("section", "key")
    assert retValue = "" ' �폜�����L�[�Ȃ̂ő��݂��Ȃ�
    
    retValue = im.GetValue("�Z�N�V����", "�L�[")
    assert retValue = "�l" ' �������Ă��Ȃ��̂ő��݂���
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.count = 0 ' �Z�N�V�������ƍ폜
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' �t�@�C���ēǂݍ���
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    retValue = im.GetValue("section", "key2")
    assert retValue = ""
    
    retValue = im.GetValue("�Z�N�V����", "�L�[")
    assert retValue = "�l"
    
    retValue = im.GetValue("�Z�N�V����", "�L�[=" & vbCr & vbLf)
    assert retValue = "�l=" & vbCr & vbLf
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.count = 0
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' �t�@�C������ύX
    testFileName = "test2.ini"
    
    ' ---------------------------------------------------------
    ' �t�@�C����������
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    ' ���Ƀu�b�N�͑��݂��邪�f�[�^�͂Ȃ��ꍇ
    assert Not im.isExistsData
    
    im.setValue "section", "key", "value"
    im.setValue "section", "key2", ""
    im.setValue "�Z�N�V����", "�L�[", "�l"
    im.setValue "�Z�N�V����", "�L�[=" & vbCr & vbLf, "�l=" & vbCr & vbLf
    
    values.setItem Array("key", "value")
    values.setItem Array("�L�[", "�l" & ChrW(&H9EB4))
    values.setItem Array("key3", "")
    im.setValues "sectionArray", values
    
    im.writeSheet
    
    ' �f�[�^�����݂���
    assert im.isExistsData
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' �t�@�C���ǂݍ���
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
    
    retValue = im.GetValue("�Z�N�V����", "�L�[")
    assert retValue = "�l"
    
    retValue = im.GetValue("�Z�N�V����", "�L�[=" & vbCr & vbLf)
    assert retValue = "�l=" & vbCr & vbLf
    
    Set retValueArray = im.getValues("sectionArray")
    assert retValueArray.getItemByIndex(1, vbVariant)(1) = "key"
    assert retValueArray.getItemByIndex(1, vbVariant)(2) = "value"
    assert retValueArray.getItemByIndex(2, vbVariant)(1) = "�L�["
    assert retValueArray.getItemByIndex(2, vbVariant)(2) = "�l" & ChrW(&H9EB4)
    assert retValueArray.getItemByIndex(3, vbVariant)(1) = "key3"
    assert retValueArray.getItemByIndex(3, vbVariant)(2) = ""
    
    im.writeSheet
    ' ---------------------------------------------------------
    
    ' �t�@�C������ύX
    testFileName = "test3.ini"
    
    ' ---------------------------------------------------------
    ' �t�@�C����������
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFileName
    
    im.setValue "section", "key", "value"
    im.setValue "section", "key2", ""
    im.setValue "�Z�N�V����", "�L�[", "�l"
    im.setValue "�Z�N�V����", "�L�[=" & vbCr & vbLf, "�l=" & vbCr & vbLf
    
    values.setItem Array("key", "value")
    values.setItem Array("�L�[", "�l" & ChrW(&H9EB4))
    values.setItem Array("key3", "")
    im.setValues "sectionArray", values
    
    im.writeSheet
    ' ---------------------------------------------------------
    
End Sub

' =========================================================
' ��Ini���[�N�V�[�g����̃p�t�H�[�}���X�e�X�g
'
' �T�v�@�@�@�F
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
    ' �t�@�C����������
    timeBegin = GetTickCount
    
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFilePath
    
    For i = 1 To 10000
        im.setValue "section", "key" & i, "value���������ށE�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E�E"
    Next
    
    im.writeSheet
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini�������݁F" & timeEnd - timeBegin & "�~���b�o��"
    ' ---------------------------------------------------------
    
    ' ---------------------------------------------------------
    ' �t�@�C���ǂݍ���
    timeBegin = GetTickCount
    
    Set im = New IniWorksheet
    im.init wb, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, testFilePath
    
    Set retValueArray = im.getValues("section")
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini�ǂݍ��݁F" & timeEnd - timeBegin & "�~���b�o��"
    
    timeBegin = GetTickCount
    
    im.delete "section"
    im.writeSheet
    
    timeEnd = GetTickCount
    
    Debug.Print "Ini�������݁F" & timeEnd - timeBegin & "�~���b�o��"
    ' ---------------------------------------------------------

End Sub

#End If

