Attribute VB_Name = "ZTest"
Option Explicit

#If DEBUG_MODE = 1 Then

#If VBA7 And Win64 Then
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Declare Function GetTickCount Lib "kernel32" () As Long
#End If

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
    
    Set a = New RegistryManipulator: a.init WinAPI_ADVAP.HKEY_CURRENT_USER, "Software\SandSoft\Sut", WinAPI_ADVAP.KEY_ALL_ACCESS, True

    Dim value As String
    Set b = a.getKeyList
    a.setValue "key1", "��������������������"
    Debug.Print a.getValue("key1", value)
    Debug.Print a.getValue("key2", value)
    
    a.deleteValue "key1"
    
    Set a = New RegistryManipulator: a.init WinAPI_ADVAP.HKEY_CURRENT_USER, "Software\SandSoft", WinAPI_ADVAP.KEY_ALL_ACCESS, True
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
    For i = 1 To c.listCount
    
        ' ���X�g�̕�������o��
        Debug.Print c.list(i)
    Next
    
    ' �t�H���g�T�C�Y���X�g���擾����
    Set c = Application.CommandBars.FindControl(Id:=10000)
    
    Debug.Print c.BuiltIn
    
    ' ���X�g�̓��e��S�ĕ\������
    For i = 1 To c.listCount
    
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

Private Sub test14()

    Dim password As String
    Dim passwordLen As Long
    
    password = "password"
    passwordLen = Len(password)

    Dim buffer(0 To 1) As Byte
    Dim bufferLen      As Long
    
    Dim resultBuffer() As Byte
    Dim resultLen      As Long
    
    buffer(0) = Asc("a")
    buffer(1) = Asc("b")
    bufferLen = 2
    
    resultLen = 0
    If SutGray.Encrypt(password _
                      , passwordLen _
                      , buffer(0) _
                      , bufferLen _
                      , 0 _
                      , resultLen) = 0 Then
    
    End If

    ReDim resultBuffer(0 To resultLen - 1)

    If SutGray.Encrypt(password _
                      , passwordLen _
                      , buffer(0) _
                      , bufferLen _
                      , resultBuffer(0) _
                      , resultLen) = 0 Then
    
    End If

End Sub

Private Sub test15()

    Dim buffer         As String
    Dim bufferLen      As Long
    
    Dim resultBuffer() As Byte
    Dim resultLen      As Long
    
    buffer = "0102030405060708090a0b0c0d0e0f10"
    bufferLen = Len(buffer)
    resultLen = 0
    If SutGray.ConvertHexToBinaryData(buffer _
                                    , bufferLen _
                                    , 0 _
                                    , resultLen) = 0 Then
    
    
    End If

    ReDim resultBuffer(0 To resultLen - 1)

    If SutGray.ConvertHexToBinaryData(buffer _
                                    , bufferLen _
                                    , resultBuffer(0) _
                                    , resultLen) = 0 Then
    
    End If

End Sub

Private Sub test16()

    Dim buffer()       As Byte
    Dim bufferLen      As Long
    
    Dim resultBuffer   As String
    Dim resultLen      As Long
    
    ReDim buffer(0 To 16)
    Dim i As Long
    For i = 0 To 16
    
        buffer(i) = i
    Next
    bufferLen = 17
    
    resultLen = 0
    If SutGray.ConvertBinaryDataToHex(buffer(0) _
                                    , bufferLen _
                                    , 0 _
                                    , resultLen) = 0 Then
    
    
    End If

    resultBuffer = Space(resultLen)

    If SutGray.ConvertBinaryDataToHex(buffer(0) _
                                    , bufferLen _
                                    , resultBuffer _
                                    , resultLen) = 0 Then
    
    End If

End Sub

Private Sub test17()

    On Error GoTo err
    
    Dim a As New ExeAuthenticateLicence

    Dim b As Date

    b = a.getProbationDate

    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

Private Sub test18()

    On Error GoTo err
    
    Dim a As New ExeAuthenticateLicence
    
    Dim b As Date
    
    Dim i As Long
    
    For i = 0 To 20
    
        b = DateAdd("d", -1 * (20 - i), Now)
    
        If a.isRangeProbation(b) = True Then
        
            Debug.Print b & "�͈͓� " & a.getRemainderProbationDay(b)
        Else
        
            Debug.Print b & "�͈͊O " & a.getRemainderProbationDay(b)
        End If
        
    Next
    
    Exit Sub
err:

End Sub

Private Sub test19()

    Dim tmp As IPictureDisp
    Set tmp = SutYellow.LoadIconAndGetPictureDisp(101)
    
End Sub

Private Sub test20()

    Dim aaa As New ValApplicationSettingShortcut
    aaa.init
    
    Dim temp As New ValCollection
    temp.setItem "final_clash"
    temp.setItem "final_atack"

    Set aaa.rclickMenuItemList = temp
    
    aaa.readForRegistryForRClick
    aaa.writeForRegistryForRClick
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

    a.init


End Sub

Private Sub testValAppSettingColFormatW()

    Dim a As New ValApplicationSettingColFormat

    Set a.dbList = New ValCollection
    
    Dim i As Long
    
    Dim dbInfo1 As New sutredlib.DbColumnTypeDbInfo
    dbInfo1.dbName = "oracle"

    Dim colList1(0 To 10) As sutredlib.DbColumnTypeColInfo
    
    For i = LBound(colList1) To UBound(colList1)
    
        Set colList1(i) = New sutredlib.DbColumnTypeColInfo
        
        colList1(i).columnName = "�J�������F" & i
        colList1(i).formatUpdate = "�t�H�[�}�b�gU�F" & i
        colList1(i).formatSelect = "�t�H�[�}�b�gS�F" & i
    
    Next
    
    Set dbInfo1.columnList = colList1

    Dim dbInfo2 As New sutredlib.DbColumnTypeDbInfo
    dbInfo2.dbName = "postgre"

    a.dbList.setItem dbInfo1
    a.dbList.setItem dbInfo2
    
    a.writeForRegistry

End Sub

Private Sub testExeDataTypeReader()

    Dim ct As sutredlib.DbColumnTypeDbInfo

    Dim impl As IDbColumnType
    Dim fac As New DbObjectFactory
    Set impl = fac.createColumnType(DbmsType.Oracle): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.MySQL): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.PostgreSQL): Set ct = impl.getDefaultColumnFormat
    Set impl = fac.createColumnType(DbmsType.Symfoware): Set ct = impl.getDefaultColumnFormat
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

    Main.getApplicationSettingColFormat

End Sub

Private Sub S()
    
End Sub

Public Sub pallet()

    Application.CommandBars("Fill Color").visible = True
    Application.Dialogs.item(xlDialogColorPalette).Show
End Sub

Private Sub StringTest()

    Dim str  As String
    Dim tmp  As String
    
    Dim str2 As New SutString
    Dim str3 As New SutString
    
    Dim i   As Long
    
    Dim timeBegin As Long
    Dim timeEnd   As Long
    
    ' -----------------------------------------------------------
    ' ���̂Q
    ' -----------------------------------------------------------
    timeBegin = GetTickCount
    For i = 0 To 1000000
        str2.append "a"
    Next
    tmp = str2.str
    timeEnd = GetTickCount
    ' -----------------------------------------------------------

    MsgBox "���̂Q�F" & (timeEnd - timeBegin) & "�~���b" & tmp
    
    ' -----------------------------------------------------------
    ' ���̂R
    ' -----------------------------------------------------------
    timeBegin = GetTickCount
    str3.reserve 1000000
    For i = 0 To 1000000
        str3.append "a"
    Next
    tmp = str3.str
    timeEnd = GetTickCount
    ' -----------------------------------------------------------

    MsgBox "���̂R�F" & (timeEnd - timeBegin) & "�~���b" & tmp
    
    ' -----------------------------------------------------------
    ' ���̂P
    ' -----------------------------------------------------------
    timeBegin = GetTickCount
    For i = 0 To 100000
        str = str & "a"
    Next
    timeEnd = GetTickCount
    ' -----------------------------------------------------------

    MsgBox "���̂P�F" & (timeEnd - timeBegin) & "�~���b" & str
    
End Sub

Private Sub SutStringTest()

    Dim str As New SutString
    
    Debug.Print str.str
    str.Erase
    Debug.Print str.str
    str.append "��bc"
    Debug.Print str.str
    str.Erase 0, 1
    Debug.Print str.str
    str.Erase
    Debug.Print str.str
    str.append "abc"
    str.reserve 10000
    Debug.Print str.str
    
    str.Erase 2, 1
    Debug.Print str.str
    str.Erase 10, 10
    Debug.Print str.str
    
End Sub

Private Sub SutStringTest2()

    Dim str As New SutString
    
    str.append "����������"
    str.replace "", ""
    Debug.Print "result: " & str.str
    str.replace "��", "��"
    Debug.Print "result: " & str.str
    str.replace "��", "a"
    Debug.Print "result: " & str.str
    str.replace "a", "��"
    Debug.Print "result: " & str.str
    str.replace "��", ""
    Debug.Print "result: " & str.str
    
End Sub

Private Sub SutStringTest3()

    Dim str As New SutString
    
    str.assign "a"
    Debug.Print "result: " & str.str
    
    str.append("a").append ("b")
    Debug.Print str.substr
    Debug.Print str.substr(0)
    Debug.Print str.substr(1)
    Debug.Print str.substr(1, 0)
    Debug.Print str.substr(1, 2)
    Debug.Print str.substr(1, 3)
    Debug.Print str.substr(1, 4)
    Debug.Print str.substr(4, 1)
    Debug.Print str.substr(5, 1)
    Debug.Print str.substr(-2, -2)
    
    
    str.assign "ab��": str.insert -1, "��": Debug.Print str.str
    str.assign "ab��": str.insert 0, "��": Debug.Print str.str
    str.assign "ab��": str.insert 1, "��": Debug.Print str.str
    str.assign "ab��": str.insert 2, "��": Debug.Print str.str
    str.assign "ab��": str.insert 3, "��": Debug.Print str.str
    
    str.assign "ab��": str.insert 3, "����", 0: Debug.Print str.str
    str.assign "ab��": str.insert 3, "����", 0, 1: Debug.Print str.str
    
End Sub

Public Sub StringBuilderTest()

    Dim str As StringBuilder
    Set str = New StringBuilder
    
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.clear
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.remove 0, 0
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.append "����������"
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.append "����������"
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.remove 1, 5
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.remove 1, 5
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.append "����������"
    str.append "����������"
    str.remove 1, 11
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.clear
    str.append "����������"
    str.insert "����������", 0
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.insert "����������", 7
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.insert "����������", 6
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
  
    str.insert "����������", 1
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
  
    str.insert "�^��", 2
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
  
    str.replace "�^��", "���X"
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
  
    str.replace "���X", "�^��"
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
  
    str.replace "�^��", "�^��"
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
    
    str.clear
    str.append String(1050, "��")
    Debug.Print "result: " & str.length & ":" & str.capacity & ":" & str.str
  
End Sub


#End If


Private Sub testCsvParser()

    Dim ret  As ValCollection
    Dim var  As Variant
    Dim var2 As Variant
    Dim varBuff As New StringBuilder
    
    Dim csvp As New CsvParser: csvp.init
    
    Set ret = csvp.parse("a,b,c,d")
    For Each var In ret.col
        varBuff.clear
        For Each var2 In var.col
            varBuff.append var2 & ":"
        Next
        Debug.Print varBuff.str
    Next
    Debug.Print ""

    Set ret = csvp.parse("a,b,c,d" & vbNewLine & "e,f,g,h")
    For Each var In ret.col
        varBuff.clear
        For Each var2 In var.col
            varBuff.append var2 & ":"
        Next
        Debug.Print varBuff.str
    Next
    Debug.Print ""

    Set ret = csvp.parse("a,b,c,d" & vbCr & "e,f,g,h")
    For Each var In ret.col
        varBuff.clear
        For Each var2 In var.col
            varBuff.append var2 & ":"
        Next
        Debug.Print varBuff.str
    Next
    Debug.Print ""

    Set ret = csvp.parse("a,b,c,d" & vbLf & "e,f,g,h")
    For Each var In ret.col
        varBuff.clear
        For Each var2 In var.col
            varBuff.append var2 & ":"
        Next
        Debug.Print varBuff.str
    Next
    Debug.Print ""

    Set ret = csvp.parse("""a"",b,c,d,""e,e"",""""""e,e""""""")
    For Each var In ret.col
        varBuff.clear
        For Each var2 In var.col
            varBuff.append var2 & ":"
        Next
        Debug.Print varBuff.str
    Next
    Debug.Print ""

    Set ret = csvp.parse("""a"",b,c,d,""e,e"",""""""e,e""""""" & vbNewLine & "����������,����������,����������" & vbNewLine & "��""��������,��""������""��,��""������""��")
    For Each var In ret.col
        varBuff.clear
        For Each var2 In var.col
            varBuff.append var2 & ":"
        Next
        Debug.Print varBuff.str
    Next
    Debug.Print ""

    Set ret = csvp.parse("""param 1"",""""")
    For Each var In ret.col
        varBuff.clear
        For Each var2 In var.col
            varBuff.append var2 & ":"
        Next
        Debug.Print varBuff.str
    Next
    Debug.Print ""

    Set ret = csvp.parse(""""",""""")
    For Each var In ret.col
        varBuff.clear
        For Each var2 In var.col
            varBuff.append var2 & ":"
        Next
        Debug.Print varBuff.str
    Next
    Debug.Print ""


End Sub

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
