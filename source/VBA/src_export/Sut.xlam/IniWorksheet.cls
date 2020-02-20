VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' INI�t�@�C������N���X�iExcel���[�N�V�[�g�Łj
'
' ��������
' �E�Z�N�V�����̏ȗ��s��
' �E�R�����g�ݒ�s��
' �E�f�[�^�{�����A���s�R�[�h��C�R�[���L��������Ȑ���R�[�h�ɕϊ������
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/02/12�@�V�K�쐬
'
' ���L�����F
' *********************************************************
' �u�b�N
Private book As Workbook
' �V�[�g
Private sheet As Worksheet
' �V�[�g��
Private sheetName As String
' �t�@�C����
Private fileName As String
' �t�@�C�����Ɋ֘A������ԍ�
Private fileNameColumn As Long
' INI�f�[�^
Private inid As IniData

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Public Sub Class_Initialize()
    
    Set inid = New IniData
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Public Sub Class_Terminate()
    
    destroy
    
End Sub

' =========================================================
' �������������B
'
' �T�v�@�@�@�F
' �����@�@�@�Fbook_      �u�b�N
'     �@�@�@�FsheetName_ �V�[�g��
'     �@�@�@�FfileName_  �t�@�C����
' �߂�l�@�@�F
'
' =========================================================
Public Sub init(ByVal book_ As Workbook, ByVal sheetName_ As String, ByVal fileName_ As String)

    ' �����o�ݒ�
    Set book = book_
    sheetName = sheetName_
    fileName = fileName_
    
    ' ���ɃV�[�g�����݂���ꍇ�́A�V�[�g�I�u�W�F�N�g���擾�������o�ɏ���ݒ肷��
    initSheet True
    
    readSheet

End Sub

' =========================================================
' ���V�[�g�̏����������B
'
' �T�v�@�@�@�F
' �����@�@�@�FisGetOnly �V�[�g�I�u�W�F�N�g�̎擾�̂݃t���O
' �߂�l�@�@�F
'
' =========================================================
Public Sub initSheet(Optional ByVal isGetOnly As Boolean = False)

    ' �V�[�g�̐���
    Set sheet = createSheet(book, sheetName, isGetOnly)
    If sheet Is Nothing Then
        Exit Sub
    End If
    
    ' �V�[�g���̗�ԍ��擾
    fileNameColumn = findColumnFilePath(sheet, fileName, False)
    
End Sub

' =========================================================
' ���j�������B
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub destroy()

    writeSheet

End Sub

' =========================================================
' ���V�[�g�̓ǂݍ��ݏ����B
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub readSheet()
    
    Dim contents As String
    If sheet Is Nothing Then
        contents = ""
    Else
        contents = readContents(sheet, fileNameColumn)
    End If
    
    inid.datas = inid.readContents(contents)
        
    ' �ǂݍ��ݒ���Ȃ̂ŕύX�ς݂��I�t�ɂ���
    inid.isModified = False

End Sub

' =========================================================
' ���V�[�g�̏������ݏ����B
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub writeSheet()
    
    If inid.isModified = False Then
        ' �ύX�Ȃ��Ȃ̂ŏ������܂Ȃ�
        Exit Sub
    End If

    Dim contents As String
    contents = inid.createWriteContents(inid.datas)
    
    If sheet Is Nothing Then
        ' �V�[�g�I�u�W�F�N�g������������Ă��Ȃ��ꍇ�́A���������݂�
        initSheet
    End If
    
    ' �w�b�_�Ƃ��ăt�@�C�������L��
    sheet.Cells(ConstantsIniData.ROW_FILE_NAME, fileNameColumn) = fileName
    
    writeContents sheet, fileNameColumn, contents
        
    ' �������񂾂̂ŕύX�ς݂��I�t�ɂ���
    inid.isModified = False

End Sub

' =========================================================
' ���f�[�^�����݂��邩���m�F���郁�\�b�h
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�FTrue �f�[�^�����݂���AFalse �f�[�^�����݂��Ȃ�
'
' =========================================================
Public Function isExistsData() As Boolean

    isExistsData = False

    ' �V�����f�[�^�̏ꍇ
    Dim isNew As Boolean
    
    ' �V�[�g���������i�擾�����݂�j
    initSheet True

    If Not sheet Is Nothing Then
        ' �ΏۃV�[�g���擾�ł����ꍇ
        
        findColumnFilePath sheet, fileName, isNew
        If isNew = False Then
            ' ���Ƀf�[�^�����݂���
            isExistsData = True
        End If
        
    End If

End Function

' =========================================================
' ���ݒ���V�[�g�𐶐����郁�\�b�h
'
' �T�v�@�@�@�F�V�[�g�����ɑ��݂���ꍇ�͉������Ȃ��B���݂��Ȃ��ꍇ�͑}�����s���B
' �����@�@�@�Fbook      �u�b�N
'     �@�@�@�FsheetName �V�[�g��
'     �@�@�@�FisGetOnly �V�[�g�I�u�W�F�N�g�̎擾�̂݃t���O
' �߂�l�@�@�F�}�������V�[�g
'
' =========================================================
Private Function createSheet(ByVal book As Workbook, ByVal sheetName As String, Optional ByVal isGetOnly As Boolean = False) As Worksheet

    On Error Resume Next

    ' ------------------------------------------
    ' ��ʂ̍X�V��}�~����
    Application.screenUpdating = False

    Dim returnSheet As Worksheet
    Set returnSheet = ActiveSheet
    ' ------------------------------------------

    ' �ΏۃV�[�g
    Dim targetSheet     As Worksheet
    ' �ΏۃV�[�g���擾����
    Set targetSheet = book.Worksheets(sheetName)
    
    If isGetOnly = True Then
        ' �擾�̂�
        Set createSheet = targetSheet
        Exit Function
    End If

    ' �ΏۃV�[�g���擾�ł����ꍇ�i���X���݂��Ă��Ȃ��ꍇ������j
    If targetSheet Is Nothing Then
        Set targetSheet = book.Worksheets.Add
        targetSheet.name = sheetName
    End If

    ' ���V�V�[�g�Ƀe�[�u����`���𖄂ߍ���
    With targetSheet

        ' �V�[�g�����ɂ���
        .visible = xlSheetVisible

    End With
    
    ' ------------------------------------------
    ' ��ʂ̍X�V�}�~���������āA���O�̃V�[�g���A�N�e�B�u�ɂ�����
    returnSheet.activate
    Application.screenUpdating = True
    ' ------------------------------------------
    
    On Error GoTo 0

    If targetSheet Is Nothing Then
        ' �V�[�g�̍쐬�Ɏ��s�����ꍇ
        err.Raise ConstantsError.ERR_NUMBER_CREATE_WORKSHEET_FAILED, err.Source, ConstantsError.ERR_DESC_CREATE_WORKSHEET_FAILED, err.HelpFile, err.HelpContext
    End If

    Set createSheet = targetSheet

End Function

' =========================================================
' ���t�@�C���������ԍ�����������B
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet    �V�[�g��
'     �@�@�@�FfileName �t�@�C����
'     �@�@�@�FoutIsNew �i�߂�l�j�V�K���ǂ����̃t���O
' �߂�l�@�@�F��ԍ�
'
' =========================================================
Private Function findColumnFilePath(ByVal sheet As Worksheet, ByVal fileName As String, ByRef outIsNew As Boolean) As Long

    outIsNew = False
    findColumnFilePath = 0

    ' ���ɐݒ�ς݂̏ꍇ�ɂ�Find�Ō�������
    Dim findRange As Range
    Set findRange = sheet.Range(ConstantsIniData.ROW_FILE_NAME & ":" & ConstantsIniData.ROW_FILE_NAME).Find(what:=fileName, MatchCase:=False, MatchByte:=True, SearchFormat:=False)
    
    If Not findRange Is Nothing Then
        ' ���������ꍇ
        findColumnFilePath = findRange.column
        Exit Function
    End If
    
    ' ���݂��Ȃ��ꍇ�́A�V�����ꏊ�̗�ԍ���ԋp����
    findColumnFilePath = 1
    Do While sheet.Cells(2, findColumnFilePath) <> ""
        findColumnFilePath = findColumnFilePath + 2
    Loop
    
    outIsNew = True

End Function

' =========================================================
' ���R���e���c��ǂݍ��ށB
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet    �V�[�g��
'     �@�@�@�Fcolumn   ��ԍ�
' �߂�l�@�@�F�R���e���c
'
' =========================================================
Private Function readContents(ByVal sheet As Worksheet, ByVal column As Long) As String

    Dim ret As New StringBuilder

    If sheet.Cells(ConstantsIniData.ROW_FIRST, column) = "" Then
        ' �擪�������͂̏ꍇ�́A��̓��e�Ƃ݂Ȃ�
        readContents = ret.str
        Exit Function
    End If
    
    ' �ŏI���͍s���擾
    Dim lastInputtedRow As Long
    lastInputtedRow = ExcelUtil.getCellEndOfLastInputtedRow(sheet, column)

    Dim i As Long
    Dim findContentCells As Variant

    ' �f�[�^�͈͂��擾����
    Dim findRange As Range
    Set findRange = sheet.Range(sheet.Cells(ConstantsIniData.ROW_FIRST, column), sheet.Cells(lastInputtedRow, column + 1))

    If Not findRange Is Nothing Then
    
        ' �����͈͂���z����擾����
        findContentCells = ExcelUtil.copyCellsToArray(sheet, ConstantsIniData.ROW_FIRST, findRange.End(xlDown).row - findRange.row + 1, column, 2)
        
        ' �z�񂩂�R���e���c���擾����
        For i = LBound(findContentCells, 1) To UBound(findContentCells, 1)
            
            If Mid(findContentCells(i, 1), 1, 1) = "[" And _
               Mid(findContentCells(i, 1), Len(findContentCells(i, 1)), 1) = "]" And _
               findContentCells(i, 2) = "" Then
               ' �Z�N�V�����̏ꍇ
                ret.append findContentCells(i, 1) & vbNewLine
               
            Else
                ' �ʏ�̒l
                ret.append findContentCells(i, 1) & "=" & findContentCells(i, 2) & vbNewLine
            
            End If

        Next
        
    End If
    
    readContents = ret.str

End Function

' =========================================================
' ���R���e���c���������ށB
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet    �V�[�g��
'     �@�@�@�Fcolumn   ��ԍ�
'     �@�@�@�Fcontents �R���e���c
' �߂�l�@�@�F
'
' =========================================================
Private Sub writeContents(ByVal sheet As Worksheet, ByVal column As Long, ByRef contents As String)

    If sheet.Cells(ConstantsIniData.ROW_FIRST, column) = "" Then
        ' �擪�������͂̏ꍇ�́A��̓��e�Ƃ݂Ȃ�
    
    Else
        
        ' �ŏI���͍s���擾
        Dim lastInputtedRow As Long
        lastInputtedRow = ExcelUtil.getCellEndOfLastInputtedRow(sheet, column)

        ' ���ɐݒ�ς݂̏ꍇ�ɂ́A��U�Z���̓��e����ɂ���
        Dim findRange As Range
        Set findRange = sheet.Range(sheet.Cells(ConstantsIniData.ROW_FIRST, column), sheet.Cells(lastInputtedRow, column + 1))
        findRange.delete xlShiftUp
    
    End If
    
    If contents = "" Then
        Exit Sub
    End If

    Dim i As Long
    Dim contentsArray As Variant
    Dim contentsKeyValueArray As Variant
    
    ' �R���e���c�����s��؂�ɂ���i��s���Z���ɏ������݂����̂Łj
    contentsArray = Split(contents, vbNewLine)
    
    Dim contentsCells() As Variant
    ReDim contentsCells(0 To UBound(contentsArray) - LBound(contentsArray), 0 To 1)
    
    For i = LBound(contentsArray) To UBound(contentsArray)
    
        ' key=value�`���Ȃ̂Ŕz��ɕ�������i�������Z�N�V�����̏ꍇ������̂ŕK������key=value�ł͂Ȃ��j
        contentsKeyValueArray = Split(contentsArray(i), "=")
        
        If contentsArray(i) = "" Then
            ' contentsArray(i)��""�̏ꍇ�́ASplit�֐��Ɏ��s���Ă���̂ŉ������Ȃ�
            
        ElseIf VBUtil.arraySize(contentsKeyValueArray) > 1 Then
            ' key=value�`��
            contentsCells(i, 0) = contentsKeyValueArray(0)
            contentsCells(i, 1) = contentsKeyValueArray(1)
        Else
            ' key�̂݁i�܂��̓Z�N�V�����j
            contentsCells(i, 0) = contentsKeyValueArray(0)
            contentsCells(i, 1) = ""
        End If
        
    Next
    
    ' ���f�O�ɏ����𕶎���ɂ���
    ExcelUtil.getArrayRange(contentsCells, sheet, ConstantsIniData.ROW_FIRST, column).NumberFormatLocal = "@"
    ' �Z���ɔ��f����
    ExcelUtil.copyArrayToCells contentsCells, sheet, ConstantsIniData.ROW_FIRST, column

End Sub

' ------------------------------------------------------------------
' IniData�����b�v���Ċe���\�b�h���Ăяo��
' ------------------------------------------------------------------
Public Function getValue _
            (ByVal SECTION As String, _
             ByVal key As String) As String

    getValue = inid.getValue(SECTION, key)
End Function

Public Sub setValue _
            (ByVal SECTION As String, _
             ByVal key As String, _
             ByVal value As String)

    inid.setValue SECTION, key, value
End Sub

Public Function getValues _
            (ByVal SECTION As String) As ValCollection

    Set getValues = inid.getValues(SECTION)
End Function

Public Sub setValues _
            (ByVal SECTION As String, _
             ByRef values As ValCollection)

    inid.setValues SECTION, values
End Sub

Public Sub delete _
            (ByVal SECTION As String, _
             Optional ByVal key As String = "")

    inid.delete SECTION, key
End Sub
