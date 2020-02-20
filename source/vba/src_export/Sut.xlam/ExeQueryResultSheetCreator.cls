VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeQueryResultSheetCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �N�G�����ʃV�[�g�𐶐�����N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2013/06/23 �V�K�쐬
'
' ���L�����F
' *********************************************************

' =========================================================
' �v���p�e�B
' =========================================================
' �A�v���P�[�V�����ݒ���
Private m_applicationSetting As ValApplicationSetting
' �����Ώۃ��[�N�u�b�N
Private m_book               As Workbook
' �����Ώۃ��[�N�u�b�N
Private m_sheet              As Worksheet

Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
End Property

Public Property Let sheet(ByVal vNewValue As Variant)

    Set m_sheet = vNewValue
    
End Property

Public Property Let book(ByVal vNewValue As Variant)

    Set m_book = vNewValue
    
End Property

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ���N�G�����ʃV�[�g����
'
' �T�v�@�@�@�F�ΏۃV�[�g��
' �@�@�@�@�@�@�N�G�����ʃV�[�g�ł��邩�𔻒肷��B
'
' �����@�@�@�Fsheet �ΏۃV�[�g
' �߂�l�@�@�FTrue �e�[�u���V�[�g�̏ꍇ
' =========================================================
Public Function isQueryResultSheet(ByRef sheet As Worksheet) As Boolean

    ' �e�[�u���V�[�g����p�̃I�[�g�V�F�C�v
    Dim s As Shape

    ' �߂�l������������
    isQueryResultSheet = False
    
    ' ------------------------------------------------
    ' �G���[���������Ă��p������
    On Error Resume Next

    Set s = sheet.Shapes(ConstantsTable.QUERY_RESULT_SHEET_CHECK_IMAGE)
    
    ' �擾�ł����ꍇ
    If Not s Is Nothing Then
    
        isQueryResultSheet = True
        Exit Function
    End If
    
    On Error GoTo 0
    ' ------------------------------------------------

End Function

' =========================================================
' ���V�[�g�𐶐�����
'
' �T�v�@�@�@�F�V�[�g�𐶐�����
' �����@�@�@�F
'
' =========================================================
Public Function createSheet() As Worksheet

    Dim sheetName As String
    sheetName = ExcelUtil.convertUniqSheetName(m_book, ConstantsTable.QUERY_RESULT_SHEET_DEFAULT_NAME)
    
    ' �e�[�u����
    Dim tableName     As String
    ' �V�[�g�I�u�W�F�N�g
    Dim sheet         As Worksheet
    
    ' �V�[�g�̍폜�i���݂��Ă�����j�Ƒ}�������s����
    deleteAndInsertSheet sheetName
    
    ' �V���ɃR�s�[���ꂽ�V�[�g�I�u�W�F�N�g���擾����
    Set sheet = m_book.Worksheets(sheetName)
    
    ' ���V�V�[�g�Ƀe�[�u����`���𖄂ߍ���
    With sheet

        ' �V�[�g�����ɂ���
        .visible = xlSheetVisible

    End With

    Set m_sheet = sheet
    Set m_book = sheet.parent
    Set createSheet = sheet
    
End Function

' =========================================================
' ���V�[�g�̍폜�E�}��
'
' �T�v�@�@�@�F�V�[�g�����ɑ��݂���ꍇ�͍폜���A�}�����s���B
' �����@�@�@�FtableSheet      �e�[�u�����
' �@�@�@�@�@�@targetSheetName �ΏۃV�[�g��
'
' =========================================================
Private Sub deleteAndInsertSheet(ByRef targetSheetName As String)

    ' �R�s�[�V�[�g
    Dim copySheetName As String
    copySheetName = ConstantsTable.QUERY_RESULT_SHEET_NAME_TEMPLATE

    ' �폜���ꂽ�V�[�g�̈ʒu�ɍĂуV�[�g���쐬�������̂�
    ' �폜�ΏۃV�[�g�̈ʒu�𒲂ח׍��킹�ɂȂ��Ă���V�[�g�I�u�W�F�N�g���擾����B
    ' �i�V�[�g�̃R�s�[�֐���
    ' �@����V�[�g����_�Ƃ��ăR�s�[����Ƃ��������@�ɂ����Ή����Ă��Ȃ����߂��̂悤�ȏ������K�v�ɂȂ�j
    On Error Resume Next

    ' �폜�ΏۃV�[�g
    Dim targetSheet     As Worksheet
    ' �폜�ΏۃV�[�g�̍��ׂɂ���V�[�g
    Dim targetSheetPrev As Worksheet
    ' �폜�ΏۃV�[�g�̉E�ׂɂ���V�[�g
    Dim targetsheetNext As Worksheet

    ' �폜�ΏۃV�[�g���擾����
    Set targetSheet = m_book.Worksheets(targetSheetName)

    ' �폜�ΏۃV�[�g���擾�ł����ꍇ�i���X���݂��Ă��Ȃ��ꍇ������j
    If Not targetSheet Is Nothing Then

        ' �u�b�N���ɃV�[�g�̐���1�����̏ꍇ�i�폜�ΏۃV�[�g�܂߂āj
        If m_book.Worksheets.count = 1 Then
        
            ' �������Ȃ�
        ' �폜�ΏۃV�[�g�����[�̏ꍇ
        ElseIf targetSheet.index = 1 Then

            Set targetsheetNext = m_book.Worksheets(2)

        ' �폜�ΏۃV�[�g���E�[�̏ꍇ
        ElseIf targetSheet.index = m_book.Worksheets.count Then

            Set targetSheetPrev = m_book.Worksheets(targetSheet.index - 1)

        ' �폜�ΏۃV�[�g�����ԁi�[�ɂ͂Ȃ��j
        Else
        
            Set targetSheetPrev = m_book.Worksheets(targetSheet.index - 1)
        End If


    End If

    ' Excel�ł̓V�[�g�̐���0�ɂȂ�̂������Ȃ��̂ŁA���̏�ԂŃV�[�g���폜�����
    ' �_�~�[�̃V�[�g������ɑ}�����Ă��܂��B
    ' ���������āA�Ώۂ̃u�b�N�̃V�[�g����1�̏ꍇ�́A�O�����ăV�[�g��}�����Ă����A�����I����ɍ폜����悤�ɂ���
    Dim dummySheetName As String
    
    If m_book.Worksheets.count = 1 Then
    
        m_book.Worksheets.Add
        dummySheetName = m_book.ActiveSheet.name
    End If
    
    ' �V�[�g�폜���s
    ExcelUtil.deleteSheet m_book, targetSheetName
    
    ' ���ׂ̃V�[�g���擾�ł��Ă���ꍇ
    If Not targetSheetPrev Is Nothing Then
    
        ExcelUtil.copySheet SutWorkbook _
                          , copySheetName _
                          , m_book _
                          , targetSheetName _
                          , targetSheetPrev.name _
                          , "after"
    
    ' �E�ׂ̃V�[�g���擾�ł��Ă���ꍇ
    ElseIf Not targetsheetNext Is Nothing Then
    
        ExcelUtil.copySheet SutWorkbook _
                          , copySheetName _
                          , m_book _
                          , targetSheetName _
                          , targetsheetNext.name _
                          , "before"
    
    ' �폜�ΏۃV�[�g�����݂��Ă��Ȃ��ꍇ
    Else
    
        ExcelUtil.copySheetAppend SutWorkbook _
                                , copySheetName _
                                , m_book _
                                , targetSheetName
        
    
    End If
    
    ' �_�~�[�V�[�g���}������Ă���ꍇ�A�_�~�[�V�[�g���폜����
    If dummySheetName <> "" Then
    
        m_book.Worksheets(dummySheetName).delete
    End If
    
    On Error GoTo 0

End Sub

' =========================================================
' ����̓��͗̈��}��
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@columnOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�F
' =========================================================
Public Sub insertEmptyCellForResultCount(ByVal recordOffset As Long, ByVal columnOffset As Long)

    ' ------------------------------------------------------------
    ' �e���v���[�g�V�[�g�̃��R�[�h�̈���R�s�[����i���������R�s�[�j
    ' ------------------------------------------------------------
    ' �R�s�[���Z��
    Dim srcRange As Range
    ' �R�s�[��Z��
    Dim desRange As Range

    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_RESULT_ROW _
                          , ConstantsTable.QUERY_RESULT_COL)
                  

    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange
    
    ' �������R�s�[����
    ExcelUtil.changeRowHeight desRange, srcRange.RowHeight
    
End Sub
' =========================================================
' ����̓��͗̈��}��
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@columnOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�F
' =========================================================
Public Sub insertEmptyCellForError(ByVal recordOffset As Long, ByVal columnOffset As Long)

    ' ------------------------------------------------------------
    ' �e���v���[�g�V�[�g�̃��R�[�h�̈���R�s�[����i���������R�s�[�j
    ' ------------------------------------------------------------
    ' �R�s�[���Z��
    Dim srcRange As Range
    ' �R�s�[��Z��
    Dim desRange As Range

    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_ERROR_ROW _
                          , ConstantsTable.QUERY_ERROR_COL)

    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange
    
    ' �������R�s�[����
    ExcelUtil.changeRowHeight desRange, srcRange.RowHeight

End Sub

' =========================================================
' ����̓��͗̈��}��
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@columnOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�F
' =========================================================
Public Sub insertEmptyCellForTitle(ByVal recordOffset As Long, ByVal columnOffset As Long)

    ' ------------------------------------------------------------
    ' �e���v���[�g�V�[�g�̃��R�[�h�̈���R�s�[����i���������R�s�[�j
    ' ------------------------------------------------------------
    ' �R�s�[���Z��
    Dim srcRange As Range
    ' �R�s�[��Z��
    Dim desRange As Range

    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_TITLE_ROW _
                          , ConstantsTable.QUERY_TITLE_COL)
                  

    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange
    
    ' �������R�s�[����
    ExcelUtil.changeRowHeight desRange, srcRange.RowHeight
    
End Sub

' =========================================================
' ����̓��͗̈��}��
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@columnOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�F
' =========================================================
Public Sub insertEmptyCellForResultSet(ByVal recordOffset As Long, ByVal columnOffset As Long)

    ' ------------------------------------------------------------
    ' �e���v���[�g�V�[�g�̃��R�[�h�̈���R�s�[����i���������R�s�[�j
    ' ------------------------------------------------------------
    ' �R�s�[���Z��
    Dim srcRange As Range
    ' �R�s�[��Z��
    Dim desRange As Range

    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_RESULTSET_ROW _
                          , ConstantsTable.QUERY_RESULTSET_COL)
                  

    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange
    
    ' �������R�s�[����
    ExcelUtil.changeRowHeight desRange, srcRange.RowHeight
    
End Sub

' =========================================================
' ����̓��͗̈��}��
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@recordSize   ���R�[�h�T�C�Y
' �@�@�@�@�@�@columnOffset �J�������f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@columnSize   �J�����T�C�Y
' �@�@�@�@�@�@isGroup      �O���[�v���̗L��
'
' �߂�l�@�@�F
' =========================================================
Public Sub insertEmptyCellForRecord(ByVal recordOffset As Long, ByVal recordSize As Long, ByVal columnOffset As Long, ByVal columnSize As Long _
                                    , Optional ByVal isGroup As Boolean = True)

    ' �J�����̐�
    Dim columnCount As Long
    ' �V�[�g�̍ő��
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(m_sheet)
    
    ' �J�����̐����V�[�g�̍ő�񐔂𒴂��Ă��Ȃ������`�F�b�N����
    If columnSize > sheetColSize - columnOffset + 1 Then
    
        ' �ő�񐔂𒴂��Ă���ꍇ�A�񐔂𒲐�����
        columnCount = sheetColSize - columnOffset + 1
        
    Else
    
        ' �ő�񐔂𒴂��Ă��Ȃ��ꍇ�A���̂܂ܐݒ�
        columnCount = columnSize
    End If

    ' ------------------------------------------------------------
    ' �e���v���[�g�V�[�g�̃��R�[�h�̈���R�s�[����i���������R�s�[�j
    ' ------------------------------------------------------------
    ' �R�s�[���Z��
    Dim srcRange As Range
    ' �R�s�[��Z��
    Dim desRange As Range

    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_RECORD_ROW _
                          , ConstantsTable.QUERY_RECORD_COL)
                  
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset + recordSize - 1 _
                                    , columnOffset + columnCount - 1).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange

    ' ------------------------------------------------------------
    ' �Z�������̐ݒ�
    ' ------------------------------------------------------------
    ' �t�H���g��ݒ�
    desRange.Font.name = m_applicationSetting.cellFontName
    ' �t�H���g�T�C�Y��ݒ�
    desRange.Font.size = m_applicationSetting.cellFontSize
    ' �܂�Ԃ���ݒ�
    desRange.WrapText = m_applicationSetting.cellWordwrap
    ' �񕝂�ݒ�
    ExcelUtil.changeColWidth desRange, m_applicationSetting.cellWidth
    ' �s����ݒ�
    If m_applicationSetting.lineHeightAutoAdjust Then
        ExcelUtil.changeRowHeight desRange, -1
    Else
        ExcelUtil.changeRowHeight desRange, m_applicationSetting.cellHeight
    End If
    

    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_ROWNUMBER_ROW _
                          , ConstantsTable.QUERY_ROWNUMBER_COL)
                  
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                     , ConstantsTable.QUERY_COLUMN_TITLE_COL).Address & ":" & _
                        m_sheet.Cells(recordOffset + recordSize - 1 _
                                    , ConstantsTable.QUERY_COLUMN_TITLE_COL).Address)

    ' ���R�[�h�ԍ����R�s�[����
    srcRange.copy Destination:=desRange

    desRange.Formula = "=ROW() - " & (recordOffset - 1)
    
    If isGroup Then
        ' �O���[�v������
        desRange.Rows.Group
    End If

End Sub


' =========================================================
' ����̓��͗̈��}��
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@recordSize   ���R�[�h�T�C�Y
' �@�@�@�@�@�@columnOffset �J�������f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@columnSize   �J�����T�C�Y
'
' �߂�l�@�@�F
' =========================================================
Public Sub insertEmptyCellForColumn(ByVal recordOffset As Long, ByVal columnOffset As Long, ByVal columnSize As Long)

    ' �J�����̐�
    Dim columnCount As Long
    ' �V�[�g�̍ő��
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(m_sheet)
    
    ' �J�����̐����V�[�g�̍ő�񐔂𒴂��Ă��Ȃ������`�F�b�N����
    If columnSize > sheetColSize - columnOffset + 1 Then
    
        ' �ő�񐔂𒴂��Ă���ꍇ�A�񐔂𒲐�����
        columnCount = sheetColSize - columnOffset + 1
        
    Else
    
        ' �ő�񐔂𒴂��Ă��Ȃ��ꍇ�A���̂܂ܐݒ�
        columnCount = columnSize
    End If

    ' ------------------------------------------------------------
    ' �e���v���[�g�V�[�g�̃��R�[�h�̈���R�s�[����i���������R�s�[�j
    ' ------------------------------------------------------------
    ' �R�s�[���Z��
    Dim srcRange As Range
    ' �R�s�[��Z��
    Dim desRange As Range

    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsTable.QUERY_SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsTable.QUERY_HEADER_ROW _
                          , ConstantsTable.QUERY_HEADER_COL)
                  
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = m_sheet _
                     .Range( _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset).Address & ":" & _
                        m_sheet.Cells(recordOffset _
                                    , columnOffset + columnCount - 1).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange

    ' ------------------------------------------------------------
    ' �Z�������̐ݒ�
    ' ------------------------------------------------------------
    ' �t�H���g��ݒ�
    desRange.Font.name = m_applicationSetting.cellFontName
    ' �t�H���g�T�C�Y��ݒ�
    desRange.Font.size = m_applicationSetting.cellFontSize
    ' �܂�Ԃ���ݒ�
    desRange.WrapText = m_applicationSetting.cellWordwrap
    ' �񕝂�ݒ�
    ExcelUtil.changeColWidth desRange, m_applicationSetting.cellWidth
    ' �s����ݒ�
    If m_applicationSetting.lineHeightAutoAdjust Then
        ExcelUtil.changeRowHeight desRange, -1
    Else
        ExcelUtil.changeRowHeight desRange, m_applicationSetting.cellHeight
    End If

End Sub

' =========================================================
' �����R�[�h���X�g�̃V�[�g���f�i���R�[�h�p�j
'
' �T�v�@�@�@�F���R�[�h���X�g���V�[�g�ɔ��f����
' �����@�@�@�FrecordList   ���R�[�h���X�g
' �@�@�@�@�@�@recordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�FTrue  ����
' �@�@�@�@�@�@False �V�[�g�ւ̔��f���s
'
' =========================================================
Public Function reflectSheetForRecord(ByRef recordList As Variant _
                                    , ByVal recordOffset As Long) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = True

    ' ���R�[�h�T�C�Y
    Dim rowSize As Long
    ' �J�����T�C�Y
    Dim colSize As Long

    ' �z�񂩂烌�R�[�h�T�C�Y���擾����
    rowSize = VBUtil.arraySize(recordList)
    ' �z�񂩂�J�����T�C�Y���擾����
    colSize = VBUtil.arraySize(recordList, 2)

    ' ------------------------------------------------------------
    ' �V�[�g�͈̔̓`�F�b�N
    ' ------------------------------------------------------------
    If ExcelUtil.checkOverMaxRow(m_sheet, recordOffset, rowSize) = False Then
    
        ' �T�C�Y���I�[�o�[���Ă���̂Œ�������
        rowSize = ExcelUtil.getSizeOfSheetRow(m_sheet) - recordOffset + 1
    
        ' �S���R�[�h���f�ł��Ȃ��̂Ŗ߂�l��False�ɐݒ�
        ret = False
    End If
        
    ' ------------------------------------------------------------
    ' �z��̓��e���V�[�g�ɔ��f����
    ' ------------------------------------------------------------
    ExcelUtil.copyArrayToCells recordList, m_sheet, recordOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL, rowSize, colSize

    reflectSheetForRecord = ret
    
    ' �߂�l�� False �̏ꍇ�A�G���[�𔭍s����
    If ret = False Then
    
        err.Raise ERR_NUMBER_OUT_OF_RANGE_SHEET _
                , _
                , ERR_DESC_OUT_OF_RANGE_SHEET
    End If

End Function


' =========================================================
' �����R�[�h���X�g�̃V�[�g���f�i�J�����p�j
'
' �T�v�@�@�@�F���R�[�h���X�g���V�[�g�ɔ��f����
' �����@�@�@�FheaderList   ���R�[�h���X�g
' �@�@�@�@�@�@headerOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�FTrue  ����
' �@�@�@�@�@�@False �V�[�g�ւ̔��f���s
'
' =========================================================
Public Function reflectSheetForColumn(ByRef headerList As Variant _
                                    , ByVal headerOffset As Long) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = True

    ' ���R�[�h�T�C�Y
    Dim rowSize As Long: rowSize = 1
    ' �J�����T�C�Y
    Dim colSize As Long

    ' �z�񂩂�J�����T�C�Y���擾����
    colSize = VBUtil.arraySize(headerList)

    ' ------------------------------------------------------------
    ' �V�[�g�͈̔̓`�F�b�N
    ' ------------------------------------------------------------
    If ExcelUtil.checkOverMaxCol(m_sheet, headerOffset, colSize) = False Then
    
        ' �T�C�Y���I�[�o�[���Ă���̂Œ�������
        colSize = ExcelUtil.getSizeOfSheetCol(m_sheet) - headerOffset + 1
    
        ' �S���R�[�h���f�ł��Ȃ��̂Ŗ߂�l��False�ɐݒ�
        ret = False
    End If

    ' ------------------------------------------------------------
    ' �z��̓��e���V�[�g�ɔ��f����
    ' ------------------------------------------------------------
    ExcelUtil.copyArrayToCellsForColumns headerList, m_sheet, headerOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL, colSize

    reflectSheetForColumn = ret

    
    ' �߂�l�� False �̏ꍇ�A�G���[�𔭍s����
    If ret = False Then
    
        err.Raise ERR_NUMBER_OUT_OF_RANGE_SHEET _
                , _
                , ERR_DESC_OUT_OF_RANGE_SHEET
    End If

End Function

' =========================================================
' �����R�[�h�̃I�t�Z�b�g�ʒu�擾
'
' �T�v�@�@�@�F�擾�������R�[�h�𔽉f����I�t�Z�b�g�ʒu���擾����B
'
' �����@�@�@�FappendSheet �Ō���ւ̒ǉ��L��
' �߂�l�@�@�F���R�[�h�̃I�t�Z�b�g�ʒu
'
' =========================================================
Public Function getRecordOffset(ByVal appendSheet As Boolean) As Long

    ' �߂�l
    Dim recordOffset As Long
    
    ' �����̍s���폜�����ɁA�Ō�̈ʒu���烌�R�[�h��ǉ�����
    If appendSheet = True Then
        
        ' �ŏI���͍s���擾���{�P����
        recordOffset = ExcelUtil.getCellEndOfLastInputtedRow(m_sheet, ConstantsTable.QUERY_COLUMN_TITLE_COL) + 1
        
        ' ��L�Ŏ擾�����l���V�[�g�͈̔͂𒴂��Ă��邩���`�F�b�N����
        If ExcelUtil.checkOverMaxRow(m_sheet, recordOffset) = False Then
        
            ' �G���[�𔭍s����
            err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SHEET _
                    , _
                    , ConstantsError.ERR_DESC_OUT_OF_RANGE_SHEET
        End If
                
    ' �����̍s���폜���āA�ŏ��̈ʒu���烌�R�[�h��ǉ�����
    Else
    
        ' �萔���烌�R�[�h�I�t�Z�b�g���擾����
        recordOffset = ConstantsTable.QUERY_COLUMN_OFFSET_ROW
    
    End If

    ' �߂�l��ݒ肷��
    getRecordOffset = recordOffset

End Function
