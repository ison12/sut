VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSnapDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�i�b�v�V���b�g�̔�r���ʂ��o�͂���N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/01/06�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ________________________________________________________
' �C�x���g
' ________________________________________________________

' =========================================================
' ���V�[�g�쐬�C�x���g�i�J�n�j
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event createSnapshotDiffSheetBegin()

' =========================================================
' ���V�[�g�쐬�C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet �V�[�g
'
' =========================================================
Public Event createSnapshotDiffSheet(ByRef sheet As Worksheet)

' =========================================================
' �����R�[�h�Z�b�g�̔�r�C�x���g�i�J�n�j
'
' �T�v�@�@�@�F
' �����@�@�@�Fcount ����
'
' =========================================================
Public Event diffRecordSetBegin(ByVal count As Long)

' =========================================================
' �����R�[�h�Z�b�g�̔�r�C�x���g�i1�������j
'
' �T�v�@�@�@�F
' �����@�@�@�Fcount ����
'
' =========================================================
Public Event diffRecordSetOne()

' =========================================================
' �����R�[�h�Z�b�g�̔�r�C�x���g�i�����j
'
' �T�v�@�@�@�F
' �����@�@�@�Fcount ����
'
' =========================================================
Public Event diffRecordSet()

' ---------------------------------------------------------
' ���v���p�e�B
' ---------------------------------------------------------
' �L�����Z���t���O
Public canceled      As Boolean

' �����Ώۃ��[�N�u�b�N
Private m_book       As Workbook
' �A�v���P�[�V�����ݒ���
Private m_applicationSetting As ValApplicationSetting

' =========================================================
' ���v���p�e�B �u�b�N
' =========================================================
Public Property Let book(ByVal vNewValue As Variant)

    Set m_book = vNewValue
End Property

' =========================================================
' ���A�v���P�[�V�����ݒ���
' =========================================================
Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
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
' �����s
'
' �T�v�@�@�@�F�ΏۂƂȂ�N�G���̎��s���ʂ��V�[�g�ɔ��f����
' �����@�@�@�Fdiff1  �N�G�����X�g
' �@�@  �@�@�Fdiff2  �N�G�����X�g
' �߂�l�@�@�F���R�[�h�擾����
'
' =========================================================
Public Function execute(ByRef diff1 As ValSnapRecordsSet, ByRef diff2 As ValSnapRecordsSet) As Long

    On Error GoTo err
    
    Dim diffSheet As Worksheet
    
    If diff1.list.count <> diff2.list.count Then
        ' ���ʃZ�b�g�̌������Ⴄ�̂̓G���[�i�������Ȃ��͂��j
        err.Raise _
            ConstantsError.ERR_NUMBER_SNAP_DIFF__EXEC_ERROR, _
            err.Source, _
            ConstantsError.ERR_DESC_SNAP_DIFF__EXEC_ERROR & vbNewLine & _
            "���ʃZ�b�g�̌������قȂ邽�߃G���[�ɂȂ�܂����B", _
            err.HelpFile, _
            err.HelpContext
    End If
    
    ' �V�[�g�𐶐�����
    Set diffSheet = createSheet
    
    ' �L�����Z���t���O���`�F�b�N����
    checkCancel
    
    ' ��r���s
    diffRecordsSet diffSheet, diff1, diff2

    Exit Function

err:

    ' �G���[�n���h���ŕʂ̊֐����Ăяo���ƃG���[��񂪏����Ă��܂����Ƃ�����̂�
    ' �\���̂ɃG���[����ۑ����Ă���
    Dim errT As errInfo: errT = VBUtil.swapErr

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' �����R�[�h�Z�b�g��r
'
' �T�v�@�@�@�F���R�[�h�Z�b�g�̔�r�����s����
' �����@�@�@�Fsheet  �V�[�g
'           �Fdiff1  ��r�Ώ�1
' �@�@  �@�@�Fdiff2  ��r�Ώ�2
' �߂�l�@�@�F���R�[�h�擾����
'
' =========================================================
Private Function diffRecordsSet(ByRef sheet As Worksheet, ByRef diff1 As ValSnapRecordsSet, ByRef diff2 As ValSnapRecordsSet)

    Dim recordSetIndex As Long
    
    Dim diff1Rec As ValSnapRecords
    Dim diff2Rec As ValSnapRecords
    
    Dim offsetRow As Long
    offsetRow = ConstantsSnapDiff.RESULT_START_ROW
    
    ' �����R�[�h�Z�b�g�̔�r�C�x���g�J�n
    RaiseEvent diffRecordSetBegin(diff1.list.count)
    
    ' ���O���Z���ɏ�������
    sheet.Cells(ConstantsSnapDiff.NAME_ROW, ConstantsSnapDiff.NAME_COL) = diff1.name
    ' �S�̂̕ύX�L�����Z���ɏ�������
    setModifyCell sheet, ConstantsSnapDiff.MODIFY_ALL_ROW, ConstantsSnapDiff.MODIFY_ALL_COL, False ' �S��

    recordSetIndex = 0
    For Each diff1Rec In diff1.list.col
    
        ' �L�����Z���t���O���`�F�b�N����
        checkCancel
    
        Set diff2Rec = diff2.list.getItemByIndex(recordSetIndex + 1)
        
        offsetRow = diffRecords(recordSetIndex, sheet, offsetRow, diff1Rec, diff2Rec)

        recordSetIndex = recordSetIndex + 1
        
        ' �����R�[�h�Z�b�g�̔�r�C�x���g
        RaiseEvent diffRecordSetOne

    Next
        
    ' �����R�[�h�Z�b�g�̔�r�C�x���g����
    RaiseEvent diffRecordSet

End Function

' =========================================================
' �����R�[�h���X�g��r
'
' �T�v�@�@�@�F���R�[�h�Z�b�g�̔�r�����s����iSQL�̎��s�P�ʁj
' �����@�@�@�FrecordSetIndex  ���R�[�h�Z�b�g�J�n�s
'           �Fsheet           �V�[�g
' �@�@  �@�@�FoffsetRow       �J�n�s
' �@�@  �@�@�Fdiff1Rec        ��r�Ώ�1
' �@�@  �@�@�Fdiff2Rec        ��r�Ώ�2
' �߂�l�@�@�F���R�[�h�擾����
'
' =========================================================
Private Function diffRecords(ByVal recordSetIndex As Long, ByRef sheet As Worksheet, ByVal offsetRow As Long, ByRef diff1Rec As ValSnapRecords, ByRef diff2Rec As ValSnapRecords) As Long

    Dim totalCount    As Long
    Dim noChangeCount As Long
    Dim insertCount   As Long
    Dim updateCount   As Long
    Dim deleteCount   As Long

    Dim offsetRowHeader As Long: offsetRowHeader = offsetRow + ConstantsSnapDiff.RESULT_START_OFFSET_HEADER
    Dim offsetRowReocrd As Long: offsetRowReocrd = offsetRow + ConstantsSnapDiff.RESULT_START_OFFSET_RECORD

    Dim endOffsetRow As Long

    Dim header1Index As Long
    Dim header2Index As Long
    
    Dim recordIndex As Long
    Dim recordNum As Long
    
    Dim pkStr As String

    Dim diff1RecMap As ValCollection
    Dim diff2RecMap As ValCollection
    
    Dim diffDeletedRecord As Variant
    
    Dim diff2RecIndex As Long
    Dim headerIndex As Long
    
    Dim isUpdate As Boolean
    Dim isOutRecord As Boolean

    diffRecords = offsetRow

    If UBound(diff1Rec.headerList) - LBound(diff1Rec.headerList) <> _
       UBound(diff2Rec.headerList) - LBound(diff2Rec.headerList) Then
        ' �w�b�_�̌������Ⴄ�̂̓G���[�i�ʏ�͔������Ȃ��͂��j
        err.Raise _
            ConstantsError.ERR_NUMBER_SNAP_DIFF__EXEC_ERROR, _
            err.Source, _
            ConstantsError.ERR_DESC_SNAP_DIFF__EXEC_ERROR & vbNewLine & _
            "���ʃZ�b�g�̃w�b�_�̌������قȂ邽�߃G���[�ɂȂ�܂����B", _
            err.HelpFile, _
            err.HelpContext
    End If
    
    For header1Index = LBound(diff1Rec.headerList) To UBound(diff1Rec.headerList)
    
        If diff1Rec.headerList(header1Index) <> diff2Rec.headerList(header1Index) Then
            ' �w�b�_�����Ⴄ�̂̓G���[�i�ʏ�͔������Ȃ��͂��j
            err.Raise _
                ConstantsError.ERR_NUMBER_SNAP_DIFF__EXEC_ERROR, _
                err.Source, _
                ConstantsError.ERR_DESC_SNAP_DIFF__EXEC_ERROR & vbNewLine & _
                "���ʃZ�b�g�̃w�b�_�����قȂ邽�߃G���[�ɂȂ�܂����B", _
                err.HelpFile, _
                err.HelpContext
        End If
        
    Next
    
    ' --------------------------------------------
    ' ��r�͈͂��Z���ɏo�͂���
    ' --------------------------------------------
    insertResultRange sheet, offsetRow
    
    ' ��r�L��
    setModifyCell sheet, offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.MODIFY_COL, False ' ��
    
    ' --------------------------------------------
    ' �w�b�_���Z���ɏo�͂���
    ' --------------------------------------------
    ' ���R�[�h���X�g�̋󗓂̗̈���o��
    insertHeaderListCell sheet, offsetRowHeader, VBUtil.arraySize(diff1Rec.headerList)
    ' ���e���o��
    ExcelUtil.copyArrayToCellsForColumns diff1Rec.headerList, sheet, offsetRowHeader, ConstantsSnapDiff.HEADER_COL
    
    Set diff1RecMap = diff1Rec.convertToMapKeyIsPkValIsReocrdIndex
    Set diff2RecMap = diff2Rec.convertToMapKeyIsPkValIsReocrdIndex
    
    ' diff1����ɔ�r����
    If Not IsEmpty(diff1Rec.recordList) Then
    
        ' --------------------------------------------
        ' ���R�[�h���X�g���Z���ɏo�͂���
        ' --------------------------------------------
        ' ���R�[�h���X�g�̋󗓂̗̈���o��
        insertRecordListCell sheet, offsetRowReocrd, VBUtil.arraySize(diff1Rec.recordList), VBUtil.arraySize(diff1Rec.recordList, 2)
        ' ���e���o��
        ExcelUtil.copyArrayToCells diff1Rec.recordList, sheet, offsetRowReocrd, ConstantsSnapDiff.RECORD_COL
    
        totalCount = totalCount + VBUtil.arraySize(diff1Rec.recordList)
    
        recordIndex = 0
        For recordIndex = LBound(diff1Rec.recordList) To UBound(diff1Rec.recordList)
            
            pkStr = diff1Rec.getPkStr(recordIndex)
            
            If diff2RecMap.exist(pkStr) = False Then
                ' PK�ɊY�����郌�R�[�h�����݂��Ȃ��ꍇ
                ' INSERT
                sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_INSERT

                ' �w�i�F��ύX����
                sheet.Range( _
                    sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL).Address, _
                    sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL + VBUtil.arraySize(diff1Rec.headerList) - 1).Address _
                ).Interior.Color = _
                    RGB( _
                        ConstantsSnapDiff.INSERT_COLOR_R, _
                        ConstantsSnapDiff.INSERT_COLOR_G, _
                        ConstantsSnapDiff.INSERT_COLOR_B)
                        
                ' �ύX�L�����Z���ɏ�������
                setModifyCell sheet, ConstantsSnapDiff.MODIFY_ALL_ROW, ConstantsSnapDiff.MODIFY_ALL_COL, True ' �S��
                setModifyCell sheet, offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.MODIFY_COL, True ' ��
                
                insertCount = insertCount + 1
            Else
                ' PK�ɊY�����郌�R�[�h�����݂���ꍇ
                
                diff2RecIndex = diff2RecMap.getItem(pkStr, vbLong)
                
                isUpdate = False
                For headerIndex = LBound(diff1Rec.headerList) To UBound(diff1Rec.headerList)
                    
                    If diff1Rec.recordList(recordIndex, headerIndex) <> diff2Rec.recordList(diff2RecIndex, headerIndex) Then
                    
                        ' �l���قȂ�ꍇ�A�ύX�̑O�̒l���Z���ɃR�����g�Ƃ��Ė��ߍ���
                        ExcelUtil.addCommentForWorkSheet sheet, _
                            offsetRowReocrd, _
                            ConstantsSnapDiff.RECORD_COL + headerIndex, _
                            diff2Rec.recordList(diff2RecIndex, headerIndex)
                            
                        ' �w�i�F��ύX����
                        sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL + headerIndex).Interior.Color = _
                            RGB( _
                                ConstantsSnapDiff.UPDATE_COLOR_R, _
                                ConstantsSnapDiff.UPDATE_COLOR_G, _
                                ConstantsSnapDiff.UPDATE_COLOR_B)
                    
                        isUpdate = True
                    End If
                    
                Next
                
                If isUpdate Then
                    ' UPDATE
                    sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_UPDATE
                    ' �ύX�L�����Z���ɏ�������
                    setModifyCell sheet, ConstantsSnapDiff.MODIFY_ALL_ROW, ConstantsSnapDiff.MODIFY_ALL_COL, True ' �S��
                    setModifyCell sheet, offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.MODIFY_COL, True ' ��
                
                    updateCount = updateCount + 1
                Else
                    ' �ύX���Ȃ��ꍇ
                    sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_NOCHANGE
                    
                    noChangeCount = noChangeCount + 1
                End If
                
            End If
            
            ' �A��
            recordNum = recordNum + 1
            sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_NUM_COL) = recordNum
            
            offsetRowReocrd = offsetRowReocrd + 1
            
        Next
        
        isOutRecord = True
    
    End If
    
    ' diff2����ɔ�r����
    If Not IsEmpty(diff2Rec.recordList) Then
        
        recordIndex = 0
        For recordIndex = LBound(diff2Rec.recordList) To UBound(diff2Rec.recordList)
            
            pkStr = diff2Rec.getPkStr(recordIndex)
            
            If diff1RecMap.exist(pkStr) = False Then
                ' PK�ɊY�����郌�R�[�h�����݂��Ȃ��ꍇ
                ' DELETE
                
                diffDeletedRecord = convert2to1Array(diff2Rec.recordList, recordIndex)
                
                ' --------------------------------------------
                ' ���R�[�h���X�g���Z���ɏo�͂���
                ' 1�����z����o�͂���̂ŁAExcelUtil.copyArrayToCellsForColumns���g�p����
                ' --------------------------------------------
                ' ���R�[�h���X�g�̋󗓂̗̈���o��
                insertRecordListCell sheet, offsetRowReocrd, 1, VBUtil.arraySize(diffDeletedRecord)
                ' ���e���o��
                ExcelUtil.copyArrayToCellsForColumns diffDeletedRecord, sheet, offsetRowReocrd, ConstantsSnapDiff.RECORD_COL
                sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_DELETE
    
                deleteCount = deleteCount + 1
                totalCount = totalCount + 1
            
                ' �w�i�F��ύX����
                sheet.Range( _
                    sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL).Address, _
                    sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_COL + VBUtil.arraySize(diffDeletedRecord) - 1).Address _
                ).Interior.Color = _
                    RGB( _
                        ConstantsSnapDiff.DELETE_COLOR_R, _
                        ConstantsSnapDiff.DELETE_COLOR_G, _
                        ConstantsSnapDiff.DELETE_COLOR_B)
            
                ' �ύX�L�����Z���ɏ�������
                setModifyCell sheet, ConstantsSnapDiff.MODIFY_ALL_ROW, ConstantsSnapDiff.MODIFY_ALL_COL, True ' �S��
                setModifyCell sheet, offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.MODIFY_COL, True ' ��
                
                ' �A��
                recordNum = recordNum + 1
                sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_NUM_COL) = recordNum
            
                isOutRecord = True
                offsetRowReocrd = offsetRowReocrd + 1

            End If
            
        Next
        
    End If
    
    If isOutRecord Then
        ' ���R�[�h�o�͂�����
        endOffsetRow = offsetRowReocrd + 1
    
    Else
        ' ���R�[�h�o�͂��Ȃ��ioffsetRowRecord�ϐ�����x���C���N�������g����Ă��Ȃ��̂ŁA+1�]���ɉ��Z�j
        endOffsetRow = offsetRowReocrd + 2
        
        sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL) = ConstantsSnapDiff.MODIFY_NORECORD
        
        ' �ŏ��̗�̌r������������
        sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL + 1).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
        sheet.Cells(offsetRowReocrd, ConstantsSnapDiff.RECORD_MODIFY_COL + 1).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone

    End If
    
    Dim pkListStr As String
    
    ' SQL�ԍ�
    sheet.Cells(offsetRow + RESULT_START_OFFSET_SQL, ConstantsSnapDiff.SQL_NUM_COL) = recordSetIndex + 1
    ' SQL
    sheet.Cells(offsetRow + RESULT_START_OFFSET_SQL, ConstantsSnapDiff.SQL_COL) = diff1Rec.sql
    ' PK���X�g
    sheet.Cells(offsetRow + RESULT_START_OFFSET_SQL, ConstantsSnapDiff.PKEY_COL) = Join(diff1Rec.convertPkListToArray, vbNewLine)
    ' ���l
    sheet.Cells(offsetRow + RESULT_START_OFFSET_SQL, ConstantsSnapDiff.MEMO_COL) = diff1Rec.memo
    
    ' �����Ȃǂ��Z���ɏ�������
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.TOTAL_COUNT_COL) = totalCount
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.NOCHANGE_COUNT_COL) = noChangeCount
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.INSERT_COUNT_COL) = insertCount
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.UPDATE_COUNT_COL) = updateCount
    sheet.Cells(offsetRow + RESULT_START_OFFSET_MODIFY, ConstantsSnapDiff.DELETE_COUNT_COL) = deleteCount
    
    ' �w�b�_�ƃ��R�[�h���̍s���O���[�v������
    sheet.Range( _
        sheet.Cells(offsetRow + RESULT_START_OFFSET_HEADER, 1).Address, _
        sheet.Cells(endOffsetRow - 2, 1).Address).Rows.Group
    
    diffRecords = endOffsetRow
    
End Function

' =========================================================
' ���V�[�g�𐶐�����
'
' �T�v�@�@�@�F�w�肳�ꂽ�e�[�u���̃V�[�g�𐶐�����
' �����@�@�@�F
' �߂�l    �F�V�[�g
'
' =========================================================
Private Function createSheet() As Worksheet

    ' �V�[�g��
    Dim sheetName     As String

    ' ���X�i�b�v�V���b�g��r�V�[�g�����C�x���g
    RaiseEvent createSnapshotDiffSheetBegin

    sheetName = "��r����_" & Format(Now, "yyyymmdd_hhmmss")

    ' �R�s�[�V�[�g
    Dim copySheetName As String
    copySheetName = ConstantsSnapDiff.SHEET_NAME_TEMPLATE

    ExcelUtil.copySheetAppend SutWorkbook _
                            , copySheetName _
                            , m_book _
                            , sheetName
                            
    ' �V���ɃR�s�[���ꂽ�V�[�g�I�u�W�F�N�g���擾����
    Set createSheet = m_book.Worksheets(sheetName)
                            
    ' ���V�V�[�g�Ƀe�[�u����`���𖄂ߍ���
    With createSheet

        ' �V�[�g�����ɂ���
        .visible = xlSheetVisible

    End With
    
    ' ���X�i�b�v�V���b�g��r�V�[�g���������C�x���g
    RaiseEvent createSnapshotDiffSheet(createSheet)

End Function

' =========================================================
' ����r���ʃZ���͈͂�}��
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet        ���[�N�V�[�g
' �@�@�@�@�@�@recordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
'
' �߂�l�@�@�F
' =========================================================
Private Sub insertResultRange(ByRef sheet As Worksheet, ByVal recordOffset As Long)

    ' ------------------------------------------------------------
    ' �e���v���[�g�V�[�g�̃��R�[�h�̈���R�s�[����i���������R�s�[�j
    ' ------------------------------------------------------------
    ' �R�s�[���Z��
    Dim srcRange As Range
    ' �R�s�[��Z��
    Dim desRange As Range
    
    Dim srcSheet As Worksheet
    Set srcSheet = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE)

    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = srcSheet _
                    .Range( _
                        srcSheet.Cells(ConstantsSnapDiff.RESULT_RANGE_START_ROW _
                            , ConstantsSnapDiff.RESULT_RANGE_START_COL).Address & ":" & _
                        srcSheet.Cells(ConstantsSnapDiff.RESULT_RANGE_END_ROW _
                            , ConstantsSnapDiff.RESULT_RANGE_END_COL).Address)
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = sheet _
                     .Range( _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.RESULT_RANGE_START_COL).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange

End Sub

' =========================================================
' ���w�b�_�Z�����X�g��}��
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet        ���[�N�V�[�g
' �@�@�@�@�@�@recordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@columnSize   �J�����T�C�Y
'
' �߂�l�@�@�F
' =========================================================
Private Sub insertHeaderListCell(ByRef sheet As Worksheet, ByVal recordOffset As Long, ByVal columnSize As Long)

    ' �J�����̐�
    Dim columnCount As Long
    
    ' �V�[�g�̍ő��
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(sheet)
    
    ' �J�����̐����V�[�g�̍ő�񐔂𒴂��Ă��Ȃ������`�F�b�N����
    If columnSize > sheetColSize - ConstantsSnapDiff.RECORD_COL + 1 Then
    
        ' �ő�񐔂𒴂��Ă���ꍇ�A�񐔂𒲐�����
        columnCount = sheetColSize - ConstantsSnapDiff.RECORD_COL + 1
        
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
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_HEADER _
                          , ConstantsSnapDiff.HEADER_COL)
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = sheet _
                     .Range( _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.HEADER_COL).Address & ":" & _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.HEADER_COL + columnCount - 1).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange

    ' �J�������̒���
    ExcelUtil.changeColWidth desRange, srcRange.EntireColumn.ColumnWidth

    ' ------------------------------------------------------------
    ' ���R�[�h�ԍ�
    ' ------------------------------------------------------------
    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_HEADER _
                          , ConstantsSnapDiff.RECORD_NUM_COL)
                  
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = sheet _
                     .Range( _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.RECORD_NUM_COL).Address)

    srcRange.copy Destination:=desRange

    ' ------------------------------------------------------------
    ' ���R�[�h������e
    ' ------------------------------------------------------------
    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_HEADER _
                          , ConstantsSnapDiff.RECORD_MODIFY_COL)
                  
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = sheet _
                     .Range( _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.RECORD_MODIFY_COL).Address)
                            
    srcRange.copy Destination:=desRange

End Sub

' =========================================================
' �����R�[�h�Z�����X�g��}��
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet        ���[�N�V�[�g
' �@�@�@�@�@�@recordOffset ���R�[�h���f���̃I�t�Z�b�g�ʒu
' �@�@�@�@�@�@recordSize   ���R�[�h�T�C�Y
' �@�@�@�@�@�@columnSize   �J�����T�C�Y
'
' �߂�l�@�@�F
' =========================================================
Private Sub insertRecordListCell(ByRef sheet As Worksheet, ByVal recordOffset As Long, ByVal recordSize As Long, ByVal columnSize As Long)

    ' �J�����̐�
    Dim columnCount As Long
    
    ' �V�[�g�̍ő��
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(sheet)
    
    ' �J�����̐����V�[�g�̍ő�񐔂𒴂��Ă��Ȃ������`�F�b�N����
    If columnSize > sheetColSize - ConstantsSnapDiff.RECORD_COL + 1 Then
    
        ' �ő�񐔂𒴂��Ă���ꍇ�A�񐔂𒲐�����
        columnCount = sheetColSize - ConstantsSnapDiff.RECORD_COL + 1
        
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
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_RECORD _
                          , ConstantsSnapDiff.RECORD_COL)
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = sheet _
                     .Range( _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        sheet.Cells(recordOffset + recordSize - 1 _
                                  , ConstantsSnapDiff.RECORD_COL + columnCount - 1).Address)

    ' ���͗̈���R�s�[����
    srcRange.copy Destination:=desRange

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

    ' ------------------------------------------------------------
    ' ���R�[�h�ԍ�
    ' ------------------------------------------------------------
    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_RECORD _
                          , ConstantsSnapDiff.RECORD_NUM_COL)
                  
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = sheet _
                     .Range( _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        sheet.Cells(recordOffset + recordSize - 1 _
                                  , ConstantsSnapDiff.RECORD_NUM_COL).Address)

    srcRange.copy Destination:=desRange

    ' ------------------------------------------------------------
    ' ���R�[�h������e
    ' ------------------------------------------------------------
    ' �R�s�[����Range�I�u�W�F�N�g���擾����
    Set srcRange = ThisWorkbook.Worksheets(ConstantsSnapDiff.SHEET_NAME_TEMPLATE) _
                     .Cells(ConstantsSnapDiff.RESULT_START_ROW + ConstantsSnapDiff.RESULT_START_OFFSET_RECORD _
                          , ConstantsSnapDiff.RECORD_MODIFY_COL)
                  
                  
    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set desRange = sheet _
                     .Range( _
                        sheet.Cells(recordOffset _
                                  , ConstantsSnapDiff.RECORD_COL).Address & ":" & _
                        sheet.Cells(recordOffset + recordSize - 1 _
                                  , ConstantsSnapDiff.RECORD_MODIFY_COL).Address)
                            
    srcRange.copy Destination:=desRange

End Sub

' =========================================================
' ���ύX�L���̃Z������ݒ肷��
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet    ���[�N�V�[�g
' �@�@�@�@�@�@row      �s
' �@�@�@�@�@�@col      ��
' �@�@�@�@�@�@modify   �ύX�L��
'
' �߂�l�@�@�F
' =========================================================
Private Sub setModifyCell(ByRef sheet As Worksheet, ByVal row As Long, ByVal col As Long, ByVal modify As Boolean)

    Dim targetRange As Range

    ' �R�s�[���Range�I�u�W�F�N�g���擾����
    Set targetRange = sheet _
                     .Range(sheet.Cells(row, col).Address)

    If modify Then
        targetRange.value = ConstantsSnapDiff.MODIFY_ON
        targetRange.Interior.Color = _
            RGB( _
                ConstantsSnapDiff.MODIFY_CHANGE_COLOR_R, _
                ConstantsSnapDiff.MODIFY_CHANGE_COLOR_G, _
                ConstantsSnapDiff.MODIFY_CHANGE_COLOR_B)
    Else
        targetRange.value = ConstantsSnapDiff.MODIFY_OFF
        targetRange.Interior.Color = _
            RGB( _
                ConstantsSnapDiff.MODIFY_NOCHANGE_COLOR_R, _
                ConstantsSnapDiff.MODIFY_NOCHANGE_COLOR_G, _
                ConstantsSnapDiff.MODIFY_NOCHANGE_COLOR_B)
    End If

End Sub

' =========================================================
' ���L�����Z���t���O�`�F�b�N
'
' �T�v�@�@�@�F�L�����Z���t���O���`�F�b�N���L�����Z������Ă����ꍇ��
' �@�@�@�@�@�@�G���[�𔭍s����B
'
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub checkCancel()

    If canceled = True Then

        err.Raise ConstantsError.ERR_NUMBER_PROC_CANCEL _
                , err.Source _
                , ConstantsError.ERR_DESC_PROC_CANCEL
    End If

End Sub

