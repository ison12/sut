VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeQueryExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �N�G�������s����N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2013/06/19�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ________________________________________________________
' �C�x���g
' ________________________________________________________

' =========================================================
' ��SQL���s�C�x���g
'
' �T�v�@�@�@�FSQL���s���ɌĂяo�����C�x���g
' �����@�@�@�Fsql        SQL�X�e�[�g�����g
'
' =========================================================
Public Event executeSql(ByRef sql As String)

' =========================================================
' ��SQL���s�L�����Z���C�x���g
'
' �T�v�@�@�@�FSQL�̎��s�L�����Z�����ɌĂяo�����C�x���g
' �����@�@�@�Fcount ���s����
'
' =========================================================
Public Event ExecuteSqlCancel(ByVal count As Long)

' =========================================================
' ��SQL���s�����C�x���g
'
' �T�v�@�@�@�FSQL���s�ɐ��������ۂɌĂяo�����C�x���g
' �����@�@�@�Fsql        SQL�X�e�[�g�����g
'
' =========================================================
Public Event ExecuteSqlSuccessful(ByRef sql As String)

' =========================================================
' ��SQL���s���̃G���[�C�x���g
'
' �T�v�@�@�@�FSQL���s���ɃG���[�����������ۂɌĂяo�����C�x���g
' �����@�@�@�Fsql        SQL�X�e�[�g�����g
' �@�@�@�@�@�Fsqlstate   SQL�X�e�[�^�X
' �@�@�@�@�@�Fmessage    ���b�Z�[�W
'
' =========================================================
Public Event ExecuteSqlError(ByRef sql As String, ByRef sqlState As String, ByRef message As String)

' =========================================================
' ��SQL���s���ʔ��f���̃C�x���g�i���R�[�h�擾�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordGetBegin(ByVal recordListSize As Long)

' =========================================================
' ��SQL���s���ʔ��f���̃C�x���g�i���R�[�h�擾�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordGet(ByVal recordListSizeActual As Long _
                               , ByVal recordListSize As Long _
                               , ByRef recordList As Variant)

' =========================================================
' ��SQL���s���ʔ��f��̃C�x���g�i�擾���R�[�h�ҏW�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordConvertBegin(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ��SQL���s���ʔ��f��̃C�x���g�i�擾���R�[�h�ҏW�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordConvert(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ��SQL���s���ʔ��f���̃C�x���g�i�V�[�g���f�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecordBegin(ByVal recordListSize As Long, ByRef recordList As Variant)

' =========================================================
' ��SQL���s���ʔ��f��̃C�x���g�i�V�[�g���f�j
'
' �T�v�@�@�@�F
' �����@�@�@�FrecordListSize ���R�[�h���X�g�T�C�Y
' �@�@�@�@�@�@recordList     ���R�[�h���X�g
'
' =========================================================
Public Event ReflectionRecord(ByVal recordListSize As Long, ByRef recordList As Variant)


' ---------------------------------------------------------
' ���v���p�e�B
' ---------------------------------------------------------
' DB�R�l�N�V����
Public conn          As Object
' �L�����Z���t���O
Public canceled      As Boolean
' ���[�N�V�[�g�ɒǉ����邩�ǂ���
Public appendSheet   As Boolean
' �g�����U�N�V�����J�n�L���t���O
Public isBeginTran   As Boolean

' SQL���s���t���O
Private executingSql      As Boolean
' SQL���s���ɃG���[�����������ꍇ
Private executingSqlError As Boolean

' �N�G�����ʃV�[�g����
Private queryResultSheetCreator As ExeQueryResultSheetCreator

' ���[�N�V�[�g
Private m_sheet       As Worksheet
' �A�v���P�[�V�����ݒ���
Private m_applicationSetting As ValApplicationSetting

' =========================================================
' ���v���p�e�B �V�[�g
' =========================================================
Public Property Let sheet(ByVal vNewValue As Variant)

    Set m_sheet = vNewValue
    queryResultSheetCreator.sheet = vNewValue
    queryResultSheetCreator.book = vNewValue.parent
End Property

' =========================================================
' ���A�v���P�[�V�����ݒ���
' =========================================================
Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
    queryResultSheetCreator.applicationSetting = vNewValue
End Property

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set queryResultSheetCreator = New ExeQueryResultSheetCreator
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set queryResultSheetCreator = Nothing
End Sub

' =========================================================
' ���N�G�����s
'
' �T�v�@�@�@�F�ΏۂƂȂ�N�G���̎��s���ʂ��V�[�g�ɔ��f����
' �����@�@�@�FqueryList  �N�G�����X�g
' �߂�l�@�@�F���R�[�h�擾����
'
' =========================================================
Public Function execute(ByRef queryList As ValCollection) As Long

    On Error GoTo err

    ' �e����^�������R�[�h
    Dim affectedRecord As Long
    
    Dim i, j As Long
    
    ' �w�b�_�z��
    Dim headerList As Variant
    ' �t�B�[���h
    Dim field   As Object
    
    ' ���R�[�h�̑���
    Dim recordCount As Long
    ' ���R�[�h�Z�b�g
    Dim recordSet   As Object
    ' ���R�[�h�z��
    Dim recordList  As Variant
    ' ���R�[�h�z��i�s�T�C�Y�j
    Dim recordListSize As Long

    ' ���R�[�h�Z�b�g�C���f�b�N�X
    Dim recordSetIndex As Long
    ' ���R�[�h���f�I�t�Z�b�g
    Dim recordOffset  As Long
    
    ' ���R�[�h�����P��
    Dim processCount As Long: processCount = m_applicationSetting.getRecProcessCount

    If queryList.count <= 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_SQL_EMPTY _
                , err.Source _
                , ConstantsError.ERR_DESC_SQL_EMPTY
    End If

    ' ���R�[�h�Z�b�g�C���f�b�N�X
    recordSetIndex = 1
    ' ���R�[�h���f�I�t�Z�b�g�ʒu���擾����
    recordOffset = queryResultSheetCreator.getRecordOffset(appendSheet)
    
    ' �ΏۃV�[�g���N�G�����ʃV�[�g�ł͂Ȃ��ꍇ
    If Not queryResultSheetCreator.isQueryResultSheet(m_sheet) Then
    
        Set m_sheet = queryResultSheetCreator.createSheet
    End If

    If appendSheet = False Then
        
        ' �s���폜����
        ExcelUtil.deleteRowEndOfLastInputted m_sheet, ConstantsTable.QUERY_COLUMN_OFFSET_ROW, ConstantsTable.QUERY_COLUMN_TITLE_COL
        
        ' �A�E�g���C�����N���A����
        m_sheet.Cells.ClearOutline
        
    End If
    
    m_sheet.activate
    m_sheet.Cells(recordOffset _
              , ConstantsTable.QUERY_COLUMN_OFFSET_COL).activate
    
    ' �L�����Z���t���O���`�F�b�N����
    checkCancel
    
    ' �g�����U�N�V�����I�u�W�F�N�g
    Dim transaction As DBTransaction
    Set transaction = Nothing
    
    If isBeginTran = False Then
    
        Set transaction = New DBTransaction
        transaction.init conn
        
    End If
        
    Dim var As Variant
    Dim sql As String
    
    executingSql = False
    executingSqlError = False
    
    For Each var In queryList.col
    
        sql = var.query
    
        ' -------------------------------------------------------------
        ' SQL�����s����
        ' -------------------------------------------------------------
        ' �L�����Z���t���O���`�F�b�N����
        checkCancel
        
        ' �����̂͂��߂Ƀt���O���I�t�ɂ��Ă���
        executingSql = False
    
        RaiseEvent executeSql(sql)
    
        ' SQL���s���t���O��ON�ɂ���
        executingSql = True
    
        ' SQL���s�t���O��ON�ɂ���
        executingSql = True
        executingSqlError = False
        
        ' �N�G���[�����s����
        Set recordSet = ADOUtil.queryBatch(conn, sql, affectedRecord)
        
        ' SQL���s���t���O��OFF�ɂ���
        executingSql = False
                
        ' SQL���s���G���[�t���O��ON�̏ꍇ
        ' �G���[�t���O�́A�N�G�����s���ɃG���[�����������ꍇ��
        ' �{���\�b�h�̃G���[�n���h�����O��ON�ɐݒ肳���B
        If executingSqlError = True Then
        
            GoTo loop_end

        End If
    
        RaiseEvent ExecuteSqlSuccessful(sql)
        
        ' -------------------------------------------------------------
        ' ���R�[�h���擾����
        ' -------------------------------------------------------------
        
        ' UPDATE�Ȃǂ́A���R�[�h��Ԃ��Ȃ��N�G���̏ꍇ State = 0 (adStateClosed) �ɂȂ�
        If recordSet.state = 0 Then
        
            ' ���R�[�h�Z�b�g�̔ԍ����Z���ɖ��ߍ���
            queryResultSheetCreator.insertEmptyCellForTitle recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
            m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "No." & recordSetIndex
            
            ' sql���R�����g�Ƃ��Ė��ߍ���
            ExcelUtil.addComment m_sheet.name, recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL, sql
        
            queryResultSheetCreator.insertEmptyCellForResultCount recordOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL
            m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL).value = "����Ɏ��s����܂����i" & affectedRecord & "���̃��R�[�h�ɉe����^���܂����j�B"
            recordOffset = recordOffset + 1
            
            recordCount = recordCount + affectedRecord
        Else
        
            ' ���R�[�h�Z�b�g�̔ԍ����Z���ɖ��ߍ���
            queryResultSheetCreator.insertEmptyCellForTitle recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
            m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "No." & recordSetIndex
            
            ' sql���R�����g�Ƃ��Ė��ߍ���
            ExcelUtil.addComment m_sheet.name, recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL, sql
        
            ' ---------------------------------------------------
            ' �w�b�_���̎擾
            ' ---------------------------------------------------
            ReDim headerList(recordSet.Fields.count - 1)
            
            ' �w�b�_���̎擾
            i = 0
            For Each field In recordSet.Fields
            
                headerList(i) = field.name
                
                i = i + 1
                
            Next field

            ' �w�b�_���X�g�̘g�g�݂�}������
            queryResultSheetCreator.insertEmptyCellForColumn recordOffset, ConstantsTable.QUERY_COLUMN_OFFSET_COL, VBUtil.arraySize(headerList)
            ' ���[�N�V�[�g�Ƀ��R�[�h���X�g�̓��e�𔽉f����
            queryResultSheetCreator.reflectSheetForColumn headerList, recordOffset
            
            ' �w�b�_�𔽉f�����̂ŃI�t�Z�b�g�ʒu�����Z����
            recordOffset = recordOffset + 1
            ' ---------------------------------------------------
            
            ' ---------------------------------------------------
            ' ���R�[�h���̎擾
            ' ---------------------------------------------------
            ' �擾���R�[�h��0���ɂȂ�܂Ń��R�[�h�Z�b�g����������
            Do While Not recordSet.EOF
            
                ' �L�����Z���t���O���`�F�b�N����
                checkCancel
                
                ' �����R�[�h�擾�C�x���g
                RaiseEvent ReflectionRecordGetBegin(processCount)
                
                ' ���R�[�h�Z�b�g���烌�R�[�h���X�g���擾����
                recordList = getRecordList(recordSet, processCount)
                
                ' ���R�[�h���X�g�̃T�C�Y���擾����
                ' ADO�̃��R�[�h�Z�b�g����GetRows�Ŏ擾�����z���(��~�s)�Ńf�[�^���i�[����Ă���
                recordListSize = VBUtil.arraySize(recordList, 2)
        
                ' �����R�[�h�擾�C�x���g
                RaiseEvent ReflectionRecordGet(processCount, recordListSize, recordList)
                
                ' ���R�[�h���X�g���z��ł͂Ȃ��ꍇ
                If IsArray(recordList) = False Then
                
                    ' ���̃u���b�N�ɓ��邱�Ƃ͂���̂��H
                    GoTo loop_end
                End If
                
                ' �����R�[�h�ϊ��C�x���g
                RaiseEvent ReflectionRecordConvertBegin(recordListSize, recordList)
                
                ' ���R�[�h���X�g��ϊ�����
                convertRecordList recordList
                
                ' �����R�[�h�ϊ��C�x���g
                RaiseEvent ReflectionRecordConvert(recordListSize, recordList)
                
                ' �����R�[�h���f�C�x���g
                RaiseEvent ReflectionRecordBegin(recordListSize, recordList)
                
                ' ���R�[�h���X�g�̘g�g�݂�}������
                queryResultSheetCreator.insertEmptyCellForRecord recordOffset, recordListSize, ConstantsTable.QUERY_COLUMN_OFFSET_COL, VBUtil.arraySize(headerList)
                ' ���[�N�V�[�g�Ƀ��R�[�h���X�g�̓��e�𔽉f����
                queryResultSheetCreator.reflectSheetForRecord recordList, recordOffset
                
                ' �����R�[�h���f�C�x���g
                RaiseEvent ReflectionRecord(recordListSize, recordList)
                
                ' ����̃��R�[�h���X�g���f���̃I�t�Z�b�g�ʒu�̌v�Z
                ' ���R�[�h���X�g�̔��f��Ƀ��R�[�h���X�g�̃T�C�Y�����Z����
                recordOffset = recordOffset + recordListSize
                
                ' ���R�[�h�擾�������J�E���g����
                recordCount = recordCount + recordListSize
            Loop
    
        End If
            
loop_end:
            
        recordSetIndex = recordSetIndex + 1
        
        ' ���R�[�h�Z�b�g���������
        ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    Next
    
    ' �O�̂��߃��R�[�h�Z�b�g���������
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    ' �g�����U�N�V�������R�~�b�g����
    If Not transaction Is Nothing Then
        transaction.commit
    End If
    
    ' -------------------------------------------------------

    ' ���R�[�h�擾������Ԃ�
    execute = recordCount
    
    ' ���R�[�h���X�g�̘g�g�݂�}������
    queryResultSheetCreator.insertEmptyCellForResultSet recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
    ' �ŏI�s�Ɍ��ݓ��t�𖄂ߍ���
    m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).WrapText = False
    m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "'" & Now & " : �N�G�����s�����c" & (recordSetIndex - 1)

    ' �O���[�v�����ꂽ�̈��S�ĕ���
    m_sheet.Outline.ShowLevels 1

    Exit Function

err:

    ' �G���[�n���h���ŕʂ̊֐����Ăяo���ƃG���[��񂪏����Ă��܂����Ƃ�����̂�
    ' �\���̂ɃG���[����ۑ����Ă���
    Dim errT As errInfo: errT = VBUtil.swapErr

    ' ���R�[�h�Z�b�g���������
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    ' SQL���s���ɃG���[������
    If executingSql = True Then

        ' SQL�X�e�[�^�X
        Dim sqlState As String
        
        If conn.Errors.count > 0 Then
        
            ' �G���[�I�u�W�F�N�g���擾����
            Dim error As Object
            Set error = conn.Errors.item(0)
            sqlState = error.sqlState
        End If

        ' ��SQL���s���G���[
        RaiseEvent ExecuteSqlError(sql, sqlState, err.Description)
        
        ' ���R�[�h�Z�b�g�̔ԍ����Z���ɖ��ߍ���
        queryResultSheetCreator.insertEmptyCellForTitle recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
        m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "No." & recordSetIndex
            
        ' sql���R�����g�Ƃ��Ė��ߍ���
        ExcelUtil.addComment m_sheet.name, recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL, sql

        queryResultSheetCreator.insertEmptyCellForError recordOffset, QUERY_COLUMN_OFFSET_COL
        m_sheet.Cells(recordOffset, QUERY_COLUMN_OFFSET_COL).value = sqlState & " " & errT.Description
        m_sheet.Cells(recordOffset, QUERY_COLUMN_OFFSET_COL).WrapText = False
        
        ' �G���[���𖄂ߍ��񂾂̂ŃI�t�Z�b�g�ʒu�����Z����
        recordOffset = recordOffset + 1
        
        ' �O���[�v�����ꂽ�̈��S�ĕ���
        m_sheet.Outline.ShowLevels 1
        
        ' SQL���s���G���[�t���O�𗧂Ă�
        executingSqlError = True
        Resume Next
        
        ' ----------------------------------------------------
        ' Resume next �ŃG���[�����ꏊ�ɖ߂�̂ŁA�ȉ��̏����͎��s���Ȃ�
        ' ----------------------------------------------------
'        ' ���R�[�h���X�g�̘g�g�݂�}������
'        queryResultSheetCreator.insertEmptyCellForResultSet recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL
'        ' �ŏI�s�Ɍ��ݓ��t�𖄂ߍ���
'        m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).WrapText = False
'        m_sheet.Cells(recordOffset, ConstantsTable.QUERY_COLUMN_TITLE_COL).value = "'" & Now & " : �N�G�������f����܂����B"
'
'        ' �O���[�v�����ꂽ�̈��S�ĕ���
'        m_sheet.Outline.ShowLevels 1
        
    Else

        err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext
    
    End If

End Function

' =========================================================
' �����R�[�h�擾
'
' �T�v�@�@�@�F���R�[�h��RecordSet�I�u�W�F�N�g����擾���A�z��Ɋi�[���ĕԂ��B
'
' �����@�@�@�FrecordSet ���R�[�h�Z�b�g
' �@�@�@�@�@�@fetchSize �t�F�b�`�T�C�Y
' �߂�l�@�@�F���R�[�h�z��
'
' =========================================================
Private Function getRecordList(ByRef recordSet As Object _
                             , Optional ByVal fetchSize As Long = 0) As Variant

    ' �߂�l
    Dim result         As Variant
    
    ' ���R�[�h�Z�b�g��EOF�ł͂Ȃ��ꍇ
    If Not recordSet.EOF Then
    
        ' �t�F�b�`�T�C�Y���`�F�b�N����
        If fetchSize = 0 Then
        
            ' ���R�[�h�Z�b�g����S���R�[�h���擾����
            result = recordSet.getRows
        Else
        
            ' ���R�[�h�Z�b�g���烌�R�[�h���t�F�b�`�T�C�Y���擾����
            result = recordSet.getRows(fetchSize)
        End If
    
        getRecordList = result
    Else
    
        ' ���Ԃ�
        getRecordList = Empty
    End If

End Function

' =========================================================
' �����R�[�h���X�g�ϊ�
'
' �T�v�@�@�@�F���R�[�h���X�g�̃f�[�^��ϊ�����B
' �����@�@�@�FtableSheet   �e�[�u���V�[�g
'             recordList   ���R�[�h���X�g
'
' =========================================================
Private Function convertRecordList(ByRef recordList As Variant)

    ' ���R�[�h�z��̗v�f����ւ��ɂ���ĐV���ɐ��������z��
    Dim temp As Variant
    
    ' ���R�[�h�z��̃f�[�^�ϊ�����
    ' RecordSet.GetRows���Ԃ��ꂽ�z��́A(col,rec)�ƂȂ��Ă��邽��(rec,col)�ƂȂ�悤�ɓ���ւ������{����
    temp = convertArrayValuesToUnder(recordList)
    
    ' ���R�[�h�z����������
    Erase recordList

    ' ���R�[�h���X�g�����ւ���
    recordList = temp

End Function

' =========================================================
' ���z��v�f�̕ϊ�
'
' �T�v�@�@�@�F2�����z��̗v�f��(x,y)����(y,x)�ɐݒ肵�Ȃ����B
' �@�@�@�@�@�@����ɗv�f���Z���`���̒l�ɕϊ�����B
' �����@�@�@�Fv 2�����z��
'
' �߂�l�@�@�F2�����z��
' ���L�����@�FVBUtil.transposeDim�֐����J�X�^�}�C�Y�����֐��B
' �@�@�@�@�@�@�����͂قړ����B
'
' =========================================================
Private Function convertArrayValuesToUnder(ByRef v As Variant) As Variant
    
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
    
    Dim value As Variant
    
    For X = Xlower To Xupper
        For Y = Ylower To Yupper
        
            ' ��U�l���擾����
            value = v(Y, X)
            
            ' ������̏ꍇ�A�Z���`���̒l�ւƕϊ�����
            tempArray(X, Y) = ExcelUtil.convertCellValue(value)
        
        Next Y
    Next X
    
    convertArrayValuesToUnder = tempArray

End Function

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