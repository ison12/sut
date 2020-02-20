VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���R�[�h���X�V����N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/02/01�@�V�K�쐬
'
' ���L�����F
' *********************************************************
'
' ________________________________________________________
' �C�x���g
' ________________________________________________________

' =========================================================
' ��ExeTableSheetReader�Œ�`���Ă���C�x���g���Ē�`
' �@��ʃ��W���[���ɁA���������C�x���g�����̂܂ܓ`�d���邽�߂ɍĒ�`
' =========================================================
Public Event readColumnListBegin(ByVal cnt As Long)
Public Event readColumn(ByRef column As ValDbDefineColumn)
Public Event readColumnList(ByRef columnList As ValCollection)
Public Event readTableInfoBegin()
Public Event readTableInfo(ByRef table As ValTableWorksheet)
Public Event readRecordListSize(ByVal cnt As Long)
Public Event readRecordListBegin(ByVal cnt As Long)
Public Event readRecordList(ByRef recordList As ValRecordList)
Public Event readRecordListBySelectionBegin(ByVal cnt As Long)
Public Event readRecordListBySelectionOne(ByRef recordList As ValRecordList)
Public Event readRecordListBySelection(ByRef recordLists As ValCollection)

' =========================================================
' ��SQL���s�J�n�C�x���g
'
' �T�v�@�@�@�FSQL���s�J�n�O�ɌĂяo�����C�x���g
' �����@�@�@�Fcount ���s����
'
' =========================================================
Public Event ExecuteSqlBegin(ByVal count As Long)

' =========================================================
' ��SQL���s�C�x���g
'
' �T�v�@�@�@�FSQL���s���ɌĂяo�����C�x���g
' �@�@�@�@�@�@����Ɏ��s���ꂽ�Ƃ��Ă����ʂ�0���̏ꍇ�����邽��
' �@�@�@�@�@�@�����������ꍇ�ɂ́Acount���Q�Ƃ��ĉ����X�V����Ă��邩���m�F����B
'
' �����@�@�@�FtableSheet       �e�[�u���V�[�g
' �@�@�@�@�@�@recordList       ���R�[�h���X�g
' �@�@�@�@�@�@recordListIndex  ���R�[�h���X�g�C���f�b�N�X
'             sql              SQL�X�e�[�g�����g
' �@�@�@�@�@�@count            ���ʌ���
'
' =========================================================
Public Event executeSql(ByRef tableSheet As ValTableWorksheet _
                      , ByRef recordList As ValRecordList _
                      , ByVal recordListIndex As Long _
                      , ByRef sql As String _
                      , ByVal count As Long)

' =========================================================
' ��SQL���s�����C�x���g
'
' �T�v�@�@�@�FSQL���s�ɐ��������ۂɌĂяo�����C�x���g
' �����@�@�@�Fcount ���s����
'
' =========================================================
Public Event ExecuteSqlSuccessful(ByVal count As Long)

' =========================================================
' ��SQL���s���G���[�C�x���g
'
' �T�v�@�@�@�FSQL���s���ɃG���[�����������ۂɌĂяo�����C�x���g
' �����@�@�@�FtableSheet       �e�[�u���V�[�g
' �@�@�@�@�@�@recordList       ���R�[�h���X�g
' �@�@�@�@�@�@recordListIndex  ���R�[�h���X�g�C���f�b�N�X
'             sql              SQL�X�e�[�g�����g
'             sqlstate         SQL�X�e�[�^�X
' �@�@�@�@�@�@errorDesc        �G���[���
'
' =========================================================
Public Event ExecuteSqlError(ByRef tableSheet As ValTableWorksheet _
                           , ByRef recordList As ValRecordList _
                           , ByVal recordListIndex As Long _
                           , ByRef sql As String _
                           , ByRef sqlState As String _
                           , ByRef errorDesc As String)

' =========================================================
' ��SQL���s�L�����Z���C�x���g
'
' �T�v�@�@�@�FSQL�̎��s�L�����Z�����ɌĂяo�����C�x���g
' �����@�@�@�Fcount ���s����
'
' =========================================================
Public Event ExecuteSqlCancel(ByVal count As Long)


' ---------------------------------------------------------
' ���v���p�e�B
' ---------------------------------------------------------
' �e�[�u���V�[�g�ǂݍ���
Private WithEvents tableSheetReader As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1

' DB�R�l�N�V����
Public conn          As Object
' ���[�N�V�[�g
Public sheet         As Worksheet
' �L�����Z���t���O
Public canceled      As Boolean

' �A�v���P�[�V�����ݒ���
Public applicationSetting As ValApplicationSetting
' �A�v���P�[�V�����ݒ�i�J�����������j
Public applicationSettingColFmt As ValApplicationSettingColFormat

' �J�����������X�g
Public columnFormatList As ValCollection

' �N�G���[���
Public queryType     As Long

' �N�G���[�����I�u�W�F�N�g
Private queryCreator As IDbQueryCreator
' �N�G���ϊ�
Private queryConverter As New DbQueryConverter
' �N�G���p�����[�^���X�g
Private queryParameterList_ As ValQueryParameterList

' SQL���s���t���O
Private executingSql      As Boolean
' SQL���s���ɃG���[�����������ꍇ
Private executingSqlError As Boolean

' �t�@�C���o��
Public isFileOutput  As Boolean
' �t�@�C���p�X
Public filePath      As String
' �����R�[�h
Public charSet       As String
' ���s�R�[�h
Public newline       As String

' �߂�l�p�@�ǂݍ��܂ꂽ�e�[�u���V�[�g
' delete all ���ɌĂяo�����ɏ���ԋp���邽�߂Ɏg�p����
Public readTableSheet As ValTableWorksheet
' �߂�l�p�@�������ꂽSQL
' delete all ���ɌĂяo�����ɏ���ԋp���邽�߂Ɏg�p����
Public createdSql As String

' �t�@�C���o�̓I�u�W�F�N�g
Private writer As FileWriter

' =========================================================
' ���N�G���[��� INSERT + UPDATE
' =========================================================
Public Property Get QUERY_TYPE_INSERT_UPDATE() As Variant
    
    QUERY_TYPE_INSERT_UPDATE = 4
End Property

' =========================================================
' ���N�G���[��� INSERT
' =========================================================
Public Property Get QUERY_TYPE_INSERT() As Variant
    
    QUERY_TYPE_INSERT = 0
End Property

' =========================================================
' ���N�G���[��� UPDATE
' =========================================================
Public Property Get QUERY_TYPE_UPDATE() As Variant

    QUERY_TYPE_UPDATE = 1
End Property

' =========================================================
' ���N�G���[��� DELETE
' =========================================================
Public Property Get QUERY_TYPE_DELETE() As Variant

    QUERY_TYPE_DELETE = 2
End Property

' =========================================================
' ���N�G���[��� DELETE�i�S���R�[�h�폜�j
' =========================================================
Public Property Get QUERY_TYPE_DELETE_ALL() As Variant

    QUERY_TYPE_DELETE_ALL = 3
End Property

' =========================================================
' ���N�G���p�����[�^���X�g�擾�v���p�e�B
' =========================================================
Public Property Get queryParameterList() As ValQueryParameterList

    If Not queryParameterList_ Is Nothing Then
        Set queryParameterList = queryParameterList_
        Exit Property
    End If

    Set queryParameterList_ = Nothing
        
    If sheet Is Nothing Then
        Exit Property
    End If
    
    Set queryParameterList_ = New ValQueryParameterList
    queryParameterList_.init sheet.parent
    queryParameterList_.readForData
    
    Set queryParameterList = queryParameterList_
    
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
' ���X�V���s
'
' �T�v�@�@�@�F�e�[�u���V�[�g���烌�R�[�h����ǂݍ���
' �@�@�@�@�@�@���R�[�h���X�g�̐�����SQL�̐����E���s���J��Ԃ��s���B
' �����@�@�@�F
' �߂�l�@�@�F���s����
'
' =========================================================
Public Function execute() As Long

    On Error GoTo err
    
    ' �e�[�u���V�[�g
    Dim tableSheet As ValTableWorksheet
    ' ���R�[�h���X�g
    Dim recordList As ValRecordList
    
    ' ���R�[�h��������
    Dim processCount As Long: processCount = applicationSetting.getRecProcessCount

    ' �t�@�C�����o�͂���ꍇ
    If isFileOutput = True Then
        
        ' �t�@�C���̑��݃`�F�b�N���s���A���݂���ꍇ�t�@�C�����폜����
        If VBUtil.isExistFile(filePath) = True Then
        
            Kill filePath
        End If
        
        ' �t�@�C���������݃I�u�W�F�N�g������������
        Set writer = New FileWriter
        
        writer.init filePath _
                  , charSet _
                  , newline _
                  , False
    
    End If

    ' �e�[�u���V�[�g�ǂݍ��݃I�u�W�F�N�g������������
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = sheet
    
    Set tableSheet = tableSheetReader.readTableInfo

    ' ���s����
    Dim ret     As Long
    
    ' ���R�[�h�T�C�Y
    Dim recSize As Long
    ' ���R�[�h�T�C�Y���擾����
    recSize = tableSheetReader.getRecordSize(tableSheet)

    ' ���R�[�h���X�g�̃T�C�Y���C�x���g�Ƃ��đ��M����
    RaiseEvent readRecordListSize(recSize)
    
    ' �����P�ʂ�1���ȏ�̏ꍇ
    If processCount > 0 Then
    
        ' �����P�ʁi�ꎞ�ϐ��j
        Dim tmpProcessCount As Long: tmpProcessCount = processCount
        
        Dim i As Long
    
        ' �����l��1�Ƃ���
        i = 1
        
        ' �����P�ʂ��Ƃɕ������ď���
        Do While i <= recSize
            
            ' �����P�ʂ����R�[�h�T�C�Y�������
            ' �㑱�̏����ŃG���[���������Ă��܂��̂ŁA�����P�ʂ̐��𒲐�����
            If i + tmpProcessCount - 1 > recSize Then
            
                tmpProcessCount = recSize - i + 1
            End If
        
            
            ' ���R�[�h���X�g���擾����
            Set recordList = tableSheetReader.readRecordList(tableSheet _
                                                           , i _
                                                           , tmpProcessCount)
                                                                 
            ' �N�G���[�����s����
            If isFileOutput = False Then
                ' �N�G���[�����s����
                ret = ret + executeQuery(tableSheet, recordList)
            
            ' �t�@�C�����o�͂���
            Else
            
                ret = ret + executeOutputQueryFile(tableSheet, recordList)
            End If
            
            i = i + tmpProcessCount
            
        Loop
        
    ' �����P�ʂ���L�ȊO
    Else
    
        ' �S���R�[�h�܂Ƃ߂ď�������
        
        ' ���R�[�h���X�g���擾����
        Set recordList = tableSheetReader.readRecordListAll(tableSheet)
                                                             
        ' �N�G���[�����s����
        If isFileOutput = False Then
        
            ret = executeQuery(tableSheet, recordList)
        
        ' �t�@�C�����o�͂���
        Else
            
            ret = executeOutputQueryFile(tableSheet, recordList)
        End If
    
    End If
    
    ' �߂�l�Ɏ��s������ݒ�
    execute = ret

    ' �t�@�C�����o�͂���
    If Not writer Is Nothing Then
    
        writer.destroy
        Set writer = Nothing
    End If
    
    Exit Function
err:

    ' �t�@�C���̏o�͂��L�����Z������
    If Not writer Is Nothing Then
    
        writer.Cancel
        Set writer = Nothing
    End If
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Function

' =========================================================
' ���I��̈�̃��R�[�h�̍X�V�����s����
'
' �T�v�@�@�@�F�e�[�u���V�[�g���烌�R�[�h����ǂݍ���
' �@�@�@�@�@�@���R�[�h���X�g�̐�����SQL�̐����E���s���J��Ԃ��s���B
' �����@�@�@�F
' �߂�l�@�@�F���s����
'
' =========================================================
Public Function executeSelection() As Long

    On Error GoTo err
    
    ' ���s����
    Dim ret     As Long
    
    ' �e�[�u���V�[�g
    Dim tableSheet  As ValTableWorksheet

    ' ���R�[�h���X�g�y�R���N�V�����z
    Dim recordLists As ValCollection
    ' ���R�[�h���X�g
    Dim recordList As ValRecordList

    ' �t�@�C�����o�͂���ꍇ
    If isFileOutput = True Then
        
        ' �t�@�C���̑��݃`�F�b�N���s���A���݂���ꍇ�t�@�C�����폜����
        If VBUtil.isExistFile(filePath) = True Then
        
            Kill filePath
        End If
        
        ' �t�@�C���������݃I�u�W�F�N�g������������
        Set writer = New FileWriter
        
        writer.init filePath _
                  , charSet _
                  , newline _
                  , False
    
    End If

    ' �e�[�u���V�[�g�ǂݍ��݃I�u�W�F�N�g������������
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = sheet
    
    Set tableSheet = tableSheetReader.readTableInfo

    ' ���R�[�h���X�g�y�R���N�V�����z���擾����
    Set recordLists = tableSheetReader.readRecordListBySelectionArea(tableSheet)
    
    ' ���R�[�h���X�g�y�R���N�V�����z��1������������
    For Each recordList In recordLists.col
    
        ' �N�G���[�����s����
        If isFileOutput = False Then
        
            ret = ret + executeQuery(tableSheet, recordList)
        
        ' �t�@�C�����o�͂���
        Else
            
            ret = ret + executeOutputQueryFile(tableSheet, recordList)
        End If
    
    Next
        
    ' �߂�l�Ɏ��s������ݒ�
    executeSelection = ret

    ' �t�@�C�����o�͂���
    If Not writer Is Nothing Then

        writer.destroy
        Set writer = Nothing
    End If
    
    Exit Function
err:

    ' �t�@�C���̏o�͂��L�����Z������
    If Not writer Is Nothing Then
        
        writer.Cancel
        Set writer = Nothing
    End If
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
        
End Function

' =========================================================
' �����R�[�h�X�V
'
' �T�v�@�@�@�F�e�[�u����`�ƃ��R�[�h���X�g����N�G���[�𐶐������s����B
' �@�@�@�@�@�@�{���\�b�h���Ńg�����U�N�V�������������Ă���
' �@�@�@�@�@�@����I�������ꍇ�́A�R�~�b�g�B�ُ�I�������ꍇ�́A���[���o�b�N�����B
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �@�@�@�@�@�@recordList ���R�[�h���X�g
' �߂�l�@�@�F���s����
'
' =========================================================
Private Function executeQuery(ByRef tableSheet As ValTableWorksheet _
                            , ByRef recordList As ValRecordList) As Long

    On Error GoTo err

    ' �e�[�u����`
    Dim table As ValDbDefineTable
    ' �e�[�u����`���擾����
    Set table = tableSheet.table
    
    executingSql = False
    executingSqlError = False
    
    ' �N�G���[�����I�u�W�F�N�g�𐶐�����Factory�I�u�W�F�N�g
    Dim dbObjFactory As New DbObjectFactory
    Set queryCreator = dbObjFactory.createQueryCreator(conn _
                                                            , applicationSetting.emptyCellReading _
                                                            , applicationSetting.getDirectInputChar _
                                                            , applicationSettingColFmt.getDbColFormatListByDbConn(conn) _
                                                            , applicationSetting.schemaUse _
                                                            , applicationSetting.getTableColumnEscapeByDbConn(conn))

    ' �g�����U�N�V�����I�u�W�F�N�g
    Dim transaction As New DBTransaction: transaction.init conn
    
    ' ���R�[�h����
    Dim recordCount As Long

    ' ���R�[�h���X�g�z��
    Dim records    As Variant
    ' ���R�[�h�z��
    Dim record     As Variant
    
    ' SQL
    Dim sql       As String
    ' ���ʌ���
    Dim resultCnt   As Long
    ' ���ʌ����i1�N�G���[�j
    Dim resultCntOne As Long

    ' �C���f�b�N�X
    Dim i   As Long
    
    ' �L�����Z���t���O���`�F�b�N����
    checkCancel

    ' ���R�[�h�z����擾����
    records = recordList.values
    
    ' ���R�[�h�z��iVariant�^�j���{���ɔz�񂩂ǂ������`�F�b�N����
    If IsArray(records) = True Then
    
        ' ���R�[�h�������擾����
        recordCount = VBUtil.arraySize(records)
    
        ' ��SQL�J�n�C�x���g
        RaiseEvent ExecuteSqlBegin(recordCount)
    
        ' �ŏ��C���f�b�N�X
        Dim min As Long: min = LBound(records, 1)
        ' �ő�C���f�b�N�X
        Dim max As Long: max = UBound(records, 1)

        ' 1������������
        For i = min To max

            ' �L�����Z���t���O���`�F�b�N����
            checkCancel

            ' 2�����z�񂩂� i �̕��������o��1�����z��ɕϊ�����
            record = VBUtil.convert2to1Array(records, i)

            ' SQL�𐶐�
            sql = createSql(table, record)
    
            ' SQL�����s
            If sql <> "" Then
            
                ' SQL���s�t���O��ON�ɂ���
                executingSql = True
                executingSqlError = False

                resultCntOne = ADOUtil.queryAction(conn, sql)
                
                ' SQL���s���G���[�t���O��ON�̏ꍇ
                ' �G���[�t���O�́AqueryAction���s���ɃG���[�����������ꍇ��
                ' �{���\�b�h�̃G���[�n���h�����O��ON�ɐݒ肳���B
                If executingSqlError = False Then
                
                    ' �G���[���������Ȃ������ꍇ�ɃJ�E���g����
                    resultCnt = resultCnt + resultCntOne
                    
                    ' ��SQL���s
                    RaiseEvent executeSql(tableSheet, recordList, i, sql, resultCntOne)
    
                End If

                ' SQL���s�t���O��OFF�ɂ���
                executingSql = False

            End If

        Next

        ' ��SQL�����C�x���g
        RaiseEvent ExecuteSqlSuccessful(resultCnt)

    End If

    ' �m�F���s��
    If applicationSetting.commitConfirm = applicationSetting.COMMIT_CONFIRM_YES Then
    
        ' ���b�Z�[�W�{�b�N�X�̖߂�l
        Dim resultMsg As Long
        
        resultMsg = VBUtil.showMessageBoxForYesNoCancel("�R�~�b�g���Ă���낵���ł����H" _
                           & vbNewLine _
                           & vbNewLine & "�E[�͂�] �� �R�~�b�g" _
                           & vbNewLine & "�E[������] �� ���[���o�b�N" _
                           & vbNewLine & "�E[�L�����Z��] �� ���[���o�b�N���ď����𒆒f" _
                        , ConstantsCommon.APPLICATION_NAME)

        ' �L�����Z�������i���[���o�b�N���Ē��f�j
        If resultMsg = WinAPI_User.IDCANCEL Then
        
            ' ��SQL�L�����Z���C�x���g
            RaiseEvent ExecuteSqlCancel(resultCnt)
            
            ' �������L�����Z������
            canceled = True: checkCancel
            
        ' �����������i���[���o�b�N���Čp���j
        ElseIf resultMsg = WinAPI_User.IDNO Then
        
            ' ��SQL�L�����Z���C�x���g
            RaiseEvent ExecuteSqlCancel(resultCnt)
            
            ' �߂�l�Ɏ��s������ݒ肵�Ċ֐��𔲂���
            executeQuery = 0
            
            Exit Function
        End If
        
    End If

    ' �g�����U�N�V�������R�~�b�g����
    transaction.commit

    ' �߂�l�Ɏ��s������ݒ�
    executeQuery = resultCnt

    Exit Function

err:

    ' �G���[�n���h���ŕʂ̊֐����Ăяo���ƃG���[��񂪏����Ă��܂����Ƃ�����̂�
    ' �\���̂ɃG���[����ۑ����Ă���
    Dim errT As errInfo: errT = VBUtil.swapErr

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
        RaiseEvent ExecuteSqlError(tableSheet, recordList, i, sql, sqlState, err.Description)

        ' SQL�G���[�����������ꍇ�ɏ������p������
        If applicationSetting.sqlErrorHandling = applicationSetting.SQL_ERROR_HANDLING_RESUME Then
        
            ' SQL���s���G���[�t���O�𗧂Ă�
            executingSqlError = True
            Resume Next
        End If
        
        ' �G���[���b�Z�[�W��ݒ�
        errT.Description = ConstantsError.ERR_DESC_SQL_EXECUTE_FAILED _
                        & vbNewLine _
                        & vbNewLine _
                        & "[�ڍ׏��]" & vbNewLine _
                        & errT.Number & vbNewLine & errT.Description
                          
        ' �G���[�ԍ���ݒ�
        errT.Number = ConstantsError.ERR_NUMBER_SQL_EXECUTE_FAILED
        
    End If
    
    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' �����R�[�h�X�V
'
' �T�v�@�@�@�F�e�[�u����`�ƃ��R�[�h���X�g����N�G���[�𐶐����t�@�C���ɏo�͂���B
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �@�@�@�@�@�@recordList ���R�[�h���X�g
' �߂�l�@�@�F���s����
'
' =========================================================
Private Function executeOutputQueryFile(ByRef tableSheet As ValTableWorksheet _
                                      , ByRef recordList As ValRecordList) As Long

    On Error GoTo err

    ' �e�[�u����`
    Dim table As ValDbDefineTable
    ' �e�[�u����`���擾����
    Set table = tableSheet.table

    ' �N�G���[�����I�u�W�F�N�g�𐶐�����Factory�I�u�W�F�N�g
    Dim dbObjFactory As New DbObjectFactory
    Set queryCreator = dbObjFactory.createQueryCreator(conn _
                                                            , applicationSetting.emptyCellReading _
                                                            , applicationSetting.getDirectInputChar _
                                                            , applicationSettingColFmt.getDbColFormatListByDbConn(conn) _
                                                            , applicationSetting.schemaUse _
                                                            , applicationSetting.getTableColumnEscapeByDbConn(conn))

    ' ���R�[�h����
    Dim recordCount As Long

    ' ���R�[�h���X�g�z��
    Dim records    As Variant
    ' ���R�[�h�z��
    Dim record     As Variant
    
    ' SQL
    Dim sql       As String
    ' ���ʌ���
    Dim resultCnt   As Long
    ' ���ʌ����i1�N�G���[�j
    Dim resultCntOne As Long

    ' �C���f�b�N�X
    Dim i   As Long
    
    ' �L�����Z���t���O���`�F�b�N����
    checkCancel

    ' ���R�[�h�z����擾����
    records = recordList.values
    
    ' ���R�[�h�z��iVariant�^�j���{���ɔz�񂩂ǂ������`�F�b�N����
    If IsArray(records) = True Then
    
        ' ���R�[�h�������擾����
        recordCount = VBUtil.arraySize(records)
    
        ' ��SQL�J�n�C�x���g
        RaiseEvent ExecuteSqlBegin(recordCount)
    
        ' �ŏ��C���f�b�N�X
        Dim min As Long: min = LBound(records, 1)
        ' �ő�C���f�b�N�X
        Dim max As Long: max = UBound(records, 1)

        ' 1������������
        For i = min To max

            ' �L�����Z���t���O���`�F�b�N����
            checkCancel

            ' 2�����z�񂩂� i �̕��������o��1�����z��ɕϊ�����
            record = VBUtil.convert2to1Array(records, i)

            ' SQL�𐶐�
            sql = createSql(table, record)

            ' SQL�����s
            If sql <> "" Then
            
                ' �J�E���g����
                resultCnt = resultCnt + 1
                
                ' ���s�R�[�h��ϊ�����
                sql = replace(sql, vbLf, newline)
                ' �Z�~�R������t������
                sql = sql & ";" & newline
                
                ' �t�@�C�����o�͂���
                writer.writeText sql, True
                
                
                ' ��SQL���s
                RaiseEvent executeSql(tableSheet, recordList, i, sql, resultCntOne)

            End If

        Next

        ' ��SQL�����C�x���g
        RaiseEvent ExecuteSqlSuccessful(resultCnt)

    End If

    ' �߂�l�Ɏ��s������ݒ�
    executeOutputQueryFile = resultCnt

    Exit Function

err:

    ' �G���[�n���h���ŕʂ̊֐����Ăяo���ƃG���[��񂪏����Ă��܂����Ƃ�����̂�
    ' �\���̂ɃG���[����ۑ����Ă���
    Dim errT As errInfo: errT = VBUtil.swapErr

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' �����R�[�h�폜�X�V
'
' �T�v�@�@�@�F�e�[�u����̑S���R�[�h���폜����SQL�𔭍s����B
' �@�@�@�@�@�@�{���\�b�h�Ńg�����U�N�V�������������Ă���
' �@�@�@�@�@�@����I�������ꍇ�́A�R�~�b�g�B�ُ�I�������ꍇ�́A���[���o�b�N�����B
' �����@�@�@�F
' �߂�l�@�@�F���s����
'
' =========================================================
Public Function executeQueryDeleteOfAllTable() As Long

    On Error GoTo err

    ' �e�[�u���V�[�g
    Dim tableSheet  As ValTableWorksheet

    ' �e�[�u���V�[�g�ǂݍ��݃I�u�W�F�N�g������������
    Dim tsr As New ExeTableSheetReader
    Set tsr.conn = conn
    Set tsr.sheet = sheet
    
    Set tableSheet = tsr.readTableInfo

    Set readTableSheet = tableSheet

    executingSql = False
    
    ' �N�G���[�����I�u�W�F�N�g�𐶐�����Factory�I�u�W�F�N�g
    Dim dbObjFactory As New DbObjectFactory
    Set queryCreator = dbObjFactory.createQueryCreator(conn _
                                                            , applicationSetting.emptyCellReading _
                                                            , applicationSetting.getDirectInputChar _
                                                            , applicationSettingColFmt.getDbColFormatListByDbConn(conn) _
                                                            , applicationSetting.schemaUse _
                                                            , applicationSetting.getTableColumnEscapeByDbConn(conn))

    ' SQL
    Dim sql       As String
    ' ���ʌ���
    Dim resultCnt As Long

    ' �g�����U�N�V�����I�u�W�F�N�g
    Dim transaction As New DBTransaction: transaction.init conn
    
    ' �L�����Z���t���O���`�F�b�N����
    checkCancel

    ' SQL�𐶐�
    sql = createSql(tableSheet.table, Nothing)
    createdSql = sql
    
    ' SQL�����s
    If sql <> "" Then
    
        ' SQL���s�t���O��ON�ɂ���
        executingSql = True

        resultCnt = resultCnt + ADOUtil.queryAction(conn, sql)

        ' SQL���s�t���O��OFF�ɂ���
        executingSql = False

    End If

    ' �폜�m�F���s��
    If VBUtil.showMessageBoxForYesNo("�S�Ẵ��R�[�h���폜���܂����B�R�~�b�g���Ă���낵���ł����H" _
     , ConstantsCommon.APPLICATION_NAME) = WinAPI_User.IDNO Then
    
        ' �������L�����Z������
        canceled = True: checkCancel
    End If
    
    ' �g�����U�N�V�������R�~�b�g����
    transaction.commit

    ' �߂�l�Ɏ��s������ݒ�
    executeQueryDeleteOfAllTable = resultCnt

    Exit Function

err:

    ' �G���[�n���h���ŕʂ̊֐����Ăяo���ƃG���[��񂪏����Ă��܂����Ƃ�����̂�
    ' �\���̂ɃG���[����ۑ����Ă���
    Dim errT As errInfo: errT = VBUtil.swapErr

    ' SQL���s���ɃG���[������
    If executingSql = True Then

        ' �G���[���b�Z�[�W��ݒ�
        errT.Description = ConstantsError.ERR_DESC_SQL_EXECUTE_FAILED _
                        & vbNewLine _
                        & vbNewLine _
                        & "[�ڍ׏��]" & vbNewLine _
                        & errT.Number & vbNewLine & errT.Description
                          
        ' �G���[�ԍ���ݒ�
        errT.Number = ConstantsError.ERR_NUMBER_SQL_EXECUTE_FAILED
        
    End If
    
    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ��SQL����
'
' �T�v�@�@�@�F�N�G���[��ނɑΉ�����SQL�𐶐�����B
'
' �����@�@�@�Ftable  �e�[�u����`
' �@�@�@�@�@�@record ���R�[�h�z��
' �߂�l�@�@�FSQL
'
' =========================================================
Public Function createSql(ByRef table As ValDbDefineTable, ByRef record As Variant) As String

    ' �N�G���ϊ��I�u�W�F�N�g������������
    queryConverter.initIfFirst Me.conn, Me.queryParameterList
    ' �N�G���p�����[�^�Ń��R�[�h�̓��e��ϊ�����
    ' ---record�̓��e��ϊ�����������l�������ASELECT�Ɋւ��Ă̓N�G���S�̂�ϊ�����������̗p���Ă��邽��
    ' -- �X�V�n�������悤�ɃN�G���S�̂�ϊ�����������̗p����
    ' -- ���Ƀe�[�u������J�������ɃN�G���p�����[�^�Ɠ����ϐ���������ꍇ�ɂ̓N�G�����������Ȃ��ƂɂȂ邪�A���̂悤�ȃP�[�X�͂قƂ�ǂȂ��ƍl���l�����Ȃ�
    'convertRecordValueByQueryParameter record

    Select Case queryType
    
        Case QUERY_TYPE_INSERT_UPDATE
        
            On Error GoTo errExecuteSelect
            
            Dim recordSet As Object
            
            ' ---------------------------------------------------------
            ' PK�������쐬����
            ' ---------------------------------------------------------
            Dim condition As New ValSelectCondition
            
            Dim i As Long: i = 1
            
            ' �J����
            Dim column     As ValDbDefineColumn
            ' �J�������X�g
            Dim columnList As ValCollection
            
            ' �e�[�u��������(PK)
            Dim tableConstPk    As New ValDbDefineTableConstraints
            ' PK�J�����ł��邩������킷�t���O
            Dim isColumnPk      As Boolean
            
            Dim tableConstTmp   As ValDbDefineTableConstraints
            ' �e�[�u�����񃊃X�g����PK������擾����
            For Each tableConstTmp In table.constraintsList.col
            
                If tableConstTmp.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk Then
                
                    Set tableConstPk = tableConstTmp
                    Exit For
                End If
            Next
            
            ' �J�������X�g���擾����
            Set columnList = table.columnList
            
            ' �J�������X�g��1������������
            For Each column In columnList.col
                    
                ' PK����ł��邩�ǂ����𔻒肷��
                If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
                
                    isColumnPk = False
                Else
                
                    isColumnPk = True
                End If
                
                ' �J������PK�ł���ꍇ
                If isColumnPk = True Then
                
                    If record(i) = "" Then
                    
                        condition.setCondition column.columnName, "NULL", Empty
                    Else
                    
                        condition.setCondition column.columnName, record(i), Empty
                    End If
                    
                End If
                
                i = i + 1
            Next
            ' ---------------------------------------------------------
            
            Dim selectSql As String
            selectSql = queryCreator.createSelect(table, condition)
            selectSql = queryConverter.convertSql(selectSql)
            
            Set recordSet = ADOUtil.querySelect(conn, selectSql)

            ' �擾���R�[�h��0���̏ꍇ
            If recordSet.EOF Then
                
                createSql = queryCreator.createInsert(table, record)
                
            ' �擾���R�[�h��1���̏ꍇ
            Else
                
                createSql = queryCreator.createUpdate(table, record)
                
            End If
            
            ' ���R�[�h�Z�b�g���������
            ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing

            GoTo exitInsertUpdate
            
errExecuteSelect:

            ' ���R�[�h�Z�b�g���������
            ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
            
            createSql = queryCreator.createInsert(table, record)
    
exitInsertUpdate:
    
        Case QUERY_TYPE_INSERT
        
            createSql = queryCreator.createInsert(table, record)
        
        Case QUERY_TYPE_UPDATE
        
            createSql = queryCreator.createUpdate(table, record)
        
        Case QUERY_TYPE_DELETE
        
            createSql = queryCreator.createDelete(table, record)
        
        Case QUERY_TYPE_DELETE_ALL
        
            createSql = queryCreator.createDeleteOfAllTable(table)
    
    End Select

    If queryParameterList.list.count > 0 Then
        ' �N�G���p�����[�^������ꍇ�ɕϊ�����
        createSql = queryConverter.convertSql(createSql)
    End If

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

' =========================================================
' �����R�[�h�̒l���N�G���p�����[�^�ɂ���ĕϊ����鏈��
'
' �T�v�@�@�@�F
' �����@�@�@�Frecord ���R�[�h
' �߂�l�@�@�F
'
' =========================================================
Private Sub convertRecordValueByQueryParameter(ByRef record As Variant)

    If queryParameterList.list.count <= 0 Then
        ' �N�G���p�����[�^���Ȃ��ꍇ�́A�������Ȃ�
        Exit Sub
    End If

    Dim i As Long
    
    For i = LBound(record) To UBound(record)
    
        record(i) = queryConverter.convertSql(CStr(record(i)))
    Next

End Sub

' --------------------������
' �e�[�u���V�[�g�ǂݍ��݃N���X�̃C�x���g��{�N���X�̃C�x���g�Ƃ��Ă��̂܂ܔ��s����
Private Sub tableSheetReader_readColumnListBegin(ByVal cnt As Long)

    RaiseEvent readColumnListBegin(cnt)
End Sub

Private Sub tableSheetReader_readColumnList(columnList As ValCollection)

    RaiseEvent readColumnList(columnList)
End Sub

Private Sub tableSheetReader_readColumn(column As ValDbDefineColumn)

    RaiseEvent readColumn(column)
End Sub

Private Sub tableSheetReader_readTableInfoBegin()

    RaiseEvent readTableInfoBegin
End Sub

Private Sub tableSheetReader_readTableInfo(table As ValTableWorksheet)

    RaiseEvent readTableInfo(table)
End Sub

Private Sub tableSheetReader_readRecordListBegin(ByVal cnt As Long)

    RaiseEvent readRecordListBegin(cnt)
End Sub

Private Sub tableSheetReader_readRecordList(recordList As ValRecordList)

    RaiseEvent readRecordList(recordList)
End Sub

Private Sub tableSheetReader_readRecordListBySelectionBegin(ByVal cnt As Long)

    RaiseEvent readRecordListBySelectionBegin(cnt)
End Sub

Private Sub tableSheetReader_readRecordListBySelectionOne(recordList As ValRecordList)

    RaiseEvent readRecordListBySelectionOne(recordList)
End Sub

Private Sub tableSheetReader_readRecordListBySelection(recordLists As ValCollection)

    RaiseEvent readRecordListBySelection(recordLists)
End Sub
' --------------------������
