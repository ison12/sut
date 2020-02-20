VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSnapQueryExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�i�b�v�V���b�g�̃N�G�������s����N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/01/05�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ________________________________________________________
' �C�x���g
' ________________________________________________________
' =========================================================
' ��SQL���X�g���s�C�x���g�i�J�n�j
'
' �T�v�@�@�@�FSQL���s���ɌĂяo�����C�x���g
' �����@�@�@�Fcount ����
'
' =========================================================
Public Event ExecuteSqlListBegin(ByVal count As Long)

' =========================================================
' ��SQL���X�g���s�C�x���g�i�����j
'
' �T�v�@�@�@�FSQL���s���ɌĂяo�����C�x���g
' �����@�@�@�F
'
' =========================================================
Public Event ExecuteSqlListComplete()

' =========================================================
' ��SQL���s�C�x���g
'
' �T�v�@�@�@�FSQL���s����O�ɌĂяo�����C�x���g
' �����@�@�@�Fsql        SQL�X�e�[�g�����g
'
' =========================================================
Public Event ExecuteSqlBegin(ByRef sql As String)

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
' �����@�@�@�Fsheet          ���[�N�V�[�g
' �@�@�@�@�@�FrecordSetIndex ���R�[�h�Z�b�g�C���f�b�N�X
' �@�@�@�@�@�Fsql            SQL�X�e�[�g�����g
' �@�@�@�@�@�Fsqlstate       SQL�X�e�[�^�X
' �@�@�@�@�@�Fmessage        ���b�Z�[�W
'
' =========================================================
Public Event ExecuteSqlError(ByRef sheet As Worksheet, ByVal recordSetIndex As Long, ByRef sql As String, ByRef sqlState As String, ByRef message As String)

' ---------------------------------------------------------
' ���v���p�e�B
' ---------------------------------------------------------
' DB�R�l�N�V����
Public conn          As Object
' �L�����Z���t���O
Public canceled      As Boolean

' SQL���s���t���O
Private executingSql      As Boolean

' �A�v���P�[�V�����ݒ���
Private m_applicationSetting As ValApplicationSetting

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
' ���N�G�����s
'
' �T�v�@�@�@�F�ΏۂƂȂ�N�G�������s����
' �����@�@�@�FsqlDefineWorkSheet  �N�G�����X�g
' �߂�l�@�@�F���R�[�h�擾���
'
' =========================================================
Public Function execute(ByRef sqlDefineWorkSheet As ValSnapSqlDefineWorkSheet) As ValSnapRecordsSet

    On Error GoTo err

    ' �e����^�������R�[�h
    Dim affectedRecord As Long
    
    Dim i, j As Long
    
    ' �w�b�_�z��
    Dim headerList As Variant
    ' �t�B�[���h
    Dim field   As Object
    
    ' ���R�[�h�Z�b�g
    Dim recordSet   As Object
    ' ���R�[�h�z��
    Dim recordList  As Variant
    ' ���R�[�h�z��i�s�T�C�Y�j
    Dim recordListSize As Long
    ' ���R�[�h������
    Dim recordAllCount As Long

    ' ���R�[�h�Z�b�g�C���f�b�N�X
    Dim recordSetIndex As Long
    
    If sqlDefineWorkSheet.snapSqlDefine.sqlList.count <= 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_SQL_EMPTY _
                , err.Source _
                , ConstantsError.ERR_DESC_SQL_EMPTY
    End If

    ' ���R�[�h�Z�b�g�C���f�b�N�X
    recordSetIndex = 1
    
    ' �L�����Z���t���O���`�F�b�N����
    checkCancel
    
    ' �g�����U�N�V�����I�u�W�F�N�g
    Dim transaction As DBTransaction
    Set transaction = Nothing
    
    Set transaction = New DBTransaction
    transaction.init conn
        
    Dim varOfSqlList   As Variant
    Dim varOfParamList As Variant
    
    Dim sqlDefineSql   As ValSnapSqlDefineSql
    Dim sqlDefineParam As ValSnapSqlDefineParam
    
    Dim sqlTemplate As String
    Dim sql As String
    
    executingSql = False
    
    RaiseEvent ExecuteSqlListBegin(sqlDefineWorkSheet.snapSqlDefine.sqlList.count)
    
    ' �߂�l��������
    Set execute = New ValSnapRecordsSet
    recordAllCount = 0
    
    Dim snapRecords As ValSnapRecords
    
    For Each varOfSqlList In sqlDefineWorkSheet.snapSqlDefine.sqlList.col
    
        ' �L�����Z���t���O���`�F�b�N����
        checkCancel
    
        Set sqlDefineSql = varOfSqlList
        
        ' -------------------------------------------------------------
        ' SQL�i������j�𐶐�����i�p�����[�^�̒u�������j
        ' -------------------------------------------------------------
        sqlTemplate = sqlDefineSql.sql
        sql = sqlTemplate
        
        For Each varOfParamList In sqlDefineWorkSheet.snapSqlDefine.paramList.col
        
            Set sqlDefineParam = varOfParamList
            sql = replace$(sql, sqlDefineParam.name, DBUtil.convertQueryLiteral(ADOUtil.getDBMSType(dbConn), String_, sqlDefineParam.value))
        Next
    
        ' -------------------------------------------------------------
        ' SQL�����s����
        ' -------------------------------------------------------------
        ' �L�����Z���t���O���`�F�b�N����
        checkCancel
        
        ' �����̂͂��߂Ƀt���O���I�t�ɂ��Ă���
        executingSql = False
    
        RaiseEvent ExecuteSqlBegin(sql)
    
        ' SQL���s���t���O��ON�ɂ���
        executingSql = True
        
        ' �N�G���[�����s����
        Set recordSet = ADOUtil.queryBatch(conn, sql, affectedRecord)
        
        ' SQL���s���t���O��OFF�ɂ���
        executingSql = False
        
        ' -------------------------------------------------------------
        ' ���R�[�h���擾����
        ' -------------------------------------------------------------
        
        ' UPDATE�Ȃǂ́A���R�[�h��Ԃ��Ȃ��N�G���̏ꍇ State = 0 (adStateClosed) �ɂȂ�
        If recordSet.state = 0 Then
        
        Else
        
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
            
            ' ---------------------------------------------------
            
            ' ---------------------------------------------------
            ' ���R�[�h���̎擾
            ' ---------------------------------------------------
            ' �擾���R�[�h��0���ɂȂ�܂Ń��R�[�h�Z�b�g����������
            recordList = Empty
            Do While Not recordSet.EOF
            
                ' �L�����Z���t���O���`�F�b�N����
                checkCancel
                
                ' ���R�[�h�Z�b�g���烌�R�[�h���X�g���擾����
                recordList = getRecordList(recordSet, 0)
                
                ' ���R�[�h���X�g�̃T�C�Y���擾����
                ' ADO�̃��R�[�h�Z�b�g����GetRows�Ŏ擾�����z���(��~�s)�Ńf�[�^���i�[����Ă���
                recordListSize = VBUtil.arraySize(recordList, 2)
        
                ' ���R�[�h���X�g���z��ł͂Ȃ��ꍇ
                If IsArray(recordList) = False Then
                
                    ' ���̃u���b�N�ɓ��邱�Ƃ͂���̂��H
                    GoTo loop_end
                End If
                
                ' ���R�[�h���X�g��ϊ�����
                convertRecordList recordList
                
                recordAllCount = recordAllCount + recordListSize
            Loop
    
            Set snapRecords = New ValSnapRecords
            snapRecords.sql = sql
            snapRecords.memo = sqlDefineSql.memo
            snapRecords.headerList = headerList
            snapRecords.recordList = recordList
            Set snapRecords.pkList = sqlDefineSql.pkList

            execute.list.setItem snapRecords

        End If
            
loop_end:
            
        recordSetIndex = recordSetIndex + 1
        
        ' ���R�[�h�Z�b�g���������
        ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
        RaiseEvent ExecuteSqlSuccessful(sql)

    Next
    
    ' �O�̂��߃��R�[�h�Z�b�g���������
    ADOUtil.closeRecordSet recordSet: Set recordSet = Nothing
    
    ' �g�����U�N�V�������R�~�b�g����
    If Not transaction Is Nothing Then
        transaction.commit
    End If
    
    ' -------------------------------------------------------
    
    execute.name = sqlDefineWorkSheet.snapSqlDefine.name
    execute.recordCount = recordAllCount
    execute.getDate = Now
    
    RaiseEvent ExecuteSqlListComplete
    
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
        RaiseEvent ExecuteSqlError(sqlDefineWorkSheet.sheet, recordSetIndex, sql, sqlState, err.Description)
        
    Else

        ' ��SQL���s���G���[
        RaiseEvent ExecuteSqlError(sqlDefineWorkSheet.sheet, recordSetIndex, sql, Empty, err.Description)
    
    End If

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

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
' �����@�@�@�FrecordList   ���R�[�h���X�g
'
' =========================================================
Private Function convertRecordList(ByRef recordList As Variant)

    ' ���R�[�h�z��̗v�f����ւ��ɂ���ĐV���ɐ��������z��
    Dim temp As Variant
    
    ' ���R�[�h�z��̃f�[�^�ϊ�����
    ' RecordSet.GetRows���Ԃ��ꂽ�z��́A(col,rec)�ƂȂ��Ă��邽��(rec,col)�ƂȂ�悤�ɓ���ւ������{����
    temp = convertArrayValues(recordList)
    
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
Private Function convertArrayValues(ByRef v As Variant) As Variant
    
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
            tempArray(X, Y) = ExcelUtil.convertCellStrValue(value)
        
        Next Y
    Next X
    
    convertArrayValues = tempArray

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

