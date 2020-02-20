Attribute VB_Name = "ADOUtil"
Option Explicit

' *********************************************************
' ADO���ȕւɗ��p���邽�߂̃��[�e�B���e�B���W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2007/12/01�@�V�K�쐬
'
' ���L�����FMicrosoft ActiveX Data Objects Library���Q��
' �@�@�@�@�@Ver 2.5 �œ���m�F
' *********************************************************

' ObjectStateEnum�̃R�s�[
' �I�u�W�F�N�g���J���Ă��邩���Ă��邩��f�[�^ �\�[�X�ɐڑ�������R�}���h�����s������܂��̓f�[�^���擾�����ǂ�����\���܂��
' http://msdn.microsoft.com/ja-jp/library/cc389847.aspx
Public Enum ADOConnectStatusConstants

    adStateClosed = 0        ' �I�u�W�F�N�g�����Ă��邱�Ƃ������܂��B
    adStateOpen = 1          ' �I�u�W�F�N�g���J���Ă��邱�Ƃ������܂��B
    adStateConnecting = 2    ' �I�u�W�F�N�g���ڑ����Ă��邱�Ƃ������܂��B
    adStateExecuting = 4     ' �I�u�W�F�N�g���R�}���h�����s���ł��邱�Ƃ������܂��B
    adStateFetching = 8      ' �I�u�W�F�N�g�̍s���擾����Ă��邱�Ƃ������܂��B

End Enum

' CursorTypeEnum�̃R�s�[
' http://msdn.microsoft.com/ja-jp/library/cc389787.aspx
Public Enum ADOCursorTypeEnum

    adOpenDynamic = 2         ' ���I�J�[�\�����g���܂��B�ق��̃��[�U�[�ɂ��ǉ��A�ύX�A����э폜���m�F�ł��܂��B�v���o�C�_���u�b�N�}�[�N���T�|�[�g���Ă��Ȃ��ꍇ�������ARecordset ���ł̂��ׂĂ̓���������܂��B
    adOpenForwardOnly = 0     ' ����l�ł��B�O����p�J�[�\�����g���܂��B���R�[�h�̃X�N���[���������O�����Ɍ��肳��Ă��邱�Ƃ������A�ÓI�J�[�\���Ɠ������������܂��BRecordset �̃X�N���[���� 1 �񂾂��ŏ\���ȏꍇ�́A����ɂ���ăp�t�H�[�}���X������ł��܂��B
    adOpenKeyset = 1          ' �L�[�Z�b�g �J�[�\�����g���܂��B�ق��̃��[�U�[���ǉ��������R�[�h�͕\���ł��Ȃ��_�������A���I�J�[�\���Ɠ������A������ Recordset ����ق��̃��[�U�[���폜�������R�[�h�̓A�N�Z�X�ł��܂���B�ق��̃��[�U�[���ύX�����f�[�^�͕\���ł��܂��B
    adOpenStatic = 3          ' �L�[�Z�b�g �J�[�\�����J���܂��B�f�[�^�̌����܂��̓��|�[�g�̍쐬�Ɏg�p���邽�߂́A���R�[�h�̐ÓI�R�s�[�ł��B�ق��̃��[�U�[�ɂ��ǉ��A�ύX�A�܂��͍폜�͕\������܂���B
    adOpenUnspecified = -1    '  �J�[�\���̎�ނ��w�肵�܂���B

End Enum

' =========================================================
' ��DB�ڑ��֐�
'
' �T�v�@�@�@�FDB�ɐڑ�����
' �����@�@�@�FconnString �ڑ�������
'
' �߂�l�@�@�F�R�l�N�V�����I�u�W�F�N�g
'
' =========================================================
Public Function connectDb(ByVal connString As String) As Object

    Dim conn As Object
    
    Set conn = CreateObject("ADODB.Connection")
    
    conn.ConnectionString = connString
        
    conn.Open
    
    Set connectDb = conn
    
End Function

' =========================================================
' ��DB�ؒf�֐�
'
' �T�v�@�@�@�F�A�N�e�B�u��DB��ؒf����
' �����@�@�@�Fconn �R�l�N�V�����I�u�W�F�N�g
'
' =========================================================
Public Sub closeDB(ByRef conn As Object)

    If Not conn Is Nothing Then
    
        conn.Close
    End If
    
    Set conn = Nothing

End Sub

' =========================================================
' ��DB���擾
'
' �T�v�@�@�@�F�f�[�^�\�[�X����DB�����擾����
' �����@�@�@�FconnStr �ڑ�������
'
' �߂�l�@�@�FDB��
'
' =========================================================
Public Function getDBName(ByRef conn As Object) As String

    ' �f�[�^�x�[�X�����擾����
    getDBName = conn.defaultdatabase
    
End Function

' =========================================================
' ���N�G���ꊇ���s
'
' �T�v�@�@�@�F�N�G�����ꊇ���s����
' �����@�@�@�Fconn �R�l�N�V�����I�u�W�F�N�g
' �@�@�@�@�@�@sql  SQL�X�e�[�g�����g
'             cnt  �擾����
'             cursorType �J�[�\���^�C�v
'
' �߂�l�@�@�F���R�[�h�Z�b�g�I�u�W�F�N�g
'
' =========================================================
Public Function queryBatch(ByRef conn As Object _
                          , ByVal sql As String _
                          , Optional ByRef cnt As Long _
                          , Optional ByVal cursorType As ADOCursorTypeEnum = ADOCursorTypeEnum.adOpenForwardOnly) As Object

    On Error GoTo err
    
    ' ���ϐ���`
    Dim cmd As Object
    Dim rec As Object
    
    ' ���R�}���h�I�u�W�F�N�g��������
    Set cmd = CreateObject("ADODB.Command")
    
    cmd.ActiveConnection = conn
    cmd.CommandType = 1 ' adCmdText
    cmd.CommandText = sql
    
    ' �N�G���[�����s����
    Set rec = cmd.execute(cnt)
    
    ' �����R�[�h�Z�b�g��������
    'Set rec = CreateObject("ADODB.Recordset")
    
    ' �N�G���[�����s����
    'rec.Open Source:=cmd, cursorType:=cursorType
    
    ' ���߂�l��ݒ肷��
    Set queryBatch = rec
   
    Set cmd = Nothing
    Set rec = Nothing

    Exit Function
err:

    ' �G���[�n���h���ŕʂ̊֐����Ăяo���ƃG���[��񂪏����Ă��܂����Ƃ�����̂�
    ' �\���̂ɃG���[����ۑ����Ă���
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    ADOUtil.closeRecordSet rec
    Set rec = Nothing
    Set cmd = Nothing
        
    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ��Select���s
'
' �T�v�@�@�@�FSelect�������s����
' �����@�@�@�Fconn �R�l�N�V�����I�u�W�F�N�g
' �@�@�@�@�@�@sql  SQL�X�e�[�g�����g
'             cnt  �擾����
'             cursorType �J�[�\���^�C�v
'
' �߂�l�@�@�F���R�[�h�Z�b�g�I�u�W�F�N�g
'
' =========================================================
Public Function querySelect(ByRef conn As Object _
                          , ByVal sql As String _
                          , Optional ByRef cnt As Long _
                          , Optional ByVal cursorType As ADOCursorTypeEnum = ADOCursorTypeEnum.adOpenForwardOnly) As Object

    On Error GoTo err
    
    ' ���ϐ���`
    Dim cmd As Object
    Dim rec As Object
    
    ' ���R�}���h�I�u�W�F�N�g��������
    Set cmd = CreateObject("ADODB.Command")
    
    cmd.ActiveConnection = conn
    cmd.CommandType = 1 ' adCmdText
    cmd.CommandText = sql
    
    ' �N�G���[�����s����
    'Set rec = cmd.execute(cnt)
    
    ' �����R�[�h�Z�b�g��������
    Set rec = CreateObject("ADODB.Recordset")
    
    ' �N�G���[�����s����
    rec.Open Source:=cmd, cursorType:=cursorType
    
    ' ���߂�l��ݒ肷��
    Set querySelect = rec
   
    Set cmd = Nothing
    Set rec = Nothing

    Exit Function
err:

    ' �G���[�n���h���ŕʂ̊֐����Ăяo���ƃG���[��񂪏����Ă��܂����Ƃ�����̂�
    ' �\���̂ɃG���[����ۑ����Ă���
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    ADOUtil.closeRecordSet rec
    Set rec = Nothing
    Set cmd = Nothing
        
    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Function

' =========================================================
' ���A�N�V�����N�G���[���s
'
' �T�v�@�@�@�FInsert�EUpdate�EDelete�������s����
' �����@�@�@�Fconn �R�l�N�V�����I�u�W�F�N�g
' �@�@�@�@�@�@sql  SQL�X�e�[�g�����g
'
' �߂�l�@�@�F�X�V����
'
' =========================================================
Public Function queryAction(ByRef conn As Object, ByVal sql As String) As Long

    ' ���ϐ���`
    Dim cmd As Object
    Dim rec As Object
    
    Dim cnt As Long
    
    ' ���R�}���h�I�u�W�F�N�g��������
    Set cmd = CreateObject("ADODB.Command")
    
    cmd.ActiveConnection = conn
    cmd.CommandType = 1 ' adCmdText
    cmd.CommandText = sql
    
    
    ' ���N�G���[�����s����
    Set rec = cmd.execute(cnt)
    
    ' ���߂�l��ݒ肷��
    queryAction = cnt
    
    Set cmd = Nothing
    Set rec = Nothing

 End Function

' =========================================================
' �����R�[�h�Z�b�g���
'
' �T�v�@�@�@�F�A�N�e�B�u�ȃ��R�[�h�Z�b�g���������
' �����@�@�@�Frec ���R�[�h�Z�b�g�I�u�W�F�N�g
'
' =========================================================
Public Sub closeRecordSet(ByRef rec As Object)

    If Not rec Is Nothing Then
    
        ' ���R�[�h�Z�b�g���J���Ă���ꍇ�̂݃N���[�Y���s��
        If rec.state <> ADOConnectStatusConstants.adStateClosed Then
        
            rec.Close
        End If
        
    End If
    
    Set rec = Nothing
    
    Exit Sub
    
End Sub

' =========================================================
' ��DBMS��ގ擾
'
' �T�v�@�@�@�F�R�l�N�V�����I�u�W�F�N�g����DBMS�̎�ނ��擾����
' �����@�@�@�Fconn �R�l�N�V�����I�u�W�F�N�g
'
' =========================================================
Public Function getDBMSType(ByRef conn As Object) As DbmsType

    ' �f�[�^�x�[�X��
    Dim dbmsName As String
    
    ' �f�[�^�x�[�X�����擾
    dbmsName = conn.properties.item("DBMS Name")
    
    ' MySQL�f�[�^�x�[�X
    If InStr(LCase$(dbmsName), "mysql") > 0 Then
    
        getDBMSType = DbmsType.MySQL
    
    ' PostgreSQL�f�[�^�x�[�X
    ElseIf InStr(LCase$(dbmsName), "postgresql") > 0 Then
    
        getDBMSType = DbmsType.PostgreSQL
    
    ' Oracle�f�[�^�x�[�X
    ElseIf InStr(LCase$(dbmsName), "oracle") > 0 Then
    
        getDBMSType = DbmsType.Oracle
    
    ' SQL Server�f�[�^�x�[�X
    ElseIf InStr(LCase$(dbmsName), "microsoft sql server") > 0 Then
    
        getDBMSType = DbmsType.MicrosoftSqlServer
        
    
    ' Access�f�[�^�x�[�X
    ElseIf InStr(LCase$(dbmsName), "access") > 0 Or InStr(LCase$(dbmsName), "ms jet") > 0 Then
    
        getDBMSType = DbmsType.MicrosoftAccess
        
    ' Symfoware�f�[�^�x�[�X
    ElseIf InStr(LCase$(dbmsName), "symfoware") > 0 Then
    
        getDBMSType = DbmsType.Symfoware
    ' ���ʂł��Ȃ��ꍇ
    Else
    
        getDBMSType = DbmsType.Other
    
    End If

End Function

' =========================================================
' ��DBMS��ގ擾
'
' �T�v�@�@�@�F�R�l�N�V�����I�u�W�F�N�g����DBMS�̎�ނ��擾����
' �����@�@�@�Fconn �R�l�N�V�����I�u�W�F�N�g
'
' =========================================================
Public Function getDBMSTypeByConnStr(ByVal connStr As String) As DbmsType

    On Error GoTo err
    
    Dim conn As Object
    
    ' DB�ɐڑ�����
    Set conn = ADOUtil.connectDb(connStr)
    
    getDBMSTypeByConnStr = ADOUtil.getDBMSType(conn)
    
    ' DB��ؒf����
    ADOUtil.closeDB conn
    
    Exit Function
    
err:

    ' ����n��
    ' DB��ؒf����
    ADOUtil.closeDB conn
    
    Set conn = Nothing
    
End Function

' =========================================================
' ���R�l�N�V�����̃X�e�[�^�X�擾
'
' �T�v�@�@�@�F�R�l�N�V�����I�u�W�F�N�g����X�e�[�^�X���擾����
' �����@�@�@�Fconn �R�l�N�V�����I�u�W�F�N�g
' �߂�l�@�@�FADOConnectStatusConstants
'
' =========================================================
Public Function getConnectionStatus(ByRef conn As Object) As ADOConnectStatusConstants

    If conn Is Nothing Then
    
        getConnectionStatus = adStateClosed
        Exit Function
    End If

    ' �R�l�N�V�����̃X�e�[�^�X���擾����
    getConnectionStatus = conn.state

End Function
