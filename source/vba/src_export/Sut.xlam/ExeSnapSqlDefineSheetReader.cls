VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSnapSqlDefineSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �V�[�g����X�i�b�v�V���b�gSQL��`����ǂݎ�鏈�����`�����N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/01/04�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���[�N�V�[�g
Public sheet     As Worksheet
' DB�R�l�N�V����
Public conn      As Object
' �L�����Z���t���O
Public canceled  As Boolean

' =========================================================
' ��SQL��`�Ǎ��J�n�O�C�x���g
'
' �T�v�@�@�@�FSQL��`�̓Ǎ����J�n����O�ɌĂяo�����C�x���g
' �����@�@�@�F
'
' =========================================================
Public Event readSqlDefineInfoBegin()

' =========================================================
' ��SQL��`�Ǎ��C�x���g
'
' �T�v�@�@�@�FSQL��`�̓Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�FsqlDefineSheet SQL��`�V�[�g
'
' =========================================================
Public Event readSqlDefineInfo(ByRef sqlDefineSheet As ValSnapSqlDefineWorkSheet)

' =========================================================
' ���p�����[�^���X�g�Ǎ��J�n�O�C�x���g
'
' �T�v�@�@�@�F�p�����[�^���X�g�̓Ǎ����J�n����O�ɌĂяo�����C�x���g
' �����@�@�@�Fcnt ����
'
' =========================================================
Public Event readParamListBegin(ByVal cnt As Long)

' =========================================================
' ���p�����[�^���Ǎ��C�x���g
'
' �T�v�@�@�@�F�p�����[�^���Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�Fparam �p�����[�^���
'
' =========================================================
Public Event readParam(ByRef param As ValSnapSqlDefineParam)

' =========================================================
' ���p�����[�^���X�g�Ǎ��C�x���g
'
' �T�v�@�@�@�F�p�����[�^���X�g�Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�F
'
' =========================================================
Public Event readParamList()

' =========================================================
' ��SQL���X�g�Ǎ��J�n�O�C�x���g
'
' �T�v�@�@�@�FSQL���X�g�̓Ǎ����J�n����O�ɌĂяo�����C�x���g
' �����@�@�@�Fcnt ����
'
' =========================================================
Public Event readSqlListBegin(ByVal cnt As Long)

' =========================================================
' ��SQL���Ǎ��C�x���g
'
' �T�v�@�@�@�FSQL���Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�Fsql SQL���
'
' =========================================================
Public Event readSql(ByRef sql As ValSnapSqlDefineSql)

' =========================================================
' ��SQL���X�g�Ǎ��C�x���g
'
' �T�v�@�@�@�FSQL���X�g�Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�F
'
' =========================================================
Public Event readSqlList()

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
' ��SQL��`�V�[�g����
'
' �T�v�@�@�@�F�ΏۃV�[�g�i�����o��sheet�j��
' �@�@�@�@�@�@SQL��`�V�[�g�ł��邩�𔻒肷��B
'
' �����@�@�@�F
' �߂�l�@�@�FTrue SQL��`�V�[�g�̏ꍇ
' =========================================================
Public Function isSqlDefineSheet() As Boolean

    ' �e�[�u���V�[�g����p�̃I�[�g�V�F�C�v
    Dim s As Shape

    ' �߂�l������������
    isSqlDefineSheet = False
    
    ' ------------------------------------------------
    ' �G���[���������Ă��p������
    On Error Resume Next

    Set s = sheet.Shapes(ConstantsSnapSqlDefine.SHEET_CHECK_IMAGE)
    
    ' �擾�ł����ꍇ
    If Not s Is Nothing Then
    
        isSqlDefineSheet = True
        Exit Function
    End If
        
    On Error GoTo 0
    ' ------------------------------------------------

End Function

' =========================================================
' ���e�[�u���V�[�g����
'
' �T�v�@�@�@�F�ΏۃV�[�g�i�����o��sheet�jSQL��`�V�[�g�ł��邩�𔻒肷��B
' �@�@�@�@�@�@SQL��`�V�[�g�ł͂Ȃ��ꍇ�A�G���[�𔭍s����B
'
' �����@�@�@�F
' �߂�l�@�@�F
' =========================================================
Public Sub validSqlDefineSheet()

    ' �e�[�u���V�[�g�ł͂Ȃ��ꍇ
    If isSqlDefineSheet = False Then
    
        err.Raise ConstantsError.ERR_NUMBER_IS_NOT_SQL_DEFINE_SHEET _
            , _
            , ConstantsError.ERR_DESC_IS_NOT_SQL_DEFINE_SHEET
    End If

End Sub

' =========================================================
' ��SQL��`�ǂݍ���
'
' �T�v�@�@�@�F�V�[�g���SQL��`��ǂݍ����
' �@�@�@�@�@�@SQL��`�V�[�g�I�u�W�F�N�g�𐶐�����B
'
' �����@�@�@�FSqlDefineInfoOnly �e�[�u�����̂ݎ擾�i�J�������͎擾���Ȃ��j
' �߂�l�@�@�F�e�[�u���V�[�g
'
' =========================================================
Public Function readSqlDefineInfo() As ValSnapSqlDefineWorkSheet

    ' SQL��`
    Dim sqlDefine As ValSnapSqlDefine
    ' SQL��`���𐶐�����
    Set sqlDefine = New ValSnapSqlDefine
    ' SQL��`�V�[�g
    Dim sqlDefineSheet As ValSnapSqlDefineWorkSheet
    ' SQL��`�V�[�g���𐶐�����
    Set sqlDefineSheet = New ValSnapSqlDefineWorkSheet
    Set sqlDefineSheet.snapSqlDefine = sqlDefine
    
    ' ���O
    Dim name As String
    
    ' �e�[�u�������J�n�C�x���g
    RaiseEvent readSqlDefineInfoBegin
    
    With sheet

        ' �e�[�u�������擾����
        name = .Cells(ConstantsSnapSqlDefine.NAME_ROW _
                         , ConstantsSnapSqlDefine.NAME_COL)
        sqlDefine.name = name
        
        ' �e�[�u�����擾���V�[�g��ݒ肷��
        Set sqlDefineSheet.sheet = sheet
        sqlDefineSheet.sheetName = sheet.name
        
        readSqlDefineParamsInfo sqlDefineSheet
        readSqlDefineSqlsInfo sqlDefineSheet
        
    End With ' sheet

    ' SQL��`���������C�x���g
    RaiseEvent readSqlDefineInfo(sqlDefineSheet)

    ' �߂�l��ݒ肷��
    Set readSqlDefineInfo = sqlDefineSheet

End Function

' =========================================================
' ��SQL��`�̃p�����[�^���X�g�ǂݍ���
'
' �T�v�@�@�@�F�V�[�g���SQL��`�̃p�����[�^���X�g��ǂݍ��ށB
'
' �����@�@�@�FsqlDefineSheet SQL��`�V�[�g
' �߂�l�@�@�F
'
' =========================================================
Private Function readSqlDefineParamsInfo(ByRef sqlDefineSheet As ValSnapSqlDefineWorkSheet)

    Dim i As Long
    
    ' ���s�R�[�h�i�Z���̉��s�R�[�h��LF�j
    Dim newline As String: newline = vbLf

    ' SQL��`
    Dim sqlDefine      As ValSnapSqlDefine
    Set sqlDefine = sqlDefineSheet.snapSqlDefine
    
    ' �p�����[�^
    Dim param     As ValSnapSqlDefineParam
    
    ' ���[�v�p���t���O
    Dim continue   As Variant
    
    ' ���ݍs
    Dim currentRow As Long
    
    ' �p�����[�^���
    Dim paramName   As String    ' ����
    Dim paramValue  As String    ' �l
    Dim paramMemo   As String    ' ���l
    
    ' �p�����[�^�T�C�Y
    Dim paramSize As Long
    ' �p�����[�^�T�C�Y���擾����
    paramSize = getParamRowSize()
    
    ' �p�����[�^�T�C�Y��0�̏ꍇ
    If paramSize <= 0 Then
    
        Exit Function
    End If

    ' �p�����[�^�����J�n�C�x���g
    RaiseEvent readParamListBegin(paramSize)
    
    ' ---------------------------------------------------------------------
    With sheet
    
        ' �s�̍ŏ��̈ʒu���擾
        currentRow = ConstantsSnapSqlDefine.PARAM_START_ROW
    
        ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
        continue = .Cells(currentRow, ConstantsSnapSqlDefine.PARAM_ROW_COL)
        
        ' ���[�v�p�������icontinue����ł͂Ȃ��ꍇ�p���j
        Do While continue <> Empty
        
            ' ���R�[�h�ɃJ��������ǉ�����
            Set param = New ValSnapSqlDefineParam
            
            ' �p�����[�^�����擾
            paramName = .Cells(currentRow, ConstantsSnapSqlDefine.PARAM_NAME_COL)
            param.name = paramName
            
            ' �p�����[�^�l���擾
            paramValue = .Cells(currentRow, ConstantsSnapSqlDefine.PARAM_VALUE_COL)
            param.value = paramValue
            
            ' ���l���擾
            paramMemo = .Cells(currentRow, ConstantsSnapSqlDefine.PARAM_MEMO_COL)
            param.memo = paramMemo

            sqlDefine.addParam param
            
            ' �p�����[�^���������C�x���g
            RaiseEvent readParam(param)
        
            ' �s����i�߂�
            currentRow = currentRow + 1
            
            ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
            continue = .Cells(currentRow, ConstantsSnapSqlDefine.PARAM_ROW_COL)
        
        Loop
        ' ---------------------------------------------------------------------
    
        ' �p�����[�^���X�g���������C�x���g
        RaiseEvent readParamList
    
    End With ' sheet

End Function

' =========================================================
' ��SQL��`��SQL���X�g�ǂݍ���
'
' �T�v�@�@�@�F�V�[�g���SQL��`��SQL���X�g��ǂݍ��ށB
'
' �����@�@�@�FsqlDefineSheet SQL��`�V�[�g
' �߂�l�@�@�F
'
' =========================================================
Private Function readSqlDefineSqlsInfo(ByRef sqlDefineSheet As ValSnapSqlDefineWorkSheet)

    Dim i As Long
    
    ' ���s�R�[�h�i�Z���̉��s�R�[�h��LF�j
    Dim newline As String: newline = vbLf

    ' SQL��`
    Dim sqlDefine      As ValSnapSqlDefine
    Set sqlDefine = sqlDefineSheet.snapSqlDefine
    
    ' SQL
    Dim sql   As ValSnapSqlDefineSql
    Dim sqlPkObj As ValSnapSqlDefineSqlPk
    
    ' ���[�v�p���t���O
    Dim continue   As Variant
    
    ' ���ݍs
    Dim currentRow As Long
    
    ' SQL���
    Dim sqlText    As String    ' SQL
    Dim sqlPk      As String    ' Primary Key
    Dim sqlPkArr() As String
    Dim sqlMemo    As String    ' ���l
    
    ' ---------------------------------------------------------------------
    ' �p�����[�^���X�g���ŏ��ɓǂݍ��ݏI�_��T������
    With sheet
    
        ' �s�̍ŏ��̈ʒu���擾
        currentRow = ConstantsSnapSqlDefine.PARAM_START_ROW
    
        ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
        continue = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_ROW_COL)
        
        ' ���[�v�p�������icontinue����ł͂Ȃ��ꍇ�p���j
        Do While continue <> Empty
        
            ' �s����i�߂�
            currentRow = currentRow + 1
            
            ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
            continue = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_ROW_COL)
        
        Loop
        
        ' �I�_����N�s���Z����ƁASQL��`���X�g�̍ŏ��̈ʒu�ɂȂ�
        currentRow = currentRow + 3
        
    End With
    ' ---------------------------------------------------------------------
    
    ' SQL�T�C�Y
    Dim sqlSize As Long
    ' SQL�T�C�Y���擾����
    sqlSize = getSqlRowSize(currentRow)
    
    ' SQL�T�C�Y��0�̏ꍇ
    If sqlSize <= 0 Then
    
        Exit Function
    End If

    ' SQL�����J�n�C�x���g
    RaiseEvent readSqlListBegin(sqlSize)
    
    ' ---------------------------------------------------------------------
    With sheet
    
        ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
        continue = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_ROW_COL)
        
        ' ���[�v�p�������icontinue����ł͂Ȃ��ꍇ�p���j
        Do While continue <> Empty
        
            ' ���R�[�h�ɃJ��������ǉ�����
            Set sql = New ValSnapSqlDefineSql
            
            ' �p�����[�^�����擾
            sqlText = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_SQL_COL)
            sql.sql = sqlText
            
            ' �p�����[�^�l���擾
            sqlPk = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_PRIMARY_KEY_COL)
            sqlPkArr = Split(sqlPk, newline)
            
            For i = LBound(sqlPkArr) To UBound(sqlPkArr)
                Set sqlPkObj = New ValSnapSqlDefineSqlPk
                sqlPkObj.pk = sqlPkArr(i)
                sql.addPk sqlPkObj
            Next i
            
            ' ���l���擾
            sqlMemo = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_MEMO_COL)
            sql.memo = sqlMemo
            
            If sqlText <> "" And sqlPk <> "" Then
            
                sqlDefine.addSql sql
                
                ' SQL���������C�x���g
                RaiseEvent readSql(sql)
                
            End If
        
            ' �s����i�߂�
            currentRow = currentRow + 1
            
            ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
            continue = .Cells(currentRow, ConstantsSnapSqlDefine.SQL_DEFINE_ROW_COL)
        
        Loop
        ' ---------------------------------------------------------------------
    
        ' SQL���X�g���������C�x���g
        RaiseEvent readSqlList
    
    End With ' sheet

End Function

' =========================================================
' ���V�[�g�̃p�����[�^���X�g�T�C�Y�擾
'
' �T�v�@�@�@�F�p�����[�^���X�g�T�C�Y�̎擾���s��
'
' �����@�@�@�F
' �߂�l�@�@�F�p�����[�^���X�g�T�C�Y
'
' =========================================================
Public Function getParamRowSize() As Long

    ' �߂�l�ϐ�
    Dim ret As Long
    ret = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsSnapSqlDefine.PARAM_ROW_COL) - ConstantsSnapSqlDefine.PARAM_START_ROW + 1
    
    ' �߂�l��ݒ肷��
    getParamRowSize = ret
    
End Function

' =========================================================
' ���V�[�g��SQL���X�g�T�C�Y�擾
'
' �T�v�@�@�@�FSQL���X�g�T�C�Y�̎擾���s��
'
' �����@�@�@�F
' �߂�l�@�@�FSQL���X�g�T�C�Y
'
' =========================================================
Public Function getSqlRowSize(ByVal startRow As Long) As Long

    ' �߂�l�ϐ�
    Dim ret As Long
    ret = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsSnapSqlDefine.SQL_DEFINE_ROW_COL) - startRow + 1
    
    ' �߂�l��ݒ肷��
    getSqlRowSize = ret
    
End Function
