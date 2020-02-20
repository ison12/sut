VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbQueryConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�̃N�G���𐶐�����C���^�[�t�F�[�X�B
' �SDB�ŋ��ʂ���ėp�I�Ȏ����C���^�[�t�F�[�X�B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/03/28�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************
' DB�R�l�N�V����
Public conn          As Object

' �͂ݕ��� �J�n
Public encloseStart As String
' �͂ݕ��� �I��
Public encloseEnd   As String

' �N�G���p�����[�^���X�g
Private queryParamList As ValCollection
' �ϊ��Ώۃp�����[�^���X�g
Private convertedParamList As ValCollection
' ���[�N�V�[�g
Private sheet As Worksheet
' �������t���O
Private isInit As Boolean

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    encloseStart = ConstantsTable.QUERY_PARAMETER_ENCLOSE_START
    encloseEnd = ConstantsTable.QUERY_PARAMETER_ENCLOSE_END
    
    Set queryParamList = New ValCollection
    Set convertedParamList = New ValCollection
    
    isInit = False

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' ���������������\�b�h
' =========================================================
Public Sub init(ByRef conn As Object, ByRef queryParamList_ As ValQueryParameterList)

    ' DB�R�l�N�V����
    Set Me.conn = conn
    ' �N�G���p�����[�^���X�g
    Dim queryParameter As ValQueryParameter
    For Each queryParameter In queryParamList_.list.col
    
        queryParamList.setItem queryParameter, queryParameter.name
    
    Next
    
    isInit = True
    
End Sub

' =========================================================
' ���������������\�b�h�i����̂ݎ��s�j
' =========================================================
Public Sub initIfFirst(ByRef conn As Object, ByRef queryParamList_ As ValQueryParameterList)

    If isInit = False Then
        Call init(conn, queryParamList_)
    End If
    
End Sub


' =========================================================
' ��SQL�ϊ����s
'
' �T�v�@�@�@�FSQL�ϊ������s����B
' �����@�@�@�Fsql SQL
'
' �߂�l�@�@�F�ϊ����SQL
'
' ���L�����@�F
'
'
' =========================================================
Public Function convertSql(ByRef sql As String) As String

    If queryParamList.count <= 0 Then
        ' �N�G���p�����[�^���X�g���Ȃ��ꍇ�́A�������Ȃ�
        convertSql = sql
        Exit Function
    End If

    ' �ϊ��Ώۃp�����[�^
    Dim convertedParam           As Variant
    Dim convertedQueryParamName  As String
    Dim convertedQueryParamValue As String
    
    Dim queryParameter  As ValQueryParameter
    
    Dim mm         As Object
    
    Dim i As Long
    Dim match      As Object
    Dim matches    As Object
    Dim matcheSub  As Object
    
    Dim matchParamName As String

    ' ���K�\���֘A�I�u�W�F�N�g
    ' RegExp�I�u�W�F�N�g�̍쐬
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    ' ���K�\���̎w��
    With reg
        ' �p�^�[�����w��
        .Pattern = "\" & encloseStart & "(.+?)" & "\" & encloseEnd
        ' �啶���Ə���������ʂ��邩(False)�A���Ȃ���(True)
        .IgnoreCase = False
        ' ������S�̂��������邩(True)�A���Ȃ���(False)
        .Global = True
    End With
    
    ' ���K�\���ł̃}�b�`���O�����s
    Set matches = reg.execute(sql)
    
    If matches.count <= 0 Then
        convertSql = sql
        Exit Function
    End If
    
    For i = 0 To matches.count - 1
    
        Set match = matches.item(i)
        Set matcheSub = match.submatches
        matchParamName = matcheSub.item(0)
        
        If queryParamList.exist(matchParamName) = False Then
            ' �Y������N�G���p�����[�^�����݂��Ȃ�
            GoTo continue
        End If
    
        Set queryParameter = queryParamList.getItem(matchParamName)
    
        If convertedParamList.exist(matchParamName) = False Then
            ' �ϊ��ς݃N�G���p�����[�^�����݂��Ȃ�
        
            If isShouldExecuteSql(queryParameter.value) Then
                convertedQueryParamValue = executeSql(queryParameter.value)
            End If
            
            convertedParam = Array(queryParameter.name, convertedQueryParamValue)
            convertedParamList.setItem convertedParam, queryParameter.name
            
        Else
            ' ���ɕϊ��ς݃N�G���p�����[�^�����݂���
            convertedParam = convertedParamList.getItem(matchParamName, vbVariant)
            
        End If
        
        ' �ϊ��Ώۃp�����[�^�ŕϊ������s����
        ' SQL��ϊ�����
        convertedQueryParamName = convertedParam(0)
        convertedQueryParamValue = convertedParam(1)
        
        sql = replace(sql, encloseStart & convertedQueryParamName & encloseEnd, convertedQueryParamValue)
        
continue:
    
    Next
    
    convertSql = sql
    
End Function

' =========================================================
' ��SQL���s
'
' �T�v�@�@�@�FSQL�����s����B
' �����@�@�@�Fsql SQL
'
' �߂�l�@�@�F���s����
'
' ���L�����@�F
'
'
' =========================================================
Public Function executeSql(ByRef sql As String) As String

    ' ���R�[�h�Z�b�g
    Dim recordSet   As Object
    ' �N�G���[�����s����
    Set recordSet = ADOUtil.querySelect(conn, sql, 0)
    
    ' �߂�l
    Dim result         As Variant

    ' ���R�[�h�Z�b�g��EOF�ł͂Ȃ��ꍇ
    If Not recordSet.EOF Then
    
        ' ���R�[�h�Z�b�g����S���R�[�h���擾����
        result = recordSet.getRows(1)
    
        executeSql = result(0, 0)
    Else
        ' ���Ԃ�
        executeSql = Empty
    End If

End Function

' =========================================================
' ��SELECT�����s���邩�ǂ����̔��菈��
'
' �T�v�@�@�@�FSELECT�����s���邩�ǂ����̔��菈�����s���B
'             �擪������"SELECT"�Ŏn�܂�ꍇ
' �����@�@�@�Fsql SQL
'
' �߂�l�@�@�FSELECT�����s���邩�ǂ����̔���t���O
'
' ���L�����@�F
'
'
' =========================================================
Public Function isShouldExecuteSql(ByRef sql As String) As Boolean

    ' ���K�\���֘A�I�u�W�F�N�g
    ' RegExp�I�u�W�F�N�g�̍쐬
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    ' ���K�\���̎w��
    With reg
        .Pattern = "SELECT "         ' �p�^�[�����w��
        .IgnoreCase = True           ' �啶���Ə���������ʂ��邩(False)�A���Ȃ���(True)
        .Global = True               ' ������S�̂��������邩(True)�A���Ȃ���(False)
    End With
    
    Dim matches    As Object
    Dim matcheSub  As Object
    Dim matcheText As String
    
    Set matches = reg.execute(sql) ' ���K�\���ł̃}�b�`���O�����s
    
    If matches.count > 0 Then
        isShouldExecuteSql = True
    Else
        isShouldExecuteSql = False
    End If

End Function