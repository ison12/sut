VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbMetaInfoQueryCreatorHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB���^���N�G�������w���p�[�N���X
' DbMetaInfoQueryCreator�̊e�N���X�ŋ��ʎg�p���郁�\�b�h�Ȃǂ��`
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/15�@�V�K�쐬
'
' ���L�����F
' *********************************************************

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

' �X�L�[�}���X�g�擾�p�N�G���̃f�B���N�g���p�X���擾
Private Function getQueryDir(ByVal dbms As DbmsType) As String

    Dim dbmsName As String
    dbmsName = DBUtil.getDbmsTypeName(dbms)

    Dim queryDir As String
    queryDir = VBUtil.concatFilePath(SutWorkbook.path, "resource\dbMetaInfoQuery\" & dbmsName)

    getQueryDir = queryDir

End Function

' �����Ȃ��N�G���̐���
Public Function getSqlForNoCondition(ByVal dbms As DbmsType _
                                   , ByVal sqlId As String) As String

    ' DBMS�����擾
    Dim dbmsName As String
    dbmsName = DBUtil.getDbmsTypeName(dbms)

    ' �N�G���f�B���N�g�����擾
    Dim queryDir As String
    queryDir = getQueryDir(dbms)

    ' �N�G���t�@�C���p�X���擾
    Dim queryFilePath As String
    queryFilePath = VBUtil.concatFilePath(queryDir, LCase$(dbmsName) & "_" & sqlId & ".sql")

    ' �t�@�C���̓��e��ǂݍ���
    Dim query     As String
    
    Dim fr As New FileReader
    fr.init queryFilePath, "UTF-8", vbCr & vbLf
    query = fr.readText
    
    getSqlForNoCondition = query

End Function

' �����t���N�G���̐���
Public Function getSqlForCondition(ByVal dbms As DbmsType _
                                    , ByVal sqlId As String _
                                    , ByVal sqlCondId As String _
                                    , Optional ByVal schemaList As ValCollection = Nothing _
                                    , Optional ByVal schemaName As String = Empty _
                                    , Optional ByVal tableName As String = Empty) As String

    Dim schemaVar As Variant
    Dim schemaStr As StringBuilder

    ' DBMS�����擾
    Dim dbmsName As String
    dbmsName = DBUtil.getDbmsTypeName(dbms)

    ' �N�G���f�B���N�g�����擾
    Dim queryDir As String
    queryDir = getQueryDir(dbms)

    ' �{�̃N�G���t�@�C���p�X���擾
    Dim queryFilePath As String
    queryFilePath = VBUtil.concatFilePath(queryDir, LCase$(dbmsName) & "_" & sqlId & ".sql")

    ' �����N�G���t�@�C���p�X���擾
    Dim queryCondFilePath As String
    queryCondFilePath = VBUtil.concatFilePath(queryDir, LCase$(dbmsName) & "_" & sqlCondId & ".sql")

    ' �t�@�C���̓��e��ǂݍ���
    Dim query     As String
    Dim queryCond As String
    
    Dim fr As New FileReader
    fr.init queryFilePath, "UTF-8", vbCr & vbLf
    query = fr.readText

    Set fr = New FileReader
    fr.init queryCondFilePath, "UTF-8", vbCr & vbLf
    queryCond = fr.readText
    
    ' �t�@�C���̓��e�ɏ����𖄂ߍ���
    If schemaName <> Empty And _
        tableName <> Empty Then
        
        queryCond = replace(queryCond, "${table_schema}", DBUtil.encloseSingleQuart(DBUtil.escapeValueForSinglequart(dbms, schemaName)))
        queryCond = replace(queryCond, "${table_name}", DBUtil.encloseSingleQuart(DBUtil.escapeValueForSinglequart(dbms, tableName)))
        
        query = replace(query, "${condition}", queryCond)
        
    ElseIf Not schemaList Is Nothing Then
    
        Set schemaStr = New StringBuilder
        For Each schemaVar In schemaList.col
            If schemaStr.length <> 0 Then
                schemaStr.append ", "
            End If
            schemaStr.append DBUtil.encloseSingleQuart(DBUtil.escapeValueForSinglequart(dbms, schemaVar.schemaName))
        Next
    
        queryCond = replace(queryCond, "${schemaList}", schemaStr.str)
        query = replace(query, "${condition}", queryCond)

    Else
        
        queryCond = replace(queryCond, "${schemaList}", DBUtil.encloseSingleQuart(DBUtil.escapeValueForSinglequart(dbms, schemaName)))
        
        query = replace(query, "${condition}", queryCond)
    End If
    
    getSqlForCondition = query

End Function
