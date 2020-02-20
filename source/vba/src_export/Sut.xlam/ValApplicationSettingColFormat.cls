VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValApplicationSettingColFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �A�v���P�[�V�����̃I�v�V�������iDB�J���������ݒ�j
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/03/14�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' =========================================================
' �������o
' =========================================================
Public dbList As ValCollection

' =========================================================
' ���萔�v���p�e�B
' �@Class���W���[���ł́APublic Const����`�ł��Ȃ��̂Ńv���p�e�B�ő�p����
' =========================================================
' �L�[�l
Public Property Get KEY_COLUMN() As String

    KEY_COLUMN = "column"
End Property

' �L�[�l
Public Property Get KEY_UPDATE() As String

    KEY_UPDATE = "update"
End Property

' �L�[�l
Public Property Get KEY_SELECT() As String

    KEY_SELECT = "select"
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

Public Sub init()

    readForData
    
End Sub

' =========================================================
' ��DB�ŗL�̃J������������ݒ肷��
'
' �T�v�@�@�@�FdbColumnFmtInfo DB�J���������ݒ���
'
' =========================================================
Public Function setDbColFormatInfo(ByRef dbColumnFmtInfo As ValDbColumnFormatInfo)

    ' �ݒ���i�����̏����N���[������j
    Dim setInfo As New ValDbColumnFormatInfo
    setInfo.dbName = dbColumnFmtInfo.dbName

    ' �J�������X�g�𐶐����Ȃ����i�J���������L�[�ɂ������̂Łj
    Dim v As ValDbColumnTypeColInfo
    For Each v In dbColumnFmtInfo.columnList.col
        setInfo.columnList.setItem v, v.columnName
    Next

    ' DB���X�g�ɍĐݒ肷��
    dbList.setItem setInfo, setInfo.dbName

End Function

' =========================================================
' ��DB�ŗL�̃J����������񃊃X�g���擾����
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function getDbColFormatInfo(ByRef dbmsT As DbmsType) As ValDbColumnFormatInfo

    ' DB�J�����������
    Dim dbColumnFmtInfo As ValDbColumnFormatInfo
    
    Set dbColumnFmtInfo = dbList.getItem(CStr(dbmsT))

    If dbColumnFmtInfo Is Nothing Then
    
        Set dbColumnFmtInfo = New ValDbColumnFormatInfo
        dbColumnFmtInfo.dbName = dbmsT
        
        #If DEBUG_MODE = 1 Then
        
            Debug.Print "�I�I�I�J����������񂪌�����܂���I�I�I"
        #End If

    End If
    
    ' �߂�l��ݒ肷��
    Set getDbColFormatInfo = dbColumnFmtInfo

End Function

Public Function getDbColFormatListByDbConn(ByRef dbConn As Object) As ValCollection

    ' DBMS��ނ��擾����
    Dim dmt As DbmsType
    dmt = ADOUtil.getDBMSType(dbConn)
    
    ' �߂�l��ݒ肷��
    Set getDbColFormatListByDbConn = getDbColFormatList(dmt)

End Function

Public Function getDbColFormatList(ByRef dbmsT As DbmsType) As ValCollection

    ' DB�J�����������
    Dim dbColumnFmtInfo As ValDbColumnFormatInfo
    ' DB�J����������񃊃X�g
    Dim columnFormatList As ValCollection
    
    Set dbColumnFmtInfo = getDbColFormatInfo(dbmsT)

    ' �߂�l��ݒ肷��
    Set getDbColFormatList = dbColumnFmtInfo.columnList

End Function

' =========================================================
' ���ݒ���̐���
' =========================================================
Private Function createApplicationProperties(ByVal keyNameDb As String) As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingColFormat_" & keyNameDb & ".ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ������ǂݍ���
' =========================================================
Public Sub readForData()

    On Error GoTo err
    
    Set dbList = New ValCollection
    
    Dim dbColumnFmtInfo As ValDbColumnFormatInfo
    
    Dim dbt As DbmsType
    
    ' �eDB�̃J�������������擾����
    ' ----------------------------------------------
    ' Oracle
    ' ----------------------------------------------
    dbt = DbmsType.Oracle
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' MySQL
    ' ----------------------------------------------
    dbt = DbmsType.MySQL
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' PostgreSQL
    ' ----------------------------------------------
    dbt = DbmsType.PostgreSQL
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' SQL Server
    ' ----------------------------------------------
    dbt = DbmsType.MicrosoftSqlServer
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' Access
    ' ----------------------------------------------
    dbt = DbmsType.MicrosoftAccess
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------

    ' ----------------------------------------------
    ' Symfoware
    ' ----------------------------------------------
    dbt = DbmsType.Symfoware
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ������ǂݍ��� DB��
' =========================================================
Private Function readForDataDbInfo(ByVal keyNameDb As String) As ValDbColumnFormatInfo

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties(keyNameDb)
                
    ' �f�[�^��ǂݍ���
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    Dim val As Variant
    
    Dim dbColumnFmtInfo As New ValDbColumnFormatInfo
    
    If values.count > 0 Then
    
        ' DB���̂�ݒ肷��
        dbColumnFmtInfo.dbName = keyNameDb
    
        ' -------------------------------------------------------
        ' �S�Ă̏����擾����
        ' -------------------------------------------------------
        Dim i As Long: i = 1
        Do While True
        
            val = values.getItem(i & "_" & KEY_COLUMN, vbVariant)
            If Not IsArray(val) Then
                Exit Do
            End If
                        
            Dim column As ValDbColumnTypeColInfo
            Set column = New ValDbColumnTypeColInfo
                        
            val = values.getItem(i & "_" & KEY_COLUMN, vbVariant): If IsArray(val) Then column.columnName = UCase$(val(2)) ' �J���������������ɑ啶���݂̂ŏ�����v�����邽�߂ɁA�����ł͑啶���ɕϊ����Đݒ肷��
            val = values.getItem(i & "_" & KEY_UPDATE, vbVariant): If IsArray(val) Then column.formatUpdate = val(2)
            val = values.getItem(i & "_" & KEY_SELECT, vbVariant): If IsArray(val) Then column.formatSelect = val(2)
            
            dbColumnFmtInfo.columnList.setItem column, UCase$(column.columnName)
        
            i = i + 1
        Loop
        
    Else
    
        ' DB�I�u�W�F�N�g�����N���X
        Dim dbObjFactory As New DbObjectFactory
        ' DB�J�������擾�I�u�W�F�N�g
        Dim dbColumnType As IDbColumnType
        Set dbColumnType = dbObjFactory.createColumnType(keyNameDb)
        
        ' DB���̂�ݒ肷��
        dbColumnFmtInfo.dbName = keyNameDb
        
        Set dbColumnFmtInfo.columnList = dbColumnType.getDefaultColumnFormat
        
    End If
    
    ' �߂�l��ݒ肷��
    Set readForDataDbInfo = dbColumnFmtInfo
        
    Exit Function
    
err:

    Main.ShowErrorMessage

End Function

' =========================================================
' ��������������
' =========================================================
Public Sub writeForData()

    On Error GoTo err
    
    ' DB���
    Dim dbInfo As Variant
    
    For Each dbInfo In dbList.col
    
        writeForDataDbInfo dbInfo
    
    Next
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' �������������� DB��
' =========================================================
Public Sub writeForDataDbInfo(ByVal dbInfo As ValDbColumnFormatInfo)

    On Error GoTo err
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties(dbInfo.dbName)
                
    ' �������݃f�[�^
    Dim val As Variant
    Dim values As New ValCollection
    
    Dim i As Long: i = 1
    For Each val In dbInfo.columnList.col
    
        values.setItem Array(i & "_" & KEY_COLUMN, val.columnName)
        values.setItem Array(i & "_" & KEY_UPDATE, val.formatUpdate)
        values.setItem Array(i & "_" & KEY_SELECT, val.formatSelect)
    
        i = i + 1
    Next
    
    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData

    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub