VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValNamingTableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���̂�\������N���X (�e�[�u�����)
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/06/29�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �e�[�u���_����
Private m_tableNameLog  As String
' �e�[�u��������
Private m_tableName     As String
' �J�������X�g
Private m_columnList As ValCollection

' =========================================================
' ���v���p�e�B�錾
Public Property Get columnList() As Variant

    Set columnList = m_columnList
End Property

Public Property Let columnList(ByVal vNewValue As Variant)

    Set m_columnList = vNewValue
End Property

Public Property Get tableNameLog() As Variant

    tableNameLog = m_tableNameLog
End Property

Public Property Let tableNameLog(ByVal vNewValue As Variant)

    m_tableNameLog = vNewValue
End Property

Public Property Get tableName() As Variant

    tableName = m_tableName
End Property

Public Property Let tableName(ByVal vNewValue As Variant)

    m_tableName = vNewValue
End Property
' ���v���p�e�B�錾
' =========================================================


' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set m_columnList = New ValCollection

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set m_columnList = Nothing
End Sub

' =========================================================
' ���N���X���������\�b�h (�����t��)
' =========================================================
Public Sub initial(ByVal tableName As String, ByVal tableNameLog As String)

    Me.tableName = tableName
    Me.tableNameLog = tableNameLog
End Sub

' =========================================================
' ���J��������ǉ�
' =========================================================
Public Sub addColumnInfo(ByVal columnName As String, ByVal columnNameLog As String)

    ' �J�������
    Dim columnInfo As New ValNamingColumnInfo
    
    ' �J��������������
    columnInfo.initial columnName, columnNameLog, Me
    
    ' �J�������X�g�ɃJ��������ǉ�
    m_columnList.setItem columnInfo, columnName
    
End Sub

' =========================================================
' ���J���������擾
' =========================================================
Public Function getColumnInfo(ByVal columnName As String) As ValNamingColumnInfo

    ' �߂�l��������
    Set getColumnInfo = Nothing
    
    ' �J�������X�g��������擾
    If m_columnList.exist(columnName) = True Then
    
        ' �L�[�����݂��Ă���ꍇ�ɏ����擾����
        Set getColumnInfo = m_columnList.getItem(columnName)
    
    End If
    
End Function
