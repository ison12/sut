VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbMetaInfoQueryCreatorOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB���^���N�G�������N���X
' Oracle�p
'
' IDbMetaInfoQueryCreator�����������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/15�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Implements IDbMetaInfoQueryCreator

' DBMS
Private dbms As DbmsType
' �w���p�[�I�u�W�F�N�g
Private helper As DbMetaInfoQueryCreatorHelper

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    dbms = DbmsType.Oracle
    Set helper = New DbMetaInfoQueryCreatorHelper
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    Set helper = Nothing
End Sub

' �X�L�[�}���X�g�擾�p�N�G���̐���
Private Function IDbMetaInfoQueryCreator_createQuerySelectSchemaList() As String

    ' SQL���擾����
    IDbMetaInfoQueryCreator_createQuerySelectSchemaList = helper.getSqlForNoCondition(dbms, "schema_list")

End Function

' �e�[�u�����X�g�擾�p�N�G���̐���
Private Function IDbMetaInfoQueryCreator_createQuerySelectTableList(ByVal schemaList As ValCollection) As String

    ' SQL���擾����
   IDbMetaInfoQueryCreator_createQuerySelectTableList = helper.getSqlForCondition(dbms, "table", "table_cond_sl", schemaList)
End Function

' �e�[�u�����񃊃X�g�擾�p�N�G���̐���
Private Function IDbMetaInfoQueryCreator_createQuerySelectTableConstList(ByVal schemaList As ValCollection) As String

    ' SQL���擾����
   IDbMetaInfoQueryCreator_createQuerySelectTableConstList = helper.getSqlForCondition(dbms, "table_const", "table_const_cond_sl", schemaList)
End Function

' �e�[�u�����擾�p�N�G���̐���
Private Function IDbMetaInfoQueryCreator_createQuerySelectTable(ByVal schemaName As String, ByVal tableName As String) As String

    ' SQL���擾����
   IDbMetaInfoQueryCreator_createQuerySelectTable = helper.getSqlForCondition(dbms, "table", "table_cond_t", Nothing, schemaName, tableName)
End Function

' �J�������擾�p�N�G���̐���
Private Function IDbMetaInfoQueryCreator_createQuerySelectColumn(ByVal schemaName As String, ByVal tableName As String) As String

    ' SQL���擾����
   IDbMetaInfoQueryCreator_createQuerySelectColumn = helper.getSqlForCondition(dbms, "column", "table_cond_t", Nothing, schemaName, tableName)
End Function

' �e�[�u��������擾�p�N�G���̐���
Private Function IDbMetaInfoQueryCreator_createQuerySelectTableConst(ByVal schemaName As String, ByVal tableName As String) As String

    ' SQL���擾����
   IDbMetaInfoQueryCreator_createQuerySelectTableConst = helper.getSqlForCondition(dbms, "table_const", "table_const_cond_t", Nothing, schemaName, tableName)
End Function

