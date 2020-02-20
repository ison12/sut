VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbMetaInfoGetterOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB���^���擾�N���X
' Oracle�p
'
' IDbMetaInfoGetter�����������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/15�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Implements IDbMetaInfoGetter

' DB�R�l�N�V����
Private m_dbConn As Object
' DBMS���
Private m_dbms   As DbmsType
' �N�G�������I�u�W�F�N�g
Private m_getterHelper As DbMetaInfoGetterHelper

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
' ��DB�R�l�N�V�����ݒ�
'
' �T�v�@�@�@�FDB�R�l�N�V������ݒ肷��B
' �����@�@�@�FvNewValue DB�R�l�N�V����
'
' ���L�����@�F
'
'
' =========================================================
Private Property Let IDbMetaInfoGetter_dbConn(ByVal RHS As Variant)

    ' DB�R�l�N�V������ݒ�
    Set m_dbConn = RHS
    ' DB��ނ�ݒ�
    m_dbms = ADOUtil.getDBMSType(m_dbConn)
    
    Set m_getterHelper = New DbMetaInfoGetterHelper
    m_getterHelper.dbConn = RHS
    
End Property

' =========================================================
' ���X�L�[�}��񃊃X�g�擾
'
' �T�v�@�@�@�F�ڑ����Ă���DB�̑S�X�L�[�}�����擾����
' �����@�@�@�F
'
' �߂�l�@�@�F�X�L�[�}���X�g�B
'
' ���L�����@�F
'
'
' =========================================================
Private Function IDbMetaInfoGetter_getSchemaList() As ValCollection

    On Error GoTo err
    
    Set IDbMetaInfoGetter_getSchemaList = m_getterHelper.getSchemaList()
    
    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
        
End Function

' =========================================================
' ���e�[�u�����擾�i�{�J�������j
'
' �T�v�@�@�@�F�C�ӂ̃e�[�u������S�Ď擾����B
' �����@�@�@�FschemaList �X�L�[�}���X�g
'
' �߂�l�@�@�F�e�[�u����`���B
'
' ���L�����@�F
'
'
' =========================================================
Private Function IDbMetaInfoGetter_getTableList(ByRef schemaList As ValCollection) As ValCollection

    On Error GoTo err
    
    Set IDbMetaInfoGetter_getTableList = m_getterHelper.getTableList(schemaList)

    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
        
End Function

' =========================================================
' ���e�[�u�����X�g�擾�i�{�J�������j
'
' �T�v�@�@�@�F�C�ӂ̃X�L�[�}�E�e�[�u�������L�[�ɂ��ăe�[�u�������擾����B
' �����@�@�@�FschemaName �X�L�[�}��
' �@�@�@�@�@�@ tableName �e�[�u����
'
' �߂�l�@�@�F�e�[�u����`���X�g�B
'
' ���L�����@�F
'
'
' =========================================================
Private Function IDbMetaInfoGetter_getTable(schemaName As String, tableName As String) As ValDbDefineTable

    On Error GoTo err
    
    Set IDbMetaInfoGetter_getTable = m_getterHelper.getTable(schemaName, tableName)
    
    Exit Function

err:
    
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext

End Function

