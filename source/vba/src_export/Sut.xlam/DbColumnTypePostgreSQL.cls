VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbColumnTypePostgreSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDbColumnType

' *********************************************************
' DB�̃J������������\������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/08/14�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

Private m_dbConn As Object

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
Private Property Let IDbColumnType_dbConn(ByVal RHS As Variant)

    Set m_dbConn = RHS

End Property

' =========================================================
' ��DB�̃f�t�H���g�̃J�������������擾����
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' �߂�l�@�@�FDB�J�����������
'
' ���L�����@�F
'
'
' =========================================================
Private Function IDbColumnType_getDefaultColumnFormat() As ValCollection

    ' �f�[�^�^�ǂݍ��݃I�u�W�F�N�g�𐶐�����
    Dim dataTypeReader As New ExeDataTypeReader
    Set dataTypeReader.sheet = SutWorkbook.Worksheets("data_type")
    dataTypeReader.dbName = DBUtil.getDbmsTypeName(DbmsType.PostgreSQL)
    
    ' �߂�l��ݒ肷��
    Set IDbColumnType_getDefaultColumnFormat = dataTypeReader.execute

End Function
