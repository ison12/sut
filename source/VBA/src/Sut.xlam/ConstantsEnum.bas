Attribute VB_Name = "ConstantsEnum"
Option Explicit

' *********************************************************
' �񋓌^�萔���W���[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' =========================================================
' ���e�[�u��������
'
' �T�v�@�@�@�F�e�[�u��������
'
' =========================================================
Public Enum TABLE_CONSTANTS_TYPE

    tableConstPk = 0
    tableConstUk = 1
    tableConstFk = 2
    tableConstUnknown = -1

End Enum

' =========================================================
' ���s�t�H�[�}�b�g���
'
' �T�v�@�@�@�F�s�t�H�[�}�b�g���
'
' =========================================================
Public Enum REC_FORMAT

    recFormatToUnder = 0
    recFormatToRight = 1

End Enum

' =========================================================
' ���ꊇ�N�G�����s���
'
' �T�v�@�@�@�F�ꊇ�N�G�����s���
'
' =========================================================
Public Enum BATCH_QUERY_TYPE

    none = 0
    insertUpdate = 1
    insert = 2
    update = 3
    deleteOnSheet = 4
    deleteAll = 5
    selectAll = 6
    selectCondition = 7
    selectReExec = 8

End Enum


