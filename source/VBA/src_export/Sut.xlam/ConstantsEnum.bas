Attribute VB_Name = "ConstantsEnum"
Option Explicit

' *********************************************************
' �񋓌^�萔���W���[��
'
' �쐬�ҁ@�FIson
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
' ��DB�N�G���o�b�`���
'
' �T�v�@�@�@�FDB�N�G���o�b�`���
'
' =========================================================
Public Enum DB_QUERY_BATCH_TYPE

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

' =========================================================
' ���A�C�R�����\�[�X
'
' �T�v�@�@�@�F�A�C�R�����\�[�X
'
' =========================================================
Public Enum RESOURCE_ICON

    Add = 1
    addFile = 2
    addFolder = 3
    addFolder2 = 4
    alert = 5
    alertMessage = 6
    book = 7
    buttonHelp = 8
    database = 9
    databaseSetting = 10
    databaseSearch = 11
    delete = 12
    deleteDatabase = 13
    devil = 14
    Edit = 15
    remove = 16
    Run = 17
    SaveAs = 18
    Search = 19
    searchWindow = 20
    settings = 21
    smile = 22
    windowImport = 23
    flagGreen = 24
    flagBlue = 25
    flagRed = 26
    areaAdd = 27
    areaEdit = 28
    areaRemove = 29
    areaSearch = 30
    bug = 31
    Paste = 32
    Forward = 33
    
End Enum

' =========================================================
' ��DB�ڑ������
' =========================================================
Public Enum DB_CONNECT_INFO_TYPE

    favorite = 1
    history = 2

End Enum

' �ꊇ�N�G�����s��ޖ���
Private dbQueryTypeNames As ValCollection

' =========================================================
' ��DB�N�G���o�b�`��ޖ��̂��擾����B
'
' �T�v�@�@�@�F
' �����@�@�@�Fd �ꊇ�N�G�����s���
' �߂�l�@�@�F�ꊇ�N�G�����s��ޖ���
' ���L�����@�F
'
' =========================================================
Public Function getDbQueryBatchTypeName(ByVal d As DB_QUERY_BATCH_TYPE) As String

    If dbQueryTypeNames Is Nothing Then
        ' ���񎞂̂ݎ��s
    
        Set dbQueryTypeNames = New ValCollection
        
        ' ��ޖ��̂̐ݒ�
        dbQueryTypeNames.setItem "", DB_QUERY_BATCH_TYPE.none
        dbQueryTypeNames.setItem "INSERT + UPDATE", DB_QUERY_BATCH_TYPE.insertUpdate
        dbQueryTypeNames.setItem "INSERT", DB_QUERY_BATCH_TYPE.insert
        dbQueryTypeNames.setItem "UPDATE", DB_QUERY_BATCH_TYPE.update
        dbQueryTypeNames.setItem "DELETE", DB_QUERY_BATCH_TYPE.deleteOnSheet
        dbQueryTypeNames.setItem "DELETE �e�[�u����̑S���R�[�h", DB_QUERY_BATCH_TYPE.deleteAll
        dbQueryTypeNames.setItem "SELECT", DB_QUERY_BATCH_TYPE.selectAll
        dbQueryTypeNames.setItem "SELECT �����w��", DB_QUERY_BATCH_TYPE.selectCondition
        dbQueryTypeNames.setItem "SELECT �Ď��s", DB_QUERY_BATCH_TYPE.selectReExec
        
    End If
    
    ' ��ޖ��̂̓���
    getDbQueryBatchTypeName = dbQueryTypeNames.getItem(d, vbVariant)

End Function
