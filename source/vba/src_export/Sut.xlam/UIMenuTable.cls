VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���j���[���쎞�̏��� - �e�[�u��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/01/24�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ________________________________________________________
' �����o�ϐ�
' ________________________________________________________
' DB�G�N�X�v���[��
Private WithEvents frmDBExplorerVar        As frmDBExplorer
Attribute frmDBExplorerVar.VB_VarHelpID = -1
' �e�[�u���V�[�g�쐬�t�H�[��
Private WithEvents frmTableSheetCreatorVar As frmTableSheetCreator
Attribute frmTableSheetCreatorVar.VB_VarHelpID = -1
' �e�[�u���V�[�g���X�g�t�H�[��
Private WithEvents frmTableSheetListVar    As frmTableSheetList
Attribute frmTableSheetListVar.VB_VarHelpID = -1
' �e�[�u���V�[�g���X�g�t�H�[��
Private WithEvents frmTableSheetUpdateVar  As frmTableSheetUpdate
Attribute frmTableSheetUpdateVar.VB_VarHelpID = -1
' �e�[�u���V�[�g����
Private WithEvents tableSheetCreator       As ExeTableSheetCreator
Attribute tableSheetCreator.VB_VarHelpID = -1
' �v���O���X�t�H�[��
Private WithEvents frmProgressVar          As frmProgress
Attribute frmProgressVar.VB_VarHelpID = -1
' ���R�[�h��񐶐�
Private WithEvents tableSheetReader        As ExeTableSheetReader
Attribute tableSheetReader.VB_VarHelpID = -1

' DB�R�l�N�V����
Private dbConn As Object
' �A�v���P�[�V�����ݒ���
Private applicationSetting As ValApplicationSetting

' ________________________________________________________
' �C�x���g
' ________________________________________________________
' =========================================================
' ��DB�G�N�X�v���[�� Excel�V�[�g�ɏo��
' =========================================================
Private Sub frmDBExplorerVar_export(ByVal tableList As ValCollection, ByVal recFormat As REC_FORMAT)

    createEmptyBook
    
    Dim tableSheetList As ValCollection
    Set tableSheetList = createDataOfTableSheetList(tableList, recFormat)
    
    frmTableSheetCreatorVar_complete tableSheetList
    
End Sub

' =========================================================
' ��DB�G�N�X�v���[�� ����
' =========================================================
Private Sub frmDBExplorerVar_closed()

End Sub

' =========================================================
' ���e�[�u���V�[�g�쐬 - �e�[�u���V�[�g�쐬�J�n�i�����j
' =========================================================
Private Sub tableSheetCreator_createTableSheetsBegin(tableSheetList As ValCollection)

    frmProgressVar.initPri tableSheetList.count, "�e�[�u���V�[�g�쐬��"
    
    DoEvents
End Sub

' =========================================================
' ���e�[�u���V�[�g�쐬 - �e�[�u���V�[�g�쐬�����i�����j
' =========================================================
Private Sub tableSheetCreator_createTableSheets(tableSheetList As ValCollection)

    frmProgressVar.priMessage = "�e�[�u���V�[�g�쐬����"
    frmProgressVar.priCount = tableSheetList.count
    
    DoEvents
End Sub

' =========================================================
' ���e�[�u���V�[�g�쐬 - �e�[�u���V�[�g�쐬�J�n
' =========================================================
Private Sub tableSheetCreator_createTableSheetBegin(tableSheet As ValTableWorksheet)

    frmProgressVar.priMessage = "�e�[�u���V�[�g�쐬���E�E�E" & tableSheet.table.schemaTableName
    
    DoEvents
End Sub

' =========================================================
' ���e�[�u���V�[�g�쐬 - �e�[�u���V�[�g�쐬����
' =========================================================
Private Sub tableSheetCreator_createTableSheet(tableSheet As ValTableWorksheet)

    frmProgressVar.inclimentPri
    
    DoEvents
End Sub

' =========================================================
' ���������L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
' =========================================================
Private Sub frmProgressVar_cancel()

    If Not tableSheetCreator Is Nothing Then
        tableSheetCreator.canceled = True
    End If
    
    If Not tableSheetReader Is Nothing Then
        tableSheetReader.canceled = True
    End If

End Sub

' =========================================================
' ���e�[�u���V�[�g�쐬�t�H�[�� - �L�����Z�����ɌĂяo�����C�x���g
' =========================================================
Private Sub frmTableSheetCreatorVar_cancel()

End Sub

' =========================================================
' ���e�[�u���V�[�g�쐬�t�H�[�� - �������ɌĂяo�����C�x���g
' =========================================================
Private Sub frmTableSheetCreatorVar_complete(createTargetTable As ValCollection)

    On Error GoTo err

    ' Excel.Application�v���p�e�B�𑀍삷��I�u�W�F�N�g
    ' �֐��𔲂���Ǝ�����Application�v���p�e�B�����������
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init

    ' �u�b�N�I�u�W�F�N�g
    Dim book As Workbook
    
    ' �u�b�N�I�u�W�F�N�g���擾����
    Set book = ActiveWorkbook
    
    ' �u�b�N�I�u�W�F�N�g�̃`�F�b�N
    If book Is Nothing Then
    
        Set book = Application.Workbooks.Add
        book.activate
    End If
    
    frmProgressVar.ShowExt vbModeless, False
    frmProgressVar.title = "�e�[�u���V�[�g�쐬"
    
    DoEvents

    Set tableSheetCreator = New ExeTableSheetCreator
    tableSheetCreator.applicationSetting = applicationSetting
    tableSheetCreator.book = ActiveWorkbook
    tableSheetCreator.createSheets createTargetTable

    ' ExcelLongTimeProcessing�ɂāAApplication.interactive�𖳌��ɂ�����Ԃ�
    ' �_�C�A���O�����J����������肷��ƁA�E�B���h�E��Z�I�[�_�[������Ă��܂��B�i�����I����ɁAExcel���A�N�e�B�u�ɂȂ�Ȃ��j
    ' ���̂��߁AExcelLongTimeProcessing�I�u�W�F�N�g�𖾎��I�ɔj������
    Set longTimeProcessing = Nothing
    
    frmProgressVar.HideExt
    
    Exit Sub
err:

    Set longTimeProcessing = Nothing
    
    frmProgressVar.HideExt
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ���e�[�u���V�[�g�X�V�t�H�[�� - OK�{�^���������ɌĂяo�����C�x���g
' =========================================================
Private Sub frmTableSheetUpdateVar_ok(ByVal recFormat As REC_FORMAT)

    On Error GoTo err

    ' Excel.Application�v���p�e�B�𑀍삷��I�u�W�F�N�g
    ' �֐��𔲂���Ǝ�����Application�v���p�e�B�����������
    Dim longTimeProcessing As New ExcelLongTimeProcessing: longTimeProcessing.init
    
    ' ----------------------------------------------
    ' �e�[�u���V�[�g�����x�e�[�u������ǂݍ���
    Dim srcTable As ValTableWorksheet
    
    Dim tableSheetReader As ExeTableSheetReader
    
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = dbConn
    Set tableSheetReader.sheet = ActiveSheet
    
    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
    
    Set srcTable = tableSheetReader.readTableInfo
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' DB����ŐV�̃e�[�u������ǂݍ���
    ' DB���^���擾�p�I�u�W�F�N�g�����N���X
    Dim dbObjFactory As New DbObjectFactory
    ' DB���^���擾�p�N���X
    Dim DbMetaInfoGetter  As IDbMetaInfoGetter
    
    Set DbMetaInfoGetter = dbObjFactory.createMetaInfoGetterObject(dbConn)
    
    ' �e�[�u���V�[�g
    Dim desTableSheet As New ValTableWorksheet
    ' �e�[�u�����
    Dim desTable      As ValDbDefineTable
    ' �e�[�u�������擾����
    Set desTable = DbMetaInfoGetter.getTable(srcTable.table.schemaName, srcTable.table.tableName)
    
    If desTable Is Nothing Then
    
        err.Raise ConstantsError.ERR_NUMBER_NOT_EXIST_TABLE_INFO _
                , _
                , ConstantsError.ERR_DESC_NOT_EXIST_TABLE_INFO

        Exit Sub
    End If
    
    ' �ŐV�̃e�[�u�����̍s�t�H�[�}�b�g���t�H�[���Ŏw�肵�����̂ɐݒ肷��
    Set desTableSheet.table = desTable
    Set desTableSheet.sheet = srcTable.sheet
    desTableSheet.recFormat = recFormat
    desTableSheet.omitsSchema = srcTable.omitsSchema
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' �e�[�u���V�[�g�𐶐�����
    
    ' �e�[�u���V�[�g�����p�̃I�u�W�F�N�g�𐶐�����
    ' �C�x���g����M���Ȃ��悤�ɂ��邽�߂ɁA�����o�ϐ��itableSheetCreator�j�ł͂Ȃ����[�J���ϐ��𗘗p����B
    ' �����ł̏����́A�v���O���X�o�[�ɐi���󋵂�\�����Ȃ��B
    Dim tsCreator As ExeTableSheetCreator
    
    Set tsCreator = New ExeTableSheetCreator
    tsCreator.applicationSetting = applicationSetting
    tsCreator.book = ActiveWorkbook
    tsCreator.update = True
    tsCreator.createSheet desTableSheet
    ' ----------------------------------------------
    
    ' ExcelLongTimeProcessing�ɂāAApplication.interactive�𖳌��ɂ�����Ԃ�
    ' �_�C�A���O�����J����������肷��ƁA�E�B���h�E��Z�I�[�_�[������Ă��܂��B�i�����I����ɁAExcel���A�N�e�B�u�ɂȂ�Ȃ��j
    ' ���̂��߁AExcelLongTimeProcessing�I�u�W�F�N�g�𖾎��I�ɔj������
    Set longTimeProcessing = Nothing
    
    Exit Sub

err:

    Set longTimeProcessing = Nothing
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���e�[�u���V�[�g�ꗗ�t�H�[�� - �V�[�g�I�����ɌĂяo�����C�x���g
' =========================================================
Private Sub frmTableSheetListVar_selected(table As ValTableWorksheet)

    On Error GoTo err

    table.sheet.activate

    Exit Sub

err:

        

End Sub

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    
    Set frmDBExplorerVar = Nothing
    Set frmTableSheetCreatorVar = Nothing
    Set frmTableSheetListVar = Nothing
    Set frmTableSheetUpdateVar = Nothing
    Set frmProgressVar = Nothing
    Set tableSheetCreator = Nothing
    Set tableSheetReader = Nothing

    Set applicationSetting = Nothing
    Set dbConn = Nothing
End Sub

' =========================================================
' ��������
'
' �T�v�@�@�@�F���������������s����B
' �����@�@�@�FappSetting        �A�v���P�[�V�����ݒ���
' �@�@�@�@�@�@conn               DB�R�l�N�V����
'
' =========================================================
Public Sub init(ByRef appSetting As ValApplicationSetting _
              , ByRef conn As Object)

    ' �t�H�[����ݒ肷��
    If VBUtil.unloadFormIfChangeActiveBook(frmProgress) Then Unload frmProgress
    Load frmProgress
    ' �t�H�[����ݒ肷��
    Set frmProgressVar = frmProgress

    ' DB�G�N�X�v���[���t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmDBExplorer) Then Unload frmDBExplorer
    Load frmDBExplorer
    ' �C�x���g��M�p�̕ϐ���DB�G�N�X�v���[���t�H�[����ݒ肷��
    Set frmDBExplorerVar = frmDBExplorer

    ' �e�[�u���V�[�g�쐬�t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmTableSheetCreator) Then Unload frmTableSheetCreator
    Load frmTableSheetCreator
    ' �C�x���g��M�p�̕ϐ��Ƀe�[�u���V�[�g�쐬�t�H�[����ݒ肷��
    Set frmTableSheetCreatorVar = frmTableSheetCreator
    
    ' �e�[�u���V�[�g�ꗗ�t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmTableSheetList) Then Unload frmTableSheetList
    Load frmTableSheetList
    ' �C�x���g��M�p�̕ϐ��Ƀe�[�u���V�[�g�ꗗ�t�H�[����ݒ肷��
    Set frmTableSheetListVar = frmTableSheetList
    
    ' �e�[�u���V�[�g�X�V�t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmTableSheetUpdate) Then Unload frmTableSheetUpdate
    Load frmTableSheetUpdate
    ' �C�x���g��M�p�̕ϐ��Ƀe�[�u���V�[�g�X�V�t�H�[����ݒ肷��
    Set frmTableSheetUpdateVar = frmTableSheetUpdate
    
    ' ���R�[�h��񐶐��I�u�W�F�N�g�𐶐�����
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.conn = conn
    Set tableSheetReader.sheet = ActiveSheet
    
    ' �A�v���P�[�V�����ݒ����ݒ肷��
    Set applicationSetting = appSetting
    ' DB�R�l�N�V������ݒ肷��
    Set dbConn = conn
    
End Sub

' =========================================================
' ��DB�R�l�N�V�����̍X�V
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub updateDbConn(ByRef dbConn As Object)

    On Error GoTo err
    
    ' DB�G�N�X�v���[���̃R�l�N�V�������Đݒ肷��
    If Not frmDBExplorer Is Nothing Then
    
        If VBUtil.unloadFormIfChangeActiveBook(frmDBExplorer) Then Unload frmDBExplorer
        Load frmDBExplorer
        Set frmDBExplorerVar = frmDBExplorer
        
        frmDBExplorerVar.DbConnection = dbConn
    End If

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ��DB�G�N�X�v���[���\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub showDbExplorer()

    On Error GoTo err
    
    ' DB�G�N�X�v���[���t�H�[����\������
    frmDBExplorerVar.ShowExt vbModeless, dbConn

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���e�[�u���V�[�g�ꗗ�\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub showTableSheetList()

    On Error GoTo err
    
    ' �e�[�u���V�[�g�ꗗ�t�H�[����\������
    frmTableSheetListVar.ShowExt vbModeless

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���e�[�u���V�[�g�쐬
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub createTableSheet()

    On Error GoTo err

    createEmptyBook
    ' �e�[�u���쐬�t�H�[����\������
    frmTableSheetCreatorVar.ShowExt vbModal, applicationSetting, dbConn

    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���e�[�u���V�[�g�X�V
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub updateTableSheet()

    On Error GoTo err
    
    ' �e�[�u���V�[�g���ǂ������m�F����B�i���s�����ꍇ�A�G���[�����s�����j
    tableSheetReader.validTableSheet
    
    ' �e�[�u���V�[�g�X�V�t�H�[����\������
    frmTableSheetUpdateVar.ShowExt vbModal
    
    Exit Sub
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���e�[�u���V�[�g�̃f�[�^�I�u�W�F�N�g�𐶐�����B
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function createDataOfTableSheet(ByRef table As ValDbDefineTable _
                                      , ByVal recFormat As REC_FORMAT) As ValTableWorksheet

    ' �e�[�u���V�[�g�𐶐�����
    Dim tableSheet As ValTableWorksheet
    Set tableSheet = New ValTableWorksheet
    
    ' �X�L�[�}�̏ȗ��L����ݒ肷��
    If applicationSetting.schemaUse = applicationSetting.SCHEMA_USE_ONE Then
        tableSheet.omitsSchema = True
    Else
        tableSheet.omitsSchema = False
    End If
    ' �s�t�H�[�}�b�g�̐���
    tableSheet.recFormat = recFormat
    ' �e�[�u����`�̐ݒ�
    Set tableSheet.table = table

    Set createDataOfTableSheet = tableSheet

End Function

' =========================================================
' ���e�[�u����`���X�g���i�[���ꂽ���X�g�I�u�W�F�N�g����Ƀe�[�u���V�[�g�̃��X�g�I�u�W�F�N�g�𐶐�����B
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function createDataOfTableSheetList(ByRef tableList As ValCollection _
                                          , ByVal recFormat As REC_FORMAT) As ValCollection
                                          
    Dim tableSheetList As ValCollection
    Set tableSheetList = New ValCollection
    
    Dim table      As ValDbDefineTable
    Dim tableSheet As ValTableWorksheet
    
    ' �e�[�u�����X�g����e�[�u����`�����o���e�[�u���V�[�g���ɕϊ�����
    ' ���X�g�ɐݒ肷��
    For Each table In tableList.col
    
        Set tableSheet = createDataOfTableSheet(table, recFormat)
        tableSheetList.setItem tableSheet
    Next
    
    Set createDataOfTableSheetList = tableSheetList

End Function

' =========================================================
' ���u�b�N���\������Ă��Ȃ��ꍇ�A��̃u�b�N��\������B
'
' �T�v�@�@�@�F
'
' =========================================================
Private Function createEmptyBook()

    ' �u�b�N�I�u�W�F�N�g
    Dim book As Workbook
    
    ' �u�b�N�I�u�W�F�N�g���擾����
    Set book = ActiveWorkbook
    
    ' �u�b�N�I�u�W�F�N�g�̃`�F�b�N
    If book Is Nothing Then
    
        Set book = Application.Workbooks.Add
        book.activate
    End If

End Function