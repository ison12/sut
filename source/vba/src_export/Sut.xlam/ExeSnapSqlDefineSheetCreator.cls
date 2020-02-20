VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeSnapSqlDefineSheetCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�i�b�v�V���b�gSQL��`�V�[�g�𐶐�����N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/01/03 �V�K�쐬
'
' ���L�����F
' *********************************************************

' ________________________________________________________
' �C�x���g
' ________________________________________________________

' =========================================================
' ���X�i�b�v�V���b�gSQL��`�V�[�g���������J�n�O�C�x���g
'
' �T�v�@�@�@�F�X�i�b�v�V���b�gSQL��`�V�[�g�̐����O�ɌĂяo�����C�x���g
' �����@�@�@�F
'
' =========================================================
Public Event createSnapshotSqlDefineSheetBegin()

' =========================================================
' ���X�i�b�v�V���b�gSQL��`�V�[�g�����C�x���g
'
' �T�v�@�@�@�F�X�i�b�v�V���b�gSQL��`�V�[�g�̐����������ɌĂяo�����C�x���g
' �����@�@�@�F
'
' =========================================================
Public Event createSnapshotSqlDefineSheet()

' =========================================================
' �v���p�e�B
' =========================================================
' �A�v���P�[�V�����ݒ���
Private m_applicationSetting As ValApplicationSetting
' �����Ώۃ��[�N�u�b�N
Private m_book               As Workbook
' �L�����Z���t���O
Private m_canceled           As Boolean

Public Property Let applicationSetting(ByVal vNewValue As Variant)

    Set m_applicationSetting = vNewValue
End Property

Public Property Let book(ByVal vNewValue As Variant)

    Set m_book = vNewValue
    
End Property

Public Property Let canceled(ByVal vNewValue As Variant)

    m_canceled = vNewValue
    
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

' =========================================================
' ���V�[�g�𐶐�����
'
' �T�v�@�@�@�F�w�肳�ꂽ�e�[�u���̃V�[�g�𐶐�����
' �����@�@�@�F
' �߂�l    �F�V�[�g
'
' =========================================================
Public Function createSheet() As Worksheet

    ' �V�[�g��
    Dim sheetName     As String

    ' ���X�i�b�v�V���b�gSQL��`�V�[�g�����C�x���g
    RaiseEvent createSnapshotSqlDefineSheetBegin

    sheetName = "SQL_" & Format(Now, "yyyymmdd_hhmmss")

    ' �V�[�g�̍폜�i���݂��Ă�����j�Ƒ}�������s����
    deleteAndInsertSheet sheetName
    
    ' �V���ɃR�s�[���ꂽ�V�[�g�I�u�W�F�N�g���擾����
    Set createSheet = m_book.Worksheets(sheetName)
    
    ' ���V�V�[�g�Ƀe�[�u����`���𖄂ߍ���
    With createSheet

        ' �V�[�g�����ɂ���
        .visible = xlSheetVisible
        .Cells(ConstantsSnapSqlDefine.NAME_ROW, ConstantsSnapSqlDefine.NAME_COL) = sheetName

    End With
    
    ' ���X�i�b�v�V���b�gSQL��`�V�[�g���������C�x���g
    RaiseEvent createSnapshotSqlDefineSheet

End Function


' =========================================================
' ���V�[�g�̍폜�E�}��
'
' �T�v�@�@�@�F�V�[�g�����ɑ��݂���ꍇ�͍폜���A�}�����s���B
' �����@�@�@�FSnapshotSqlDefineSheet      �e�[�u�����
' �@�@�@�@�@�@targetSheetName �ΏۃV�[�g��
'
' =========================================================
Private Sub deleteAndInsertSheet(ByRef targetSheetName As String)

    ' �R�s�[�V�[�g
    Dim copySheetName As String
    copySheetName = ConstantsSnapSqlDefine.SHEET_NAME_TEMPLATE

    ExcelUtil.copySheetAppend SutWorkbook _
                            , copySheetName _
                            , m_book _
                            , targetSheetName

End Sub

