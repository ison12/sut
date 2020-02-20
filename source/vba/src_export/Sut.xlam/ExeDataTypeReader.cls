VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeDataTypeReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �C�ӂ̃V�[�g����f�[�^��ނ̒�`��ǂݍ��ރN���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/08/16�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���[�N�V�[�g
Public sheet         As Worksheet
' DB��
Public dbName        As String

' =========================================================
' ���萔�v���p�e�B
' �@Class���W���[���ł́APublic Const����`�ł��Ȃ��̂Ńv���p�e�B�ő�p����
' =========================================================
' �Z���s�ԍ� DB��
Public Property Get CELL_ROW_DBNAME() As Long

    CELL_ROW_DBNAME = 1
End Property

' �Z���s�ԍ� �I�t�Z�b�g
Public Property Get CELL_ROW_OFFSET() As Long

    CELL_ROW_OFFSET = 3
End Property

' �Z����ԍ� ���e����
Public Property Get CELL_COL_LITERAL() As Long

    CELL_COL_LITERAL = 2
End Property

' �Z����ԍ� �f�[�^���
Public Property Get CELL_COL_DATA_TYPE() As Long

    CELL_COL_DATA_TYPE = 3
End Property

' �Z����ԍ� �X�V���̏���
Public Property Get CELL_COL_FORMAT_UPDATE() As Long

    CELL_COL_FORMAT_UPDATE = 4
End Property

' �Z����ԍ� �Q�Ǝ��̏���
Public Property Get CELL_COL_FORMAT_SELECT() As Long

    CELL_COL_FORMAT_SELECT = 5
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
' ��DB�������ԍ�����������B
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet    �V�[�g��
'     �@�@�@�FdbName   DB��
' �߂�l�@�@�F��ԍ�
'
' =========================================================
Private Function findColumnDbName(ByVal sheet As Worksheet, ByVal dbName As String) As Long

    findColumnDbName = 0

    ' Find�Ō�������
    Dim findRange As Range
    Set findRange = sheet.Range(CELL_ROW_DBNAME & ":" & CELL_ROW_DBNAME).Find(what:=dbName, MatchCase:=False, MatchByte:=True, SearchFormat:=False)
    
    If Not findRange Is Nothing Then
        ' ���������ꍇ
        findColumnDbName = findRange.column
        Exit Function
    End If
    
End Function

' =========================================================
' ���C�ӂ̃V�[�g����f�[�^��ނ̒�`��ǂݎ��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Function execute() As Variant

    ' �J������񃊃X�g
    Dim colList As ValCollection
    ' �J�������
    Dim colType As ValDbColumnTypeColInfo

    ' �J������񃊃X�g�̃T�C�Y���m�ۂ���
    Set colList = New ValCollection
    
    ' ��ʒu�ƂȂ��ԍ�����������
    Dim offsetColumn As Long
    offsetColumn = findColumnDbName(sheet, dbName)
    
    If offsetColumn <= 0 Then
        ' ������Ȃ��ꍇ
        Set execute = colList
        Exit Function
    End If
    
    ' ��`���̃T�C�Y
    Dim size As Long
    ' ��`���̃T�C�Y���擾����
    size = ExcelUtil.getCellEndOfLastInputtedRow(sheet, offsetColumn - 2 + CELL_COL_LITERAL) - CELL_ROW_OFFSET + 1
    
    Dim i As Long
    
    For i = 0 To size - 1
    
        ' �J�������𐶐�����
        Set colType = New ValDbColumnTypeColInfo
    
        ' �f�[�^��ނ��擾����i��ɑ啶���ŃJ���������������邽�߁A�啶���ɕϊ�����j
        colType.columnName = UCase$(sheet.Cells(i + CELL_ROW_OFFSET, offsetColumn - 2 + CELL_COL_DATA_TYPE))
        ' �X�V���̏��������擾����
        colType.formatUpdate = sheet.Cells(i + CELL_ROW_OFFSET, offsetColumn - 2 + CELL_COL_FORMAT_UPDATE)
        ' �Q�Ǝ��̏��������擾����
        colType.formatSelect = sheet.Cells(i + CELL_ROW_OFFSET, offsetColumn - 2 + CELL_COL_FORMAT_SELECT)
        
        colList.setItem colType, colType.columnName
    Next
    
    Set execute = colList
    
End Function