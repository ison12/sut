VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExeTableSheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �V�[�g���烌�R�[�h����ǂݎ�鏈�����`�����N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/02/01�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���[�N�V�[�g
Public sheet     As Worksheet
' DB�R�l�N�V����
Public conn      As Object
' �L�����Z���t���O
Public canceled  As Boolean

' =========================================================
' ���J�������X�g�Ǎ��J�n�O�C�x���g
'
' �T�v�@�@�@�F�J�������X�g�̓Ǎ����J�n����O�ɌĂяo�����C�x���g
' �����@�@�@�Fcnt ����
'
' =========================================================
Public Event readColumnListBegin(ByVal cnt As Long)

' =========================================================
' ���J�������Ǎ��C�x���g
'
' �T�v�@�@�@�F�J�������Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�Fcolumn �J�������
'
' =========================================================
Public Event readColumn(ByRef column As ValDbDefineColumn)

' =========================================================
' ���J�������X�g�Ǎ��C�x���g
'
' �T�v�@�@�@�F�J�������X�g�Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�FcolumnList �J�������X�g
'
' =========================================================
Public Event readColumnList(ByRef columnList As ValCollection)

' =========================================================
' ���e�[�u����`�Ǎ��J�n�O�C�x���g
'
' �T�v�@�@�@�F�e�[�u����`�̓Ǎ����J�n����O�ɌĂяo�����C�x���g
' �����@�@�@�F
'
' =========================================================
Public Event readTableInfoBegin()

' =========================================================
' ���e�[�u����`�Ǎ��C�x���g
'
' �T�v�@�@�@�F�e�[�u����`�̓Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�FtableSheet �e�[�u���V�[�g
'
' =========================================================
Public Event readTableInfo(ByRef tableSheet As ValTableWorksheet)

' =========================================================
' �����R�[�h���X�g�Ǎ��J�n�O�C�x���g
'
' �T�v�@�@�@�F���R�[�h���X�g�̐������J�n����O�ɌĂяo�����C�x���g
' �����@�@�@�Fcnt ����
'
' =========================================================
Public Event readRecordListBegin(ByVal cnt As Long)

' =========================================================
' �����R�[�h���X�g�Ǎ��C�x���g
'
' �T�v�@�@�@�F���R�[�h���X�g�������ɌĂяo�����C�x���g
' �����@�@�@�FrecordList ���R�[�h
'
' =========================================================
Public Event readRecordList(ByRef recordList As ValRecordList)

' =========================================================
' ���͈͑I�����̃��R�[�h���X�g�y�R���N�V�����z�Ǎ��J�n�O�C�x���g
'
' �T�v�@�@�@�F���R�[�h���X�g�y�R���N�V�����z�̐������J�n����O�ɌĂяo�����C�x���g
' �����@�@�@�Fcnt ����
'
' =========================================================
Public Event readRecordListBySelectionBegin(ByVal cnt As Long)

' =========================================================
' ���͈͑I�����̃��R�[�h���X�g�Ǎ��C�x���g
'
' �T�v�@�@�@�F���R�[�h���X�g�y�R���N�V�����z�Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�FrecordList ���R�[�h���X�g
'
' =========================================================
Public Event readRecordListBySelectionOne(ByRef recordList As ValRecordList)

' =========================================================
' ���͈͑I�����̃��R�[�h���X�g�y�R���N�V�����z�Ǎ��C�x���g
'
' �T�v�@�@�@�F���R�[�h���X�g�y�R���N�V�����z�Ǎ����ɌĂяo�����C�x���g
' �����@�@�@�FrecordLists ���R�[�h���X�g�y�R���N�V�����z
'
' =========================================================
Public Event readRecordListBySelection(ByRef recordLists As ValCollection)

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
' ���e�[�u���V�[�g����
'
' �T�v�@�@�@�F�ΏۃV�[�g�i�����o��sheet�j��
' �@�@�@�@�@�@�e�[�u���V�[�g�ł��邩�𔻒肷��B
'
' �����@�@�@�F
' �߂�l�@�@�FTrue �e�[�u���V�[�g�̏ꍇ
' =========================================================
Public Function isTableSheet() As Boolean

    ' �e�[�u���V�[�g����p�̃I�[�g�V�F�C�v
    Dim s As Shape

    ' �߂�l������������
    isTableSheet = False
    
    ' ------------------------------------------------
    ' �G���[���������Ă��p������
    On Error Resume Next

    Set s = sheet.Shapes(ConstantsTable.U_SHEET_CHECK_IMAGE)
    
    ' �擾�ł����ꍇ
    If Not s Is Nothing Then
    
        isTableSheet = True
        Exit Function
    End If
    
    Set s = sheet.Shapes(ConstantsTable.R_SHEET_CHECK_IMAGE)
    
    ' �擾�ł����ꍇ
    If Not s Is Nothing Then
    
        isTableSheet = True
        Exit Function
    End If
        
    On Error GoTo 0
    ' ------------------------------------------------

End Function

' =========================================================
' ���e�[�u���V�[�g����
'
' �T�v�@�@�@�F�ΏۃV�[�g�i�����o��sheet�j���e�[�u���V�[�g�ł��邩�𔻒肷��B
' �@�@�@�@�@�@�e�[�u���V�[�g�ł͂Ȃ��ꍇ�A�G���[�𔭍s����B
'
' �����@�@�@�F
' �߂�l�@�@�F
' =========================================================
Public Sub validTableSheet()

    ' �e�[�u���V�[�g�ł͂Ȃ��ꍇ
    If isTableSheet = False Then
    
        err.Raise ConstantsError.ERR_NUMBER_IS_NOT_TABLE_SHEET _
            , _
            , ConstantsError.ERR_DESC_IS_NOT_TABLE_SHEET
    End If

End Sub
' =========================================================
' ���s�t�H�[�}�b�g�擾
'
' �T�v�@�@�@�F�ΏۃV�[�g�i�����o��sheet�j�̍s�t�H�[�}�b�g���擾����
'
' �����@�@�@�F
' �߂�l�@�@�FrecFormatToUnder
' �@�@�@�@�@�@recFormatToRight
' �@�@�@�@�@�@-1
'
' =========================================================
Public Function getRowFormat() As REC_FORMAT

    ' �߂�l
    Dim ret As REC_FORMAT
    
    ' �e�[�u���V�[�g����p�̃I�[�g�V�F�C�v
    Dim s As Shape

    ' �߂�l������������
    ret = -1
    
    ' ------------------------------------------------
    ' �G���[���������Ă��p������
    On Error Resume Next

    Set s = sheet.Shapes("SUT_WORKSHEET_MARK_TO_UNDER")
    
    ' �擾�ł����ꍇ
    If Not s Is Nothing Then
    
        ' �߂�l�ϐ��ɍs�t�H�[�}�b�g��ݒ肷��
        ret = REC_FORMAT.recFormatToUnder
        
        ' �߂�l��ݒ肷��
        getRowFormat = ret
        
        Exit Function
    End If
    
    Set s = sheet.Shapes("SUT_WORKSHEET_MARK_TO_RIGHT")
    
    ' �擾�ł����ꍇ
    If Not s Is Nothing Then
    
        ' �߂�l�ϐ��ɍs�t�H�[�}�b�g��ݒ肷��
        ret = REC_FORMAT.recFormatToRight
        
        ' �߂�l��ݒ肷��
        getRowFormat = ret
        
        Exit Function
    End If
        
    On Error GoTo 0
    ' ------------------------------------------------

    getRowFormat = ret

End Function

' =========================================================
' ���e�[�u����`�ǂݍ���
'
' �T�v�@�@�@�F�V�[�g��̃e�[�u����`��ǂݍ����
' �@�@�@�@�@�@�e�[�u����`�V�[�g�I�u�W�F�N�g�𐶐�����B
'
' �����@�@�@�FtableInfoOnly �e�[�u�����̂ݎ擾�i�J�������͎擾���Ȃ��j
' �߂�l�@�@�F�e�[�u���V�[�g
'
' =========================================================
Public Function readTableInfo(Optional tableInfoOnly As Boolean = False) As ValTableWorksheet

    ' �e�[�u����`
    Dim table      As ValDbDefineTable
    ' �e�[�u�����𐶐�����
    Set table = New ValDbDefineTable
    ' �e�[�u���V�[�g
    Dim tableSheet As ValTableWorksheet
    ' �e�[�u���V�[�g���𐶐�����
    Set tableSheet = New ValTableWorksheet
    Set tableSheet.table = table
    
    ' �e�[�u����
    Dim tableName As String
    
    ' �e�[�u�������J�n�C�x���g
    RaiseEvent readTableInfoBegin
    
    With sheet

        ' �e�[�u�������擾����
        tableName = .Cells(ConstantsTable.TABLE_NAME_ROW _
                         , ConstantsTable.TABLE_NAME_COL)
        
        
        ' �e�[�u�����擾���V�[�g��ݒ肷��
        Set tableSheet.sheet = sheet
        tableSheet.sheetName = sheet.name
        
        If InStr(tableName, ".") > 0 Then
        
            ' �X�L�[�}�i�ꎞ�ϐ��j
            Dim tempSchema As String
            ' �e�[�u���i�ꎞ�ϐ��j
            Dim tempTable  As String
            
            ' �X�L�[�}�E�e�[�u�����𒊏o���ݒ肷��
            DBUtil.extractSchemaTable tableName, tempSchema, tempTable
            
            table.schemaName = tempSchema
            table.tableName = tempTable
            
            ' �X�L�[�}�����t������Ă���ꍇ�́A�X�L�[�}�����ȗ�����Ă��Ȃ��Ƃ݂Ȃ�
            tableSheet.omitsSchema = False
        Else
        
            ' �e�[�u������ݒ肷��
            table.tableName = tableName
            ' �X�L�[�}����ݒ肷��i�B���Z������擾�j
            table.schemaName = .Cells(ConstantsTable.SCHEMA_NAME_ROW _
                                    , ConstantsTable.SCHEMA_NAME_COL)
            
            ' �X�L�[�}�����t������Ă��Ȃ��ꍇ�́A�X�L�[�}�����ȗ�����Ă���Ƃ݂Ȃ�
            tableSheet.omitsSchema = True
        End If
        
        ' �e�[�u���R�����g��ݒ肷��
        table.tableComment = .Cells(TABLE_NAME_LOG_ROW _
                                  , TABLE_NAME_LOG_COL)
        
        ' �s�t�H�[�}�b�g��ݒ肷��
        tableSheet.recFormat = getRowFormat
        
        If tableInfoOnly = False Then
        
            If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
            
                readTableInfoToUnder tableSheet
            
            ElseIf tableSheet.recFormat = REC_FORMAT.recFormatToRight Then
            
                readTableInfoToRight tableSheet
                
            End If
        
        End If
        
    End With ' sheet

    ' �e�[�u����`���������C�x���g
    RaiseEvent readTableInfo(tableSheet)

    ' �߂�l��ݒ肷��
    Set readTableInfo = tableSheet

End Function

' =========================================================
' ���e�[�u����`�ǂݍ��݁iX�����j
'
' �T�v�@�@�@�F�V�[�g��̃e�[�u����`��X�����̏ꍇ������𗘗p����B
' �@�@�@�@�@�@��ɃJ�������X�g��ǂݍ��ށB
'
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �߂�l�@�@�F
'
' =========================================================
Private Function readTableInfoToUnder(ByRef tableSheet As ValTableWorksheet)

    Dim i As Long
    
    ' ���s�R�[�h�i�Z���̉��s�R�[�h��LF�j
    Dim newline As String: newline = vbLf

    ' �V�[�g�̍ő��
    Dim sheetColSize As Long: sheetColSize = ExcelUtil.getSizeOfSheetCol(tableSheet.sheet)

    ' �e�[�u����`
    Dim table      As ValDbDefineTable
    Set table = tableSheet.table
    ' �J����
    Dim column     As ValDbDefineColumn
    
    ' �e�[�u�������񃊃X�g
    Dim tableConstList As ValCollection
    Set tableConstList = tableSheet.table.constraintsList
    
    ' �e�[�u��������iPK�j
    Dim tableConstPk   As New ValDbDefineTableConstraints
    
    ' �e�[�u��������i�ėp�I�Ɏg�p�j
    Dim tableConst     As ValDbDefineTableConstraints
    
    ' ���[�v�p���t���O
    Dim continue   As Variant
    
    ' ���ݗ�
    Dim currentCol As Long
    
    ' �J�������
    Dim columnName  As String    ' ����
    Dim columnType  As String    ' �^
    Dim columnNull  As String    ' NULL��
    Dim columnDef   As String    ' �f�t�H���g
    Dim columnPk    As String    ' PK
    Dim columnUk    As String    ' UNIQ
    Dim columnFk    As String    ' FK
    
    Dim columnUkArray()  As String    ' UNIQ�z��
    Dim columnFkArray()  As String    ' FK  �z��
    
    ' �J�����T�C�Y
    Dim columnSize As Long
    ' �J�����T�C�Y���擾����
    columnSize = getColumnSize(tableSheet)
    
    ' �J�����T�C�Y��0�̏ꍇ
    If columnSize <= 0 Then
    
        Exit Function
    End If
    
    ' �J���������J�n�C�x���g
    RaiseEvent readColumnListBegin(columnSize)
    
    ' ---------------------------------------------------------------------
    With sheet
    
        ' ��̍ŏ��̈ʒu���擾
        currentCol = ConstantsTable.U_COLUMN_OFFSET_COL
    
        ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
        continue = .Cells(ConstantsTable.U_COLUMN_NAME_ROW, currentCol)
        
        ' ���[�v�p�������icontinue����ł͂Ȃ��ꍇ�p���j
        Do While continue <> Empty
        
            ' ���R�[�h�ɃJ��������ǉ�����
            Set column = New ValDbDefineColumn
            
            ' �J���������擾
            columnName = .Cells(ConstantsTable.U_COLUMN_NAME_ROW, currentCol)
            column.columnName = columnName
            
            
            ' �J�����^���擾
            columnType = .Cells(ConstantsTable.U_COLUMN_TYPE_ROW, currentCol)
            
            column.dataTypeFormal = columnType
            
            ' �J�����^�������ϊ��i�����񒆂̊��ʂ���菜���j
            columnType = convertColumnType(columnType)

            column.dataType = columnType
            
            ' �J������NULL�ł��邩���擾
            columnNull = .Cells(ConstantsTable.U_COLUMN_NULL_ROW, currentCol)
            
            If columnNull <> "" Then
            
                column.isNull = True
            Else
            
                column.isNull = False
            End If
            
            ' �J�����̃f�t�H���g�l���擾
            columnDef = .Cells(ConstantsTable.U_COLUMN_DEF_ROW, currentCol)
            column.default = columnDef
            
            ' �J������PK���ǂ������擾
            columnPk = .Cells(ConstantsTable.U_COLUMN_PK_ROW, currentCol)
            
            If columnPk <> "" Then
            
                tableConstPk.addColumn column
            End If
            
            ' �J������UK���ǂ������擾
            columnUk = .Cells(ConstantsTable.U_COLUMN_UK_ROW, currentCol)
            
            If columnUk <> "" Then
            
                columnUk = replace(columnUk, vbCr, "")
                columnUkArray = Split(columnUk, newline)
                
                For i = LBound(columnUkArray) To UBound(columnUkArray)
                
                    ' �e�[�u�������񂪃��X�g�Ɋi�[����Ă��邩���m�F����
                    Set tableConst = tableConstList.getItem(columnUkArray(i))
                    ' �e�[�u�������񂪑��݂��Ȃ��ꍇ
                    If tableConst Is Nothing Then
                    
                        ' �V����������𐶐�����
                        Set tableConst = New ValDbDefineTableConstraints
                        tableConst.schemaName = table.schemaName
                        tableConst.tableName = table.tableName
                        tableConst.constraintType = tableConstUk
                        tableConst.constraintName = columnUkArray(i)
                    End If
                    
                    tableConst.addColumn column
                Next
                
            End If
            
            ' �J������FK���ǂ������擾
            columnFk = .Cells(ConstantsTable.U_COLUMN_REFER_ROW, currentCol)
            If columnFk <> "" Then
            
                columnFk = replace(columnFk, vbCr, "")
                columnFkArray = Split(columnFk, newline)
                
                For i = LBound(columnFkArray) To UBound(columnFkArray)
                
                    ' �e�[�u�������񂪃��X�g�Ɋi�[����Ă��邩���m�F����
                    Set tableConst = tableConstList.getItem(columnFkArray(i))
                    ' �e�[�u�������񂪑��݂��Ȃ��ꍇ
                    If tableConst Is Nothing Then
                    
                        ' �V����������𐶐�����
                        Set tableConst = New ValDbDefineTableConstraints
                        tableConst.schemaName = table.schemaName
                        tableConst.tableName = table.tableName
                        tableConst.constraintType = tableConstFk
                        tableConst.constraintName = columnFkArray(i)
                    End If
                    
                    tableConst.addColumn column
                Next
                
            End If
            
            table.addColumn column
            
            ' �J�������������C�x���g
            RaiseEvent readColumn(column)
        
            ' �����i�߂�
            currentCol = currentCol + 1
            
            If sheetColSize < currentCol Then
            
                Exit Do
            End If
            
            ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
            continue = .Cells(ConstantsTable.U_COLUMN_NAME_ROW, currentCol)
        
        Loop
        ' ---------------------------------------------------------------------
    
        ' PK����ƂȂ�J�������P���ȏ゠��ꍇ�A�e�[�u��������Ƃ��Ēǉ�����
        If tableConstPk.columnList.count > 0 Then
        
            tableConstPk.schemaName = table.schemaName
            tableConstPk.tableName = table.tableName
            tableConstPk.constraintName = "PrimaryKey"
            tableConstPk.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk
            
            tableConstList.setItem tableConstPk
        End If
        
        ' �J�������X�g���������C�x���g
        RaiseEvent readColumnList(table.columnList)
    
    End With ' sheet

End Function

' =========================================================
' ���e�[�u����`�ǂݍ��݁iY�����j
'
' �T�v�@�@�@�F�V�[�g��̃e�[�u����`��Y�����̏ꍇ������𗘗p����B
' �@�@�@�@�@�@��ɃJ�������X�g��ǂݍ��ށB
'
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �߂�l�@�@�F
'
' =========================================================
Private Function readTableInfoToRight(ByRef tableSheet As ValTableWorksheet)

    ' ���s�R�[�h�i�Z���̉��s�R�[�h��LF�j
    Dim newline As String: newline = vbLf
    
    Dim i As Long
    
    ' �V�[�g�̍ő�s��
    Dim sheetRowSize As Long: sheetRowSize = ExcelUtil.getSizeOfSheetRow(tableSheet.sheet)

    ' �e�[�u����`
    Dim table      As ValDbDefineTable
    Set table = tableSheet.table
    ' �J������`
    Dim columnList As ValCollection
    ' �J����
    Dim column     As ValDbDefineColumn
    
    ' �e�[�u�������񃊃X�g
    Dim tableConstList As ValCollection
    Set tableConstList = tableSheet.table.constraintsList
    
    ' �e�[�u��������iPK�j
    Dim tableConstPk   As New ValDbDefineTableConstraints
    
    ' �e�[�u��������i�ėp�I�Ɏg�p�j
    Dim tableConst     As ValDbDefineTableConstraints
    
    ' ���[�v�p���t���O
    Dim continue   As Variant
    
    ' ���ݍs
    Dim currentRow As Long
    
    ' �J�������
    Dim columnName  As String    ' ����
    Dim columnType  As String    ' �^
    Dim columnNull  As String    ' NULL��
    Dim columnDef   As String    ' �f�t�H���g
    Dim columnPk    As String    ' PK
    Dim columnUk    As String    ' UNIQ
    Dim columnFk    As String    ' FK
    
    Dim columnUkArray()  As String    ' UNIQ�z��
    Dim columnFkArray()  As String    ' FK  �z��
    
    ' �J�����T�C�Y
    Dim columnSize As Long
    ' �J�����T�C�Y���擾����
    columnSize = getColumnSize(tableSheet)
    
    ' �J�����T�C�Y��0�̏ꍇ
    If columnSize <= 0 Then
    
        Exit Function
    End If
    
    ' �J���������J�n�C�x���g
    RaiseEvent readColumnListBegin(columnSize)
    
    ' ---------------------------------------------------------------------
    With sheet
    

        ' ��̍ŏ��̈ʒu���擾
        currentRow = ConstantsTable.R_COLUMN_OFFSET_ROW
    
        ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
        continue = .Cells(currentRow, ConstantsTable.R_COLUMN_NAME_COL)
        
        ' ���[�v�p�������icontinue����ł͂Ȃ��ꍇ�p���j
        Do While continue <> Empty
        
            ' ���R�[�h�ɃJ��������ǉ�����
            Set column = New ValDbDefineColumn
            
            ' �J���������擾
            columnName = .Cells(currentRow, ConstantsTable.R_COLUMN_NAME_COL)
            column.columnName = columnName
            
            
            ' �J�����^���擾
            columnType = .Cells(currentRow, ConstantsTable.R_COLUMN_TYPE_COL)
            
            column.dataTypeFormal = columnType
            
            ' �J�����^�������ϊ��i�����񒆂̊��ʂ���菜���j
            columnType = convertColumnType(columnType)
            
            column.dataType = columnType
            ' �J������NULL�ł��邩���擾
            columnNull = .Cells(currentRow, ConstantsTable.R_COLUMN_NULL_COL)
            
            If columnNull <> "" Then
            
                column.isNull = True
            Else
            
                column.isNull = False
            End If
            
            ' �J�����̃f�t�H���g�l���擾
            columnDef = .Cells(currentRow, ConstantsTable.R_COLUMN_DEF_COL)
            column.default = columnDef
            
            ' �J������PK���ǂ������擾
            columnPk = .Cells(currentRow, ConstantsTable.R_COLUMN_PK_COL)
            
            If columnPk <> "" Then
            
                tableConstPk.addColumn column
            End If
            
            ' �J������UK���ǂ������擾
            columnUk = .Cells(currentRow, ConstantsTable.R_COLUMN_UK_COL)
            
            If columnUk <> "" Then
            
                columnUk = replace(columnUk, vbCr, "")
                columnUkArray = Split(columnUk, newline)
                
                For i = LBound(columnUkArray) To UBound(columnUkArray)
                
                    ' �e�[�u�������񂪃��X�g�Ɋi�[����Ă��邩���m�F����
                    Set tableConst = tableConstList.getItem(columnUkArray(i))
                    ' �e�[�u�������񂪑��݂��Ȃ��ꍇ
                    If tableConst Is Nothing Then
                    
                        ' �V����������𐶐�����
                        Set tableConst = New ValDbDefineTableConstraints
                        tableConst.schemaName = table.schemaName
                        tableConst.tableName = table.tableName
                        tableConst.constraintType = tableConstUk
                        tableConst.constraintName = columnUkArray(i)
                    End If
                    
                    tableConst.addColumn column
                Next
                
            End If
            
            ' �J������FK���ǂ������擾
            columnFk = .Cells(currentRow, ConstantsTable.R_COLUMN_REFER_COL)
            If columnFk <> "" Then
            
                columnFk = replace(columnFk, vbCr, "")
                columnFkArray = Split(columnFk, newline)
                
                For i = LBound(columnFkArray) To UBound(columnFkArray)
                
                    ' �e�[�u�������񂪃��X�g�Ɋi�[����Ă��邩���m�F����
                    Set tableConst = tableConstList.getItem(columnFkArray(i))
                    ' �e�[�u�������񂪑��݂��Ȃ��ꍇ
                    If tableConst Is Nothing Then
                    
                        ' �V����������𐶐�����
                        Set tableConst = New ValDbDefineTableConstraints
                        tableConst.schemaName = table.schemaName
                        tableConst.tableName = table.tableName
                        tableConst.constraintType = tableConstFk
                        tableConst.constraintName = columnFkArray(i)
                    End If
                    
                    tableConst.addColumn column
                Next
                
            End If
            
            table.addColumn column
            
            ' �J�������������C�x���g
            RaiseEvent readColumn(column)
        
            ' �s����i�߂�
            currentRow = currentRow + 1
            
            If sheetRowSize < currentRow Then
            
                Exit Do
            End If
            
            ' ���[�v���J��Ԃ��s�����ǂ����𔻒肷�邽�߂̕�������擾����
            continue = .Cells(currentRow, ConstantsTable.R_COLUMN_NAME_COL)
        
        Loop
        ' ---------------------------------------------------------------------
    
        ' PK����ƂȂ�J�������P���ȏ゠��ꍇ�A�e�[�u��������Ƃ��Ēǉ�����
        If tableConstPk.columnList.count > 0 Then
        
            tableConstPk.schemaName = table.schemaName
            tableConstPk.tableName = table.tableName
            tableConstPk.constraintName = "PrimaryKey"
            tableConstPk.constraintType = TABLE_CONSTANTS_TYPE.tableConstPk
            
            tableConstList.setItem tableConstPk
        End If
        
        ' �J�������X�g���������C�x���g
        RaiseEvent readColumnList(columnList)
    
    End With ' sheet

End Function

' =========================================================
' ���V�[�g��̃��R�[�h����S�ēǂݍ���
'
' �T�v�@�@�@�F�V�[�g��̃��R�[�h����S�ēǂݍ���
'
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �߂�l�@�@�F���R�[�h���X�g
'
' =========================================================
Public Function readRecordListAll(ByRef tableSheet As ValTableWorksheet) As ValRecordList

    ' �߂�l
    Dim ret As ValRecordList
    
    ' ���R�[�h�T�C�Y
    Dim recordSize As Long
    
    ' ���R�[�h�T�C�Y���擾����
    recordSize = getRecordSize(tableSheet)
    
    ' �T�C�Y�� 0 �ȉ��̏ꍇ�A��̖߂�l��Ԃ��I������
    If recordSize <= 0 Then
    
        ' �߂�l��ݒ肷��
        Set readRecordListAll = New ValRecordList
        Exit Function
    End If
            
    ' �߂�l��ݒ肷��
    Set ret = readRecordList(tableSheet, 1, recordSize)
    
    Set readRecordListAll = ret

End Function

' =========================================================
' ���V�[�g��̃��R�[�h����C�ӂ͈̔͂œǂݍ���
'
' �T�v�@�@�@�F�V�[�g��̃��R�[�h����C�ӂ͈̔͂œǂݍ���
'
' �@�@�@�@�@�@����
' �@�@�@�@�@�@�E1���ǂݍ��ޏꍇ
' �@�@�@�@�@�@�@row     = 1
' �@�@�@�@�@�@�@rowSize = 1
'
' �@�@�@�@�@�@�E5�`10���ǂݍ��ޏꍇ
' �@�@�@�@�@�@�@row     = 5
' �@�@�@�@�@�@�@rowSize = 10
'
' �@�@�@�@�@�@�E1000���ǂݍ��ޏꍇ
' �@�@�@�@�@�@�@row     = 1
' �@�@�@�@�@�@�@rowSize = 1000
'
'
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �@�@�@�@�@�@row        �ǂݍ��݊J�n�s
' �@�@�@�@�@�@rowSize    �ǂݍ��݃T�C�Y
' �߂�l�@�@�F���R�[�h���X�g
'
' =========================================================
Public Function readRecordList(ByRef tableSheet As ValTableWorksheet _
                             , ByVal row As Long _
                             , ByVal rowSize As Long) As ValRecordList

    ' �e�[�u����`
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    
    ' �߂�l
    Dim recordList As ValRecordList
    ' �߂�l�𐶐�����
    Set recordList = New ValRecordList
        
    ' ���R�[�h
    Dim records As Variant
    
    ' �J�����T�C�Y
    Dim columnSize As Long
    
    ' �J�����T�C�Y���擾����
    columnSize = table.columnList.count
    ' �T�C�Y�� 0 �ȉ��̏ꍇ�A��̖߂�l��Ԃ��I������
    If columnSize <= 0 Then
    
        ' �߂�l��ݒ肷��
        Set readRecordList = recordList
        Exit Function
    End If
    
    ' �����R�[�h���X�g�����������J�n����邱�Ƃ��C�x���g�Ƃ��Ĕ��M����
    RaiseEvent readRecordListBegin(rowSize)
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        row = row + ConstantsTable.U_RECORD_OFFSET_ROW - 1
    
        ' �Z�����烌�R�[�h���擾����
        records = ExcelUtil.copyCellsToArray( _
                                 sheet _
                               , row _
                               , rowSize _
                               , ConstantsTable.U_COLUMN_OFFSET_COL _
                               , columnSize)
        
    
        ' ���R�[�h�z��擾���Ƃ̃I�t�Z�b�gRange�I�u�W�F�N�g��ݒ肷��
        Set recordList.valuesCellOffset = sheet.Cells(row, ConstantsTable.U_COLUMN_OFFSET_COL)
    
        ' ���R�[�h�z���ݒ肷��
        recordList.values = records
    
    Else
    
        row = row + ConstantsTable.R_RECORD_OFFSET_COL - 1
        
        ' �Z�����烌�R�[�h���擾����
        records = ExcelUtil.copyCellsToArray( _
                                 sheet _
                               , ConstantsTable.R_COLUMN_OFFSET_ROW _
                               , columnSize _
                               , row _
                               , rowSize)
                               
        ' X�EY�������t�ɓ���ւ���
        Dim swapArray As Variant
        
        swapArray = VBUtil.transposeDim(records)
        
        ' ��x records �z����폜����
        Erase records
        
        ' records �z��ɓ���ւ����z����Đݒ肷��
        records = swapArray
    
        ' ���R�[�h�z��擾���Ƃ̃I�t�Z�b�gRange�I�u�W�F�N�g��ݒ肷��
        Set recordList.valuesCellOffset = sheet.Cells(ConstantsTable.R_COLUMN_OFFSET_ROW, row)
    
        ' ���R�[�h�z���ݒ肷��
        recordList.values = records
            
    End If
    
    ' �߂�l��ݒ肷��
    Set readRecordList = recordList
    
    ' �����R�[�h���X�g���������������������Ƃ��C�x���g�Ƃ��Ĕ��M����
    RaiseEvent readRecordList(recordList)

End Function

' =========================================================
' �����R�[�h���i�͈͑I���j�̓Ǎ�
'
' �T�v�@�@�@�F���R�[�h�����Z������擾���ăN���X�Ɋi�[����
'
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �߂�l�@�@�F���R�[�h���N���X
'
' =========================================================
Public Function readRecordListBySelectionArea(ByRef tableSheet As ValTableWorksheet) As ValCollection

    ' �e�[�u����`
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    
    ' �߂�l�iValRecordList�̃R���N�V�����j
    Dim recordLists As ValCollection
    ' �߂�l������������
    Set recordLists = New ValCollection

    ' �I��̈�̈�Ӑ��`�F�b�N�p�̃R���N�V�����I�u�W�F�N�g
    ' �����̈悪�I������Ă��Ȃ������`�F�b�N����
    ' Selectin.Areas�́A�����Z����������I�����ꂽ�ꍇ�ł������Z����񂪕����i�[�����B
    ' �Ⴆ�΁AA1�Z����2�� Ctrl+�}�E�X���N���b�N �őI�������ƁAA1�Z����2���Ƃ��Ċi�[�����B���̂��߂Ƀ`�F�b�N���s���K�v������B
    Dim uniqueCheck    As New ValCollection
    ' �s�ԍ��܂��͗�ԍ����L�[�Ƃ��Ċi�[����
    Dim uniqueCheckKey As String

    ' ���R�[�h���X�g
    Dim recordList As ValRecordList
        
    ' ���R�[�h
    Dim records As Variant
    
    ' �I��̈�
    Dim selectionArea  As Areas
    ' �����W�I�u�W�F�N�g
    Dim selectionRange As Range
    
    ' �I��̈�̑���
    Dim cnt As Long
    
    Dim rowsCount As Long
    Dim colsCount As Long
    
    Dim i As Long
            

    
    ' �J�����T�C�Y
    Dim columnSize As Long: columnSize = table.columnList.count
        
    Set selectionArea = ActiveWindow.RangeSelection.Areas
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ' �I��̈�Ɋւ���`�F�b�N���s���A���R�[�h�̑������J�E���g����
        For Each selectionRange In selectionArea
    
            cnt = cnt + selectionRange.Rows.count
            
            ' �I��͈͂����͗̈�͈͓̔������`�F�b�N
            If selectionRange.row < ConstantsTable.U_RECORD_OFFSET_ROW Then
            
                ' �I��͈͂����R�[�h�J�n�I�t�Z�b�g��菬�����ꍇ
                ' �G���[�𔭍s����
                err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SELECTION _
                        , _
                        , ConstantsError.ERR_DESC_OUT_OF_RANGE_SELECTION
            End If

            ' �I��̈悩��Z�����P�������o���ԍ��Ƃ��ă`�F�b�N�p�R���N�V�����Ɋi�[
            ' �`�F�b�N�p�R���N�V�����Ɋ��Ɋi�[�ς݂ł���΃G���[�Ƃ���
            For i = 1 To selectionRange.Rows.count
                
                ' �L�[���擾����
                uniqueCheckKey = CStr(selectionRange.Offset.row + i - 1)
                
                ' �d�����Ă���ꍇ
                If uniqueCheck.exist(uniqueCheckKey) = True Then
                
                    ' �G���[�𔭍s����
                    err.Raise ERR_NUMBER_DUPLICATE_SELECTION_CELL _
                            , _
                            , ERR_DESC_DUPLICATE_SELECTION_CELL
                ' �d�����Ă��Ȃ��ꍇ
                Else
                
                    ' �i�[
                    uniqueCheck.setItem True, uniqueCheckKey
                End If
                
            Next
        Next
        
        ' �����R�[�h���X�g�����������J�n����邱�Ƃ��C�x���g�Ƃ��Ĕ��M����
        RaiseEvent readRecordListBySelectionBegin(cnt)
    
        For Each selectionRange In selectionArea
        
            records = ExcelUtil.copyCellsToArray( _
                                 ActiveWindow.ActiveSheet _
                               , selectionRange.row _
                               , selectionRange.Rows.count _
                               , ConstantsTable.U_COLUMN_OFFSET_COL _
                               , columnSize)
    
            ' ���R�[�h���X�g������������
            Set recordList = New ValRecordList
            
            ' ���R�[�h���X�g�ɔz���ݒ肷��
            recordList.values = records
            ' ���R�[�h���X�g��Range�I�u�W�F�N�g��ݒ肷��
            Set recordList.valuesCellOffset = selectionRange
            
            ' �߂�l�Ƀ��R�[�h���X�g��ݒ肷��
            recordLists.setItem recordList
            
            ' �����R�[�h���X�g��1�������ꂽ���Ƃ��C�x���g�Ƃ��Ĕ��M����
            RaiseEvent readRecordListBySelectionOne(recordList)
        Next
    
    Else
    
        ' �I��̈�Ɋւ���`�F�b�N���s���A���R�[�h�̑������J�E���g����
        For Each selectionRange In selectionArea
    
            cnt = cnt + selectionRange.Columns.count
            
            ' �I��͈͂����͗̈�͈͓̔������`�F�b�N
            If selectionRange.column < ConstantsTable.R_RECORD_OFFSET_COL Then
            
                ' �I��͈͂����R�[�h�J�n�I�t�Z�b�g��菬�����ꍇ
                ' �G���[�𔭍s����
                err.Raise ConstantsError.ERR_NUMBER_OUT_OF_RANGE_SELECTION _
                        , _
                        , ConstantsError.ERR_DESC_OUT_OF_RANGE_SELECTION
            End If

            ' �I��̈悩��Z�����P�������o���ԍ��Ƃ��ă`�F�b�N�p�R���N�V�����Ɋi�[
            ' �`�F�b�N�p�R���N�V�����Ɋ��Ɋi�[�ς݂ł���΃G���[�Ƃ���
            For i = 1 To selectionRange.Columns.count
                
                ' �L�[���擾����
                uniqueCheckKey = CStr(selectionRange.Offset.column + i - 1)
                
                ' �d�����Ă���ꍇ
                If uniqueCheck.exist(uniqueCheckKey) = True Then
                
                    ' �G���[�𔭍s����
                    err.Raise ERR_NUMBER_DUPLICATE_SELECTION_CELL _
                            , _
                            , ERR_DESC_DUPLICATE_SELECTION_CELL
                ' �d�����Ă��Ȃ��ꍇ
                Else
                
                    ' �i�[
                    uniqueCheck.setItem True, uniqueCheckKey
                End If
                
            Next

        Next
        
        ' �����R�[�h���X�g�����������J�n����邱�Ƃ��C�x���g�Ƃ��Ĕ��M����
        RaiseEvent readRecordListBySelectionBegin(cnt)
    
        For Each selectionRange In selectionArea
        
            records = ExcelUtil.copyCellsToArray( _
                                 ActiveWindow.ActiveSheet _
                               , ConstantsTable.R_COLUMN_OFFSET_ROW _
                               , columnSize _
                               , selectionRange.column _
                               , selectionRange.Columns.count)
    
            ' X�EY�������t�ɓ���ւ���
            Dim swapArray As Variant
            
            swapArray = VBUtil.transposeDim(records)
            
            ' ��x records �z����폜����
            Erase records
            
            ' records �z��ɓ���ւ����z����Đݒ肷��
            records = swapArray
    
            ' ���R�[�h���X�g������������
            Set recordList = New ValRecordList
            
            ' ���R�[�h���X�g�ɔz���ݒ肷��
            recordList.values = records
            ' ���R�[�h���X�g��Range�I�u�W�F�N�g��ݒ肷��
            Set recordList.valuesCellOffset = selectionRange
            
            ' �߂�l�Ƀ��R�[�h���X�g��ݒ肷��
            recordLists.setItem recordList
            
            ' �����R�[�h���X�g��1�������ꂽ���Ƃ��C�x���g�Ƃ��Ĕ��M����
            RaiseEvent readRecordListBySelectionOne(recordList)
        Next
    
    End If

    ' �����R�[�h���X�g���������������������Ƃ��C�x���g�Ƃ��Ĕ��M����
    RaiseEvent readRecordListBySelection(recordLists)

    ' �߂�l��ݒ肷��
    Set readRecordListBySelectionArea = recordLists
End Function

' =========================================================
' ���e�[�u���V�[�g�̃J�����T�C�Y�擾
'
' �T�v�@�@�@�F�e�[�u���V�[�g����J�����T�C�Y�̎擾���s��
'
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �߂�l�@�@�F�J�����T�C�Y
'
' =========================================================
Public Function getColumnSize(ByRef tableSheet As ValTableWorksheet) As Long

    ' �߂�l�ϐ�
    Dim ret As Long
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ret = ExcelUtil.getCellEndOfLastInputtedCol(sheet, ConstantsTable.U_COLUMN_NAME_ROW) - ConstantsTable.U_COLUMN_OFFSET_COL + 1
    
    Else
    
        ret = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsTable.R_COLUMN_NAME_COL) - ConstantsTable.R_COLUMN_OFFSET_ROW + 1
    End If
    
    ' �߂�l��ݒ肷��
    getColumnSize = ret
    
End Function

' =========================================================
' ���e�[�u���V�[�g�̍s�T�C�Y�擾
'
' �T�v�@�@�@�F�e�[�u���V�[�g����s�T�C�Y�̎擾���s��
'
' �����@�@�@�FtableSheet �e�[�u���V�[�g
' �߂�l�@�@�F�s�T�C�Y
'
' =========================================================
Public Function getRecordSize(ByRef tableSheet As ValTableWorksheet) As Long


    ' �߂�l�ϐ�
    Dim ret As Long
    
    If tableSheet.recFormat = REC_FORMAT.recFormatToUnder Then
    
        ret = ExcelUtil.getCellEndOfLastInputtedRow(sheet, ConstantsTable.U_RECORD_NUM_COL)
        ret = ret - ConstantsTable.U_RECORD_OFFSET_ROW + 1
    
    Else
    
        ret = ExcelUtil.getCellEndOfLastInputtedCol(sheet, ConstantsTable.R_RECORD_NUM_ROW)
        ret = ret - ConstantsTable.R_RECORD_OFFSET_COL + 1
    End If
    
    ' �߂�l��ݒ肷��
    getRecordSize = ret

End Function

Public Function convertColumnType(ByVal columnType As String) As String

    ' ���K�\���I�u�W�F�N�g�𐶐�����
    Static reg As Object
    
    If reg Is Nothing Then
    
        ' IE5.0�ȏ�K�{
        Set reg = CreateObject("VBScript.RegExp")
        
        If reg Is Nothing Then
        
            err.Raise ConstantsError.ERR_NUMBER_REG_EXP_NOT_CREATED _
                    , "" _
                    , ConstantsError.ERR_DESC_REG_EXP_NOT_CREATED
        
        End If
    
        With reg
            ' �p�^�[��������
            .Pattern = "\(.*?\)"
            ' �啶���Ə���������ʂ��Ȃ�
            .IgnoreCase = True
            ' ������S�̂�����
            .Global = True
            
        End With
    
    End If
    
    convertColumnType = reg.replace(columnType, "")
    
End Function