VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectConditionCreator 
   Caption         =   "SELECT"
   ClientHeight    =   8805.001
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7935
   OleObjectBlob   =   "frmSelectConditionCreator.frx":0000
End
Attribute VB_Name = "frmSelectConditionCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' SELECT���������t�H�[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2009/04/03�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����������������ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�Fsql SELECT SQL
'
' =========================================================
Public Event ok(ByVal sql As String, ByVal append As Boolean)

' =========================================================
' ���������L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' ---------------------------------------------------------
' INI�t�@�C���֘A�萔
' ---------------------------------------------------------
' ���Ō�ɑ��삳�ꂽ���
Private Const REG_SUB_KEY_SELECT_CONDITION As String = "select_condition"

' �ȈՐݒ�y�[�W
Private Const PAGE_SIMPLE_SETTING As Long = 0
' �ڍאݒ�y�[�W
Private Const PAGE_DETAIL_SETTING As Long = 1

' �����w�萔�̍ŏ��l
Private Const COLUMN_COND_MIN As Long = 1
' �����w�萔�̍ő�l
Private Const COLUMN_COND_MAX As Long = 10

' �����l ����
Private Const ORDER_BY_VALUE_ASC  As Variant = True
' �����l �~��
Private Const ORDER_BY_VALUE_DESC As Variant = False
' �����l �w��Ȃ�
Private Const ORDER_BY_VALUE_NONE As Variant = Null

' �A�v���P�[�V�����ݒ�
Private applicationSetting As ValApplicationSetting
' �A�v���P�[�V�����ݒ�i�J�����������j
Private applicationSettingColFmt As ValApplicationSettingColFormat

' DB�R�l�N�V����
Private dbConn As Object
' DBMS���
Private dbms   As DbmsType
' �e�[�u����`
Private tableSheet As ValTableWorksheet

' ���������@�z�񃓃g���[���@�J����
Private columnCondList()   As CntListBox
' ���������@�z�񃓃g���[���@�l
Private valueCondList()    As control
' ���������@�z�񃓃g���[���@����
Private orderCondList()    As control

' SQL�ҏW�t���O
Private editedSql As Boolean

' =========================================================
' ���t�H�[���\���i�g�������j
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants _
                 , ByRef apps As ValApplicationSetting _
                 , ByRef appsColFmt As ValApplicationSettingColFormat _
                 , ByRef conn As Object)

    ' �G���[���b�Z�[�W���N���A����
    lblErrorMessage.Caption = ""

    ' �A�v���P�[�V�����ݒ����ݒ�
    Set applicationSetting = apps
    Set applicationSettingColFmt = appsColFmt
    ' DB�R�l�N�V������ݒ�
    Set dbConn = conn
    ' DBMS��ނ��擾����
    dbms = ADOUtil.getDBMSType(dbConn)
    
    ' �A�N�e�B�u���̏���
    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal

End Sub

' =========================================================
' ���t�H�[����\���i�g�������j
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub HideExt()

    ' �f�B�A�N�e�B�u���̏���
    deactivate
    
    Main.storeFormPosition Me.name, Me
    Me.Hide

End Sub

' =========================================================
' ���e�X�g�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdTest_Click()

    On Error GoTo err

    SutWhite.showHourglassWindowOnCenterPt Me
    
    Dim resultSet   As Object
    Dim resultCount As Long

    Set resultSet = ADOUtil.querySelect(dbConn, txtSqlEditor.value, resultCount, adOpenStatic)
    resultCount = resultSet.recordCount
    
    ADOUtil.closeRecordSet resultSet

    lblResultCount.Caption = resultCount & "��"

    SutWhite.closeHourglassWindow
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

    SutWhite.closeHourglassWindow
End Sub

' =========================================================
' ���t�H�[�����������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
    ' ���������������s����
    initial

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���t�H�[���j�����̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Terminate()

    On Error GoTo err
    
    ' �j�����������s����
    unInitial
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���t�H�[���A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Activate()

    ' �őO�ʕ\���ɂ���
    ExcelUtil.setUserFormTopMost Me

End Sub

' =========================================================
' ������������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub initial()

End Sub

' =========================================================
' ����n������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub unInitial()

End Sub

' =========================================================
' ���A�N�e�B�u���̏���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    ' �e�[�u���V�[�g�Ǎ��I�u�W�F�N�g
    Dim tableSheetReader As ExeTableSheetReader
    Set tableSheetReader = New ExeTableSheetReader
    Set tableSheetReader.sheet = ActiveSheet
    Set tableSheetReader.conn = dbConn
            
    ' �e�[�u����`��ǂݍ���
    Set tableSheet = tableSheetReader.readTableInfo

    Dim table As ValDbDefineTable
    Set table = tableSheet.table

    Dim i As Long
    
    ' -----------------------------------------------
    ' �J������
    ' -----------------------------------------------
    ' �R���g���[���z����������
    Erase columnCondList
    ' �R���g���[���z����m�ۂ���
    ReDim columnCondList(COLUMN_COND_MIN To COLUMN_COND_MAX)
    
    i = COLUMN_COND_MIN
    
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond1: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond2: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond3: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond4: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond5: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond6: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond7: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond8: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond9: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
    Set columnCondList(i) = New CntListBox: columnCondList(i).control = cboColumnCond10: columnCondList(i).addAllByGenericCollection table.columnList, "ColumnName": i = i + 1
        
    ' -----------------------------------------------
    ' �l
    ' -----------------------------------------------
    ' �R���g���[���z����������
    Erase valueCondList
    ' �R���g���[���z����m�ۂ���
    ReDim valueCondList(COLUMN_COND_MIN To COLUMN_COND_MAX)
    
    i = COLUMN_COND_MIN
        
    ' TextBox�̃I�u�W�F�N�g�����̂܂ܑ�����悤�Ƃ���Ɖ��̂�String�^�ɕϊ������̂�
    ' Controls���X�g����ԐړI�Ɏ擾���đ������
    Set valueCondList(i) = Controls.item(txtCondValue1.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue2.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue3.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue4.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue5.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue6.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue7.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue8.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue9.name): i = i + 1
    Set valueCondList(i) = Controls.item(txtCondValue10.name): i = i + 1
    
    ' -----------------------------------------------
    ' ����
    ' -----------------------------------------------
    ' �R���g���[���z����������
    Erase orderCondList
    ' �R���g���[���z����m�ۂ���
    ReDim orderCondList(COLUMN_COND_MIN To COLUMN_COND_MAX)
    
    i = COLUMN_COND_MIN
        
    Set orderCondList(i) = tglOrderCond1: i = i + 1
    Set orderCondList(i) = tglOrderCond2: i = i + 1
    Set orderCondList(i) = tglOrderCond3: i = i + 1
    Set orderCondList(i) = tglOrderCond4: i = i + 1
    Set orderCondList(i) = tglOrderCond5: i = i + 1
    Set orderCondList(i) = tglOrderCond6: i = i + 1
    Set orderCondList(i) = tglOrderCond7: i = i + 1
    Set orderCondList(i) = tglOrderCond8: i = i + 1
    Set orderCondList(i) = tglOrderCond9: i = i + 1
    Set orderCondList(i) = tglOrderCond10: i = i + 1
    
    
    ' �t�@�C������e�R���g���[���̏���ǂݍ���
    restoreSelectCondition
    
    ' �y�[�W��؂�ւ�����
    ' SQL�G�f�B�^�Ƀe�L�X�g���ݒ肳��Ă��Ȃ��ꍇ
    If txtSqlEditor.value = "" Then
    
        ' �ȈՃy�[�W��
        mpageCondition.value = PAGE_SIMPLE_SETTING
        
    ' SQL�G�f�B�^�Ƀe�L�X�g���ݒ肳��Ă���ꍇ
    Else
    
        ' �ڍ׃y�[�W��
        mpageCondition.value = PAGE_DETAIL_SETTING
        
        ' �ҏW�t���O��true�ɐݒ肵�Ă���
        editedSql = True
    End If
    
    ' ���ʌ����\�����x��������������
    lblResultCount.Caption = ""

    
End Sub

' =========================================================
' ���m���A�N�e�B�u���̏���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub deactivate()

End Sub

' =========================================================
' �������w��g�O���{�^���ύX���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub tglOrderCond1_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond1
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond2_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond2
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond3_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond3
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond4_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond4
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond5_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond5
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond6_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond6
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond7_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond7
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond8_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond8
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond9_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond9
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub
Private Sub tglOrderCond10_Change()

    On Error GoTo err:

    changeLabelOrderControl tglOrderCond10
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �������w��g�O���{�^���̃��x���ύX
'
' �T�v�@�@�@�F�����w��g�O���{�^���̏�Ԃɉ����ă��x����ύX���邽�߂̏���
' �����@�@�@�Fcontrol �g�O���{�^��
' �߂�l�@�@�F
'
' =========================================================
Private Sub changeLabelOrderControl(ByRef control As ToggleButton)

    If control.value = ORDER_BY_VALUE_ASC Then
    
        control.Caption = "����"
    
    ElseIf control.value = ORDER_BY_VALUE_DESC Then
    
        control.Caption = "�~��"
    Else
    
        control.Caption = "�Ȃ�"
    End If

End Sub

' =========================================================
' ��PK�ŏ����w��N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdPkCondition_Click()

    On Error GoTo err:

    ' ��x���Z�b�g����
    resetWhereOrderby

    Dim i As Long: i = COLUMN_COND_MIN
    
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    ' �J����
    Dim column     As ValDbDefineColumn
    ' �J�������X�g
    Dim columnList As ValCollection
    
    ' �e�[�u��������(PK)
    Dim tableConstPk    As New ValDbDefineTableConstraints
    ' PK�J�����ł��邩������킷�t���O
    Dim isColumnPk      As Boolean
    
    Dim tableConstTmp   As ValDbDefineTableConstraints
    ' �e�[�u�����񃊃X�g����PK������擾����
    For Each tableConstTmp In table.constraintsList
    
        If tableConstTmp.constraintType = sutredlib.tableConstPk Then
        
            Set tableConstPk = tableConstTmp
            Exit For
        End If
    Next
    
    ' �J�������X�g���擾����
    Set columnList = table.columnList
    
    ' �J�������X�g��1������������
    For Each column In columnList.col
            
        ' PK����ł��邩�ǂ����𔻒肷��
        If tableConstPk.columnList.getItem(column.columnName) Is Nothing Then
        
            isColumnPk = False
        Else
        
            isColumnPk = True
        End If
        
        ' �J������PK�ł���ꍇ
        If isColumnPk = True Then
        
            ' PK�̐����R���g���[���̐��������Ă���̂Ń��b�Z�[�W��\�����ďI������
            If i > COLUMN_COND_MAX Then
            
                err.Raise ConstantsError.ERR_NUMBER_OVER_SELECT_COND_CONTROL _
                        , _
                        , ConstantsError.ERR_DESC_OVER_SELECT_COND_CONTROL
                Exit Sub
            End If
            
            ' �J��������PK�񖼂�ݒ肷��
            columnCondList(i).control.value = column.columnName
            ' �����ɏ�����ݒ肷��
            orderCondList(i).value = ORDER_BY_VALUE_ASC
            i = i + 1
        End If
        
                
    Next

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �����R�[�h�擾�͈́@�J�n�e�L�X�g�{�b�N�X�̃`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtRecRangeStart_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' ��̏ꍇ�A����
    If txtRecRangeStart.text = "" Then
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecRangeStart

    ' �e�L�X�g�{�b�N�X�̒l�����������`�F�b�N����
    ElseIf validInteger(txtRecRangeStart.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        ' �A���[�g��\������
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
        
        changeControlPropertyByValidFalse txtRecRangeStart
    
    ' ���l�͈̓`�F�b�N
    ElseIf CDbl(txtRecRangeStart.text) < 1 Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse txtRecRangeStart
    
    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecRangeStart
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �����R�[�h�擾�͈́@�I���e�L�X�g�{�b�N�X�̃`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtRecRangeEnd_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err:

    ' ��̏ꍇ�A����
    If txtRecRangeEnd.text = "" Then
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecRangeEnd

    ' �e�L�X�g�{�b�N�X�̒l�����������`�F�b�N����
    ElseIf validInteger(txtRecRangeEnd.text) = False Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        ' �A���[�g��\������
        lblErrorMessage.Caption = ConstantsError.VALID_ERR_INTEGER
    
        changeControlPropertyByValidFalse txtRecRangeEnd
        
    ' ���l�͈̓`�F�b�N
    ElseIf CDbl(txtRecRangeEnd.text) < 1 Then
    
        ' �X�V���L�����Z������
        Cancel = True
    
        lblErrorMessage.Caption = replace(ConstantsError.VALID_ERR_AND_OVER, "{1}", 1)
        
        ' �R���g���[���̃v���p�e�B��ύX����
        changeControlPropertyByValidFalse txtRecRangeEnd

    ' ����ȏꍇ
    Else
    
        lblErrorMessage.Caption = ""
        
        changeControlPropertyByValidTrue txtRecRangeEnd
    End If
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���e�L�X�g�{�b�N�X�`�F�b�N�������̃R���g���[���ύX����
'
' �T�v�@�@�@�F
' �����@�@�@�Fcnt �R���g���[��
' �߂�l�@�@�F
'
' =========================================================
Public Sub changeControlPropertyByValidTrue(ByRef cnt As MSForms.control)

    With cnt
        .BackColor = &H80000005
        .ForeColor = &H80000012
    
    End With

End Sub

' =========================================================
' ���e�L�X�g�{�b�N�X�`�F�b�N���s���̃R���g���[���ύX����
'
' �T�v�@�@�@�F
' �����@�@�@�Fcnt �R���g���[��
' �߂�l�@�@�F
'
' =========================================================
Public Sub changeControlPropertyByValidFalse(ByRef cnt As MSForms.control)

    With cnt
        ' �e�L�X�g�S�̂�I������
        .SelStart = 0
        .SelLength = Len(.text)
        
        .BackColor = RGB(&HFF, &HFF, &HCC)
        .ForeColor = reverseRGB(&HFF, &HFF, &HCC)
        
    End With

End Sub

' =========================================================
' ���擪100�����擾����{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdLimit100_Click()

    On Error GoTo err:

    ' ���R�[�h�͈� �J�n��ݒ肷��
    txtRecRangeStart.value = 1
    ' ���R�[�h�͈� �I����ݒ肷��
    txtRecRangeEnd.value = 100
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �����փ{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F�y�[�W��؂�ւ���
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdNext_Click()

    On Error GoTo err:

    ' SQL�𐶐����ASQL�ҏW�e�L�X�g�{�b�N�X�ɓ��e��\��
    ' �y�[�W��؂�ւ���O�ɕύX���s��
    txtSqlEditor.value = createSql

    ' �y�[�W��؂�ւ���
    mpageCondition.value = PAGE_DETAIL_SETTING
    
    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���߂�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F�y�[�W��؂�ւ���
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdReturn_Click()

    On Error GoTo err:

    ' SQL�ҏW�t���O�̊m�F
    ' ���e���ҏW����Ă���ꍇ
    If editedSql = True Then
    
        ' ���b�Z�[�W�{�b�N�X�̖߂�l
        Dim ret As Long
    
        ' �ҏW��ɖ߂�ꍇ�́A�x����\������
        ret = VBUtil.showMessageBoxForYesNo("�߂�ƕҏW���e�������Ă��܂��܂����A��낵���ł����H", ConstantsCommon.APPLICATION_NAME)
        
        ' �͂���I�������ꍇ
        If ret = WinAPI_User.IDYES Then
        
            ' �y�[�W��؂�ւ���
            mpageCondition.value = PAGE_SIMPLE_SETTING
            txtSqlEditor.value = ""
            editedSql = False
        End If
        
    ' ���e���ҏW����Ă��Ȃ��ꍇ
    Else
    
        ' �y�[�W��؂�ւ���
        mpageCondition.value = PAGE_SIMPLE_SETTING
        txtSqlEditor.value = ""
    
    End If

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �����Z�b�g�N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdReset_Click()

    On Error GoTo err:

    ' �����E���я������Z�b�g
    resetWhereOrderby
    ' ���R�[�h�͈͎w������Z�b�g
    resetRecRange

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �������E���я��̐ݒ�����Z�b�g
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub resetWhereOrderby()

    Dim i As Long
    
    ' �R���g���[���z���1������������
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        ' �J����������ɐݒ�
        columnCondList(i).control.value = ""
        ' �l����ɐݒ�
        valueCondList(i).value = ""
        ' �������Ȃ��ɐݒ�
        orderCondList(i).value = ORDER_BY_VALUE_NONE
    Next
    
End Sub

' =========================================================
' �����R�[�h�擾�͈͎̔w������Z�b�g
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub resetRecRange()

    txtRecRangeStart.value = ""
    txtRecRangeEnd.value = ""
    
End Sub

' =========================================================
' ��SQL�G�f�B�^ �ύX�C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub txtSqlEditor_Change()

    ' �ڍ׃y�[�W�ŁAChange�C�x���g�����������ꍇ�A�ҏW�t���O��ON�ɂ���
    If mpageCondition.value = PAGE_DETAIL_SETTING Then
    
        editedSql = True
    End If
End Sub


' =========================================================
' ��OK�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err:
    
    ' SQL
    Dim sql As String
    
    ' �y�[�W���ȈՐݒ�̏ꍇ
    If mpageCondition.value = PAGE_SIMPLE_SETTING Then
    
        ' SQL���R���g���[�����琶������
        sql = createSql
    
    ' �y�[�W���ڍאݒ�̏ꍇ
    Else
    
        ' SQL���G�f�B�^����擾����
        sql = txtSqlEditor.value
    End If
    
    ' SELECT������ۑ�����
    storeSelectCondition
    
    Me.HideExt
    
    RaiseEvent ok(sql, cbxTableSheetAppend.value)

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���L�����Z���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdCancel_Click()

    On Error GoTo err:
    
    Me.HideExt
    RaiseEvent Cancel

    Exit Sub

err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ��SQL�𐶐�����
'
' �T�v�@�@�@�FSQL�𐶐�����B
' �����@�@�@�F
' �߂�l�@�@�FSELECT�N�G���[
'
' =========================================================
Private Function createSql() As String

    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    ' SELECT����
    Dim condition As ValSelectCondition
    ' �t�H�[����������𐶐�����
    Set condition = createCondition

    Dim dbObjFactory As New DbObjectFactory
    Dim queryCreator        As IDbQueryCreator

    Set queryCreator = dbObjFactory.createQueryCreator(dbConn _
                                                            , applicationSetting.emptyCellReading _
                                                            , applicationSetting.getDirectInputChar _
                                                            , applicationSettingColFmt.getDbColFormatListByDbConn(dbConn) _
                                                            , applicationSetting.schemaUse _
                                                            , applicationSetting.getTableColumnEscapeByDbConn(dbConn))

    ' SELECT���𐶐�����
    createSql = queryCreator.createSelect(table, condition)
    
End Function

' =========================================================
' ����������
'
' �T�v�@�@�@�F�R���g���[����������𐶐�����B
' �����@�@�@�F
' �߂�l�@�@�FSELECT�����I�u�W�F�N�g
'
' =========================================================
Private Function createCondition() As ValSelectCondition

    ' �߂�l
    Dim ret As ValSelectCondition
    ' �߂�l������������
    Set ret = New ValSelectCondition

    ' �J������
    Dim columnName  As String
    ' �l
    Dim value       As String
    ' ����
    Dim orderby     As Variant
    ' ���� (Long�^)
    Dim orderByLong As Long
    
    Dim i As Long
    
    ' �R���g���[���z���1������������
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        ' �J���������擾
        columnName = columnCondList(i).control.value
        ' �l���擾
        value = valueCondList(i).value
        ' �������擾
        orderby = orderCondList(i).value
    
        ' �J���������ݒ肳��Ă���ꍇ�̂݁A�����Ƃ��Đݒ肷��
        If columnName <> "" Then
        
            ' �R���g���[���̒l�� ValSelectCondition �̒萔�ɕϊ�����
            ' ����
            If orderby = ORDER_BY_VALUE_ASC Then
            
                orderByLong = ret.ORDER_ASC
                
            ' �~��
            ElseIf orderby = ORDER_BY_VALUE_DESC Then
            
                orderByLong = ret.ORDER_DESC
                
            ' ����
            Else
            
                orderByLong = ret.ORDER_NONE
            End If
            
            ' ������ݒ肷��
            ret.setCondition columnName, value, orderByLong
            
        End If
        
    Next
    
    ' ���R�[�h�͈͎w�� �J�n���ݒ肳��Ă���ꍇ�A�����Ƃ��Đݒ�
    If txtRecRangeStart.value <> "" Then
    
        ret.recRangeStart = txtRecRangeStart.value
    End If
    
    ' ���R�[�h�͈͎w�� �I�����ݒ肳��Ă���ꍇ�A�����Ƃ��Đݒ�
    If txtRecRangeEnd.value <> "" Then
    
        ret.recRangeEnd = txtRecRangeEnd.value
    End If

    ' �߂�l�ɐݒ�
    Set createCondition = ret

End Function

' =========================================================
' ��SELECT������ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeSelectCondition()

    On Error GoTo err
    
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    ' SELECT�������i�[���邽�߂̔z��ϐ�
    ' �����w�萔�~(�J�����E�l�E���я�)�{���R�[�h�͈͎w��i�J�n�E�I���j�{SQL�G�f�B�^ �� 10�~3�{2�{1
    Dim selectCondition(COLUMN_COND_MIN _
                    To (COLUMN_COND_MAX * 3 + 2 + 1), 0 To 1) As Variant
    
    
    Dim i As Long
    Dim j As Long
    
    j = COLUMN_COND_MIN
    
    ' �R���g���[���z���1������������
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        selectCondition(j, 0) = columnCondList(i).control.name
        selectCondition(j, 1) = columnCondList(i).control.value: j = j + 1
        selectCondition(j, 0) = valueCondList(i).name
        selectCondition(j, 1) = valueCondList(i).value: j = j + 1
        selectCondition(j, 0) = orderCondList(i).name
        ' �����R���g���[���i�g�O���{�^���j�͖��I���̏ꍇ��NULL��Ԃ��̂ŋ󕶎���ɕϊ�����
        selectCondition(j, 1) = VBUtil.convertNullToEmptyStr(orderCondList(i).value): j = j + 1
    
    Next

    selectCondition(j, 0) = txtRecRangeStart.name
    selectCondition(j, 1) = txtRecRangeStart.value: j = j + 1

    selectCondition(j, 0) = txtRecRangeEnd.name
    selectCondition(j, 1) = txtRecRangeEnd.value: j = j + 1
    
    selectCondition(j, 0) = txtSqlEditor.name
    selectCondition(j, 1) = txtSqlEditor.value: j = j + 1
    
    ' ���W�X�g������N���X
    Dim registry As New RegistryManipulator
    ' ���W�X�g������N���X������������
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SELECT_CONDITION & "\" & table.schemaTableName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True

    ' ���W�X�g���ɏ���ݒ肷��
    registry.setValues selectCondition
    
    Set registry = Nothing
    
    ' ----------------------------------------------
    ' �u�b�N�ݒ���
    Dim bookProp As New BookProperties
    bookProp.sheet = tableSheet.sheet
    bookProp.removeValueByPrefixName ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_"
    
    ' �R���g���[���z���1������������
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & columnCondList(i).control.name, columnCondList(i).control.value
        bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & valueCondList(i).name, valueCondList(i).value
        bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & orderCondList(i).name, VBUtil.convertNullToEmptyStr(orderCondList(i).value)
    
    Next
    
    bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & txtRecRangeStart.name, txtRecRangeStart.value
    bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & txtRecRangeEnd.name, txtRecRangeEnd.value
    bookProp.setValue ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName & "_" & txtSqlEditor.name, txtSqlEditor.value
    ' ----------------------------------------------
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ��SELECT������ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreSelectCondition()

    On Error GoTo err
    
    Dim i As Long

    ' ----------------------------------------------
    ' �u�b�N�ݒ���
    Dim bookProp As New BookProperties
    bookProp.sheet = tableSheet.sheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' �ݒ���V�[�g�����݂���
        
        Set bookPropVal = bookProp.getValuesByPrefixName(ConstantsBookProperties.TABLE_SELECT_CONDITION_CREATOR_DIALOG, tableSheet.sheetName)
        If bookPropVal.count > 0 Then
        
            ' �ݒ��񂪑��݂���̂ŁA�t�H�[���ɔ��f����
            ' �R���g���[���z���1������������
            For i = COLUMN_COND_MIN To COLUMN_COND_MAX
            
                columnCondList(i).control.value = bookPropVal.getItem(tableSheet.sheetName & "_" & columnCondList(i).control.name, vbString)
                valueCondList(i).value = bookPropVal.getItem(tableSheet.sheetName & "_" & valueCondList(i).name, vbString)
                orderCondList(i).value = bookPropVal.getItem(tableSheet.sheetName & "_" & orderCondList(i).name, vbString)
                
            Next
        
            txtRecRangeStart.value = bookPropVal.getItem(tableSheet.sheetName & "_" & txtRecRangeStart.name, vbString)
            txtRecRangeEnd.value = bookPropVal.getItem(tableSheet.sheetName & "_" & txtRecRangeEnd.name, vbString)
            
            txtSqlEditor.value = bookPropVal.getItem(tableSheet.sheetName & "_" & txtSqlEditor.name, vbString)

            Exit Sub
        End If
    End If
    ' ----------------------------------------------
    
    Dim table As ValDbDefineTable
    Set table = tableSheet.table
    ' ���W�X�g������N���X
    Dim registry As New RegistryManipulator
    ' ���W�X�g������N���X������������
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_SELECT_CONDITION & "\" & table.schemaTableName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
        
    Dim retColumn As String
    Dim retValue  As String
    Dim retOrder  As String
    
    Dim retRecRangeStart As String
    Dim retRecRangeEnd   As String
    
    Dim retSqlEdit As String
    
    ' �R���g���[���z���1������������
    For i = COLUMN_COND_MIN To COLUMN_COND_MAX
    
        retColumn = ""
        retValue = ""
        retOrder = ""
        
        registry.getValue columnCondList(i).control.name, retColumn
        registry.getValue valueCondList(i).name, retValue
        registry.getValue orderCondList(i).name, retOrder
    
        columnCondList(i).control.value = retColumn
        valueCondList(i).value = retValue
        orderCondList(i).value = retOrder
        
    Next

    registry.getValue txtRecRangeStart.name, retRecRangeStart
    registry.getValue txtRecRangeEnd.name, retRecRangeEnd
    
    txtRecRangeStart.value = retRecRangeStart
    txtRecRangeEnd.value = retRecRangeEnd
    
    registry.getValue txtSqlEditor.name, retSqlEdit
    
    txtSqlEditor.value = retSqlEdit
    
    Set registry = Nothing
    
    Exit Sub
    
err:
    
    Set registry = Nothing
    
    Main.ShowErrorMessage

End Sub

