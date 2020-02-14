VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSnapshot 
   Caption         =   "�X�i�b�v�V���b�g�擾"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8385
   OleObjectBlob   =   "frmSnapShot.frx":0000
End
Attribute VB_Name = "frmSnapShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�i�b�v�V���b�g�擾�t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/09/06�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' ���X�i�b�v�V���b�g�擾���s�C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�Fsheet ���[�N�V�[�g
'
' =========================================================
Public Event execSnapshot(ByRef sheet As Worksheet)

' =========================================================
' ���L�����Z���C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' =========================================================
' ��DB�ύX�C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event changeDb()

' =========================================================
' ��SQL�ύX�C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event changeSql(ByRef sheet As Worksheet)

' =========================================================
' ���X�i�b�v�V���b�g�N���A�C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event clearSnapshot(ByRef sheet As Worksheet)

' =========================================================
' ���X�i�b�v�V���b�g��r�\���C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event showSnapshotDiff(ByRef sheet As Worksheet)

' ---------------------------------------------------------
' ���W�X�g���t�@�C���L�[
' ---------------------------------------------------------
Private Const REG_SUB_KEY_SNAPSHOT As String = "snapshot"

' �A�v���P�[�V�����ݒ���
Private applicationSetting As ValApplicationSetting

' DB�R�l�N�V�����I�u�W�F�N�g
Private dbConn As Object
' DB�ڑ�������
Private dbConnStr As String

' ���sSQL���X�g
Private executeSqltList  As CntListBox
' �X�i�b�v�V���b�g���X�g
Private snapShotList     As CntListBox

' �Ώۃu�b�N
Private targetBook As Workbook
' �Ώۃu�b�N���擾����
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal    ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@aps      �A�v���P�[�V�����ݒ���
' �@�@�@�@�@�@conn     DB�R�l�N�V����
' �@�@�@�@�@�@connStr  DB�ڑ�������
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef aps As ValApplicationSetting, ByRef conn As Object, ByVal connStr As String)

    ' �A�v���P�[�V��������ݒ肷��
    Set applicationSetting = aps
    ' DB�R�l�N�V������ݒ肷��
    Set dbConn = conn
    dbConnStr = connStr
    ' �A�N�e�B�u����
    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal
    
End Sub

' =========================================================
' ���t�H�[����\��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub HideExt()

    Main.storeFormPosition Me.name, Me
    Me.Hide
    
    ' ��A�N�e�B�u����
    deactivate
    
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
    
    ' ���[�h���_�̃A�N�e�B�u�u�b�N��ێ����Ă���
    Set targetBook = ExcelUtil.getActiveWorkbook
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
' ���t�H�[������C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Main.storeFormPosition Me.name, Me

End Sub

' =========================================================
' ��DB�ύX����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDBConnectedChange_Click()

    On Error GoTo err
    
    Me.Hide
    
    RaiseEvent changeDb
    
    Me.Show vbModeless
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ��SQL�X�V����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdExecuteSqlUpdate_Click()

    On Error GoTo err
    
    Dim book  As Workbook
    Dim sheet As Worksheet
    
    Dim ExeSnapSqlDefineSheetReader As ExeSnapSqlDefineSheetReader
    
    ' ���X�g�I�u�W�F�N�g������������
    executeSqltList.removeAll
    executeSqltList.init cboExecuteSql
    
    ' �S�u�b�N��Ώۂɂ���
    For Each book In Workbooks
    
        ' �S�V�[�g��Ώۂɂ���
        For Each sheet In book.Sheets
        
            Set ExeSnapSqlDefineSheetReader = New ExeSnapSqlDefineSheetReader
            Set ExeSnapSqlDefineSheetReader.sheet = sheet
                    
            If ExeSnapSqlDefineSheetReader.isSqlDefineSheet = True Then
                ' SQL��`�V�[�g�̏ꍇ�A���X�g�ɒǉ�
                executeSqltList.addItem book.name & " " & sheet.name, sheet
            
            End If
        
        Next
        
    Next
    
    ' ���sSQ�I���R���{�{�b�N�X�ɒǉ����ꂽ���̂������
    ' �擪���f�t�H���g�I������
    If cboExecuteSql.ListCount >= 1 Then
        cboExecuteSql.ListIndex = 0
    End If
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���t�H�[����\��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cboExecuteSql_Change()

    On Error GoTo err
    
    ' ���sSQ�I���R���{�{�b�N�X�����I���̏ꍇ
    If cboExecuteSql.ListIndex = -1 Then
    
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = executeSqltList.getItem(cboExecuteSql.ListIndex)
    
    RaiseEvent changeSql(sheet)
    
    sheet.activate
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�ꗗ�N���A����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdSnapshotClear_Click()

    On Error GoTo err
    
    ' ���sSQ�I���R���{�{�b�N�X�����I���̏ꍇ
    If cboExecuteSql.ListIndex = -1 Then
    
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = executeSqltList.getItem(cboExecuteSql.ListIndex)
    
    RaiseEvent clearSnapshot(sheet)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���X�i�b�v�V���b�g��r�t�H�[���\������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdSnapshotDiff_Click()

    On Error GoTo err
    
    ' ���sSQ�I���R���{�{�b�N�X�����I���̏ꍇ
    If cboExecuteSql.ListIndex = -1 Then
    
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = executeSqltList.getItem(cboExecuteSql.ListIndex)
    
    Me.Hide
    
    RaiseEvent showSnapshotDiff(sheet)
    
    Me.Show vbModeless
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���X�i�b�v�V���b�g�擾����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdSnapshotGet_Click()

    On Error GoTo err
    
    ' ���sSQ�I���R���{�{�b�N�X�����I���̏ꍇ
    If cboExecuteSql.ListIndex = -1 Then
    
        Exit Sub
    End If
    
    Dim sheet As Worksheet
    Set sheet = executeSqltList.getItem(cboExecuteSql.ListIndex)
    
    RaiseEvent execSnapshot(sheet)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' �����鏈��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdClose_Click()

    On Error GoTo err
    
    ' �t�H�[�������
    HideExt
    
    ' �L�����Z���C�x���g�𑗐M����
    RaiseEvent Cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
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

    Set executeSqltList = New CntListBox
    Set snapShotList = New CntListBox
    snapShotList.init lstSnapshot

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

    Set executeSqltList = Nothing
    Set snapShotList = Nothing
End Sub

' =========================================================
' ���A�N�e�B�u����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    cmdExecuteSqlUpdate_Click
    
    txtDBConnected.text = dbConnStr
End Sub

' =========================================================
' ����A�N�e�B�u����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub deactivate()

End Sub

' =========================================================
' ��DB�R�l�N�V�����̍X�V
'
' �T�v�@�@�@�Fconn    DB�R�l�N�V����
'             connStr DB�ڑ�������
'
' =========================================================
Public Sub updateDbConn(ByRef conn As Object, ByVal connStr As String)

    On Error GoTo err
    
    Set dbConn = conn
    dbConnStr = connStr

    txtDBConnected.text = dbConnStr

    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���X�i�b�v�V���b�g�̍폜
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub clearSnapshot()

    On Error GoTo err
    
    snapShotList.removeAll
    snapShotList.init lstSnapshot
    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub



' =========================================================
' ���X�i�b�v�V���b�g�̒ǉ�
'
' �T�v�@�@�@�Flabel ���x��
'             value �l
'
' =========================================================
Public Sub addSnapshot(ByRef label As String, ByRef value As String)

    On Error GoTo err
    
    snapShotList.addItem label, value
    Exit Sub
    
err:
    
    Main.ShowErrorMessage
    
End Sub


