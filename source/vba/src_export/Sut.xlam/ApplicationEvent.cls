VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �A�v���P�[�V�����p�X���ꎞ�I�ɕύX����N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2013/07/16�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �A�v���P�[�V�����I�u�W�F�N�g
Public WithEvents myEvent As Application
Attribute myEvent.VB_VarHelpID = -1

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
' �����������\�b�h
' =========================================================
Public Sub init(ByRef app As Application)

    Set myEvent = app
    
End Sub

' =========================================================
' ���u�b�N�̐V�K�쐬��
'
' �T�v�@�@�@�F
'
' �����@�@�@�FWb ���[�N�u�b�N
' �߂�l�@�@�F����
' =========================================================
Private Sub myEvent_NewWorkbook(ByVal wb As Workbook)

    Debug.Print wb.name
End Sub

' =========================================================
' ���u�b�N�̊J���C�x���g
'
' �T�v�@�@�@�F
'
' �����@�@�@�FWb ���[�N�u�b�N
' �߂�l�@�@�F����
' =========================================================
Private Sub myEvent_WorkbookOpen(ByVal wb As Workbook)

    Debug.Print wb.name
End Sub

' =========================================================
' ���u�b�N�̃N���[�Y���C�x���g
'
' �T�v�@�@�@�F
'
' �����@�@�@�FWb ���[�N�u�b�N
' �߂�l�@�@�F����
' =========================================================
Private Sub myEvent_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)

    Debug.Print wb.name
    
    If Not SutWorkbook Is wb Then
    
        Main.SutRelease
    End If
    
End Sub
