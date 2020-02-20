VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelCursorWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' Excel�}�N���ŃJ�[�\���̂ݑҋ@��Ԃɂ��邽�߂̃N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/12�@�V�K�쐬

' �g�p���@�F
' �@�{�N���X�̃I�u�W�F�N�g�� New ����B
' �@�֐��𔲂���Ǝ����� Class_Terminate �Ō�n�����s����B
'
'   Dim var As New ExcelCursorWait: var.init
'
' *********************************************************

' �j������p������t���O
Private isContinueWhenDestroyed As Boolean
' �����ԏ����N���X
Private excelLongTimeProcess As ExcelLongTimeProcessing

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    On Error Resume Next
    
    Set excelLongTimeProcess = New ExcelLongTimeProcessing
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set excelLongTimeProcess = Nothing
    
End Sub

' =========================================================
' ������������
'
' �T�v�@�@�@�FisContinueWhenDestroyed_ �j������p������t���O
'
' =========================================================
Public Sub init(Optional ByVal isContinueWhenDestroyed_ As Boolean = False)
              
    ' �j������p������t���O
    isContinueWhenDestroyed = isContinueWhenDestroyed_
    
    If isContinueWhenDestroyed = True Then
        ' �j������p������̂ŋ����I�ɃJ�[�\����ҋ@��Ԃɂ���
        Application.cursor = xlWait
        
        Exit Sub
    End If
              
    If Application.cursor = xlWait Then
        ' ���ɃJ�[�\�����ҋ@��ԂȂ̂ŁA�����L���ɂ��Ȃ�
    Else
        ' �J�[�\���̂ݗL���ɂ���
        excelLongTimeProcess.init False, False, False, True, False, False, False
    End If
              

End Sub

' =========================================================
' ���j������
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub destroy()

    excelLongTimeProcess.destroy

End Sub

' =========================================================
' �������I�ɃJ�[�\�������̏�Ԃɖ߂�����
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub forceRestore()

    Application.cursor = xlDefault

End Sub
