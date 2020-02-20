VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationCurDirChanger"
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

' �A�v���P�[�V�����p�X
Private applicationPath As String

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    ' �J�����g�p�X���ꎞ�ޔ���Ԃɂ���
    applicationPath = CurDir$
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' �p�X�����ɖ߂�
    ChDrive applicationPath
    ChDir applicationPath
    
End Sub

' =========================================================
' �����������\�b�h
' =========================================================
Public Sub init(Optional ByVal path As String)
    
    ' ���݂̃p�X��C�ӂ̃p�X�ɐݒ肷��
    ChDrive path
    ChDir path

End Sub

' =========================================================
' �����������\�b�h
' =========================================================
Public Sub initByThisWorkbook()
    
    ' ���݂̃p�X�����[�N�u�b�N�̃p�X�ɐݒ肷��
    ChDrive ThisWorkbook.path
    ChDir ThisWorkbook.path

End Sub
