VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSnapRecordsSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�i�b�v�V���b�g�̃��R�[�h�Z�b�g���
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/01/06�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���O
Public name       As String
' ����
Public getDate     As Date
' ���R�[�h��
Public recordCount As Long

' ���X�g
Public list As ValCollection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set list = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    
    Set list = Nothing
End Sub