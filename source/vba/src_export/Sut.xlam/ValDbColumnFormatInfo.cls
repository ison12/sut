VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbColumnFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�̃J������������\������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/08/01�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' DB��
Public dbName    As String
' �J�������X�g
Public columnList  As ValCollection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    ' �J�������X�g������������
    Set columnList = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set columnList = Nothing
    
End Sub