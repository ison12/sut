VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValRecordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���R�[�h���X�g��\������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/02/01�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �z��
Public values           As Variant
' Range
Public valuesCellOffset As Range

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    values = Empty
    Set valuesCellOffset = Nothing
    
End Sub