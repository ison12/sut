VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryResultSetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �N�G�����ʃZ�b�g���N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public queryResultInfoList As ValCollection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    Set queryResultInfoList = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    Set queryResultInfoList = Nothing
End Sub

' =========================================================
' ���N�G�����ʏ��̒ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub addInfo(info As ValQueryResultInfo)
    
    queryResultInfoList.setItem info

End Sub

' =========================================================
' ���N�G�����ʏ��̃N���A
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub clearInfos()
    
    Set queryResultInfoList = New ValCollection

End Sub

' =========================================================
' ���N�G�����ʏ��̎擾
'
' �T�v�@�@�@�F
' �����@�@�@�Findex �C���f�b�N�X
' �߂�l�@�@�F
'
' =========================================================
Public Function getInfo(index As Long) As ValQueryResultInfo

    Set getInfo = queryResultInfoList.getItemByIndex(index)

End Function

' =========================================================
' ���N�G�����ʏ��̌����擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�N�G�����ʏ��̌���
'
' =========================================================
Public Function getInfoCount() As Long

    getInfoCount = queryResultInfoList.count

End Function