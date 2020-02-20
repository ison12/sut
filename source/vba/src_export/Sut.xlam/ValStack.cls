VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�^�b�N
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/17�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �R���N�V�����I�u�W�F�N�g
Private col As collection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set col = New collection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set col = Nothing
End Sub

' =========================================================
' ���f�[�^�ǉ�
'
' �T�v�@�@�@�F�X�^�b�N�Ƀf�[�^��ǉ�����B
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub push(ByVal value As Variant)

    col.Add value
End Sub

' =========================================================
' ���f�[�^�擾
'
' �T�v�@�@�@�F�X�^�b�N����f�[�^���擾����B
' �����@�@�@�Ft �f�[�^�^
' �߂�l�@�@�F
'
' =========================================================
Public Function pop(Optional t As VbVarType = vbVariant) As Variant

    ' �T�C�Y���O�̏ꍇ
    If col.count = 0 Then
    
        ' ��ŕԂ�
        pop = Empty
        
    ' �T�C�Y���P�ȏ�
    Else
    
        ' �f�[�^�^���I�u�W�F�N�g
        If t = vbObject Then
        
            ' Set���g���Ė߂�l�ɐݒ�
            Set pop = col.item(col.count)
            
        ' �f�[�^�^���I�u�W�F�N�g�ȊO
        Else
        
            ' �߂�l�ɐݒ�
            pop = col.item(col.count)
        End If
        
        ' �Ō���̏����폜����
        col.remove col.count
    
    End If
    
End Function

' =========================================================
' ���T�C�Y�擾
'
' �T�v�@�@�@�F�X�^�b�N�̃T�C�Y���擾����B
' �����@�@�@�F
' �߂�l�@�@�F�X�^�b�N�T�C�Y
'
' =========================================================
Public Function count() As Long

    count = col.count

End Function