VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbColumnType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�̃f�t�H���g�̃J������������\������N���X
' �Ή�����DB���ɖ{�C���^�[�t�F�[�X��Implements����B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/08/14�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' =========================================================
' ��DB�R�l�N�V�����ݒ�
'
' �T�v�@�@�@�FDB�R�l�N�V������ݒ肷��B
' �����@�@�@�FvNewValue DB�R�l�N�V����
'
' ���L�����@�F
'
'
' =========================================================
Public Property Let dbConn(ByVal vNewValue As Variant)

End Property

' =========================================================
' ��DB�̃f�t�H���g�̃J�������������擾����
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' �߂�l�@�@�FDB�J�����������
'
' ���L�����@�F
'
'
' =========================================================
Public Function getDefaultColumnFormat() As ValCollection

End Function