VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWindowProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' Callback�֐� WindowProc �̃C���^�[�t�F�[�X
' Window���T�u�N���X�����Ǝ��Ƀ��b�Z�[�W��������ꍇ�A�{�C���^�[�t�F�[�X��Implements����B
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/10/11�@�V�K�쐬
'
' ���L�����F
' �@�֘A���W���[�����ȉ��Ɏ����B
' �@�@�DIWindowProc.cls
' �@�A�DWinAPI_CallbackWindowProc.bas
' �@�B�DWinAPI_User.bas
'
' *********************************************************

' =========================================================
' �����C�����\�b�h
'
' �T�v�@�@�@�F�E�B���h�E���b�Z�[�W���������郁�\�b�h�B
' �����@�@�@�FhWnd   �E�B���h�E�n���h��
' �@�@�@�@�@�@msg    ���b�Z�[�W
' �@�@�@�@�@�@wParam �p�����[�^����1
' �@�@�@�@�@�@lParam �p�����[�^����2
' �@�@�@�@�@�@result �������ʃR�[�h
' �߂�l�@�@�F�����������ǂ�����\���t���O
'
' =========================================================
Public Function process(ByVal hWnd As Long _
                         , ByVal msg As Long _
                         , ByVal wParam As Long _
                         , ByVal lParam As Long _
                         , ByRef result As Long) As Boolean

End Function
