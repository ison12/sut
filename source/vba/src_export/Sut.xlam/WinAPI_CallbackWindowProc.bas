Attribute VB_Name = "WinAPI_CallbackWindowProc"
Option Explicit

' *********************************************************
' �E�B���h�E�v���V�[�W���̏�������������IWindowProc�ւ�
' ������U�蕪����R���g���[���̖������ʂ������W���[���B
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

' IWindowProc�I�u�W�F�N�g���i�[���郊�X�g�I�u�W�F�N�g
Private list               As ValCollection
' �ݒ�O�̃E�B���h�E�v���V�[�W�����i�[���郊�X�g�I�u�W�F�N�g
Private listPrevWindowProc As ValCollection


' =========================================================
' ��IWindowProc�I�u�W�F�N�g�o�^
'
' �T�v�@�@�@�FIWindowProc�I�u�W�F�N�g��o�^����B
' �����@�@�@�Fproc   IWindowProc�����������I�u�W�F�N�g�ϐ�
' �@�@�@�@�@�@hWnd   �E�B���h�E�n���h��
'
' �߂�l�@�@�F
'
' =========================================================
Public Sub registWindowProc(ByRef proc As IWindowProc _
                          , ByVal hwnd As Long)


    ' ���X�g������������Ă��Ȃ��ꍇ�A�����������{����
    If list Is Nothing Then
    
        ' ����������
        Set list = New ValCollection
        Set listPrevWindowProc = New ValCollection
    End If
    
    Dim prevWindowProc As Long
    
    ' �ݒ�O�̃E�B���h�E�v���V�[�W��
    prevWindowProc = WinAPI_User.GetWindowLong(hwnd, WinAPI_User.GWL_WNDPROC)
    
    ' �G���[�`�F�b�N
    If prevWindowProc = 0 Then
    
        err.Raise 5000 _
                , 0 _
                , "API�G���["
    
    End If
    
    
    ' �E�B���h�E�n���h�����L�[�ɁAIWindowProc�I�u�W�F�N�g��ݒ肷��
    list.setItem proc, hwnd
    ' �E�B���h�E�n���h�����L�[�ɁA�ݒ�O�̃E�B���h�E�v���V�[�W����ݒ肷��
    list.setItem prevWindowProc, hwnd
    
    ' �T�u�N���X���J�n
    If WinAPI_User.SetWindowLong(hwnd _
                                , WinAPI_User.GWL_WNDPROC _
                                , AddressOf windowProcedure) = 0 Then
                            
        err.Raise 5000 _
                , 0 _
                , "API�G���["
                                        
    End If

End Sub

' =========================================================
' ��IWindowProc�I�u�W�F�N�g�폜
'
' �T�v�@�@�@�FIWindowProc�I�u�W�F�N�g���폜����
' �����@�@�@�FhWnd           �E�B���h�E�n���h��
' �@�@�@�@�@�@prevWindowProc �ŏ��ɐݒ肳��Ă����E�B���h�E�v���V�[�W��
'
' �߂�l�@�@�F�����ɐ����������ǂ�����\���t���O
'
' =========================================================
Public Sub unregistWindowProc(ByVal hwnd As Long)

    ' �E�B���h�E�n���h�����L�[�ɁAIWindowProc�I�u�W�F�N�g���폜����
    list.remove hwnd
    
    ' �ݒ�O�̃E�B���h�E�v���V�[�W��
    Dim prevWindowProc As Long
    
    ' �E�B���h�E�n���h�����L�[�ɁA�ݒ�O�̃E�B���h�E�v���V�[�W�����폜����
    prevWindowProc = listPrevWindowProc.getItem(hwnd, vbLong)
    
    ' �ŏ��ɐݒ肳��Ă����E�B���h�E�v���V�[�W�����ăZ�b�g����
    WinAPI_User.SetWindowLong hwnd, WinAPI_User.GWL_WNDPROC, prevWindowProc

    listPrevWindowProc.remove hwnd

End Sub

' =========================================================
' ���E�B���h�E�v���V�[�W��
'
' �T�v�@�@�@�F���b�Z�[�W��IWindowProc�ɐU�蕪����B
' �����@�@�@�FhWnd   �E�B���h�E�n���h��
' �@�@�@�@�@�@msg    ���b�Z�[�W
' �@�@�@�@�@�@wParam �p�����[�^����1
' �@�@�@�@�@�@lParam �p�����[�^����2
'
' �߂�l�@�@�F���ʃR�[�h
'
' =========================================================
Private Function windowProcedure(ByVal hwnd As Long _
                               , ByVal msg As Long _
                               , ByVal wParam As Long _
                               , ByVal lParam As Long) As Long

    ' ���ʃR�[�h
    Dim resultCode As Long
    
    ' IWindowProc�I�u�W�F�N�g
    Dim windowProc As IWindowProc
    
    ' �ŏ��ɐݒ肳��Ă����E�B���h�E�v���V�[�W��
    Dim prevWindowProc As Long

    ' �I�u�W�F�N�g���擾����
    Set windowProc = list.getItem(hwnd)

    ' �E�B���h�E�v���V�[�W���I�u�W�F�N�g�̃`�F�b�N
    If Not windowProc Is Nothing Then
    
        ' IWindowProc�I�u�W�F�N�g�ɏ�����U�蕪����
        If windowProc.process(hwnd, msg, wParam, lParam, resultCode) = False Then
        
            ' �ŏ��ɐݒ肳��Ă����E�B���h�E�v���V�[�W�����擾����
            prevWindowProc = listPrevWindowProc.getItem(hwnd, vbLong)
            ' �f�t�H���g���b�Z�[�W����
            windowProcedure = CallWindowProc(prevWindowProc, hwnd, msg, wParam, lParam)
            
        Else
        
            ' �������ʃR�[�h��Ԃ�
            windowProcedure = resultCode
        End If

    End If
    
End Function

