VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelLongTimeProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' Excel�}�N���Œ����ԏ�������ۂɕs�K�v�ȃv���p�e�B��
' �ꎞ�I�ɃI�t�ɂ��āA�{�N���X�̃I�u�W�F�N�g���폜�����ۂɌ��ɖ߂��B
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/12�@�V�K�쐬

' �g�p���@�F
' �@�{�N���X�̃I�u�W�F�N�g�� New ����B
' �@�֐��𔲂���Ǝ����� Class_Terminate �Ō�n�����s����B
'
'   Dim var As New ExcelLongTimeProcessing: var.init
'
'
' *********************************************************

' �x���\���̗L���E�����i�V�[�g�폜���Ȃǂ̌x���_�C�A���O�Ȃǁj
Public isSetDisplayAlerts    As Boolean
' �}�N�����s����Esc�L�[�̗L���E����
Public isSetEnableCancelKey  As Boolean
' �C�x���g�̗L���E����
Public isSetEnableEvents     As Boolean
' �}�E�X�J�[�\���̌`��
Public isSetCursor           As Boolean
' ��ʍX�V�̒�~�L��
Public isSetScreenUpdating   As Boolean
' Excel���[�N�V�[�g�̌v�Z����
Public isSetCalculation      As Boolean
' Excel�ւ̃L�[�{�[�h�E�}�E�X����
Public isInteractive         As Boolean

' �v�Z�����i�ꎞ�ۑ��j
Private tmpCalculation       As XlCalculation

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    On Error Resume Next
    
    On Error GoTo 0
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' �A�v���P�[�V�����̊e��v���p�e�B�����ɖ߂�
    restoreApplicationProp
    
End Sub

' =========================================================
' ������������
'
' �T�v�@�@�@�FApplication�I�u�W�F�N�g�̊e�v���p�e�B���I�t�ɐݒ肷��B
' �@�@�@�@�@�@�Y���v���p�e�B�̃I���E�I�t��؂�ւ��邩�ǂ�����
' �@�@�@�@�@�@�{���\�b�h�̈����ɂ���Č��߂�B
' �����@�@�@�FdisplayAlerts    �x���\��
' �@�@�@�@�@�@enableCancelKey  ESC�L�[�̗L���E����
' �@�@�@�@�@�@enableEvents     �C�x���g�̗L���E����
' �@�@�@�@�@�@cursor           �}�E�X�J�[�\��
' �@�@�@�@�@�@screenUpdating   ��ʍX�V�̗L���E����
' �@�@�@�@�@�@calculation      Excel���[�N�V�[�g�̌v�Z����
' �@�@�@�@�@�@interactive      �Θb���[�h�i�L�[�{�[�h��}�E�X����̓��͗L���j
'
' =========================================================
Public Sub init(Optional ByVal displayAlerts As Boolean = True _
              , Optional ByVal enableCancelKey As Boolean = True _
              , Optional ByVal enableEvents As Boolean = True _
              , Optional ByVal cursor As Boolean = False _
              , Optional ByVal screenUpdating As Boolean = False _
              , Optional ByVal calculation As Boolean = True _
              , Optional ByVal interactive As Boolean = True)
              
    ' �����o�ɏ���ݒ肷��
    isSetDisplayAlerts = displayAlerts
    isSetEnableCancelKey = enableCancelKey
    isSetEnableEvents = enableEvents
    isSetCursor = cursor
    isSetScreenUpdating = screenUpdating
    isSetCalculation = calculation
    isInteractive = interactive
    
    ' �A�v���P�[�V�����̊e��v���p�e�B�������o�ϐ��̒l�ɏ]���Đݒ肷��
    setApplicationProp

End Sub

' =========================================================
' ���j������
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub destroy()

    ' �A�v���P�[�V�����̊e��v���p�e�B�����ɖ߂�
    restoreApplicationProp

End Sub

' =========================================================
' ���A�v���P�[�V�����̊e��v���p�e�B�������o�ϐ��̒l�ɏ]���Đݒ肷��
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub setApplicationProp()

    On Error Resume Next

    ' �v�Z����
    If isSetCalculation = True Then
        ' �v�Z�������ꎞ�ϐ��ɕۑ����Ă���
        tmpCalculation = Application.calculation
        ' �v�Z�������蓮�ɂ���
        Application.calculation = xlCalculationManual
    End If
    
    ' �x���\��
    If isSetDisplayAlerts = True Then
        Application.displayAlerts = False
    End If
    
    ' Esc�L�[�̗L���E����
    If isSetEnableCancelKey = True Then
        Application.enableCancelKey = xlDisabled
    End If
    
    ' �C�x���g�̗L���E����
    If isSetEnableEvents = True Then
        Application.enableEvents = False
    End If
    
    ' �J�[�\��
    If isSetCursor = True Then
    
        ' �f�o�b�O���̓J�[�\���������v�ɂȂ�̂�h���i�ז��Ȃ̂Łj
        #If (DEBUG_MODE <> 1) Then
            Application.cursor = xlWait
        #End If
        
    End If
    
    ' ��ʂ̎����X�V
    If isSetScreenUpdating = True Then
        Application.screenUpdating = False
    End If
    
    If isInteractive = True Then
        Application.interactive = False
    End If
    
    On Error GoTo 0

End Sub

' =========================================================
' ���A�v���P�[�V�����̊e��v���p�e�B�����ɖ߂�
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub restoreApplicationProp()

    On Error Resume Next
    
    ' �v�Z����
    If isSetCalculation = True Then
        ' �v�Z���������ɖ߂�
        Application.calculation = tmpCalculation
    End If
    
    ' �x���\��
    If isSetDisplayAlerts = True Then
        Application.displayAlerts = True
    End If
    
    ' Esc�L�[�̗L���E����
    If isSetEnableCancelKey = True Then
        Application.enableCancelKey = xlInterrupt
    End If
    
    ' �C�x���g�̗L���E����
    If isSetEnableEvents = True Then
        Application.enableEvents = True
    End If
    
    ' �J�[�\��
    If isSetCursor = True Then
    
        Application.cursor = xlDefault
    End If
    
    ' ��ʂ̎����X�V
    If isSetScreenUpdating = True Then
        Application.screenUpdating = True
    End If

    '
    If isInteractive = True Then
        Application.interactive = True
    End If
    
    On Error GoTo 0

End Sub
