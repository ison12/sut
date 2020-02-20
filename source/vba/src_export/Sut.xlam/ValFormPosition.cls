VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValFormPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �t�H�[���̈ʒu�����W�X�g���ɏo�����ꂷ��N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/06/11�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' =========================================================
' �������o
' =========================================================
' �t�H�[����
Private m_formName As String

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

End Sub

' =========================================================
' �����������\�b�h
' =========================================================
Public Sub init(ByVal formName As String)

    m_formName = formName
End Sub

' =========================================================
' ���A�v���P�[�V�����v���p�e�B�̐���
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Set createApplicationProperties = New ApplicationProperties
    createApplicationProperties.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_FORM_POSITION & "\" & m_formName & ".ini"
    
End Function

' =========================================================
' ������ǂݍ���
' =========================================================
Public Function readForData(ByRef rect As ValRectPt)

    On Error GoTo err
    
    ' �X�N���[���T�C�Y���擾����
    Dim ss As ScreenSizePt
    ss = WinAPI_User.getScreenSizePt
    
    Dim appProp As ApplicationProperties: Set appProp = createApplicationProperties
    
    Dim varX As String
    Dim varY As String
    
    varX = appProp.getValue(ConstantsApplicationProperties.INI_SECTION_DEFAULT, ConstantsApplicationProperties.INI_KEY_X)
    varY = appProp.getValue(ConstantsApplicationProperties.INI_SECTION_DEFAULT, ConstantsApplicationProperties.INI_KEY_Y)
    
    ' ���W�X�g����������擾�ł����ꍇ�́A������𗘗p����
    If _
        IsNumeric(varX) = True _
    And IsNumeric(varY) = True Then
    
        ' ����ݒ菈���𔲂���
        rect.Left = CSng(varX)
        rect.Top = CSng(varY)
        
        ' �v�Z�������S�ʒu���X�N���[�����Ɏ��܂��Ă��邩
        If VBUtil.isInnerScreen( _
                            ss.virtualSizeX _
                          , ss.virtualSizeY _
                          , ss.virtualSizeWidth _
                          , ss.virtualSizeHeight _
                          , rect.Left _
                          , rect.Top _
                          , rect.Width _
                          , rect.Height) = True Then

            ' ���W�X�g������擾������񂪃X�N���[�����Ɏ��܂��Ă���Ώ����𔲂���
            Exit Function
        End If
        
    End If
    
    Dim newX As Single
    Dim newY As Single
    
    ' Application�{�̂̃E�B���h�E�̒��S�ʒu���v�Z����
    VBUtil.calcCenterPoint _
                        Application.Left _
                      , Application.Top _
                      , Application.Width _
                      , Application.Height _
                      , newX _
                      , newY _
                      , rect.Width _
                      , rect.Height
    
    ' �v�Z�������S�ʒu���X�N���[�����Ɏ��܂��Ă��邩
    If VBUtil.isInnerScreen( _
                        ss.virtualSizeX _
                      , ss.virtualSizeY _
                      , ss.virtualSizeWidth _
                      , ss.virtualSizeHeight _
                      , newX _
                      , newY _
                      , rect.Width _
                      , rect.Height) = False Then

        ' ���܂��Ă��Ȃ��ꍇ
        ' �v���C�}���X�N���[���ɑ΂��Ă̒��S�ʒu���v�Z����
        VBUtil.calcCenterPoint _
                            0 _
                          , 0 _
                          , ss.primarySizeWidth _
                          , ss.primarySizeHeight _
                          , newX _
                          , newY _
                          , rect.Width _
                          , rect.Height
        
    End If
    
    rect.Left = newX
    rect.Top = newY
    
    Set appProp = Nothing
    
    Exit Function
    
err:

    ' �G���[����ޔ�����
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    Set appProp = Nothing

    ' �ޔ������G���[����ݒ肵�Ȃ���
    VBUtil.setErr errT

    Main.ShowErrorMessage
    
End Function

' =========================================================
' ��������������
' =========================================================
Public Sub writeForData(ByRef rect As ValRectPt)

    On Error GoTo err
    
    Dim appProp As ApplicationProperties: Set appProp = createApplicationProperties
    
    appProp.setValue ConstantsApplicationProperties.INI_SECTION_DEFAULT, ConstantsApplicationProperties.INI_KEY_X, rect.Left
    appProp.setValue ConstantsApplicationProperties.INI_SECTION_DEFAULT, ConstantsApplicationProperties.INI_KEY_Y, rect.Top
    
    Set appProp = Nothing
    
    Exit Sub
    
err:
    
    ' �G���[����ޔ�����
    Dim errT As errInfo: errT = VBUtil.swapErr
    
    Set appProp = Nothing

    ' �ޔ������G���[����ݒ肵�Ȃ���
    VBUtil.setErr errT

    Main.ShowErrorMessage

End Sub