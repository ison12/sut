Attribute VB_Name = "SutWhite"
Option Explicit
' *********************************************************
' SutWhite.dll�֘A�̃��W���[��
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2009/03/14�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Private Const HOURGLASS_WIDTH  As Single = 55
Private Const HOURGLASS_HEIGHT As Single = 64

' DLL�̃n���h��
Private libraryHandle As Variant
' DLL�̃p�X
Private libraryPath As String

' =========================================================
' �����C�u���������[�h����
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function LoadLibrary()

    ' DLL�̃p�X��ݒ肷��
    #If (DEBUG_MODE = 1) Then
    
        #If VBA7 And Win64 Then
            libraryPath = SutWorkbook.path & "\..\CPP\Sut\x64\Debug ASM\SutWhite.dll"
        
        #Else
            libraryPath = SutWorkbook.path & "\..\CPP\Sut\Debug ASM\SutWhite.dll"
        
        #End If
    #Else
        ' DLL�̃p�X��ݒ�
        libraryPath = SutWorkbook.path & "\lib\SutWhite.dll"
        
    #End If
    

    ' ���W���[���n���h��
    Dim handle As Variant
    
    ' �n���h�����擾����
    handle = WinAPI_Kernel32.GetModuleHandle _
                (libraryPath)

    ' �����[�h�̏ꍇ
    If handle = 0 Then
    
        ' dll�����[�h����
        libraryHandle = WinAPI_Kernel32.LoadLibrary _
                                    (libraryPath)
    
        ' �߂�l��NULL�̏ꍇ
        If libraryHandle = 0 Then
        
            ' �G���[�𔭍s����
            err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                    , _
                    , ConstantsError.ERR_DESC_DLL_FUNCTION_FAILED
        End If
    
    ' ���[�h�ς݂̏ꍇ
    Else
    
        libraryHandle = handle
    End If

End Function

' =========================================================
' �����C�u�������������
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function freeLibrary()

    ' �n���h���̃`�F�b�N���s��
    If libraryHandle = 0 Then
    
        ' �n���h�������蓖�Ă��Ă��Ȃ��ꍇ�A�I������
        Exit Function
    End If
    
    ' SutWhite.dll ���������
    WinAPI_Kernel32.freeLibrary (libraryHandle)
    
    ' �n���h�����[���N���A����
    libraryHandle = 0
End Function

' =========================================================
' �����C�u����������������
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function initialize()

    ' �n���h���̃`�F�b�N���s��
    If libraryHandle = 0 Then
    
        ' �n���h�������蓖�Ă��Ă��Ȃ��ꍇ�A�I������
        Exit Function
    End If

    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "Initialize")
    
    ' DLL�֐��̖߂�l
    Dim ret As Long
    
    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtr(procAddr)

    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function

' =========================================================
' �����C�u������j������
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function destroy()

    ' �n���h���̃`�F�b�N���s��
    If libraryHandle = 0 Then
    
        ' �n���h�������蓖�Ă��Ă��Ȃ��ꍇ�A�I������
        Exit Function
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "Destroy")
    
    ' DLL�֐��̖߂�l
    Dim ret As Long
    
    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtr(procAddr)
    
    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function

' =========================================================
' ���X�v���b�V���E�B���h�E��\������
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function showSplashWindow()

    ' �n���h���̃`�F�b�N���s��
    If libraryHandle = 0 Then
    
        ' �n���h�������蓖�Ă��Ă��Ȃ��ꍇ
        LoadLibrary     ' ���C�u�����̃��[�h
        initialize      ' ���C�u�����̏�����
        
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "ShowSplashWindow")
    
    ' DLL�֐��̖߂�l
    Dim ret As Long
    
    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtrParam(procAddr, ExcelUtil.getApplicationHWnd)
    
    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function

' =========================================================
' ���X�v���b�V���E�B���h�E�̕\������������܂őҋ@����
'
' �T�v�@�@�@�F
' �����@�@�@�FwaitTime �~���b�w��
' �߂�l�@�@�F0  ����
' �@�@�@�@�@�@10 �^�C���A�E�g
'
' =========================================================
Public Function waitSplashWindow(ByVal waitTime As Long) As Long

    ' �n���h���̃`�F�b�N���s��
    If libraryHandle = 0 Then
    
        ' �n���h�������蓖�Ă��Ă��Ȃ��ꍇ�A�I������
        Exit Function
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "WaitSplashWindow")
    
    ' DLL�֐��̖߂�l
    Dim ret As Long
        
    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtrParamInt(procAddr, waitTime)
    
    If ret <> 0 And ret <> 10 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
    waitSplashWindow = ret
    
End Function

' =========================================================
' �������v�E�B���h�E��\������i�t�H�[���̒��S�ɕ\������j
'
' �T�v�@�@�@�F
' �����@�@�@�FfrmObj �t�H�[���I�u�W�F�N�g
'
' =========================================================
Public Function showHourglassWindowOnCenterPt(Optional ByRef frmObj As Object = Nothing, _
                                            Optional ByVal x As Long = 0, _
                                            Optional ByVal y As Long = 0)

    Dim newX As Single
    Dim newY As Single

    ' DPI���擾����
    Dim d As DPI
    d = WinAPI_GDI.getSystemDPI
    
    If Not frmObj Is Nothing Then
    
        VBUtil.calcCenterPoint frmObj.Left _
                             , frmObj.Top _
                             , frmObj.Width _
                             , frmObj.Height _
                             , newX _
                             , newY _
                             , VBUtil.convertPixelToPoint(d.horizontal, HOURGLASS_WIDTH) _
                             , VBUtil.convertPixelToPoint(d.vertical, HOURGLASS_HEIGHT)
    Else
    
        VBUtil.calcCenterPoint Application.Left _
                             , Application.Top _
                             , Application.Width _
                             , Application.Height _
                             , newX _
                             , newY _
                             , VBUtil.convertPixelToPoint(d.horizontal, HOURGLASS_WIDTH) _
                             , VBUtil.convertPixelToPoint(d.vertical, HOURGLASS_HEIGHT)
    End If
                         
    showHourglassWindow VBUtil.convertPointToPixel(d.horizontal, newX) + x _
                      , VBUtil.convertPointToPixel(d.vertical, newY) + y

End Function

' =========================================================
' �������v�E�B���h�E��\������
'
' �T�v�@�@�@�F���W�̓s�N�Z���P�ʂŎw�肷��
' �����@�@�@�Fx �E�B���h�E�\���ʒu X
' �@�@�@�@�@�@y �E�B���h�E�\���ʒu Y
'
' =========================================================
Public Function showHourglassWindow(Optional ByVal x As Long = 0 _
                                  , Optional ByVal y As Long = 0)

    ' �n���h���̃`�F�b�N���s��
    If libraryHandle = 0 Then
    
        ' �n���h�������蓖�Ă��Ă��Ȃ��ꍇ
        LoadLibrary     ' ���C�u�����̃��[�h
        initialize      ' ���C�u�����̏�����
        
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "ShowHourglassWindow")
    
    ' DLL�֐��̖߂�l
    Dim ret As Long
    
    ' �E�B���h�E�\���ʒu
    Dim pt As point
    pt.x = x
    pt.y = y
    
    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtrParam2(procAddr, ExcelUtil.getApplicationHWnd, pt)
    
    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function

' =========================================================
' �������v�E�B���h�E���\���ɂ���
'
' �T�v�@�@�@�F
'
' =========================================================
Public Function closeHourglassWindow()

    ' �n���h���̃`�F�b�N���s��
    If libraryHandle = 0 Then
    
        ' �n���h�������蓖�Ă��Ă��Ȃ��ꍇ�A�I������
        Exit Function
    End If
    
    Dim procAddr As Variant
    procAddr = WinAPI_Kernel32.GetProcAddress(libraryHandle, "CloseHourglassWindow")
    
    ' DLL�֐��̖߂�l
    Dim ret As Long
    
    ' �f�B���N�g�����ꎞ�I�ɕύX����
    Dim appCurDirChanger As New ApplicationCurDirChanger: appCurDirChanger.initByThisWorkbook
    
    ret = SutGreen.CallByFuncPtr(procAddr)
    
    If ret <> 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "SutWhite.dll" _
                , ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED
    End If
    
End Function
