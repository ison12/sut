Attribute VB_Name = "WinAPI_GDI"
Option Explicit

' *********************************************************
' GDI�֘ADLL�̃��W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/05/04�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

#If VBA7 And Win64 Then
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As LongPtr, _
    ByVal nIndex As Long _
) As Long

Private Declare PtrSafe Function EnumFontFamiliesEx Lib "gdi32.dll" Alias "EnumFontFamiliesExW" _
        (ByVal hdc As LongPtr _
       , ByRef lpLogFont As LOGFONT _
       , ByVal lpEnumFontFamExProc As LongPtr _
       , ByVal lParam As Long _
       , ByVal dwFlags As Long) As Long
#Else
Public Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nIndex As Long _
) As Long

Private Declare Function EnumFontFamiliesEx Lib "gdi32.dll" Alias "EnumFontFamiliesExW" _
        (ByVal hdc As Long _
       , ByRef lpLogFont As LOGFONT _
       , ByVal lpEnumFontFamExProc As Long _
       , ByVal lParam As Long _
       , ByVal dwFlags As Long) As Long
#End If

' GetDeviceCaps�֘A�̒萔
Private Const LOGPIXELSX As Long = &H58&
Private Const LOGPIXELSY As Long = &H5A&

Public Type DPI

    horizontal As Long
    vertical   As Long
    
End Type

Private Const LF_FACESIZE     As Long = 32
Private Const LF_FULLFACESIZE As Long = 64

Private Type LOGFONT
    
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQualiy         As Byte
    lfPitchAndFamily As Byte
    lfFaceName       As String * LF_FACESIZE
End Type

Private Type ENUMLOGFONTEX

    elfLogFont     As LOGFONT
    elfFullName    As String * LF_FULLFACESIZE
    elfStyle       As String * LF_FACESIZE
    elfScript      As String * LF_FACESIZE
End Type

' fWeight�̒萔
Private Const FW_DONTCARE       As Long = 0
Private Const FW_THIN           As Long = 100
Private Const FW_EXTRALIGHT     As Long = 200
Private Const FW_ULTRALIGHT     As Long = FW_EXTRALIGHT
Private Const FW_LIGHT          As Long = 300
Private Const FW_NOMAL          As Long = 400
Private Const FW_REGULAR        As Long = FW_NOMAL
Private Const FW_MEDIUM         As Long = 500
Private Const FW_SEMIBOLD       As Long = 600
Private Const FW_DEMIBOLD       As Long = FW_SEMIBOLD
Private Const FW_BOLD           As Long = 700
Private Const FW_EXTRABOLD      As Long = 800
Private Const FW_ULTRABOLD      As Long = FW_EXTRABOLD
Private Const FW_HEAVY          As Long = 900
Private Const FW_BLACK          As Long = FW_HEAVY
' lfCharSet�̒萔
Private Const ANSI_CHARSET      As Long = 0
Private Const DEFAULT_CHARSET   As Long = 1
Private Const OEM_CHARSET       As Long = 255
Private Const SHIFTJIS_CHARSET  As Long = 128
Private Const SYMBOL_CHARSET    As Long = 2
Private Const BALTIC_CHARSET    As Long = 186
Private Const CHINESEBIG5_CHARSET  As Long = 136
Private Const EASTEUROPE_CHARSET   As Long = 238
Private Const GREEK_CHARSET        As Long = 161
Private Const HANGEUL_CHARSET      As Long = 129
Private Const MAC_CHARSET          As Long = 77
Private Const RUSSIAN_CHARSET      As Long = 204
Private Const TURKISH_CHARSET      As Long = 162
' lfOutPrecision�̒萔
Private Const OUT_CHARCTER_PRECIS  As Long = 2
Private Const OUT_DEFAULT_PRECIS   As Long = 0
Private Const OUT_DEVICE_PRECIS    As Long = 5
Private Const OUT_RASTER_PRECIS    As Long = 6
Private Const OUT_STRING_PRECIS    As Long = 1
Private Const OUT_STROKE_PRECIS    As Long = 3
Private Const OUT_TT_ONLY_PRECIS   As Long = 7
Private Const OUT_TT_PRECIS        As Long = 4
' lfClipPrecision�̒萔
Private Const CLIP_DEFAULT_PRECIS  As Long = 0
Private Const CLIP_CHARCTER_PRECIS As Long = 1
Private Const CLIP_STROKE_PRECIS   As Long = 2
Private Const CLIP_MASK            As Long = 15
Private Const CLIP_EMBEDDED        As Long = 128
Private Const CLIP_LH_ANGLES       As Long = 16
Private Const CLIP_TT_ALWAYS       As Long = 32
' lfPitchAndFamily�̒萔
Private Const DEFAULT_PITCH        As Long = 0
Private Const FIXED_PITCH          As Long = 1
Private Const VARIABLE_PITCH       As Long = 2
Private Const FF_DECORATIVE        As Long = 80
Private Const FF_DONTCARE          As Long = 0
Private Const FF_MODERN            As Long = 48
Private Const FF_ROMAN             As Long = 16
Private Const FF_SCRIPT            As Long = 64
Private Const FF_SWISS             As Long = 32

' FontType�̒萔
Private Const RASTER_FONTTYPE      As Long = &H1
Private Const DEVICE_FONTTYPE      As Long = &H2
Private Const TRUETYPE_FONTTYPE    As Long = &H4

' �t�H���g���X�g
' ��getFontNameList��EnumFontFamExProc�ŋ��L���邽�߂̃��X�g�ϐ������W���[�����x���Ő錾���Ă���
Private fontList       As ValCollection
Private fontListAtmark As ValCollection

' =========================================================
' ���V�X�e��DPI���擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�FDPI���
' =========================================================
Public Function getSystemDPI() As DPI

    Dim hdc As Long
    hdc = WinAPI_User.GetDC(0)
    
    If hdc = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "" _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_FAILED
    End If
    
    getSystemDPI.horizontal = GetDeviceCaps(hdc, LOGPIXELSX)
    getSystemDPI.vertical = GetDeviceCaps(hdc, LOGPIXELSY)

    If WinAPI_User.ReleaseDC(0, hdc) = 0 Then
    
        err.Raise ConstantsError.ERR_NUMBER_DLL_FUNCTION_FAILED _
                , "" _
                , ConstantsError.ERR_DESC_DLL_FUNCTION_FAILED
    End If

End Function

' =========================================================
' ���t�H���g�����X�g�̎擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�t�H���g���̃��X�g�R���N�V����
' =========================================================
Public Function getFontNameList() As ValCollection

    On Error GoTo err
    
    ' �f�o�C�X�R���e�L�X�g���擾����i�O�̂��߃[���N���A���Ă����j
    Dim hdc As Long: hdc = 0
    
    hdc = WinAPI_User.GetDC(0)
    
    ' �t�H���g�̌�������
    Dim fontSearch As LOGFONT
    
    fontSearch.lfCharSet = DEFAULT_CHARSET
    fontSearch.lfFaceName = String(LF_FACESIZE, vbNullChar)
    fontSearch.lfPitchAndFamily = DEFAULT_PITCH
    
    ' �t�H���g���X�g������������
    Set fontList = Nothing
    Set fontListAtmark = Nothing
    Set fontList = New ValCollection
    Set fontListAtmark = New ValCollection
    
    ' �t�H���g���X�g��񋓂���
    EnumFontFamiliesEx hdc, fontSearch, AddressOf EnumFontFamExProc, 0, 0

    ' �c�����Ɖ������̃t�H���g���X�g��S�Č�����
    ' �߂�l�Ƃ��ĕԂ�
    Set getFontNameList = fontListAtmark
    
    Dim temp As Variant
    
    For Each temp In fontList.col
    
        getFontNameList.setItem temp, temp
    Next

    ' �f�o�C�X�R���e�L�X�g���������
    WinAPI_User.ReleaseDC 0, hdc
    
    Exit Function
    
err:

    ' �f�o�C�X�R���e�L�X�g���������
    WinAPI_User.ReleaseDC 0, hdc
    
    err.Raise err.Number

End Function

' =========================================================
' ���t�H���g���̗񋓁iWinAPI�̃R�[���o�b�N�֐��j
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
' =========================================================
Private Function EnumFontFamExProc(ByRef lpelfe As ENUMLOGFONTEX _
                                 , ByRef lpntme As Long _
                                 , ByVal FontType As Long _
                                 , ByRef lParam As Long) As Long

    ' �t�H���g��
    Dim fontName As String
    ' ��������擾�iNULL�����̈ʒu�𔻒肷��j
    fontName = Left(lpelfe.elfFullName, InStr(lpelfe.elfFullName, vbNullChar))
    
    ' �t�H���g���̐擪�� "@" �̏ꍇ�A�c�����t�H���g�Ƃ��Ĕ��肷��
    If InStr(fontName, "@") = 1 Then
    
        If fontListAtmark.exist(fontName) = False Then
        
            ' �t�H���g����ݒ肷��
            fontListAtmark.setItem fontName, fontName
            
            #If DEBUG_MODE = 1 Then
            
                Debug.Print fontName & " : " & FontType
            #End If
    
        End If
    
    ' ��L�ȊO�A�������t�H���g
    Else
    
        If fontList.exist(fontName) = False Then
        
            ' �t�H���g����ݒ肷��
            fontList.setItem fontName, fontName
            
            #If DEBUG_MODE = 1 Then
            
                Debug.Print fontName & " : " & FontType
            #End If
    
        End If
        
    End If
    
    
    ' �񋓂𑱂��邽�߁A1��Ԃ��i�񋓂𒆒f����ꍇ�́A0��Ԃ��K�v������j
    EnumFontFamExProc = 1

End Function
