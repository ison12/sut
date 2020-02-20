Attribute VB_Name = "WinAPI_GDI"
Option Explicit

' *********************************************************
' GDI関連DLLのモジュール
'
' 作成者　：Ison
' 履歴　　：2009/05/04　新規作成
'
' 特記事項：
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

' GetDeviceCaps関連の定数
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

' fWeightの定数
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
' lfCharSetの定数
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
' lfOutPrecisionの定数
Private Const OUT_CHARCTER_PRECIS  As Long = 2
Private Const OUT_DEFAULT_PRECIS   As Long = 0
Private Const OUT_DEVICE_PRECIS    As Long = 5
Private Const OUT_RASTER_PRECIS    As Long = 6
Private Const OUT_STRING_PRECIS    As Long = 1
Private Const OUT_STROKE_PRECIS    As Long = 3
Private Const OUT_TT_ONLY_PRECIS   As Long = 7
Private Const OUT_TT_PRECIS        As Long = 4
' lfClipPrecisionの定数
Private Const CLIP_DEFAULT_PRECIS  As Long = 0
Private Const CLIP_CHARCTER_PRECIS As Long = 1
Private Const CLIP_STROKE_PRECIS   As Long = 2
Private Const CLIP_MASK            As Long = 15
Private Const CLIP_EMBEDDED        As Long = 128
Private Const CLIP_LH_ANGLES       As Long = 16
Private Const CLIP_TT_ALWAYS       As Long = 32
' lfPitchAndFamilyの定数
Private Const DEFAULT_PITCH        As Long = 0
Private Const FIXED_PITCH          As Long = 1
Private Const VARIABLE_PITCH       As Long = 2
Private Const FF_DECORATIVE        As Long = 80
Private Const FF_DONTCARE          As Long = 0
Private Const FF_MODERN            As Long = 48
Private Const FF_ROMAN             As Long = 16
Private Const FF_SCRIPT            As Long = 64
Private Const FF_SWISS             As Long = 32

' FontTypeの定数
Private Const RASTER_FONTTYPE      As Long = &H1
Private Const DEVICE_FONTTYPE      As Long = &H2
Private Const TRUETYPE_FONTTYPE    As Long = &H4

' フォントリスト
' ※getFontNameListとEnumFontFamExProcで共有するためのリスト変数をモジュールレベルで宣言しておく
Private fontList       As ValCollection
Private fontListAtmark As ValCollection

' =========================================================
' ▽システムDPIを取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：DPI情報
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
' ▽フォント名リストの取得
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：フォント名のリストコレクション
' =========================================================
Public Function getFontNameList() As ValCollection

    On Error GoTo err
    
    ' デバイスコンテキストを取得する（念のためゼロクリアしておく）
    Dim hdc As Long: hdc = 0
    
    hdc = WinAPI_User.GetDC(0)
    
    ' フォントの検索条件
    Dim fontSearch As LOGFONT
    
    fontSearch.lfCharSet = DEFAULT_CHARSET
    fontSearch.lfFaceName = String(LF_FACESIZE, vbNullChar)
    fontSearch.lfPitchAndFamily = DEFAULT_PITCH
    
    ' フォントリストを初期化する
    Set fontList = Nothing
    Set fontListAtmark = Nothing
    Set fontList = New ValCollection
    Set fontListAtmark = New ValCollection
    
    ' フォントリストを列挙する
    EnumFontFamiliesEx hdc, fontSearch, AddressOf EnumFontFamExProc, 0, 0

    ' 縦書きと横書きのフォントリストを全て結合し
    ' 戻り値として返す
    Set getFontNameList = fontListAtmark
    
    Dim temp As Variant
    
    For Each temp In fontList.col
    
        getFontNameList.setItem temp, temp
    Next

    ' デバイスコンテキストを解放する
    WinAPI_User.ReleaseDC 0, hdc
    
    Exit Function
    
err:

    ' デバイスコンテキストを解放する
    WinAPI_User.ReleaseDC 0, hdc
    
    err.Raise err.Number

End Function

' =========================================================
' ▽フォント名の列挙（WinAPIのコールバック関数）
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
' =========================================================
Private Function EnumFontFamExProc(ByRef lpelfe As ENUMLOGFONTEX _
                                 , ByRef lpntme As Long _
                                 , ByVal FontType As Long _
                                 , ByRef lParam As Long) As Long

    ' フォント名
    Dim fontName As String
    ' 文字列を取得（NULL文字の位置を判定する）
    fontName = Left(lpelfe.elfFullName, InStr(lpelfe.elfFullName, vbNullChar))
    
    ' フォント名の先頭が "@" の場合、縦書きフォントとして判定する
    If InStr(fontName, "@") = 1 Then
    
        If fontListAtmark.exist(fontName) = False Then
        
            ' フォント名を設定する
            fontListAtmark.setItem fontName, fontName
            
            #If DEBUG_MODE = 1 Then
            
                Debug.Print fontName & " : " & FontType
            #End If
    
        End If
    
    ' 上記以外、横書きフォント
    Else
    
        If fontList.exist(fontName) = False Then
        
            ' フォント名を設定する
            fontList.setItem fontName, fontName
            
            #If DEBUG_MODE = 1 Then
            
                Debug.Print fontName & " : " & FontType
            #End If
    
        End If
        
    End If
    
    
    ' 列挙を続けるため、1を返す（列挙を中断する場合は、0を返す必要がある）
    EnumFontFamExProc = 1

End Function
