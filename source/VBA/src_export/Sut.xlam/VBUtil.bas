Attribute VB_Name = "VBUtil"
Option Explicit

' *********************************************************
' VB�֘A�̋��ʊ֐����W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/08/10�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' �G���[�����i�[����\����
Public Type errInfo

    Source       As Variant
    Number       As Variant
    Description  As Variant
    LastDllError As Variant
    HelpFile     As Variant
    HelpContext  As Variant
    
End Type

Private Const KEY_CODE_CTRL  As String = "^"
Private Const KEY_CODE_SHIFT As String = "+"
Private Const KEY_CODE_ALT   As String = "%"

' Application#OnKey�̕����}�b�v
Private applicationOnKeyMap1 As ValCollection ' �_�������L�[�ɂ��Ă���
Private applicationOnKeyMap2 As ValCollection ' �R�[�h���L�[�ɂ��Ă���

' ���W�X�g���p�X - �����R�[�h�ꗗ
Private Const REG_PATH_CHARACTER_CODE_LIST As String = "MIME\Database\Charset"
' ���W�X�g���L�[ - �����R�[�h�̕ʖ�
Private Const REG_KEY_ALIAS_CHARSET As String = "AliasForCharset"

Public Const NEW_LINE_STR_CRLF As String = "CRLF"
Public Const NEW_LINE_STR_CR As String = "CR"
Public Const NEW_LINE_STR_LF As String = "LF"


Public Function getAppOnKeyNameOfShiftByCode(ByVal shiftCode As String)

    If KEY_CODE_CTRL = shiftCode Then
    
        getAppOnKeyNameOfShiftByCode = "Ctrl"
    ElseIf KEY_CODE_SHIFT = shiftCode Then
    
        getAppOnKeyNameOfShiftByCode = "Shift"
    ElseIf KEY_CODE_ALT = shiftCode Then
    
        getAppOnKeyNameOfShiftByCode = "Alt"
    Else
    
        getAppOnKeyNameOfShiftByCode = ""
    End If

End Function

' =========================================================
' ��Application#OnKey�֐��ɓK�p�\��Key�R�[�h���X�g������������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub initializeAppOnKeyMap()

    If applicationOnKeyMap1 Is Nothing Then
    
        Set applicationOnKeyMap1 = New ValCollection
        applicationOnKeyMap1.setItem "1", "1"
        applicationOnKeyMap1.setItem "2", "2"
        applicationOnKeyMap1.setItem "3", "3"
        applicationOnKeyMap1.setItem "4", "4"
        applicationOnKeyMap1.setItem "5", "5"
        applicationOnKeyMap1.setItem "6", "6"
        applicationOnKeyMap1.setItem "7", "7"
        applicationOnKeyMap1.setItem "8", "8"
        applicationOnKeyMap1.setItem "9", "9"
        applicationOnKeyMap1.setItem "0", "0"
        applicationOnKeyMap1.setItem "a", "a"
        applicationOnKeyMap1.setItem "b", "b"
        applicationOnKeyMap1.setItem "c", "c"
        applicationOnKeyMap1.setItem "d", "d"
        applicationOnKeyMap1.setItem "e", "e"
        applicationOnKeyMap1.setItem "f", "f"
        applicationOnKeyMap1.setItem "g", "g"
        applicationOnKeyMap1.setItem "h", "h"
        applicationOnKeyMap1.setItem "i", "i"
        applicationOnKeyMap1.setItem "j", "j"
        applicationOnKeyMap1.setItem "k", "k"
        applicationOnKeyMap1.setItem "l", "l"
        applicationOnKeyMap1.setItem "m", "m"
        applicationOnKeyMap1.setItem "n", "n"
        applicationOnKeyMap1.setItem "o", "o"
        applicationOnKeyMap1.setItem "p", "p"
        applicationOnKeyMap1.setItem "q", "q"
        applicationOnKeyMap1.setItem "r", "r"
        applicationOnKeyMap1.setItem "s", "s"
        applicationOnKeyMap1.setItem "t", "t"
        applicationOnKeyMap1.setItem "u", "u"
        applicationOnKeyMap1.setItem "v", "v"
        applicationOnKeyMap1.setItem "w", "w"
        applicationOnKeyMap1.setItem "x", "x"
        applicationOnKeyMap1.setItem "y", "y"
        applicationOnKeyMap1.setItem "z", "z"
        applicationOnKeyMap1.setItem "-", "-"
        applicationOnKeyMap1.setItem "{^}", "^"
        applicationOnKeyMap1.setItem "\", "\"
        applicationOnKeyMap1.setItem "@", "@"
        applicationOnKeyMap1.setItem "{[}", "["
        applicationOnKeyMap1.setItem ";", ";"
        applicationOnKeyMap1.setItem ":", ":"
        applicationOnKeyMap1.setItem "{]}", "]"
        applicationOnKeyMap1.setItem ",", ","
        applicationOnKeyMap1.setItem ".", "."
        applicationOnKeyMap1.setItem "/", "/"
        applicationOnKeyMap1.setItem "\", "\"
        applicationOnKeyMap1.setItem "{F1}", "F1"
        applicationOnKeyMap1.setItem "{F2}", "F2"
        applicationOnKeyMap1.setItem "{F3}", "F3"
        applicationOnKeyMap1.setItem "{F4}", "F4"
        applicationOnKeyMap1.setItem "{F5}", "F5"
        applicationOnKeyMap1.setItem "{F6}", "F6"
        applicationOnKeyMap1.setItem "{F7}", "F7"
        applicationOnKeyMap1.setItem "{F8}", "F8"
        applicationOnKeyMap1.setItem "{F9}", "F9"
        applicationOnKeyMap1.setItem "{F10}", "F10"
        applicationOnKeyMap1.setItem "{F11}", "F11"
        applicationOnKeyMap1.setItem "{F12}", "F12"
        applicationOnKeyMap1.setItem "{F13}", "F13"
        applicationOnKeyMap1.setItem "{F14}", "F14"
        applicationOnKeyMap1.setItem "{F15}", "F15"
        applicationOnKeyMap1.setItem "{NUMLOCK}", "Num Lock"
        applicationOnKeyMap1.setItem "{96}", "10key(0)"
        applicationOnKeyMap1.setItem "{97}", "10key(1)"
        applicationOnKeyMap1.setItem "{98}", "10key(2)"
        applicationOnKeyMap1.setItem "{99}", "10key(3)"
        applicationOnKeyMap1.setItem "{100}", "10key(4)"
        applicationOnKeyMap1.setItem "{101}", "10key(5)"
        applicationOnKeyMap1.setItem "{102}", "10key(6)"
        applicationOnKeyMap1.setItem "{103}", "10key(7)"
        applicationOnKeyMap1.setItem "{104}", "10key(8)"
        applicationOnKeyMap1.setItem "{105}", "10key(9)"
        applicationOnKeyMap1.setItem "{106}", "10key(*)"
        applicationOnKeyMap1.setItem "{107}", "10key(+)"
        applicationOnKeyMap1.setItem "{109}", "10key(-)"
        applicationOnKeyMap1.setItem "{110}", "10key(.)"
        applicationOnKeyMap1.setItem "{111}", "10key(/)"
        applicationOnKeyMap1.setItem "{ENTER}", "10key(Enter)"
        applicationOnKeyMap1.setItem "{ESC}", "Esc"
        applicationOnKeyMap1.setItem "{TAB}", "Tab"
        applicationOnKeyMap1.setItem "{CAPSLOCK}", "CapsLock"
        applicationOnKeyMap1.setItem "{BS}", "BackSpace"
        applicationOnKeyMap1.setItem "~", "Enter"
        applicationOnKeyMap1.setItem "{RETURN}", "Return"
        applicationOnKeyMap1.setItem "{SCROLLLOCK}", "ScrollLock"
        applicationOnKeyMap1.setItem "{BREAK}", "Break"
        applicationOnKeyMap1.setItem "{CLEAR}", "Clear"
        applicationOnKeyMap1.setItem "{INSERT}", "Ins"
        applicationOnKeyMap1.setItem "{DEL}", "Delete(Del)"
        applicationOnKeyMap1.setItem "{HOME}", "Home"
        applicationOnKeyMap1.setItem "{END}", "End"
        applicationOnKeyMap1.setItem "{PGUP}", "PageUp"
        applicationOnKeyMap1.setItem "{PGDN}", "PageDown"
        applicationOnKeyMap1.setItem "{HELP}", "Help"
        applicationOnKeyMap1.setItem "{UP}", "��"
        applicationOnKeyMap1.setItem "{DOWN}", "��"
        applicationOnKeyMap1.setItem "{LEFT}", "��"
        applicationOnKeyMap1.setItem "{RIGHT}", "��"

    End If
    
    If applicationOnKeyMap2 Is Nothing Then
    
        Set applicationOnKeyMap2 = New ValCollection
        applicationOnKeyMap2.setItem "1", "1"
        applicationOnKeyMap2.setItem "2", "2"
        applicationOnKeyMap2.setItem "3", "3"
        applicationOnKeyMap2.setItem "4", "4"
        applicationOnKeyMap2.setItem "5", "5"
        applicationOnKeyMap2.setItem "6", "6"
        applicationOnKeyMap2.setItem "7", "7"
        applicationOnKeyMap2.setItem "8", "8"
        applicationOnKeyMap2.setItem "9", "9"
        applicationOnKeyMap2.setItem "0", "0"
        applicationOnKeyMap2.setItem "a", "a"
        applicationOnKeyMap2.setItem "b", "b"
        applicationOnKeyMap2.setItem "c", "c"
        applicationOnKeyMap2.setItem "d", "d"
        applicationOnKeyMap2.setItem "e", "e"
        applicationOnKeyMap2.setItem "f", "f"
        applicationOnKeyMap2.setItem "g", "g"
        applicationOnKeyMap2.setItem "h", "h"
        applicationOnKeyMap2.setItem "i", "i"
        applicationOnKeyMap2.setItem "j", "j"
        applicationOnKeyMap2.setItem "k", "k"
        applicationOnKeyMap2.setItem "l", "l"
        applicationOnKeyMap2.setItem "m", "m"
        applicationOnKeyMap2.setItem "n", "n"
        applicationOnKeyMap2.setItem "o", "o"
        applicationOnKeyMap2.setItem "p", "p"
        applicationOnKeyMap2.setItem "q", "q"
        applicationOnKeyMap2.setItem "r", "r"
        applicationOnKeyMap2.setItem "s", "s"
        applicationOnKeyMap2.setItem "t", "t"
        applicationOnKeyMap2.setItem "u", "u"
        applicationOnKeyMap2.setItem "v", "v"
        applicationOnKeyMap2.setItem "w", "w"
        applicationOnKeyMap2.setItem "x", "x"
        applicationOnKeyMap2.setItem "y", "y"
        applicationOnKeyMap2.setItem "z", "z"
        applicationOnKeyMap2.setItem "-", "-"
        applicationOnKeyMap2.setItem "^", "{^}"
        applicationOnKeyMap2.setItem "\", "\"
        applicationOnKeyMap2.setItem "@", "@"
        applicationOnKeyMap2.setItem "[", "{[}"
        applicationOnKeyMap2.setItem ";", ";"
        applicationOnKeyMap2.setItem ":", ":"
        applicationOnKeyMap2.setItem "]", "{]}"
        applicationOnKeyMap2.setItem ",", ","
        applicationOnKeyMap2.setItem ".", "."
        applicationOnKeyMap2.setItem "/", "/"
        applicationOnKeyMap2.setItem "\", "\"
        applicationOnKeyMap2.setItem "F1", "{F1}"
        applicationOnKeyMap2.setItem "F2", "{F2}"
        applicationOnKeyMap2.setItem "F3", "{F3}"
        applicationOnKeyMap2.setItem "F4", "{F4}"
        applicationOnKeyMap2.setItem "F5", "{F5}"
        applicationOnKeyMap2.setItem "F6", "{F6}"
        applicationOnKeyMap2.setItem "F7", "{F7}"
        applicationOnKeyMap2.setItem "F8", "{F8}"
        applicationOnKeyMap2.setItem "F9", "{F9}"
        applicationOnKeyMap2.setItem "F10", "{F10}"
        applicationOnKeyMap2.setItem "F11", "{F11}"
        applicationOnKeyMap2.setItem "F12", "{F12}"
        applicationOnKeyMap2.setItem "F13", "{F13}"
        applicationOnKeyMap2.setItem "F14", "{F14}"
        applicationOnKeyMap2.setItem "F15", "{F15}"
        applicationOnKeyMap2.setItem "Num Lock", "{NUMLOCK}"
        applicationOnKeyMap2.setItem "10key(0)", "{96}"
        applicationOnKeyMap2.setItem "10key(1)", "{97}"
        applicationOnKeyMap2.setItem "10key(2)", "{98}"
        applicationOnKeyMap2.setItem "10key(3)", "{99}"
        applicationOnKeyMap2.setItem "10key(4)", "{100}"
        applicationOnKeyMap2.setItem "10key(5)", "{101}"
        applicationOnKeyMap2.setItem "10key(6)", "{102}"
        applicationOnKeyMap2.setItem "10key(7)", "{103}"
        applicationOnKeyMap2.setItem "10key(8)", "{104}"
        applicationOnKeyMap2.setItem "10key(9)", "{105}"
        applicationOnKeyMap2.setItem "10key(*)", "{106}"
        applicationOnKeyMap2.setItem "10key(+)", "{107}"
        applicationOnKeyMap2.setItem "10key(-)", "{109}"
        applicationOnKeyMap2.setItem "10key(.)", "{110}"
        applicationOnKeyMap2.setItem "10key(/)", "{111}"
        applicationOnKeyMap2.setItem "10key(Enter)", "{ENTER}"
        applicationOnKeyMap2.setItem "Esc", "{ESC}"
        applicationOnKeyMap2.setItem "Tab", "{TAB}"
        applicationOnKeyMap2.setItem "CapsLock", "{CAPSLOCK}"
        applicationOnKeyMap2.setItem "BackSpace", "{BS}"
        applicationOnKeyMap2.setItem "Enter", "~"
        applicationOnKeyMap2.setItem "Return", "{RETURN}"
        applicationOnKeyMap2.setItem "ScrollLock", "{SCROLLLOCK}"
        applicationOnKeyMap2.setItem "Break", "{BREAK}"
        applicationOnKeyMap2.setItem "Clear", "{CLEAR}"
        applicationOnKeyMap2.setItem "Ins", "{INSERT}"
        applicationOnKeyMap2.setItem "Delete(Del)", "{DEL}"
        applicationOnKeyMap2.setItem "Home", "{HOME}"
        applicationOnKeyMap2.setItem "End", "{END}"
        applicationOnKeyMap2.setItem "PageUp", "{PGUP}"
        applicationOnKeyMap2.setItem "PageDown", "{PGDN}"
        applicationOnKeyMap2.setItem "Help", "{HELP}"
        applicationOnKeyMap2.setItem "��", "{UP}"
        applicationOnKeyMap2.setItem "��", "{DOWN}"
        applicationOnKeyMap2.setItem "��", "{LEFT}"
        applicationOnKeyMap2.setItem "��", "{RIGHT}"
        
    End If

End Sub

' =========================================================
' ��Application#OnKey�֐��ɓK�p�\��Key�R�[�h���X�g���擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�FKey�R�[�h���X�g
'
' =========================================================
Public Function getAppOnKeyCodeList() As ValCollection

    initializeAppOnKeyMap
    Set getAppOnKeyCodeList = applicationOnKeyMap2
End Function

' =========================================================
' ��Application#OnKey�֐���Key�R�[�h���擾
'
' �T�v�@�@�@�F�_�������L�[�ɂ���Key�R�[�h���擾����B
' �����@�@�@�Fname �_����
' �߂�l�@�@�FKey�R�[�h
'
' =========================================================
Public Function getAppOnKeyCodeByName(ByVal name As String) As String

    initializeAppOnKeyMap
    getAppOnKeyCodeByName = applicationOnKeyMap1.getItem(name, vbString)
    
End Function

' =========================================================
' ��Application#OnKey�֐���Key�R�[�h�ɕR�Â��_�������擾
'
' �T�v�@�@�@�FKey�R�[�h�ɕR�Â��_�������擾����B
' �����@�@�@�Fcode Key�R�[�h
' �߂�l�@�@�F�_����
'
' =========================================================
Public Function getAppOnKeyNameByCode(ByVal code As String) As String

    initializeAppOnKeyMap
    getAppOnKeyNameByCode = applicationOnKeyMap2.getItem(code, vbString)

End Function

' =========================================================
' ��Application#OnKey�֐��̃L�[�l�����
'
' �T�v�@�@�@�FApplication#OnKey�֐��̃L�[�l����͂�
' �@�@�@�@�@�@�߂�l�p�̈����ɕԂ��B
' �����@�@�@�FkeyCode    �L�[�R�[�h
' �@�@�@�@�@�@shiftCtrl  Ctrl�L�[
' �@�@�@�@�@�@shiftShift Shift�L�[
' �@�@�@�@�@�@shiftAlt   Alt�L�[
' �@�@�@�@�@�@keyName    �L�[�l
' �߂�l�@�@�F
'
' =========================================================
Public Function resolveAppOnKey(ByVal KeyCode As String _
                                      , ByRef shiftCtrl As Boolean _
                                      , ByRef shiftShift As Boolean _
                                      , ByRef shiftAlt As Boolean _
                                      , ByRef keyName As String)

    initializeAppOnKeyMap
    
    ' ������C���f�b�N�X
    Dim i      As Long
    ' �����񒷂�
    Dim length As Long
    ' �����񂩂璊�o����1����
    Dim char   As String
    
    ' �߂�l�p�̈���������������
    shiftCtrl = False
    shiftShift = False
    shiftAlt = False
    keyName = ""
    
    ' keyCode�̒������擾����
    length = Len(KeyCode)
    
    For i = 1 To length
    
        ' 1�������o����
        char = Mid$(KeyCode, i, 1)
        
        ' Ctrl�L�[
        If char = KEY_CODE_CTRL Then
        
            shiftCtrl = True
            
        ' Shift�L�[
        ElseIf char = KEY_CODE_SHIFT Then
        
            shiftShift = True
            
        ' Alt�L�[
        ElseIf char = KEY_CODE_ALT Then
        
            shiftAlt = True
            
        ' ���̑��̃L�[
        Else
        
            keyName = getAppOnKeyNameByCode(Mid$(KeyCode, i, length))
            Exit For
        End If
        
    Next

End Function

' =========================================================
' ��Application#OnKey�֐��ɗ^����L�[�R�[�h�̎擾
'
' �T�v�@�@�@�F����̃p�����[�^����Application#OnKey�֐��̃L�[�R�[�h���擾����B
' �����@�@�@�FshiftCtrl  Ctrl�L�[
' �@�@�@�@�@�@shiftShift Shift�L�[
' �@�@�@�@�@�@shiftAlt   Alt�L�[
' �@�@�@�@�@�@name       �L�[�̘_����
' �߂�l�@�@�F�L�[�R�[�h
'
' =========================================================
Public Function getAppOnKeyCodeBySomeParams(ByVal shiftCtrl As Boolean _
                                                  , ByVal shiftShift As Boolean _
                                                  , ByVal shiftAlt As Boolean _
                                                  , ByVal name As String) As String

    initializeAppOnKeyMap
    
    ' �߂�l
    Dim ret As String
    
    ' Ctrl�L�[
    If shiftCtrl = True Then
    
        ret = ret & KEY_CODE_CTRL
    End If
        
    ' Shift�L�[
    If shiftShift = True Then
    
        ret = ret & KEY_CODE_SHIFT
    End If
        
    ' Alt�L�[
    If shiftAlt = True Then
    
        ret = ret & KEY_CODE_ALT
    End If

    ' �L�[���擾����
    ret = ret & getAppOnKeyCodeByName(name)

    ' �߂�l��ݒ肷��
    getAppOnKeyCodeBySomeParams = ret

End Function

' =========================================================
' ��Application#OnKey�֐��ɗ^����L�[�R�[�h�̎擾
'
' �T�v�@�@�@�F����̃p�����[�^����Application#OnKey�֐��̃L�[�R�[�h���擾����B
' �����@�@�@�FshiftCtrl  Ctrl�L�[
' �@�@�@�@�@�@shiftShift Shift�L�[
' �@�@�@�@�@�@shiftAlt   Alt�L�[
' �@�@�@�@�@�@name       �L�[�̘_����
' �߂�l�@�@�F�L�[�R�[�h
'
' =========================================================
Public Function getAppOnKeyNameBySomeParams(ByVal shiftCtrl As Boolean _
                                                  , ByVal shiftShift As Boolean _
                                                  , ByVal shiftAlt As Boolean _
                                                  , ByVal name As String) As String

    initializeAppOnKeyMap
    
    ' �߂�l
    Dim ret As String
    ' ����������
    Dim juncStr As String
    
    ' Ctrl�L�[
    If shiftCtrl = True Then
    
        ret = ret & getAppOnKeyNameOfShiftByCode(KEY_CODE_CTRL)
    End If
        
    ' Shift�L�[
    If shiftShift = True Then
    
        If ret <> "" Then
            juncStr = "+"
        Else
            juncStr = ""
        End If
        
        ret = ret & juncStr & getAppOnKeyNameOfShiftByCode(KEY_CODE_SHIFT)
    End If
        
    ' Alt�L�[
    If shiftAlt = True Then
    
        If ret <> "" Then
            juncStr = "+"
        Else
            juncStr = ""
        End If
            
        ret = ret & juncStr & getAppOnKeyNameOfShiftByCode(KEY_CODE_ALT)
    End If

    If ret <> "" Then
        juncStr = "+"
    Else
        juncStr = ""
    End If
    
    ' �L�[���擾����
    ret = ret & juncStr & name

    ' �߂�l��ݒ肷��
    getAppOnKeyNameBySomeParams = ret

End Function

Public Function getAppOnKeyNameByMultipleCode(ByVal KeyCode As String) As String

    Dim a As Boolean
    Dim b As Boolean
    Dim c As Boolean
    Dim d As String

    resolveAppOnKey KeyCode _
                            , a _
                            , b _
                            , c _
                            , d
                            
    getAppOnKeyNameByMultipleCode = getAppOnKeyNameBySomeParams(a _
                                                                              , b _
                                                                              , c _
                                                                              , d)
    

End Function

' =========================================================
' ��Err�I�u�W�F�N�g�̏����\���̂ɑޔ�
'
' �T�v�@�@�@�FErr�I�u�W�F�N�g�̏����\���̂ɐݒ肵�ĕԂ��B
' �����@�@�@�F
' �߂�l�@�@�F�G���[���
'
' ���L�����@�F�G���[�n���h���ŕʂ̊֐����Ăяo����Err�I�u�W�F�N�g�̏�񂪏����Ă��܂����Ƃ�����
' �@�@�@�@�@�@���̏�ԂŁAErr.Raise����Ɛ�����������ʂ̃��W���[���ɂœ`�d�ł��Ȃ��B
' �@�@�@�@�@�@����������`�d����ꍇ�ɂ́A�{�֐��𗘗p���āA��x�G���[����ޔ����Ă���Err.Raise���Ă��Ɨǂ��B
'
' �@�@�@�@�@�@�g�p��F
' �@�@�@�@�@�@�@Dim errT As errInfo
' �@�@�@�@�@�@�@errT = VBUtil.swapErr

' �@�@�@�@�@�@�@�E�E�E�G���[���̌�n�������Ȃ�
'
' �@�@�@�@�@�@�@Err.Raise errT.Number, errT.Source�E�E�E
'
' =========================================================
Public Function swapErr() As errInfo

    swapErr.Source = err.Source
    swapErr.Number = err.Number
    swapErr.Description = err.Description
    swapErr.LastDllError = err.LastDllError
    swapErr.HelpFile = err.HelpFile
    swapErr.HelpContext = err.HelpContext

End Function

' =========================================================
' ��Err�I�u�W�F�N�g�ɏ���ݒ肷��
'
' �T�v�@�@�@�F
' �����@�@�@�FerrT �G���[���
' �߂�l�@�@�F
'
' ���L�����@�F
'
' =========================================================
Public Sub setErr(ByRef errT As errInfo)

    err.Source = errT.Source
    err.Number = errT.Number
    err.Description = errT.Description
    'err.LastDllError = errT.LastDllError
    err.HelpFile = errT.HelpFile
    err.HelpContext = errT.HelpContext

End Sub

' =========================================================
' ���ۑ��_�C�A���O�\��
'
' �T�v�@�@�@�F�ۑ��_�C�A���O��\������
' �����@�@�@�Ftitle           �_�C�A���O�̃^�C�g��
' �@�@�@�@�@�@filter          �t�B���^
' �@�@�@�@�@�@initialFileName �����t�@�C����
' �߂�l�@�@�F�ۑ��t�@�C���p�X
'
' =========================================================
Public Function openFileSaveDialog(ByVal title As String, ByVal filter As String, ByVal initialFileName As String) As String

    ' �A�v���P�[�V����
    Dim xlsApp   As Application
    
    ' �t�@�C���p�X
    Dim filePath As Variant

    ' Application�I�u�W�F�N�g�擾
    Set xlsApp = Application
    
    ' �_�C�A���O�őI�����ꂽ�t�@�C�������i�[
    filePath = xlsApp.GetSaveAsFilename(initialFileName:=initialFileName _
                                      , fileFilter:=filter _
                                      , title:=title)
                                      
    ' �L�����Z�����ꂽ���𔻒肷��
    If filePath = False Then
    
        ' �L�����Z�����ꂽ�ꍇ �󕶎����Ԃ�
        openFileSaveDialog = ""
        
    Else
        ' �ۑ���I�����ꂽ�ꍇ �t�@�C������Ԃ�
        openFileSaveDialog = filePath
    End If

End Function

' =========================================================
' ���t�H���_���J���_�C�A���O�\��
'
' �T�v�@�@�@�F�t�H���̊J���_�C�A���O��\������
' �����@�@�@�Ftitle           �_�C�A���O�̃^�C�g��
'     �@�@�@  initialFileName �����t�@�C����
' �߂�l�@�@�F�I�������t�H���̃t�@�C���p�X
'
' =========================================================
Public Function openFolderDialog(ByVal title As String, ByVal initialFileName As String) As Variant

    ' �A�v���P�[�V����
    Dim xlsApp   As Application
    
    ' �t�@�C���p�X
    Dim fileDialogObj As FileDialog
    ' �t�@�C���p�X
    Dim filePath As Variant

    ' Application�I�u�W�F�N�g�擾
    Set xlsApp = Application
    
    ' �_�C�A���O�őI�����ꂽ�t�@�C�������i�[
    Set fileDialogObj = xlsApp.FileDialog(msoFileDialogFolderPicker)
    fileDialogObj.title = title
    fileDialogObj.initialFileName = initialFileName
    fileDialogObj.Show

    If fileDialogObj.SelectedItems.count <= 0 Then
    
        ' �L�����Z�����ꂽ�ꍇ ���Ԃ�
        openFolderDialog = Empty
    Else
    
        ' �P��I�������ꍇ
        openFolderDialog = fileDialogObj.SelectedItems(1)
    
    End If

End Function

' =========================================================
' ���J���_�C�A���O�\��
'
' �T�v�@�@�@�F�J���_�C�A���O��\������
' �����@�@�@�Ftitle           �_�C�A���O�̃^�C�g��
' �@�@�@�@�@�@filter          �t�B���^
' �@�@�@�@�@�@multiSelect     �����I��
' �߂�l�@�@�F�I�������t�@�C���̃t�@�C���p�X
'
' =========================================================
Public Function openFileDialog(ByVal title As String, ByVal filter As String, Optional ByVal multiSelect As Boolean = False) As Variant

    ' �A�v���P�[�V����
    Dim xlsApp   As Application
    
    ' �t�@�C���p�X
    Dim filePath As Variant

    ' Application�I�u�W�F�N�g�擾
    Set xlsApp = Application
    
    ' �_�C�A���O�őI�����ꂽ�t�@�C�������i�[
    filePath = xlsApp.GetOpenFilename(fileFilter:=filter _
                                    , title:=title _
                                    , multiSelect:=multiSelect)

    ' �����I���̏ꍇ�A�߂�l�Ƃ��Ĕz�񂪕Ԃ����̂Ŕz�񂩂ǂ����𔻒肷��
    If IsArray(filePath) Then
    
        ' �ۑ���I�����ꂽ�ꍇ �t�@�C������Ԃ�
        openFileDialog = filePath
    
    ' �I�����L�����Z�����ꂽ�ꍇ
    ElseIf filePath = False Then
    
        ' �L�����Z�����ꂽ�ꍇ ���Ԃ�
        openFileDialog = Empty
        
    Else
        ' �ۑ���I�����ꂽ�ꍇ �t�@�C������Ԃ�
        openFileDialog = filePath
    
    End If

End Function

' =========================================================
' ���t�@�C���̊g���q�`�F�b�N
'
' �T�v�@�@�@�F�t�@�C���̊g���q���`�F�b�N����
' �����@�@�@�Ffile      �t�@�C����
' �@�@�@�@�@�@extension �g���q
' �߂�l�@�@�F�t�@�C���̊g���q���w�肳�ꂽ����extension�̏ꍇTrue��Ԃ�
'
' =========================================================
Public Function checkFileExtension(ByRef file As String _
                                 , ByRef extension As String) As Boolean

    ' �t�@�C�������璊�o�����g���q
    Dim fileExtension As String
    
    ' �C���f�b�N�X
    Dim index As Long
    
    ' �t�@�C�����Ɗg���q�̋�؂蕶���ł���h�b�g(.)����������
    index = InStrRev(file, ".")
    
    ' �h�b�g(.)��������Ȃ��ꍇ
    If index <= 0 Then
    
        Exit Function
    End If
    
    ' �t�@�C��������g���q�𒊏o����
    fileExtension = Mid$(file, index + 1, Len(file))

    If fileExtension = extension Then
    
        checkFileExtension = True
    Else
    
        checkFileExtension = False
    End If

End Function

' =========================================================
' ���t�@�C���p�X����t�@�C�������o
'
' �T�v�@�@�@�F�t�@�C���p�X����t�@�C�����𒊏o����
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�F�t�@�C����
'
' =========================================================
Public Function extractFileName(ByRef filePath As String) As String
    
    ' �t�@�C���p�X��؂蕶��
    Const FILE_SEPARATE As String = "\"

    ' �t�@�C���p�X�̉E�������͂��߂ɏo��������؂蕶���̕����ʒu
    Dim index As Long
    
    ' ��؂蕶���̈ʒu���擾����
    index = InStrRev(filePath, FILE_SEPARATE)

    ' ��؂蕶���𔭌������ꍇ
    If index > 0 Then
    
        extractFileName = Mid$(filePath, index + 1)
    
    ' ��؂蕶���𔭌��ł��Ȃ������ꍇ
    Else
        extractFileName = filePath
    
    End If

End Function

' =========================================================
' ���C���t�H���b�Z�[�W�{�b�N�X��\��
'
' �T�v�@�@�@�F�C���t�H���b�Z�[�W�{�b�N�X��\������
' �����@�@�@�FbasePrompt ��{���b�Z�[�W
'             title      ���b�Z�[�W�{�b�N�X�̃^�C�g��
' �@�@�@�@�@�@err        �G���[�I�u�W�F�N�g
'
' =========================================================
Public Sub showMessageBoxForInformation(ByRef basePrompt As String _
                                      , ByRef title As String _
                                      , Optional ByRef err As ErrObject = Nothing)

    MsgBox basePrompt, vbOKOnly, title
         
End Sub

' =========================================================
' ���G���[���b�Z�[�W�{�b�N�X��\��
'
' �T�v�@�@�@�F�G���[���b�Z�[�W�{�b�N�X��\������
' �����@�@�@�FbasePrompt ��{���b�Z�[�W
'             title      ���b�Z�[�W�{�b�N�X�̃^�C�g��
' �@�@�@�@�@�@err        �G���[�I�u�W�F�N�g
'
' =========================================================
Public Sub showMessageBoxForError(ByRef basePrompt As String _
                                , ByRef title As String _
                                , ByRef err As ErrObject)

    MsgBox basePrompt & vbNewLine & vbNewLine & _
           err.Description & vbNewLine & _
           "Error no [" & err.Number & "]" & vbNewLine & _
           "Source [" & err.Source & "]" _
           , vbOKOnly + vbCritical _
           , title

End Sub

' =========================================================
' ���x�����b�Z�[�W�{�b�N�X��\��
'
' �T�v�@�@�@�F�x�����b�Z�[�W�{�b�N�X��\������
' �����@�@�@�FbasePrompt ��{���b�Z�[�W
'             title      ���b�Z�[�W�{�b�N�X�̃^�C�g��
' �@�@�@�@�@�@err        �G���[�I�u�W�F�N�g
'
' =========================================================
Public Sub showMessageBoxForWarning(ByVal basePrompt As String _
                                  , ByVal title As String _
                                  , ByRef err As ErrObject)

    If err Is Nothing Then
    
        MsgBox basePrompt _
               , vbOKOnly + vbExclamation _
               , title
    
    ElseIf err.Number = 0 Then
    
        MsgBox basePrompt _
               , vbOKOnly + vbExclamation _
               , title
               
    Else
    
        If basePrompt <> "" Then
        
            basePrompt = basePrompt & vbNewLine & vbNewLine
        End If
        
        MsgBox basePrompt & _
               err.Description & vbNewLine & _
               "Error no [" & err.Number & "]" _
               , vbOKOnly + vbExclamation _
               , title
    End If
         
End Sub

' =========================================================
' ���͂��E�������E�L�����Z�����b�Z�[�W�{�b�N�X��\��
'
' �T�v�@�@�@�F�͂��E�������E�L�����Z�����b�Z�[�W�{�b�N�X��\������
' �����@�@�@�FbasePrompt ��{���b�Z�[�W
'             title      ���b�Z�[�W�{�b�N�X�̃^�C�g��
'
' =========================================================
Public Function showMessageBoxForYesNoCancel(ByRef basePrompt As String _
                                , ByRef title As String) As Long
    
    showMessageBoxForYesNoCancel = MsgBox(basePrompt _
           , vbYesNoCancel + vbDefaultButton2 _
           , title)

End Function

' =========================================================
' ���͂��E���������b�Z�[�W�{�b�N�X��\��
'
' �T�v�@�@�@�F�͂��E���������b�Z�[�W�{�b�N�X��\������
' �����@�@�@�FbasePrompt ��{���b�Z�[�W
'             title      ���b�Z�[�W�{�b�N�X�̃^�C�g��
'
' =========================================================
Public Function showMessageBoxForYesNo(ByRef basePrompt As String _
                                , ByRef title As String) As Long
    
    showMessageBoxForYesNo = MsgBox(basePrompt _
           , vbYesNo + vbDefaultButton2 _
           , title)

End Function

' =========================================================
' ��INI�t�@�C���p�X�擾
'
' �T�v�@�@�@�F�A�v���P�[�V������INI�t�@�C���p�X���擾����
' �����@�@�@�FfileName �t�@�C����
' �߂�l�@�@�FINI�t�@�C���p�X
'
' =========================================================
Public Function getApplicationIniFilePath(Optional ByVal fileName As String = "") As String

    ' ini�t�@�C���̃p�X���擾����
    getApplicationIniFilePath = ThisWorkbook.path & "\resource\config\" & fileName
    
End Function

' =========================================================
' �����W�X�g���p�X�擾
'
' �T�v�@�@�@�F�A�v���P�[�V�����̃��W�X�g���p�X���擾����
' �@�@�@�@�@�@���[�g�L�[�́AHKEY_CURRENT_USER
' �����@�@�@�FcompanyName ��Ж�
' �@�@�@�@�@�@appName     �A�v���P�[�V������
' �@�@�@�@�@�@suffix      ���W�X�g���p�X�̐ڔ���
' �߂�l�@�@�FINI�t�@�C���p�X
'
' =========================================================
Public Function getApplicationRegistryPath(ByVal companyName As String _
                                         , Optional ByVal suffix As String = "" _
                                         , Optional ByVal appName As String = "") As String

    ' �A�v���P�[�V���������ݒ肳��Ă��Ȃ��ꍇ
    If appName = "" Then
    
        ' �v���W�F�N�g����ݒ肷��
        appName = ConstantsCommon.APPLICATION_NAME
    End If

    ' ini�t�@�C���̃p�X���擾����
    ' �{�u�b�N�̃p�X�{�v���W�F�N�g���{".ini"
    getApplicationRegistryPath = "Software\" & companyName & "\" & appName
    
    If suffix <> "" Then
    
        getApplicationRegistryPath = getApplicationRegistryPath & "\" & suffix
    End If
    
End Function

' =========================================================
' ���z��T�C�Y�擾
'
' �T�v�@�@�@�F�z��̃T�C�Y���擾����
' �����@�@�@�Fvar       �z��
' �@�@�@�@�@�@dimension ����
'
' =========================================================
Public Function arraySize(ByRef var As Variant, Optional ByVal dimension As Long = 1) As Long

    If IsArray(var) = True Then
    
        arraySize = UBound(var, dimension) - LBound(var, dimension) + 1
        
    Else
        arraySize = 0
    
    End If
    

End Function

' =========================================================
' ��2�����z��̔C�ӂ̍s��1�����z��Ƃ��ĕԂ�
'
' �T�v�@�@�@�F
' �����@�@�@�Fval �z��
'             i   �z��̃C���f�b�N�X
'
' =========================================================
Public Function convert2to1Array(ByRef val As Variant, ByVal i As Long) As Variant

    ' �߂�l
    Dim ret() As Variant

    Dim j As Long
    
    ReDim ret(LBound(val, 2) To UBound(val, 2))
    
    For j = LBound(ret) To UBound(ret)
    
        ret(j) = val(i, j)
    
    Next
    
    convert2to1Array = ret

End Function

' =========================================================
' ��2�����z����f�o�b�O�E�B���h�E�ɏo�͂���
'
' �T�v�@�@�@�F
' �����@�@�@�Fval �z��
'
' =========================================================
Public Function debugPrintArray(ByRef val As Variant)

    ' �z��̃C���f�b�N�X
    Dim i As Long
    Dim j As Long
    
    ' �f�o�b�O�E�B���h�E�ɏo�͂��镶����
    Dim str As String
    
    str = "Output Array" & vbNewLine
    
    ' -------------------------------------------------
    ' �z��Ƃ��ď���������Ă���ꍇ�ɏo�͂����{����
    ' -------------------------------------------------
    If VarType(val) = (vbArray + vbVariant) Then
    
        ' ���[�v����
        For i = LBound(val, 1) To UBound(val, 1)
        
            str = str & "+   [" & i & "] - {"
        
            For j = LBound(val, 2) To UBound(val, 2)
            
                str = str & val(i, j) & ", "
            Next
            
            str = str & "}" & vbNewLine
            
        Next
        
    Else
        str = str & "   ... Empty"
        
    End If
    ' -------------------------------------------------
    
    Debug.Print str
    
End Function

' =========================================================
' ��2�����z��̗v�f����ւ�
'
' �T�v�@�@�@�F2�����z��̗v�f��(x,y)����(y,x)�ɐݒ肵�Ȃ����B
' �����@�@�@�Fv 2�����z��
'
' �߂�l�@�@�F2�����z��
'
' =========================================================
Public Function transposeDim(ByRef v As Variant) As Variant
    
    Dim X As Long
    Dim Y As Long
    
    Dim Xlower As Long
    Dim Xupper As Long
    
    Dim Ylower As Long
    Dim Yupper As Long
    
    Dim tempArray As Variant
    
    Xlower = LBound(v, 2)
    Xupper = UBound(v, 2)
    Ylower = LBound(v, 1)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xlower To Xupper, Ylower To Yupper)
    
    For X = Xlower To Xupper
        For Y = Ylower To Yupper
        
            tempArray(X, Y) = v(Y, X)
        
        Next Y
    Next X
    
    transposeDim = tempArray

End Function

' =========================================================
' �������`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue ����
'
' =========================================================
Public Function validInteger(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = False

    ' �`�F�b�N�Ώۂ����l�Ŋ��A�����_���܂܂Ȃ��ꍇ�AOK�Ƃ���
    If _
            IsNumeric(value) = True _
        And InStr(value, ".") = 0 Then
    
        ret = True
    
    End If

    ' �߂�l��Ԃ�
    validInteger = ret

End Function

' =========================================================
' �������`�F�b�N�i�����͊܂܂Ȃ��j
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue ����
'
' =========================================================
Public Function validUnsignedInteger(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = False

    ' �`�F�b�N�Ώۂ����l�Ŋ��A�}�C�i�X�L�����܂܂������_���܂܂Ȃ��ꍇ�AOK�Ƃ���
    If _
            IsNumeric(value) = True _
        And InStr(value, ".") = 0 _
        And InStr(value, "-") = 0 _
    Then
    
        ret = True
    
    End If

    ' �߂�l��Ԃ�
    validUnsignedInteger = ret

End Function

' =========================================================
' ��16�i���`�F�b�N
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue 16�i��
'
' =========================================================
Public Function validHex(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = True

    ' �C���f�b�N�X
    Dim i    As Long
    ' �����̃T�C�Y
    Dim size As Long
    
    ' �������1������
    Dim one    As String
    ' 1��������ASCII�R�[�h
    Dim oneAsc As Long
    
    ' �����̃T�C�Y���擾����
    size = Len(value)
    
    ' �����񂩂�1���������o�����[�v�����s����
    For i = 1 To size
    
        ' 1�������o��
        one = Mid$(value, i, 1)
        ' ���o����������ASCII�R�[�h�𒲂ׂ�
        oneAsc = Asc(one)
        
        ' �����񂪈ȉ��͈͓̔��ł��邩���m�F����
        ' 0-9 a-f A-F
        If _
             (65 <= oneAsc And oneAsc <= 70) _
          Or (97 <= oneAsc And oneAsc <= 102) _
          Or (48 <= oneAsc And oneAsc <= 57) Then
        
            ' ����
            
        Else
        
            ' �G���[��
            ret = False
            Exit For
        
        End If
        
    Next

    ' �߂�l��Ԃ�
    validHex = ret

End Function

' =========================================================
' �����l�ł��邩���`�F�b�N����
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue ����
'
' =========================================================
Public Function validNumeric(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = False

    ' �`�F�b�N�Ώۂ����l�̏ꍇ�AOK�Ƃ���
    If _
            IsNumeric(value) = True Then
    
        ret = True
    
    End If

    ' �߂�l��Ԃ�
    validNumeric = ret

End Function

' =========================================================
' �����l�ł��邩���`�F�b�N����i�����͊܂܂Ȃ��j
'
' �T�v�@�@�@�F
' �����@�@�@�Fvalue �`�F�b�N������
' �߂�l�@�@�FTrue ����
'
' =========================================================
Public Function validUnsignedNumeric(ByVal value As String) As Boolean

    ' �߂�l
    Dim ret As Boolean: ret = False

    ' �`�F�b�N�Ώۂ����l�Ŋ��}�C�i�X�L�����܂܂Ȃ��ꍇ�AOK�Ƃ���
    If _
            IsNumeric(value) = True _
        And InStr(value, "-") = 0 _
    Then
    
        ret = True
    
    End If

    ' �߂�l��Ԃ�
    validUnsignedNumeric = ret

End Function

' =========================================================
' ���R�[�h�l�`�F�b�N
'
' �T�v�@�@�@�F�����ŗ^����ꂽ�R�[�h���X�g�Ɉ�v������̂����邩���`�F�b�N����B
' �����@�@�@�Fvalue    �`�F�b�N������
' �@�@�@�@�@�@codeList �R�[�h���X�g
' �߂�l�@�@�FTrue �R�[�h���X�g�Ɉ�v����l������
'
' =========================================================
Public Function validCode(ByVal value As String, ParamArray codeList() As Variant) As Boolean

    ' �`�F�b�N�Ώۂ���̏ꍇ�AOK�Ƃ���
    Dim i As Long
    
    ' value��enums�̉��ꂩ�̒l�ƈ�v���Ă��邩�ǂ������m�F����
    For i = LBound(codeList) To UBound(codeList)
    
        ' ��v���Ă���ꍇ
        If value = CStr(codeList(i)) Then
        
            ' True��Ԃ�
            validCode = True
            
            Exit Function
        End If
    
    Next
    
    ' ��v������̂��Ȃ������̂ŁAFalse��Ԃ�
    validCode = False

End Function

' =========================================================
' ��RGB���]
'
' �T�v�@�@�@�FRGB�𔽓]������B
' �����@�@�@�Fr ��
' �@�@�@�@�@�@g ��
' �@�@�@�@�@�@b ��
' �߂�l�@�@�F���]�F
'
' =========================================================
Public Function reverseRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long

    reverseRGB = (Not RGB(r, g, b)) And &HFFFFFF

End Function

' =========================================================
' ��NULL���󕶎���ϊ�
'
' �T�v�@�@�@�FNull���󕶎���ɕϊ�����B
' �����@�@�@�Fvalue VARIANT�f�[�^
' �߂�l�@�@�F�󕶎���
' ���L�����@�FNull �l�́A�f�[�^ �A�C�e�� �ɗL���ȃf�[�^��
' �@�@�@�@�@�@�i�[����Ă��Ȃ����Ƃ������̂Ɏg�p�����o���A���g�^ (Variant) �̓��������`���ł��B
'
' =========================================================
Public Function convertNullToEmptyStr(ByRef value As Variant) As String

    ' NULL�̏ꍇ
    If isNull(value) = True Then
    
        ' �󕶎���ɕϊ�
        convertNullToEmptyStr = ""
        
    ' �z��̏ꍇ
    ElseIf IsArray(value) Then
    
        ' �󕶎���ɕϊ�
        convertNullToEmptyStr = ""
        
    ' ���̑�
    Else
    
        ' ������ɕϊ����Ċi�[����
        convertNullToEmptyStr = CStr(value)
    End If
    
End Function

' =========================================================
' ���N�C�b�N�\�[�g
'
' �T�v�@�@�@�F�N�C�b�N�\�[�g���s���B�z��ϐ��̗v�f��Long�^��O��Ƃ���B
' �����@�@�@�Fa �z��
' �߂�l�@�@�F
'
' =========================================================
Public Sub quickSort(ByRef a As Variant)

    quickSortSub a, LBound(a), UBound(a)
    
End Sub

' =========================================================
' ���N�C�b�N�\�[�g
'
' �T�v�@�@�@�F�N�C�b�N�\�[�g���s���B�z��ϐ��̗v�f��Long�^��O��Ƃ���B
' �����@�@�@�Fa     �z��
' �@�@�@�@�@�@left  ���ʒu
' �@�@�@�@�@�@right �E�ʒu
' �߂�l�@�@�F
'
' =========================================================
Private Sub quickSortSub(ByRef a As Variant _
                       , ByVal Left As Long _
                       , ByVal right As Long)

    ' �X�^�b�N�I�u�W�F�N�g
    Dim stack As New ValStack
    
    ' �X�^�b�N�Ɋi�[����l
    ' �i�z��𑖍�������A���[�ƉE�[�̃C���f�b�N�X���i�[����j
    Dim stackVal As Variant
    ' �z��ϐ��𐶐�����
    ReDim stackVal(1 To 2)
    
    
    ' �x�[�X�ƂȂ�l
    Dim base As Long
    ' �ꎞ�ϐ�
    Dim temp As Long
    
    ' ���S�̃C���f�b�N�X
    Dim center As Long
    
    Dim i      As Long
    Dim j      As Long

    ' �X�^�b�N�ɍŏ��ɐݒ肷��ϐ���ݒ�
    stackVal(1) = Left
    stackVal(2) = right
    ' �X�^�b�N�Ƀv�b�V������
    stack.push stackVal

    ' �X�^�b�N�̒��g���Ȃ��Ȃ�܂Ŏ��s
    Do While stack.count > 0
        
        ' �X�^�b�N����l�����o��
        stackVal = stack.pop
        
        ' ���[���擾
        Left = stackVal(1)
        ' �E�[���擾
        right = stackVal(2)
        
        ' ��������N�C�b�N�\�[�g�̃A���S���Y���i���ȏ��ǂ���j
        If Left < right Then
        
            center = Int((Left + right) / 2)
            
            base = a(center)
            
            i = Left
            j = right
            
            Do While i <= j
            
                ' ���������召�̔�r����
                Do While a(i) < base
                
                    i = i + 1
                Loop
            
                ' ���������召�̔�r����
                Do While a(j) > base
                
                    j = j - 1
                Loop
            
                If i <= j Then
                
                    temp = a(i)
                    a(i) = a(j)
                    a(j) = temp
                    
                    i = i + 1
                    j = j - 1
                End If
            
            Loop
            
            ' �ċA�Ăяo���ł͂Ȃ��A�X�^�b�N�ɏ����l�߂�
            ' �V���ȑ����������X�^�b�N�ɋl�߂�
            
            ' �E�����̏��
            stackVal(1) = i
            stackVal(2) = right
            stack.push stackVal
            
            ' �������̏��
            stackVal(1) = Left
            stackVal(2) = j
            stack.push stackVal
            
        End If
    
    Loop
End Sub

' =========================================================
' ���t�@�C���o�͂ł��邩���m�F����
'
' �T�v�@�@�@�F
' �����@�@�@�FfolderPath �t�@�C���p�X
' �߂�l�@�@�FTrue �t�@�C���o�͉\�AFalse �t�@�C���o�͕s��
'
' =========================================================
Public Function touch(ByVal folderPath As String) As Boolean

    On Error GoTo err

    ' �o�̓t�@�C���p�X
    Dim touchedFilePath As String
    
    Dim i As Long
    Dim fw As FileWriter
    
    ' �d�������t�@�C�������݂��邱�Ƃ��l�����ă��[�v���񂵂ăJ�E���^�ϐ����t�@�C���p�X�̈ꕔ�Ɏg�p����
    ' �i�����炭100��ȓ��ɂ́A���j�[�N�ȃt�@�C�����ɂȂ�͂��j
    For i = 0 To 100
    
        touchedFilePath = VBUtil.concatFilePath(folderPath, "sut_touch________" & (i + 1))
        
        If Not VBUtil.isExistFile(touchedFilePath) Then
        
            Set fw = New FileWriter
            fw.init touchedFilePath, "Shift_JIS", vbNewLine
            fw.writeText "touch"
            fw.destroy
            
            ' �����܂ŗ����琳��Ƀt�@�C���o�͂��ꂽ�Ƃ݂Ȃ�
            Exit For
            
        End If
    Next
    
    ' �Ō�Ƀt�@�C�����폜����
    Kill touchedFilePath

    touch = True

    Exit Function
    
err:

    touch = False

End Function

' =========================================================
' ���t�@�C�������݂��邩���`�F�b�N����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�FTrue �t�@�C�������݂���ꍇ
'
' =========================================================
Public Function isExistFile(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        ' �t�@�C�������݂���ꍇ
        isExistFile = True
    Else
        ' �t�@�C�������݂��Ȃ��ꍇ
        isExistFile = False
    End If

End Function

' =========================================================
' ���t�@�C�������݂��邩���`�F�b�N����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�FTrue �t�@�C�������݂���ꍇ
'
' =========================================================
Public Function isExistDirectory(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) Then
        ' �t�@�C�������݂���ꍇ
        isExistDirectory = True
    Else
        ' �t�@�C�������݂��Ȃ��ꍇ
        isExistDirectory = False
    End If

End Function

' =========================================================
' ���t�@�C���p�X����f�B���N�g���p�X�𒊏o����
'   �f�B���N�g���̏ꍇ�A������ԋp
'       �t�@�C���̏ꍇ�A�f�B���N�g���p�X�𒊏o
'   �������݂��Ȃ��ꍇ�A�f�B���N�g���p�X�𒊏o
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�F�f�B���N�g���p�X
'
' =========================================================
Public Function extractDirPathFromFilePath(filePath As String) As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) Then
        ' �t�@�C���ł͂Ȃ��i���f�B���N�g���Ȃǂ́j�ꍇ
        extractDirPathFromFilePath = filePath
        Exit Function
    End If
    
    ' �߂�l
    Dim ret As String
    
    ' �f�B���N�g���ʒu
    Dim dirPoint As Long

    ' ������̉E�[����"\"���������A���[����̈ʒu���擾����
    dirPoint = InStrRev(filePath, "\")
    
    ' "\"��������Ȃ��ꍇ
    If dirPoint <> 0 Then
    
        ' �f�B���N�g���p�X�̎擾
        ret = Left$(filePath, dirPoint - 1)
        
        extractDirPathFromFilePath = ret
    
    Else
        extractDirPathFromFilePath = ""
    
    End If
    
End Function

' =========================================================
' ���f�B���N�g���p�X�ƃt�@�C���p�X��A������
'
' �T�v�@�@�@�F
' �����@�@�@�Fdir      �f�B���N�g���p�X
' �@�@�@�@�@�@filePath �t�@�C���p�X
' �߂�l�@�@�F�A����̕�����
'
' =========================================================
Public Function concatFilePath(ByVal dir As String, ByVal fileName As String) As String

    ' ������̍Ō���� "\" ���t���Ă��邩���m�F����
    If InStrRev(dir, "\") = Len(dir) Then
    
        concatFilePath = dir & fileName
    Else
    
        concatFilePath = dir & "\" & fileName
    End If
    
End Function

' =========================================================
' ���f�B���N�g�����쐬����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�FTrue �f�B���N�g���쐬����True��ԋp
'
' =========================================================
Public Function createDir(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) = False And _
       fso.FileExists(filePath) = False Then
        fso.CreateFolder filePath
        createDir = True
    End If

    createDir = False
        
End Function

' =========================================================
' ���f�B���N�g�����폜����
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
' �߂�l�@�@�FTrue �f�B���N�g���폜����True��ԋp
'
' =========================================================
Public Function deleteDir(ByVal filePath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) = True Then
        fso.DeleteFolder filePath
        deleteDir = True
    End If

    deleteDir = False
        
End Function

Public Function convertKeyCodeToKeyAscii(ByVal KeyCode As Long) As String

    If vbKey0 = KeyCode Then
        convertKeyCodeToKeyAscii = "0"
    ElseIf vbKey1 = KeyCode Then convertKeyCodeToKeyAscii = "1"
    ElseIf vbKey2 = KeyCode Then convertKeyCodeToKeyAscii = "2"
    ElseIf vbKey3 = KeyCode Then convertKeyCodeToKeyAscii = "3"
    ElseIf vbKey4 = KeyCode Then convertKeyCodeToKeyAscii = "4"
    ElseIf vbKey5 = KeyCode Then convertKeyCodeToKeyAscii = "5"
    ElseIf vbKey6 = KeyCode Then convertKeyCodeToKeyAscii = "6"
    ElseIf vbKey7 = KeyCode Then convertKeyCodeToKeyAscii = "7"
    ElseIf vbKey8 = KeyCode Then convertKeyCodeToKeyAscii = "8"
    ElseIf vbKey9 = KeyCode Then convertKeyCodeToKeyAscii = "9"
    ElseIf vbKeyA = KeyCode Then convertKeyCodeToKeyAscii = "A"
    ElseIf vbKeyB = KeyCode Then convertKeyCodeToKeyAscii = "B"
    ElseIf vbKeyC = KeyCode Then convertKeyCodeToKeyAscii = "C"
    ElseIf vbKeyD = KeyCode Then convertKeyCodeToKeyAscii = "D"
    ElseIf vbKeyE = KeyCode Then convertKeyCodeToKeyAscii = "E"
    ElseIf vbKeyF = KeyCode Then convertKeyCodeToKeyAscii = "F"
    ElseIf vbKeyG = KeyCode Then convertKeyCodeToKeyAscii = "G"
    ElseIf vbKeyH = KeyCode Then convertKeyCodeToKeyAscii = "H"
    ElseIf vbKeyI = KeyCode Then convertKeyCodeToKeyAscii = "I"
    ElseIf vbKeyJ = KeyCode Then convertKeyCodeToKeyAscii = "J"
    ElseIf vbKeyK = KeyCode Then convertKeyCodeToKeyAscii = "K"
    ElseIf vbKeyL = KeyCode Then convertKeyCodeToKeyAscii = "L"
    ElseIf vbKeyM = KeyCode Then convertKeyCodeToKeyAscii = "M"
    ElseIf vbKeyN = KeyCode Then convertKeyCodeToKeyAscii = "N"
    ElseIf vbKeyO = KeyCode Then convertKeyCodeToKeyAscii = "O"
    ElseIf vbKeyP = KeyCode Then convertKeyCodeToKeyAscii = "P"
    ElseIf vbKeyQ = KeyCode Then convertKeyCodeToKeyAscii = "Q"
    ElseIf vbKeyR = KeyCode Then convertKeyCodeToKeyAscii = "R"
    ElseIf vbKeyS = KeyCode Then convertKeyCodeToKeyAscii = "S"
    ElseIf vbKeyT = KeyCode Then convertKeyCodeToKeyAscii = "T"
    ElseIf vbKeyU = KeyCode Then convertKeyCodeToKeyAscii = "U"
    ElseIf vbKeyV = KeyCode Then convertKeyCodeToKeyAscii = "V"
    ElseIf vbKeyW = KeyCode Then convertKeyCodeToKeyAscii = "W"
    ElseIf vbKeyX = KeyCode Then convertKeyCodeToKeyAscii = "X"
    ElseIf vbKeyY = KeyCode Then convertKeyCodeToKeyAscii = "Y"
    ElseIf vbKeyZ = KeyCode Then convertKeyCodeToKeyAscii = "Z"
    End If

End Function

' =========================================================
' ���|�C���g����s�N�Z���ɒP�ʂ�ϊ�����
'
' �T�v�@�@�@�F
' �����@�@�@�Fd     DPI
' �@�@�@�@�@�@pixel �s�N�Z��
' �߂�l�@�@�F�|�C���g
'
' =========================================================
Public Function convertPixelToPoint(ByVal d As Long, ByVal pixel As Long) As Single

    convertPixelToPoint = CSng(pixel) / d * 72

End Function

' =========================================================
' ���s�N�Z������|�C���g�ɒP�ʂ�ϊ�����
'
' �T�v�@�@�@�F
' �����@�@�@�Fd     DPI
' �@�@�@�@�@�@pixel �s�N�Z��
' �߂�l�@�@�F�|�C���g
'
' =========================================================
Public Function convertPointToPixel(ByVal d As Long, ByVal Point As Single) As Long

    convertPointToPixel = Point * d / 72
    
End Function

' =========================================================
' ���^�U�l�������^�U�f�[�^�ɕϊ�����B
'
' �T�v�@�@�@�F
' �����@�@�@�Fstr ������
' �߂�l�@�@�F�ϊ���̐^�U�f�[�^
'
' =========================================================
Public Function convertBoolStrToBool(ByVal str As String) As Boolean

    If str = Empty Then
        ' �����͎�
        convertBoolStrToBool = False
    Else
        ' ���͎�
    
        If LCase$(str) = "true" Then
            ' �^
            convertBoolStrToBool = True
        Else
            ' �U
            convertBoolStrToBool = False
        End If
    
    End If
    
End Function

' =========================================================
' �����S���W���v�Z����
'
' �T�v�@�@�@�F�v�Z��̍��W���Adx�Edy�Ɋi�[�����
' �����@�@�@�Fsx ��ƂȂ��` ���WX
' �@�@�@�@�@�@sy ��ƂȂ��` ���WY
' �@�@�@�@�@�@sw ��ƂȂ��` ��
' �@�@�@�@�@�@sh ��ƂȂ��` ����
' �@�@�@�@�@�@dx ��r�����` ���WX
' �@�@�@�@�@�@dy ��r�����` ���WY
' �@�@�@�@�@�@dw ��r�����` ��
' �@�@�@�@�@�@dh ��r�����` ����
'
' =========================================================
Public Sub calcCenterPoint( _
                           ByVal sx As Single _
                         , ByVal sy As Single _
                         , ByVal sw As Single _
                         , ByVal sh As Single _
                         , ByRef dx As Single _
                         , ByRef dy As Single _
                         , ByVal dw As Single _
                         , ByVal dh As Single)

    ' ���S���v�Z����
    Dim newX As Single
    Dim newY As Single
    
    newX = sw / 2 - dw / 2 + sx
    newY = sh / 2 - dh / 2 + sy

    ' ���S��ݒ肷��
    dx = newX
    dy = newY

End Sub

' =========================================================
' ����`A��B���r��A��B���Ɏ��܂��Ă��邩���m�F����
'
' �T�v�@�@�@�F
' �����@�@�@�Fsx ��ƂȂ��` ���WX
' �@�@�@�@�@�@sy ��ƂȂ��` ���WY
' �@�@�@�@�@�@sw ��ƂȂ��` ��
' �@�@�@�@�@�@sh ��ƂȂ��` ����
' �@�@�@�@�@�@dx ��r�����` ���WX
' �@�@�@�@�@�@dy ��r�����` ���WY
' �@�@�@�@�@�@dw ��r�����` ��
' �@�@�@�@�@�@dh ��r�����` ����
' �߂�l�@�@�FTrue ��`A���Ɏ��܂��Ă���ꍇ
'
' =========================================================
Public Function isInnerScreen( _
                           ByVal sx As Single _
                         , ByVal sy As Single _
                         , ByVal sw As Single _
                         , ByVal sh As Single _
                         , ByRef dx As Single _
                         , ByRef dy As Single _
                         , ByRef dw As Single _
                         , ByRef dh As Single) As Boolean

    isInnerScreen = True

    ' �g���͂ݏo���Ă��Ȃ������m�F����
    If sx > dx Then
    
        isInnerScreen = False
        
    ElseIf sy > dy Then
    
        isInnerScreen = False
        
    ElseIf (sx + sw) < (dx + dw) Then
    
        isInnerScreen = False
        
    ElseIf (sy + sh) < (dy + dh) Then
    
        isInnerScreen = False
        
    End If

End Function

' =========================================================
' ���p�f�B���O�֐�
'
' �T�v�@�@�@�F������̍����ɓ���̕�����C�ӂ̌����ɂȂ�悤�ɋl�߂�
' �����@�@�@�Fvalue  �l
' �@�@�@�@�@�@length ����
' �@�@�@�@�@�@char   ����
' �߂�l�@�@�F�p�f�B���O����
'
' =========================================================
Public Function padLeft(ByVal value As String _
                      , ByVal length As Long _
                      , Optional ByVal char As String = "0") As String

    ' �p�f�B���O���錅��
    Dim padLen As Long
    padLen = length - Len(value)
    
    If padLen < 1 Then
    
        padLeft = value
        Exit Function
    End If

    padLeft = String(length - Len(value), char) & value

End Function

' =========================================================
' ���p�f�B���O�֐�
'
' �T�v�@�@�@�F������̉E���ɓ���̕�����C�ӂ̌����ɂȂ�悤�ɋl�߂�
' �����@�@�@�Fvalue  �l
' �@�@�@�@�@�@length ����
' �@�@�@�@�@�@char   ����
' �߂�l�@�@�F�p�f�B���O����
'
' =========================================================
Public Function padRight(ByVal value As String _
                       , ByVal length As Long _
                       , Optional ByVal char As String = "0") As String

    ' �p�f�B���O���錅��
    Dim padLen As Long
    padLen = length - Len(value)
    
    If padLen < 1 Then
    
        padRight = value
        Exit Function
    End If

    padRight = value & String(length - Len(value), char)

End Function

' =========================================================
' ���G���R�[�h���X�g�擾�֐�
'
' �T�v�@�@�@�F�G���R�[�h���X�g���擾����
' �����@�@�@�F
' �߂�l�@�@�F�G���R�[�h���X�g
'
' =========================================================
Public Function getEncodeList() As ValCollection

    ' �ΏۂƂ��镶���R�[�h���X�g
    Dim includeChars As New ValCollection
    includeChars.setItem "Shift_JIS", "Shift_JIS"
    includeChars.setItem "EUC-JP", "EUC-JP"
    includeChars.setItem "UTF-8", "UTF-8"
    includeChars.setItem "UTF-8 (with bom)", "UTF-8 (with bom)"
    includeChars.setItem "UNICODE", "UNICODE"
    
    Set getEncodeList = includeChars
    
End Function


' =========================================================
' ���G���R�[�h���X�g�擾�֐��i���W�X�g������擾�j
'
' �T�v�@�@�@�F�G���R�[�h���X�g���擾����
' �����@�@�@�F
' �߂�l�@�@�F�G���R�[�h���X�g
'
' =========================================================
Public Function getEncodeListFromRegistry() As ValCollection

    ' �����R�[�h���X�g�擾�p���W�X�g���I�u�W�F�N�g
    Dim regChar As New RegistryManipulator
    ' �����R�[�h���X�g�擾�p�̃��W�X�g���I�u�W�F�N�g������������
    regChar.init RegKeyConstants.HKEY_CLASS_ROOT _
               , REG_PATH_CHARACTER_CODE_LIST _
               , RegAccessConstants.KEY_READ _
               , False
               
    ' �G�C���A�X�m�F�p���W�X�g���I�u�W�F�N�g
    Dim regCharAlias As RegistryManipulator
    
    ' �����R�[�h�ꗗ
    Dim charList As ValCollection
    ' �����R�[�h���X�g���擾����
    Set charList = regChar.getKeyList
    ' �����R�[�h�ꗗ�i�G�C���A�X�����O�j
    Dim charListRemovalAlias As New ValCollection

    ' �����R�[�h
    Dim char As Variant
    ' �����R�[�h �G�C���A�X
    Dim charAlias As String
    
    For Each char In charList.col
    
        ' �G�C���A�X�m�F�p���W�X�g���I�u�W�F�N�g������
        Set regCharAlias = New RegistryManipulator
        
        regCharAlias.init RegKeyConstants.HKEY_CLASS_ROOT _
                        , REG_PATH_CHARACTER_CODE_LIST & "\" & char _
                        , RegAccessConstants.KEY_READ _
                        , False
                        
        ' �����R�[�h�̃G�C���A�X�ł��邩�𔻒肷��
        If regCharAlias.getValue(REG_KEY_ALIAS_CHARSET, charAlias) = False Then
        
            ' �G�C���A�X�ł͂Ȃ��ꍇ�A�ǉ�����
            charListRemovalAlias.setItem char, char
        End If
    
        ' �j������
        Set regCharAlias = Nothing
    Next
    
    Set getEncodeList = charListRemovalAlias
    
End Function

' =========================================================
' �����s�R�[�h���X�g�擾�֐�
'
' �T�v�@�@�@�F���s�R�[�h���X�g�擾
' �����@�@�@�F
' �߂�l�@�@�F���s�R�[�h���X�g
'
' =========================================================
Public Function getNewlineList() As ValCollection

    Set getNewlineList = New ValCollection
    
    getNewlineList.setItem NEW_LINE_STR_CRLF
    getNewlineList.setItem NEW_LINE_STR_CR
    getNewlineList.setItem NEW_LINE_STR_LF
    
End Function

' =========================================================
' �����s�R�[�h���������ۂ̉��s�R�[�h�l�ɕϊ�����֐�
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F���ۂ̉��s�R�[�h�l
'
' =========================================================
Public Function convertNewLineStrToNewLineCode(ByVal newLineStr As String) As String

    If newLineStr = NEW_LINE_STR_CRLF Then
    
        ' Windows
        convertNewLineStrToNewLineCode = vbCr & vbLf
    
    ElseIf newLineStr = NEW_LINE_STR_CR Then
    
        ' Mac
        convertNewLineStrToNewLineCode = vbCr
    
    ElseIf newLineStr = NEW_LINE_STR_LF Then
    
        ' Unix
        convertNewLineStrToNewLineCode = vbLf
        
    ' ���Ă͂܂�Ȃ��ꍇ
    Else
    
        ' Windows
        convertNewLineStrToNewLineCode = vbCr & vbLf
    
    End If

End Function

' =========================================================
' �����K�\���̌���������̃G�X�P�[�v�����֐�
'
' �T�v�@�@�@�F���K�\���̌���������̃G�X�P�[�v����
' �����@�@�@�Fkeyword ����������L�[���[�h
' �߂�l�@�@�F�ϊ���̌���������
'
' =========================================================
Public Function escapeRegExpKeyword(ByVal keyword As String) As String

    escapeRegExpKeyword = keyword
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "\", "\\")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "*", "\*")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "+", "\+")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, ".", "\.")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "?", "\?")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "{", "\{")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "}", "\}")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "(", "\(")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, ")", "\)")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "[", "\[")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "]", "\]")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "^", "\^")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "$", "\$")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "-", "\-")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "|", "\|")
    escapeRegExpKeyword = replace(escapeRegExpKeyword, "/", "\/")

End Function

' =========================================================
' �����C���h�J�[�h�t�B���^�����֐�
'
' �T�v�@�@�@�F���C���h�J�[�h�����������������������{����B
'
'             filterKeyword�Ɏg�p�ł�����ꕶ���B
'             * �� �C�ӂ̕����̘A��
'             ? �� �C�ӂ�1����
'
'             listOfElementPropName�ɂ́A�h�b�g���g�p���邱�ƂŃl�X�g�����v���p�e�B���w�肷�邱�Ƃ��\�B
'             ��j"table.tableName"�̂悤�ɂ���ƁA�܂�table�v���p�e�B���擾���āAtable�v���p�e�B����tableName���擾����
'
' �����@�@�@�Flist                  �t�B���^�Ώۃ��X�g
'     �@�@�@  listOfElementPropName �t�B���^�Ώۃ��X�g���̃v���p�e�B��
'     �@�@�@  filterKeyword         �t�B���^�Ɏg�p����L�[���[�h
' �߂�l�@�@�F�����Ώۃ��X�g���t�B���^��������
'
' =========================================================
Public Function filterWildcard(ByVal list As ValCollection, _
                               ByVal listOfElementPropName As String, _
                               ByVal filterKeyword As String) As ValCollection

    ' ---------------------------------------
    Dim convertedFilterKeyword As String
    convertedFilterKeyword = filterKeyword
    
    ' ���C���h�J�[�h������"�K���Ȑ���R�[�h"�i��ʂœ��́E�\���ł��Ȃ�������j�ɕϊ�����
    ' ����R�[�h�ɕϊ����闝�R�Ƃ��āA�㑱�Ő��K�\����������G�X�P�[�v���鏈��������A������ŃG�X�P�[�v�����{����Ȃ��悤�ɂ��邽��
    ' ���ɒ[�Șb�A����R�[�h�Ȃ牽�ł��悢�i�ȉ����g�p�j
    ' ���u����DC1 = Char(17)
    ' ���u����DC2 = Char(18)
    ' ���u����DC3 = Char(19)
    ' ���u����DC4 = Char(20)
    convertedFilterKeyword = replace(convertedFilterKeyword, "~*", Chr(19)) ' �`���_�t���Ȃ̂Œʏ�̕��������ɂ���
    convertedFilterKeyword = replace(convertedFilterKeyword, "~?", Chr(20)) ' �`���_�t���Ȃ̂Œʏ�̕��������ɂ���
    convertedFilterKeyword = replace(convertedFilterKeyword, "*", Chr(17))
    convertedFilterKeyword = replace(convertedFilterKeyword, "?", Chr(18))
    
    ' �`���_�t���̓��ꕶ���Ȃ̂Œʏ�̕����ɂ������̂Ō��̒l�ɖ߂��āA�㑱������VBUtil.escapeRegExpKeyword�ŃG�X�P�[�v���Ēʏ�̕����Ƃ��ĉ��߂����悤�ɂ���
    convertedFilterKeyword = replace(convertedFilterKeyword, Chr(19), "*")
    convertedFilterKeyword = replace(convertedFilterKeyword, Chr(20), "?")
    
    ' �L�[���[�h�Ɋ܂܂�鐳�K�\���̓��ꕶ�����G�X�P�[�v����
    convertedFilterKeyword = VBUtil.escapeRegExpKeyword(convertedFilterKeyword)
    
    ' ���C���h�J�[�h�ɑΉ���������R�[�h�𐳋K�\���ɕϊ�����
    convertedFilterKeyword = replace(convertedFilterKeyword, Chr(17), ".*") ' *�͔C�ӂ̕�����0�ȏ�̘A��
    convertedFilterKeyword = replace(convertedFilterKeyword, Chr(18), ".") ' ?�͔C�ӂ�1����
    ' ---------------------------------------
    
    Set filterWildcard = filterRegExp(list, listOfElementPropName, convertedFilterKeyword)

End Function

' =========================================================
' �����K�\���t�B���^�����֐�
'
' �T�v�@�@�@�F���K�\�������������������������{����B
'
'             filterKeyword�Ɏg�p�ł�����ꕶ����RegExp�ɏ�����B
'
'             listOfElementPropName�ɂ́A�h�b�g���g�p���邱�ƂŃl�X�g�����v���p�e�B���w�肷�邱�Ƃ��\�B
'             ��j"table.tableName"�̂悤�ɂ���ƁA�܂�table�v���p�e�B���擾���āAtable�v���p�e�B����tableName���擾����
'
' �����@�@�@�Flist                  �t�B���^�Ώۃ��X�g
'     �@�@�@  listOfElementPropName �t�B���^�Ώۃ��X�g���̃v���p�e�B��
'     �@�@�@  filterKeyword         �t�B���^�Ɏg�p����L�[���[�h
' �߂�l�@�@�F�����Ώۃ��X�g���t�B���^��������
'
' =========================================================
Public Function filterRegExp(ByVal list As ValCollection, _
                             ByVal listOfElementPropName As String, _
                             ByVal filterKeyword As String) As ValCollection

    ' �߂�l
    Set filterRegExp = New ValCollection
    
    ' ���K�\���I�u�W�F�N�g�𐶐�
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        ' �����Ώە�����
        .Pattern = "^" & filterKeyword & "$"
        ' �啶�������������t���O
        .IgnoreCase = True
        ' ������S�̂��J��Ԃ���������t���O
        .Global = False
    End With

    ' list�ϐ��̗v�f
    Dim rec As Variant
    ' �t�B���^�Ώ�
    Dim searchExpression As Variant
    ' ���X�g�̗v�f�̃v���p�e�B���̔z��
    Dim listOfElementPropNameArray As Variant
    listOfElementPropNameArray = Split(listOfElementPropName, ".")
    
    Dim i As Long

    For Each rec In list.col
    
        Set searchExpression = rec
        For i = LBound(listOfElementPropNameArray) To UBound(listOfElementPropNameArray)
            If i = UBound(listOfElementPropNameArray) Then
                ' �����̏ꍇ�́A�v���~�e�B�u�ȃf�[�^
                If Not searchExpression Is Nothing Then
                    searchExpression = CallByName(searchExpression, listOfElementPropNameArray(i), VbGet)
                End If
            Else
                ' �����ł͂Ȃ��ꍇ�́A�I�u�W�F�N�g�^
                If Not searchExpression Is Nothing Then
                    Set searchExpression = CallByName(searchExpression, listOfElementPropNameArray(i), VbGet)
                End If
            End If
        Next
    
        If Not IsObject(searchExpression) Then
            ' �I�u�W�F�N�g�^�ł͂Ȃ��i�����Ȃ̂ŕ�����^��z��j
            If reg.test(CStr(searchExpression)) Then
                filterRegExp.setItem rec
            End If
        End If
    Next

End Function

' =========================================================
' �����[�U�[�t�H�[���������[�h����
'
' �T�v�@�@�@�F���[�U�[�t�H�[���͋N�����̃A�N�e�B�u�ȃu�b�N��ێ����鐫��������B
'             �u�b�N���؂�ւ�����ꍇ�͈�x�t�H�[�����A�����[�h����K�v�����邽�߁A�����̔�������{���ēK�؂Ƀ����[�h����B
' �����@�@�@�Fobj ���[�U�[�t�H�[��
' �߂�l�@�@�F
'
' ���L�����@�F
'
' =========================================================
Public Function unloadFormIfChangeActiveBook(ByRef obj As Variant) As Boolean

    If obj Is Nothing Then
    
        unloadFormIfChangeActiveBook = False
        Exit Function
    
    End If
    
    Dim book As Workbook
    Set book = CallByName(obj, "getTargetBook", VbMethod)
    
    If ActiveWorkbook Is book Then
    
        unloadFormIfChangeActiveBook = False
        Exit Function
    
    End If
    
    unloadFormIfChangeActiveBook = True
        
End Function

