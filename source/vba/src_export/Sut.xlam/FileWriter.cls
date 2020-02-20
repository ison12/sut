VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �t�@�C���o�̓N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/04/17�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Private stream As Object

' �t�@�C���p�X
Private filePath As String
' �����R�[�h
Private charSet As String
' BOM�t��
Private isWithBom As Boolean
' ���s�R�[�h
Private newline As String

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    destroy
End Sub

' =========================================================
' �������R�[�h�擾����
'
' �T�v�@�@�@�F
' �����@�@�@�Fc            �����R�[�h
' �@�@�@�@�@�@outChar      �߂�l�c�ϊ���̕����R�[�h
' �@�@�@�@�@�@outIsWithBom �߂�l�c�����R�[�h��BOM�L�肩�ǂ���
'
' �߂�l�@�@�F
'
' =========================================================
Private Function getCharSet(ByVal c As String _
                        , ByRef outChar As String _
                        , ByRef outIsWithBom As Boolean)

    If InStr(UCase$(c), "UTF-8") > 0 Then
        ' UTF-8
        
        outChar = "UTF-8"
        
        If InStr(UCase$(c), "BOM") > 0 Then
            ' BOM����
            outIsWithBom = True
        Else
            ' BOM�Ȃ�
            outIsWithBom = False
        End If
        
    Else
        ' ��L�ȊO
        
        outChar = c
        outIsWithBom = False
        
    End If

End Function

' =========================================================
' ������������
'
' �T�v�@�@�@�F���������������s����B
' �����@�@�@�Ff      �t�@�C���p�X
' �@�@�@�@�@�@c      �����R�[�h
' �@�@�@�@�@�@n      ���s�R�[�h
' �@�@�@�@�@�@append �ǉ��L��
'
' �߂�l�@�@�F
'
' =========================================================
Public Sub init(ByVal f As String _
              , ByVal c As String _
              , ByVal n As String _
              , Optional ByVal append = False)
              
    ' �t�@�C���p�X��ݒ�
    filePath = f
    ' �����R�[�h��ݒ�
    getCharSet c, charSet, isWithBom
    ' ���s�R�[�h��ݒ�
    newline = n

    ' �X�g���[���I�u�W�F�N�g��ݒ�
    Set stream = CreateObject("ADODB.Stream")
    
    ' �����R�[�h�v���p�e�B���w��
    ' �V�X�e�����F���ł��镶���Z�b�g������̈ꗗ�ɂ��ẮAWindows ���W�X�g���� HKEY_CLASSES_ROOT\MIME\Database\Charset �̃T�u�L�[���Q�Ƃ��Ă��������B
    stream.charSet = charSet
    
    ' ���s�R�[�h�v���p�e�B���w��
    stream.LineSeparator = convertStreamConstNewLine(newline)
    
    ' �X�g���[�����J��
    stream.Open
    
    ' �ǉ����[�h
    If append = True Then
    
        ' �t�@�C�������݂��Ă���ꍇ
        If VBUtil.isExistFile(filePath) = True Then
            
            ' �w��̃t�@�C����ǂݍ���
            stream.LoadFromFile filePath
            stream.position = stream.size
            
        End If
    End If
    
End Sub

' =========================================================
' ���j������
'
' �T�v�@�@�@�F���������������s����B
' �����@�@�@�Ff      �t�@�C���p�X
' �@�@�@�@�@�@c      �����R�[�h
' �@�@�@�@�@�@n      ���s�R�[�h
' �@�@�@�@�@�@append �ǉ��L��
'
' �߂�l�@�@�F
'
' =========================================================
Public Sub destroy()

    On Error GoTo err

    If stream Is Nothing Then
    
        Exit Sub
    End If
    
    ' �X�g���[�������Ă��Ȃ��ꍇ
    If stream.state <> 0 Then
    
        If UCase$(charSet) = "UTF-8" And isWithBom = False Then
            ' UTF-8 BOM�Ȃ�
            
            ' �o�C�g�ʒu��擪�ɖ߂�
            stream.position = 0
            stream.Type = 1 ' adTypeBinary
            
            ' BOM��3Byte���X�L�b�v
            If stream.size >= 3 Then
                stream.position = 3
            End If
            
            ' BOM�𔲂����o�C�g�f�[�^����������
            Dim buf As Variant: buf = stream.read
            ' ��t�@�C���̏ꍇ�́Anull���ԋp�����̂Ń`�F�b�N����
            If isNull(buf) = False Then
                stream.position = 0
                stream.Write buf
                stream.SetEOS
            End If
            
        End If
    
        ' http://msdn.microsoft.com/ja-jp/library/cc389870.aspx
        ' SaveOptionsEnum
        '
        ' Stream �I�u�W�F�N�g����t�@�C����ۑ�����Ƃ��Ƀt�@�C�����쐬���邩�A�㏑�����邩��\���܂��B�����̒l�� AND ���Z�q�Ō������邱�Ƃ��ł��܂��B
        '
        ' �萔  �l  ����
        ' adSaveCreateNotExist  1   ����l�ł��BFileName �p�����[�^�Ŏw�肵���t�@�C�����Ȃ��ꍇ�͐V�����t�@�C�����쐬����܂��B
        ' adSaveCreateOverWrite 2   FileName �p�����[�^�Ŏw�肵���t�@�C��������ꍇ�́A���݊J����Ă��� Stream �I�u�W�F�N�g�̃f�[�^�Ńt�@�C�����㏑������܂��B
        stream.SaveToFile filePath, 2
    
        ' �X�g���[�������
        stream.Close
        Set stream = Nothing
        
    End If
    
    Exit Sub
    
err:

    Dim errT As errInfo
    errT = VBUtil.swapErr
        
    If Not stream Is Nothing Then
        ' �X�g���[�������
        stream.Close
        Set stream = Nothing
    End If

    err.Raise errT.Number, errT.Source, errT.Description, errT.HelpFile, errT.HelpContext

End Sub

' =========================================================
' ���������ݏ���
'
' �T�v�@�@�@�F�������ݏ��������{����
' �����@�@�@�Fbuffer        �������ݓ��e
' �@�@�@�@�@�@appendNewLine ���s�R�[�h�̒ǉ��L��
'
' �߂�l�@�@�F
'
' =========================================================
Public Sub writeText(ByVal buffer As String _
                   , Optional ByVal appendNewLine As Boolean = False)

    stream.writeText buffer, convertStreamConstAppendNewLine(appendNewLine)

End Sub

' =========================================================
' ���������ݏ������f
'
' �T�v�@�@�@�F�������ݏ����𒆒f����
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub Cancel()

    ' �X�g���[�������Ă��Ȃ��ꍇ
    If stream.state <> 0 Then
    
        ' �X�g���[�������
        stream.Close
        Set stream = Nothing
    End If

End Sub

Private Function convertStreamConstNewLine(ByVal n As String) As Long

    ' http://msdn.microsoft.com/ja-jp/library/cc389826.aspx
    ' LineSeparatorsEnum
    '
    ' �e�L�X�g Stream �I�u�W�F�N�g�̍s��؂�Ɏg���Ă��镶����\���܂��B
    '
    ' �萔   �l  ����
    ' adCR   13  ���s���A�������܂��B
    ' adCRLF -1  ����l�ł��B���s���A�s����������܂��B
    ' adLF   10  �s����������܂��B
    If n = vbCr & vbLf Then
    
        ' Windows
        convertStreamConstNewLine = -1
    
    ElseIf n = vbCr Then
    
        ' Mac
        convertStreamConstNewLine = 13
    
    ElseIf n = vbLf Then
    
        ' Unix
        convertStreamConstNewLine = 10
        
    ' ���Ă͂܂�Ȃ��ꍇ
    Else
    
        ' Windows
        convertStreamConstNewLine = -1
    
    End If


End Function

Private Function convertStreamConstAppendNewLine(ByVal appendNewLine As Boolean) As Long

    ' StreamWriteEnum
    '
    ' Stream �I�u�W�F�N�g�ɏ������ޕ�����ɁA�s��؂蕶����ǉ����邩�ǂ�����\���܂��B
    '
    ' �萔  �l  ����
    ' adWriteChar   0   ����l�ł��BStream �I�u�W�F�N�g�ɑ΂��āAData �p�����[�^�Ŏw�肵���e�L�X�g��������������݂܂��B
    ' adWriteLine   1   Stream �I�u�W�F�N�g�ɁA�e�L�X�g������ƍs��؂蕶�����������݂܂��BLineSeparator �v���p�e�B����`����Ă��Ȃ��ꍇ�́A���s���G���[��Ԃ��܂��B
    If appendNewLine = True Then
    
        convertStreamConstAppendNewLine = 1
    Else
    
        convertStreamConstAppendNewLine = 0
    End If

End Function
