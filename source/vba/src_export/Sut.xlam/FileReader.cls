VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �t�@�C���ǂݍ��݃N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/06/28�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Private stream As Object

' �t�@�C���p�X
Private filePath As String
' �����R�[�h
Private charSet As String
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
              , ByVal n As String)

    ' �t�@�C���p�X��ݒ�
    filePath = f
    ' �����R�[�h��ݒ�
    charSet = c
    ' ���s�R�[�h��ݒ�
    newline = n

End Sub

' =========================================================
' ���j������
'
' �T�v�@�@�@�F
' �����@�@�@�F
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
    
        ' �X�g���[�������
        stream.Close
        
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
' ���ǂݍ��ݏ���
'
' �T�v�@�@�@�F�ǂݍ��ݏ��������{����
' �����@�@�@�F
'
' �߂�l�@�@�F�ǂݍ��񂾃t�@�C���̓��e
'
' =========================================================
Public Function readText() As String

    ' �X�g���[���I�u�W�F�N�g��ݒ�
    Set stream = CreateObject("ADODB.Stream")
    
    ' �ǂݎ��ŊJ��
    stream.Mode = 3
    ' �e�L�X�g�ŊJ��
    stream.Type = 2

    ' �����R�[�h�v���p�e�B���w��
    ' �V�X�e�����F���ł��镶���Z�b�g������̈ꗗ�ɂ��ẮAWindows ���W�X�g���� HKEY_CLASSES_ROOT\MIME\Database\Charset �̃T�u�L�[���Q�Ƃ��Ă��������B
    stream.charSet = charSet
    
    ' ���s�R�[�h�v���p�e�B���w��
    stream.LineSeparator = convertStreamConstNewLine(newline)
    
    ' �X�g���[�����J��
    stream.Open
    stream.LoadFromFile filePath
    
    readText = stream.readText
    
    ' �X�g���[�������
    stream.Close
    Set stream = Nothing

End Function

' =========================================================
' ���ǂݍ��ݏ������f
'
' �T�v�@�@�@�F�ǂݍ��ݏ����𒆒f����
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub Cancel()

    If stream Is Nothing Then
    
        Exit Sub
    End If
    
    ' �X�g���[�������Ă��Ȃ��ꍇ
    If stream.state <> 0 Then
    
        ' �X�g���[�������
        stream.Close
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