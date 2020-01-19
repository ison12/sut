Attribute VB_Name = "WinAPI_Shell"
Option Explicit

' *********************************************************
' shell32.dll�Œ�`����Ă���֐��S��萔�B
'
' �쐬�ҁ@�FHideki Isobe
' �����@�@�F2008/10/11�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' ��WinAPI�̒�`
' �O���v���O�����N���֐�
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As LongPtr _
       , ByVal lpOperation As String _
       , ByVal lpFile As String _
       , ByVal lpParameters As String _
       , ByVal lpDirectory As String _
       , ByVal nShowCmd_ As Long) As Long
#Else
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long _
       , ByVal lpOperation As String _
       , ByVal lpFile As String _
       , ByVal lpParameters As String _
       , ByVal lpDirectory As String _
       , ByVal nShowCmd_ As Long) As Long
#End If

' =========================================================
' �����ϐ����擾����
'
' �T�v�@�@�@�F
' �����@�@�@�Fkey ���ϐ��̃L�[�l
' �߂�l�@�@�F���ϐ��̒l
'
' =========================================================
Public Function getEnvironmentVariable(ByVal key As String) As String

    ' ���ϐ�
    Dim environmentString As String
    ' ���ϐ��̃L�[�l����
    Dim environmentKey    As String
    ' ���ϐ��̒l����
    Dim environmentVal    As String
    
    ' ���ϐ��̃L�[�ƒl����؂蕶���i=�j�̈ʒu��ێ�����ϐ�
    Dim envKeyValSeparate As Long
    
    ' �C���f�b�N�X
    Dim i As Long
    
    ' �C���f�b�N�X�̏����l��1�Ƃ���
    i = 1
    
    ' ���[�v�����{����
    Do
        ' ���ϐ�
        environmentString = Environ(i)
        
        ' �L�[�ƒl�̋�؂蕶���̈ʒu���i�[����
        envKeyValSeparate = InStr(environmentString, "=")
        
        ' ��؂蕶�������������ꍇ
        If envKeyValSeparate <> 0 Then
        
            ' �L�[���i�[
            environmentKey = Mid$(environmentString, 1, envKeyValSeparate - 1)
            ' �l���i�[
            environmentVal = Mid$(environmentString, envKeyValSeparate + 1, Len(environmentString))
            
            If UCase$(environmentKey) = UCase$(key) Then
            
                ' ���ϐ��̒l��߂�l�Ɋi�[����
                getEnvironmentVariable = environmentVal
                
                Exit Function
            End If
            
        End If
        
        i = i + 1
        
    Loop Until Environ(i) = ""

    ' �󕶎����Ԃ�
    getEnvironmentVariable = ""

End Function
