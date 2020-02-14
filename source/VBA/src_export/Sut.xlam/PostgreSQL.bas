Attribute VB_Name = "PostgreSQL"
Option Explicit

' *********************************************************
' PostgreSQL�Ɋ֘A�������[�e�B���e�B���W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/04/27�@�V�K�쐬
'
' ���L�����@�FPostgreSQL�Ɉˑ������֐��S���`�B
' *********************************************************

' =========================================================
' ���f�[�^�l�̃`�F�b�N�֐��i���ʂȒl�j
'
' �T�v�@�@�@�FINSERT��VALUES�哙�̒l���`�F�b�N����B
' �����@�@�@�Fvalue �f�[�^�l
'
' �߂�l�@�@�FTrue NULL��������NOW�֐��̏ꍇ
'
' =========================================================
Public Function isSpecialValue(ByVal value As String) As Boolean

    isSpecialValue = False
    
    If _
           value = "NULL" _
        Or UCase$(value) = "NOW()" _
        Or UCase$(value) = "CURRENT_DATE" _
        Or UCase$(value) = "CURRENT_TIME" _
        Or UCase$(value) = "CURRENT_TIMESTAMP" _
    Then
    
        isSpecialValue = True
    
    End If

End Function

' =========================================================
' ���f�[�^�l�̃`�F�b�N�֐��i������^�j
'
' �T�v�@�@�@�FINSERT��VALUES�哙�̒l��������^�ł��邩���`�F�b�N����B
' �����@�@�@�FdataType �f�[�^�^
'
' �߂�l�@�@�FTrue ������^�̏ꍇ
'
' =========================================================
Public Function isChar(ByVal dataType As String) As Boolean

    isChar = False

    If _
           InStr(UCase$(dataType), "BIT") <> 0 _
        Or InStr(UCase$(dataType), "BIT VARYING") <> 0 _
        Or InStr(UCase$(dataType), "BOX") <> 0 _
        Or InStr(UCase$(dataType), "CHARACTER VARYING") <> 0 _
        Or InStr(UCase$(dataType), "CHARACTER") <> 0 _
        Or InStr(UCase$(dataType), "CHAR") <> 0 _
        Or InStr(UCase$(dataType), "CIDR") <> 0 _
        Or InStr(UCase$(dataType), "CIRCLE") <> 0 _
        Or InStr(UCase$(dataType), "INET") <> 0 _
        Or InStr(UCase$(dataType), "INTERVAL") <> 0 _
        Or InStr(UCase$(dataType), "LINE") <> 0 _
        Or InStr(UCase$(dataType), "LSEG") <> 0 _
        Or InStr(UCase$(dataType), "MACADDR") <> 0 _
        Or InStr(UCase$(dataType), "MONEY") <> 0 _
        Or InStr(UCase$(dataType), "PATH") <> 0 _
        Or InStr(UCase$(dataType), "POINT") <> 0 _
        Or InStr(UCase$(dataType), "POLYGON") <> 0 _
        Or InStr(UCase$(dataType), "TEXT") <> 0 _
    Then
    
        isChar = True
            
    End If

End Function

' =========================================================
' ���f�[�^�l�̃`�F�b�N�֐��i���t�E���Ԍ^�j
'
' �T�v�@�@�@�FINSERT��VALUES�哙�̒l�����t�E���Ԍ^�ł��邩���`�F�b�N����B
' �����@�@�@�FdataType �f�[�^�^
'
' �߂�l�@�@�FTrue ���t�E���Ԍ^�̏ꍇ
'
' =========================================================
Public Function isTime(ByVal dataType As String) As Boolean

    isTime = False

    If _
           InStr(UCase$(dataType), "DATE") <> 0 _
        Or InStr(UCase$(dataType), "TIME WITHOUT TIME ZONE") <> 0 _
        Or InStr(UCase$(dataType), "TIME WITH TIME ZONE") <> 0 _
        Or InStr(UCase$(dataType), "TIMESTAMP WITHOUT TIME ZONE") <> 0 _
        Or InStr(UCase$(dataType), "TIMESTAMP WITH TIME ZONE") <> 0 _
    Then
    
        isTime = True
            
    End If

End Function

Public Function escapeValue(ByRef val As String) As String

    escapeValue = replace(val, "'", "''")

End Function

