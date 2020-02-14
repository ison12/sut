Attribute VB_Name = "MySQL"
Option Explicit

' *********************************************************
' MySQL�Ɋ֘A�������[�e�B���e�B���W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2007/12/01�@�V�K�쐬
'
' ���L�����@�FMySQL�Ɉˑ������֐��S���`�B
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
           InStr(UCase$(dataType), "CHAR") <> 0 _
        Or InStr(UCase$(dataType), "VARCHAR") <> 0 _
        Or InStr(UCase$(dataType), "BLOB") <> 0 _
        Or InStr(UCase$(dataType), "TEXT") <> 0 _
        Or InStr(UCase$(dataType), "ENUM") <> 0 _
        Or InStr(UCase$(dataType), "SET") <> 0 _
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
           InStr(UCase$(dataType), "DATETIME") <> 0 _
        Or InStr(UCase$(dataType), "DATE") <> 0 _
        Or InStr(UCase$(dataType), "TIMESTAMP") <> 0 _
        Or InStr(UCase$(dataType), "TIME") <> 0 _
        Or InStr(UCase$(dataType), "YEAR") <> 0 _
    Then
    
        isTime = True
            
    End If

End Function

Public Function escapeValue(ByRef val As String) As String

    escapeValue = replace(val, "'", "''")

End Function

