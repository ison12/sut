Attribute VB_Name = "WinAPI_Ole32"
Option Explicit

' *********************************************************
' ole32.dll�Œ�`����Ă���֐��S��萔�B
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/02/23�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As Long
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As Long
    Private Declare Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As Long, ByVal cbMax As Long) As Long
#End If

' =========================================================
' ��Guid�𐶐�����
'
' �T�v�@�@�@�F
' �����@�@�@�FhasEnclose �͂ݕ����������ǂ����̃t���O
'
' �߂�l�@�@�FGuid������
'
' =========================================================
Public Function createGuid(Optional ByVal hasEnclose As Boolean = True) As String
    
    ' guid length
    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
    
    Dim retValue As Long
    
    Dim guid    As GUID_TYPE
    Dim strGuid As String
    
    retValue = CoCreateGuid(guid)
    If retValue = 0 Then
        ' ����̏ꍇ
        
        ' guid�̒������A��������m�ۂ���
        strGuid = String$(guidLength, vbNullChar)
        ' guid�^�𕶎���^�ɕϊ�����
        retValue = StringFromGUID2(guid, StrPtr(strGuid), guidLength)
        
        If retValue = guidLength Then
            ' valid GUID as a string
            createGuid = strGuid
            If hasEnclose = False Then
                createGuid = replace(createGuid, "{", "")
                createGuid = replace(createGuid, "}", "")
            End If
        End If
    
    End If
    
End Function
