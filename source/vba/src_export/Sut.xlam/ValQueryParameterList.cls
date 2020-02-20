VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryParameterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �N�G���p�����[�^���X�g��\������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/08�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���X�g
Private m_list    As ValCollection
' �u�b�N
Private m_book    As Workbook

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    Set m_list = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    Set m_list = Nothing
End Sub

' =========================================================
' ���������������\�b�h
' =========================================================
Public Sub init(ByVal book As Workbook)
    Set m_book = book
End Sub

' =========================================================
' ���v���p�e�B
' =========================================================
Public Property Get list() As ValCollection
    Set list = m_list
End Property

Public Property Let list(ByVal newValue As ValCollection)
    Set m_list = newValue
End Property

' =========================================================
' ���ݒ���̐���
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initWorksheet m_book, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "queryParameter" & ".ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ���N�G���p�����[�^����ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub readForData()
    
    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties

    ' �f�[�^��ǂݍ���
    Dim val As Variant
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
            
    Dim queryParameterList As New ValCollection
    Dim queryParameter As ValQueryParameter
    
    Dim i As Long: i = 1
    Do While True
    
        val = values.getItem(i & "_" & "name", vbVariant)
        If Not IsArray(val) Then
            Exit Do
        End If
        
        Set queryParameter = New ValQueryParameter
                    
        val = values.getItem(i & "_" & "name", vbVariant): If IsArray(val) Then queryParameter.name = val(2)
        val = values.getItem(i & "_" & "value", vbVariant): If IsArray(val) Then queryParameter.value = val(2)
        
        queryParameterList.setItem queryParameter
    
        i = i + 1
    Loop
    
    list = queryParameterList

End Sub

' =========================================================
' ���N�G���p�����[�^����ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub writeForData()

    ' �A�v���P�[�V�����v���p�e�B�𐶐�����
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties
    
    ' �������݃f�[�^
    Dim val As Variant
    Dim values As New ValCollection
    
    Dim i As Long: i = 1
    For Each val In list.col
    
        values.setItem Array(i & "_" & "name", val.name)
        values.setItem Array(i & "_" & "value", val.value)
    
        i = i + 1
    Next

    ' �f�[�^����������
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData
    
End Sub