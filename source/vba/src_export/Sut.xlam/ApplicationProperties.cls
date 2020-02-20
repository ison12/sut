VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' INI�t�@�C������N���X
'
' ��������
' �E�Z�N�V�����̏ȗ��s��
' �E�R�����g�ݒ�s��
' �E�f�[�^�{�����A���s�R�[�h��C�R�[���L��������Ȑ���R�[�h�ɕϊ������
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/02/13�@�V�K�쐬
'
' ���L�����F
' *********************************************************
' INI�f�[�^���[�N�V�[�g
Private iniw As IniWorksheet
' INI�f�[�^�t�@�C��
Private inif As IniFile

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Public Sub Class_Initialize()
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Public Sub Class_Terminate()
    
    destroy
    
End Sub

' =========================================================
' �������������B
'
' �T�v�@�@�@�F
' �����@�@�@�Fbook_      �u�b�N
'     �@�@�@�FsheetName_ �V�[�g��
'     �@�@�@�FfileName_  �t�@�C����
' �߂�l�@�@�F
'
' =========================================================
Public Sub initWorksheet(ByVal book_ As Workbook, ByVal sheetName_ As String, ByVal fileName_ As String)

    Set iniw = New IniWorksheet
    iniw.init book_, sheetName_, fileName_

End Sub

' =========================================================
' �������������B
'
' �T�v�@�@�@�F
' �����@�@�@�FfileName_ �t�@�C����
' �߂�l�@�@�F
'
' =========================================================
Public Sub initFile(ByVal fileName_ As String)

    Set inif = New IniFile
    inif.init fileName_

End Sub

' =========================================================
' ���j�������B
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub destroy()

    If Not iniw Is Nothing Then
        iniw.destroy
    End If

    If Not inif Is Nothing Then
        inif.destroy
    End If

End Sub

' =========================================================
' ���ǂݍ��ݏ����B
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub readData()

    If Not iniw Is Nothing Then
        iniw.readSheet
    End If

    If Not inif Is Nothing Then
        inif.readFile
    End If

End Sub

' =========================================================
' ���������ݏ����B
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub writeData()

    If Not iniw Is Nothing Then
        iniw.writeSheet
    End If

    If Not inif Is Nothing Then
        inif.writeFile
    End If

End Sub

' ------------------------------------------------------------------
' IniData�����b�v���Ċe���\�b�h���Ăяo��
' ------------------------------------------------------------------
Public Function getValue _
            (ByVal SECTION As String, _
             ByVal key As String) As String

    If Not iniw Is Nothing Then
        If iniw.isExistsData Then
            getValue = iniw.getValue(SECTION, key)
            Exit Function
        End If
    End If

    If Not inif Is Nothing Then
        getValue = inif.getValue(SECTION, key)
        Exit Function
    End If

End Function

Public Sub setValue _
            (ByVal SECTION As String, _
             ByVal key As String, _
             ByVal value As String)

    If Not iniw Is Nothing Then
        iniw.setValue SECTION, key, value
    End If

    If Not inif Is Nothing Then
        inif.setValue SECTION, key, value
    End If

End Sub

Public Function getValues _
            (ByVal SECTION As String) As ValCollection

    If Not iniw Is Nothing Then
        If iniw.isExistsData Then
            Set getValues = iniw.getValues(SECTION)
            Exit Function
        End If
    End If

    If Not inif Is Nothing Then
        Set getValues = inif.getValues(SECTION)
        Exit Function
    End If
    
    Set getValues = New ValCollection
    
End Function

Public Sub setValues _
            (ByVal SECTION As String, _
             ByRef values As ValCollection)

    If Not iniw Is Nothing Then
        iniw.setValues SECTION, values
    End If

    If Not inif Is Nothing Then
        inif.setValues SECTION, values
    End If

End Sub

Public Sub delete _
            (ByVal SECTION As String, _
             Optional ByVal key As String = "")

    If Not iniw Is Nothing Then
        iniw.delete SECTION, key
    End If

    If Not inif Is Nothing Then
        inif.delete SECTION, key
    End If

End Sub



