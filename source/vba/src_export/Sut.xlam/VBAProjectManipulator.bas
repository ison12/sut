Attribute VB_Name = "VBAProjectManipulator"
Option Explicit

' *********************************************************
' �A�h�C���u�b�N�̃��W���[�����G�N�X�|�[�g�E�C���|�[�g����@�\�B
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/02/17�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

' =========================================================
' ���S�t�@�C�����G�N�X�|�[�g
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath �t�@�C���p�X
'
' �߂�l�@�@�F
'
' =========================================================
Public Sub exportAll(Optional ByVal filePath As String = "")

    Dim module                  As Object      ' ���W���[��
    Dim moduleList              As Object      ' VBA�v���W�F�N�g�̑S���W���[��
    Dim extension               As String      ' ���W���[���̊g���q
    
    Dim exportFilePath          As String      ' �G�N�X�|�[�g�t�@�C���p�X
    
    Dim targetBook              As Workbook    ' �����Ώۃu�b�N�I�u�W�F�N�g
    
    Set targetBook = ThisWorkbook
    
    If filePath = "" Then
        filePath = ThisWorkbook.path & "\module"
    End If
    
    ' �G�N�X�|�[�g��̃f�B���N�g������U�폜���čēx�쐬����
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.folderexists(filePath) = True Then
        fso.DeleteFolder filePath
        fso.CreateFolder filePath
    Else
        fso.CreateFolder filePath
    End If
    
    ' �����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = targetBook.VBProject.VBComponents
    
    ' VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList
    
        ' �N���X
        If (module.Type = 2) Then
            extension = "cls"
        ' �t�H�[��
        ElseIf (module.Type = 3) Then
            ' .frx���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"
        ' �W�����W���[��
        ElseIf (module.Type = 1) Then
            extension = "bas"
        ' ���̑�
        Else
            ' �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo continue
        End If
        
        ' �G�N�X�|�[�g���{
        exportFilePath = filePath & "\" & module.name & "." & extension
        Call module.export(exportFilePath)
        
        ' �o�͐�m�F�p���O�o��
        Debug.Print exportFilePath
        
continue:

    Next
    
End Sub

' =========================================================
' ���S�t�@�C�����C���|�[�g�܂��͍폜
'
' �T�v�@�@�@�F
' �����@�@�@�FfilePath     �t�@�C���p�X
'           : isDeleteOnly �폜�̂݃t���O
'
' �߂�l�@�@�F
'
' =========================================================
Public Sub importOrDeleteAll(Optional ByVal filePath As String = "", Optional ByVal isDeleteOnly As Boolean = False)
    
    On Error Resume Next
    
    ' �Ώۃu�b�N
    Dim book As Workbook
    Set book = ThisWorkbook
    
    Dim oFso As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    Dim moduleList()     As String       ' ���W���[���t�@�C���z��
    Dim module           As Variant      ' ���W���[���t�@�C��
    Dim extension        As String       ' �g���q
    
    If filePath = "" Then
        filePath = ThisWorkbook.path & "\module"
    End If
    
    ReDim moduleList(0)
    
    ' �S���W���[���̃t�@�C���p�X���擾
    Call searchAllFile(filePath, moduleList)
    
    ' �S���W���[�������[�v
    For Each module In moduleList
        
        ' �g���q���������Ŏ擾
        extension = LCase(oFso.GetExtensionName(module))
        
        ' �g���q��cls�Afrm�Abas�̂����ꂩ�̏ꍇ
        If (extension = "cls" Or extension = "frm" Or extension = "bas") Then
            
            If oFso.getbasename(module) <> "VBAProjectManipulator.bas" Then
            
                ' �������W���[�����폜
                Call book.VBProject.VBComponents.remove(book.VBProject.VBComponents(oFso.getbasename(module)))
                
                If isDeleteOnly = False Then
                    ' ���W���[����ǉ�
                    Call book.VBProject.VBComponents.Import(module)
                End If
                ' �m�F�p���O�o��
                Debug.Print module
            End If
        
        End If
    Next
    
End Sub

' =========================================================
' ���C�ӂ̃f�B���N�g�����̑S�t�@�C�����ċA�I�Ɍ�������
'
' �T�v�@�@�@�F
' �����@�@�@�FdirPath  �f�B���N�g���p�X
'     �@�@�@�FfileList �t�@�C�����X�g
'
' �߂�l�@�@�F
'
' =========================================================
Private Sub searchAllFile(dirPath As String, fileList() As String)
    
    Dim oFso        As Object
    Dim oFolder     As Object
    Dim oSubFolder  As Object
    Dim oFile       As Object
    
    Dim i
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    ' �t�H���_���Ȃ��ꍇ
    If (oFso.folderexists(dirPath) = False) Then
        Exit Sub
    End If
    
    Set oFolder = oFso.GetFolder(dirPath)
    
    ' �T�u�t�H���_���ċA�i�T�u�t�H���_��T���K�v���Ȃ��ꍇ�͂���For�����폜���Ă��������j
    For Each oSubFolder In oFolder.SubFolders
        Call searchAllFile(oSubFolder.path, fileList)
    Next
    
    i = UBound(fileList)
    
    ' �J�����g�t�H���_���̃t�@�C�����擾
    For Each oFile In oFolder.Files
    
        If (i <> 0 Or fileList(i) <> "") Then
            i = i + 1
            ReDim Preserve fileList(i)
        End If
        
        ' �t�@�C���p�X��z��Ɋi�[
        fileList(i) = oFile.path
    Next
    
End Sub
