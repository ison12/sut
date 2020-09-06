Attribute VB_Name = "ExcelUtil"
Option Explicit

' *********************************************************
' Excel���ȕւɗ��p���邽�߂̃��[�e�B���e�B���W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2007/12/01�@�V�K�쐬
' �@�@�@�@�@2009/06/21�@Excel�̃o�[�W�����擾�֐����C��
' �@�@�@�@�@          �@����ɂ��AExcel2002������ɔF������Ȃ�(?)�o�O���C�����ꂽ�B
'
' ���L�����F
'
' *********************************************************

' Excel�̃N���X��
Private Const XLS_CLASSNAME As String = "XLMAIN"

' Excel2000�E2003�E2007�Ŋm�F�ς�
' �R�}���h�o�[�R���g���[��ID �t�H���g ���X�g
Private Const COMMAND_CONTROL_ID_FONT_LIST As Long = 1728

' Excel2000�E2003�E2007�Ŋm�F�ς�
' �R�}���h�o�[�R���g���[��ID �t�H���g�T�C�Y ���X�g
Private Const COMMAND_CONTROL_ID_FONT_SIZE As Long = 1731

' Excel���[�N�V�[�g�̋֎~����
Private Const EXCEL_SHEET_NAME_PROHIBITION_CHAR As String = "\[:]*/?"

' Excel���[�N�V�[�g�̃V�[�g���ő咷
Private Const EXCEL_SHEET_NAME_MAX_LENGTH As Long = 31

' Excel�̃o�[�W����
Public Enum ExcelVersion
    Ver2000 = 9
    Ver2002 = 10
    Ver2003 = 11
    Ver2007 = 12
    Ver2010 = 14
    Ver2013 = 15
    Ver2016 = 16
    VerOver = 99
    VerUnknown = -1
End Enum
    
' =========================================================
' ���G�N�Z���̃o�[�W�����擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�G�N�Z���̃o�[�W����
'
' =========================================================
Public Function getExcelVersion() As ExcelVersion

    Dim ver    As String
    
    ver = Application.version
    
    getExcelVersion = VerUnknown
    
    Select Case ver
    
        Case "9.0"
            getExcelVersion = Ver2000
        
        Case "10.0"
            getExcelVersion = Ver2002
        
        Case "11.0"
            getExcelVersion = Ver2003
        
        Case "12.0"
            getExcelVersion = Ver2007
        
        Case "14.0"
            getExcelVersion = Ver2010
        
        Case "15.0"
            getExcelVersion = Ver2013
        
        Case "16.0"
            getExcelVersion = Ver2016
            
    End Select
    
    ' ���l�ɕϊ��ł��邩�H
    If IsNumeric(ver) = False Then
    
        Exit Function
    End If
    
    Dim verSin    As Single
    verSin = CSng(ver)
    
    ' �}�C�i�[�o�[�W�������l�����Ĉȉ��̏��������s
    If getExcelVersion = VerUnknown Then
    
        If verSin >= 17 Then
        
            getExcelVersion = VerOver
        ElseIf verSin >= 16 Then
        
            getExcelVersion = Ver2016
        ElseIf verSin >= 15 Then
        
            getExcelVersion = Ver2013
        ElseIf verSin >= 14 Then
        
            getExcelVersion = Ver2010
        ElseIf verSin >= 12 Then
        
            getExcelVersion = Ver2007
        ElseIf verSin >= 11 Then
        
            getExcelVersion = Ver2003
        ElseIf verSin >= 10 Then
        
            getExcelVersion = Ver2002
        ElseIf verSin >= 9 Then
        
            getExcelVersion = Ver2000
        End If
    
    End If
End Function

' =========================================================
' ��Excel�A�v���P�[�V�����̃E�B���h�E�n���h���擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�E�B���h�E�n���h��
'
' =========================================================
#If VBA7 And Win64 Then

    Public Function getApplicationHWnd() As LongPtr
    
        Dim e As ExcelVersion
        
        Dim app As Object
        Set app = Excel.Application
        
        e = getExcelVersion
        
        If e >= Ver2002 Then
        
            getApplicationHWnd = app.hwnd
        
        Else
        
            getApplicationHWnd = WinAPI_User.FindWindow(XLS_CLASSNAME, Application.Caption)
        End If
        
    End Function
#Else

    Public Function getApplicationHWnd() As Long
    
        Dim e As ExcelVersion
        
        Dim app As Object
        Set app = Excel.Application
        
        e = getExcelVersion
        
        If e >= Ver2002 Then
        
            getApplicationHWnd = app.hwnd
        
        Else
        
            getApplicationHWnd = WinAPI_User.FindWindow(XLS_CLASSNAME, Application.Caption)
        End If
        
    End Function
#End If

' =========================================================
' ���E�B���h�E�őO�ʕ\��
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' �߂�l�@�@�F
'
' =========================================================
Public Sub setUserFormTopMost(ByVal form As Object, Optional ByVal topmost As Boolean = True)

    ' �E�B���h�E�n���h��
    #If VBA7 And Win64 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd As Long
    #End If

    Dim ret As Long

    ' �t�H�[���̃E�B���h�E�n���h�����擾����
    'ret = WinAPI_OLEACC.WindowFromAccessibleObject(form, hwnd)
    
    ' �t�H�[���L���v�V������ۑ����Ă���
    Dim formCaption As String: formCaption = form.Caption
    ' �󔒂�ݒ肵�ăt�H�[���L���v�V�������d�����Ȃ��悤�ɂ���
    form.Caption = formCaption & "                                "
    
    ' �t�H�[���̃E�B���h�E�n���h�����擾����
    hwnd = WinAPI_User.FindWindow("ThunderDFrame", form.Caption)
    
    ' �t�H�[���L���v�V���������ɖ߂�
    form.Caption = formCaption
    
    If hwnd <> 0 Then
    
        ' �t�H�[�����őO�ʕ\������
        If topmost Then
        
            ret = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            
        Else
            ret = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

        End If
    End If
    
End Sub

' =========================================================
' ���^�C�g���o�[���O�����E�B���h�E�X�^�C����ݒ肷��B
'
' �T�v�@�@�@�F
' �����@�@�@�FuForm ���[�U�[�t�H�[��
'
' �߂�l�@�@�Ftrue �����Afalse ���s
'
' =========================================================
Public Function setNonTitleBarWindowStyle(ByRef uForm As Object) As Boolean

    Dim ret As Long

    ' �E�B���h�E�n���h��
    #If VBA7 And Win64 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd As Long
    #End If
  
    ' �X�^�C���K�p�O�̃t�H�[���̃T�C�Y���擾����
    Dim formWidth  As Double
    Dim formHeight As Double
    
    formWidth = uForm.InsideWidth
    formHeight = uForm.InsideHeight
  
    ' �E�B���h�E�n���h�����擾����
    WinAPI_OLEACC.WindowFromAccessibleObject uForm, hwnd

    ' �_�C�A���O�̘g������
    ret = WinAPI_User.SetWindowLong(hwnd _
                      , GWL_EXSTYLE _
                      , WinAPI_User.GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME)
    If ret = 0 Then
        setNonTitleBarWindowStyle = False
    End If
    
    ' �^�C�g���o�[������
    ret = WinAPI_User.SetWindowLong(hwnd _
                      , GWL_STYLE _
                      , WinAPI_User.GetWindowLong(hwnd, GWL_STYLE) And Not WS_CAPTION)
    If ret = 0 Then
        setNonTitleBarWindowStyle = False
    End If
    
    ' ���j���[�o�[���ĕ`��
    ret = WinAPI_User.DrawMenuBar(hwnd)
    If ret = 0 Then
        setNonTitleBarWindowStyle = False
    End If
    
    ' �T�C�Y����
    uForm.Width = uForm.Width - uForm.InsideWidth + formWidth
    uForm.Height = uForm.Height - uForm.InsideHeight + formHeight
    
End Function

' =========================================================
' ���Z�����A�N�e�B�u�ɂ���
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̃Z�����A�N�e�B�u�ɂ���
' �����@�@�@�FsheetName �V�[�g��
'             row       �s
'             col       ��
'
' =========================================================
Public Sub activateCell(sheetName As String, row As Long, col As Long)

    Worksheets(sheetName).Cells(row, col).activate

End Sub

' =========================================================
' ���V�[�g�폜
'
' �T�v�@�@�@�F�C�ӂ̃u�b�N�̃V�[�g���폜����B�V�[�g�������ꍇ�͉������Ȃ��B
' �����@�@�@�FtargetBook  ���[�N�u�b�N
' �@�@�@�@�@�@targetSheet �폜�ΏۃV�[�g
'
'
' �߂�l�@�@�FTrue �폜����
'
' =========================================================
Public Function deleteSheet(ByRef targetBook As Workbook, ByVal targetSheet As String) As Boolean

    ' �V�[�g
    Dim sheet As Worksheet
    
    ' �폜�t���O
    deleteSheet = False
    
    ' �u�b�N���̃V�[�g�𑖍�����
    For Each sheet In targetBook.Worksheets
    
        ' �폜�ΏۃV�[�g���ǂ����𔻒f����
        If sheet.name = targetSheet Then
        
            ' �V�[�g���폜����
            sheet.delete
            
            ' �폜�t���O��ON�ɂ���
            deleteSheet = True
            
            Exit Function
        
        End If
    
    Next

End Function

' =========================================================
' ���V�[�g���R�s�[����
'
' �T�v�@�@�@�F�C�ӂ̃u�b�N�̃V�[�g���R�s�[���āA�V�������O��t����B
' �����@�@�@�FcopyBook       �R�s�[���u�b�N
' �@�@�@�@�@�@copySheetName  �R�s�[���V�[�g
' �@�@�@�@�@�@newBook        �R�s�[��u�b�N
'             newSheetName   �R�s�[��V�[�g
'             baseSheetName  �V�����V�[�g��z�u�����ƂȂ�V�[�g
'             direction      �V�����V�[�g��z�u�����ƂȂ�V�[�g�ɑ΂��đO���ɔz�u���邩����ɔz�u���邩
'
' =========================================================
Public Sub copySheet(ByRef copyBook As Workbook, _
                     ByRef copySheetName As String, _
                     ByRef newBook As Workbook, _
                     ByRef newSheetName As String, _
                     ByRef baseSheetName As String, _
                     Optional ByVal direction As String = "after")


    ' ����ɔz�u
    If direction = "after" Then
    
        copyBook.Worksheets(copySheetName).copy after:=newBook.Worksheets(baseSheetName)
        
    ' �O���ɔz�u
    Else
    
        copyBook.Worksheets(copySheetName).copy before:=newBook.Worksheets(baseSheetName)
    End If
    
    ' �R�s�[��͕K���A�N�e�B�u�V�[�g���V�����V�[�g�ɂȂ�
    ' �A�N�e�B�u�V�[�g�̖��O��ύX����
    ActiveSheet.name = newSheetName

End Sub

' =========================================================
' ���u�b�N�̍Ō���ɃV�[�g���R�s�[����
'
' �T�v�@�@�@�F�ڍׂ�copySheet���Q��
'
' =========================================================
Public Sub copySheetAppend(ByRef copyBook As Workbook _
                         , ByRef copySheetName As String _
                         , ByRef newBook As Workbook _
                         , ByRef newSheetName As String)
                     
    copySheet _
        copyBook _
      , copySheetName _
      , newBook _
      , newSheetName _
      , newBook.Worksheets(newBook.Worksheets.count).name
                     
End Sub

' =========================================================
' ���Z���R�s�[
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̃Z���͈͂��w�肵�āA���V�[�g�̕ʂ̃Z���͈͂ɓ\��t����
' �����@�@�@�FsheetName     �V�[�g��
'             srcStartRow   �R�s�[���J�n�s
'             srcStartCol   �R�s�[���J�n��
'             srcEndRow     �R�s�[���I���s
'             srcEndCol     �R�s�[���I����
'             desStartRow   �\�t����J�n�s
'             desStartCol   �\�t���������
'             pasteType     �R�s�[����l�̎�ށi�l�݂̂⏑���̂ݓ��j
'
' =========================================================
Public Sub copyCell(sheetName As String, _
                        srcStartRow As Long, _
                        srcStartCol As Long, _
                        srcEndRow As Long, _
                        srcEndCol As Long, _
                        desStartRow As Long, _
                        desStartCol As Long, _
                        Optional pasteType As Variant = xlPasteAll)

    Dim sheet    As Worksheet
    Dim srcRange As Range
    Dim desRange As Range
    
    
    Set sheet = Worksheets(sheetName)
    sheet.activate
    
    Set srcRange = sheet.Range( _
        sheet.Cells(srcStartRow, srcStartCol), _
        sheet.Cells(srcEndRow, srcEndCol) _
    )
        
    Set desRange = sheet.Range( _
        sheet.Cells(desStartRow, desStartCol), _
        sheet.Cells(desStartRow + srcEndRow - srcStartRow, desStartCol + srcEndCol - srcStartCol) _
    )
    
    srcRange.copy
    
    desRange.PasteSpecial Paste:=pasteType

    Application.CutCopyMode = False
    
End Sub

' =========================================================
' ���V�[�g�����݂��Ă��邩�ǂ����̊m�F
'
' �T�v�@�@�@�F
' �����@�@�@�Fbook          ���[�N�u�b�N
' �@�@�@�@�@�@sheetName     �V�[�g��
' �߂�l�@�@�FTrue �V�[�g�����݂���
'
' =========================================================
Public Function existSheet(ByRef book As Workbook _
                         , ByVal sheetName As String) As Boolean
                         
                         
    On Error GoTo err
    
    ' �V�[�g�I�u�W�F�N�g
    Dim sheet As Worksheet
    ' �V�[�g�I�u�W�F�N�g���擾����
    Set sheet = book.Worksheets(sheetName)
    
    ' ����ɏI���ł����ꍇ�A�V�[�g�͑��݂��Ă���
    existSheet = True
    
    Exit Function
err:

    ' �G���[�ɂȂ����ꍇ�A�V�[�g�͑��݂��Ă��Ȃ�
    existSheet = False

End Function

' =========================================================
' ����ӂȃV�[�g���ւ̕ϊ�
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�����ΏۂƂȂ�u�b�N�Ɋ��ɑ��݂���ꍇ
' �@�@�@�@�@�@��ӂȃV�[�g���ɕϊ����s���B
' �����@�@�@�Fbook          ���[�N�u�b�N
' �@�@�@�@�@�@sheetName     �V�[�g��
' �߂�l�@�@�F��ӂȃV�[�g���ւ̕ϊ�
'
' =========================================================
Public Function convertUniqSheetName(ByRef book As Workbook _
                                   , ByVal sheetName As String) As String
                         
                         
    On Error Resume Next
    
    Dim i As Long: i = 1
    
    ' �ϊ���̃V�[�g��
    Dim convertSheetName As String: convertSheetName = sheetName
    ' �ϊ���̃V�[�g���@�ڔ���
    Dim convertSheetNameSuffix As String
    
    ' �V�[�g�I�u�W�F�N�g
    Dim sheet As Worksheet
    
    ' 999��V�[�g���̕ϊ����s���Ă��A����ł��ϊ��ł��Ȃ��ꍇ�̓��[�v���I������
    Do While i < 1000
    
        i = i + 1
        
        ' �V�[�g�I�u�W�F�N�g���擾����
        Set sheet = book.Worksheets(convertSheetName)
    
        ' �G���[���������Ă��Ȃ������m�F
        If err.Number <> 0 Then
        
            ' �G���[���������Ă���i���V�[�g�͑��݂��Ȃ��j�̂ŏ����𔲂���
            convertUniqSheetName = convertSheetName
            
            Exit Function
            
        End If
    
        ' �ϊ���̃V�[�g�� �ڔ�����ݒ肷��
        convertSheetNameSuffix = " (" & i & ")"
        
        ' �V�[�g���̋K��̒����𒴂��Ă��Ȃ������m�F����
        If checkMaxLengthOfSheetName(sheetName & convertSheetNameSuffix) = False Then
        
            ' �K��̒����𒴂��Ă���ꍇ�A�����𒲐�����
            convertSheetName = Mid$(sheetName _
                                  , 1 _
                                  , EXCEL_SHEET_NAME_MAX_LENGTH - Len(convertSheetNameSuffix)) & convertSheetNameSuffix
                                  
        Else
        
            convertSheetName = sheetName & convertSheetNameSuffix
        End If
        
    
    Loop
    
    On Error GoTo 0
    
End Function

' =========================================================
' ���Z���R�s�[�i�����̂݁j
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̃Z���͈͂��w�肵�āA���V�[�g�̕ʂ̃Z���͈͂ɏ������݂̂�\��t����
' �����@�@�@�FsheetName     �V�[�g��
'             srcStartRow   �R�s�[���J�n�s
'             srcStartCol   �R�s�[���J�n��
'             srcEndRow     �R�s�[���I���s
'             srcEndCol     �R�s�[���I����
'             desStartRow   �\�t����J�n�s
'             desStartCol   �\�t����J�n��
'
' =========================================================
Public Sub copyCellFormat(sheetName As String, _
                        srcStartRow As Long, _
                        srcStartCol As Long, _
                        srcEndRow As Long, _
                        srcEndCol As Long, _
                        desStartRow As Long, _
                        desStartCol As Long)

    copyCell sheetName, _
             srcStartRow, _
             srcStartCol, _
             srcEndRow, _
             srcEndCol, _
             desStartRow, _
             desStartCol, _
             xlPasteFormats
    
End Sub

Public Sub fillBgColor(ByVal sheetName As String, _
                       ByVal startRow As Long, _
                       ByVal startCol As Long, _
                       ByVal endRow As Long, _
                       ByVal endCol As Long, _
                       ByVal colorIndex As Long)

    Dim r As Range
    Dim sheet As Worksheet
    Set sheet = Worksheets(sheetName)

    Set r = sheet.Range( _
                sheet.Cells(startRow, startCol), _
                sheet.Cells(endRow, endCol) _
            )

    ' �����̐F���m�F
    If r.Interior.colorIndex <> colorIndex Then
    
        ' �����̐F�ƐV�����F���قȂ�ꍇ�A�V�����F�ŏ㏑������
        r.Interior.colorIndex = colorIndex
    End If
    

End Sub

' =========================================================
' ���R�����g��ǉ�����
'
' �T�v�@�@�@�F�C�ӂ̃Z���ɃR�����g��ǉ�����
' �����@�@�@�Fsheet     �V�[�g��
'             row   �s
'             col   ��
'             text  �e�L�X�g
'
' =========================================================
Public Sub addCommentForWorkSheet(ByVal sheet As Worksheet, _
                      ByVal row As Long, _
                      ByVal col As Long, _
                      ByVal text As String)
    
    ' �����W�I�u�W�F�N�g
    Dim r As Range
    
    ' �����W�I�u�W�F�N�g���擾����
    Set r = sheet.Range( _
                sheet.Cells(row, col), _
                sheet.Cells(row, col) _
            )
    
    ' �R�����g�����ɑ��݂���Z���ɑ΂��āA�R�����g��ǉ�����ƃG���[����������B
    ' ���̂��߁A�G���[�������ɂ��������p������悤�ɂ���
    On Error Resume Next
    
    With r.Cells(1, 1)
    
        .ClearComments
        .addComment
        
        .comment.Shape.TextFrame.AutoSize = True
        .comment.Shape.TextFrame.Characters.Font.name = "�l�r �S�V�b�N"
        ' .Comment.Visible = True
        .comment.text text:=text
    
    End With
    
    On Error GoTo 0

End Sub

' =========================================================
' ���R�����g��ǉ�����
'
' �T�v�@�@�@�F�C�ӂ̃Z���ɃR�����g��ǉ�����
' �����@�@�@�FsheetName     �V�[�g��
'             row   �s
'             col   ��
'             text  �e�L�X�g
'
' =========================================================
Public Sub addComment(ByVal sheetName As String, _
                      ByVal row As Long, _
                      ByVal col As Long, _
                      ByVal text As String)
    
    addCommentForWorkSheet Worksheets(sheetName), row, col, text

End Sub

' =========================================================
' ���R�����g���폜����
'
' �T�v�@�@�@�F�C�ӂ̃Z���ɃR�����g���폜����
' �����@�@�@�FsheetName     �V�[�g��
'             row   �s
'             col   ��
'
' =========================================================
Public Sub deleteComment(sheetName As String, _
                         row As Long, _
                         col As Long)

    ' �����W�I�u�W�F�N�g
    Dim r As Range
    ' �R�����g�I�u�W�F�N�g
    Dim c As comment
    
    Dim sheet As Worksheet
    Set sheet = Worksheets(sheetName)
    
    ' �����W�I�u�W�F�N�g���擾����
    Set r = sheet.Range( _
                sheet.Cells(row, col), _
                sheet.Cells(row, col) _
            )
            
    ' �R�����g�I�u�W�F�N�g���擾����
    Set c = r.Cells(1, 1).comment
    
    ' �R�����g�����݂��邩�m�F����
    If Not c Is Nothing Then
    
        r.Cells(1, 1).ClearComments
    End If

End Sub

' =========================================================
' ���n�C�p�[�����N�ǉ�
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̃Z���Ƀn�C�p�[�����N��ǉ�����
' �����@�@�@�FsheetName           �V�[�g��
'             row                 �s
'             col                 ��
'             text                �e�L�X�g
'             linkTargetSheetName �����N��̃V�[�g
'             linkTargetCellRow   �����N��̃Z���i�s�j
'             linkTargetCellCol   �����N��̃Z���i��j
'             book                ���[�N�u�b�N
'
' =========================================================
Public Sub addHyperLinkInBook(ByVal sheetName As String, _
                              ByVal row As Long, _
                              ByVal col As Long, _
                              ByVal text As String, _
                              ByVal linkTargetSheetName As String, _
                              Optional ByVal linkTargetCellRow As Long = 1, _
                              Optional ByVal linkTargetCellCol As Long = 1, _
                              Optional ByRef book As Workbook)

    Dim sheet  As Worksheet
    Dim r      As Range
    Dim cellRc As String
    
    ' ����book���ȗ�����Ă����ꍇ�A�A�N�e�B�u�ȃ��[�N�u�b�N��ΏۂƂ���
    If book Is Nothing Then
    
        Set book = ActiveWorkbook
    End If
    
    Set sheet = book.Worksheets(sheetName)
    sheet.activate
    
    Set r = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col))
    
    ' R1C1�`���ŃZ���ʒu���w�肷��
    cellRc = "R" & linkTargetCellRow & "C" & linkTargetCellCol
    
    sheet.Hyperlinks.Add _
        anchor:=r, _
        Address:="", _
        SubAddress:="#" & linkTargetSheetName & "!" & cellRc, _
        TextToDisplay:=text


End Sub

' =========================================================
' ���t�H���g�T�C�Y��ύX����
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̃Z���̃t�H���g�T�C�Y��ύX����
' �����@�@�@�FsheetName  �V�[�g��
'             row        �s
'             col        ��
'             fontSize   �t�H���g�T�C�Y
'             book       ���[�N�u�b�N
'
' =========================================================
Public Sub changeFontSize(ByVal sheetName As String, ByVal row As Long, ByVal col As Long, fontSize As Long, Optional book As Workbook)

    Dim sheet As Worksheet
    Dim r     As Range
    
    ' ����book���ȗ�����Ă����ꍇ�A�A�N�e�B�u�ȃ��[�N�u�b�N��ΏۂƂ���
    If book Is Nothing Then
    
        Set book = ActiveWorkbook
    End If
    
    Set sheet = book.Worksheets(sheetName)
    sheet.activate
    
    Set r = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col))

    r.Font.size = fontSize
    
End Sub

' =========================================================
' ���Z��������ւ̕ϊ�����
'
' �T�v�@�@�@�F�C�ӂ̕�������Z��������ɕϊ�����
' �����@�@�@�Fval ������
'
' =========================================================
Public Function convertCellValue(ByRef val As Variant) As Variant

    ' �߂�l������������
    convertCellValue = val
    
    ' ������̐擪���V���O���N�H�[�e�[�V�����ł��邩���m�F����
    If Mid(val, 1, 1) = "'" Then
    
        ' ������̐擪���V���O���N�H�[�e�[�V�����̏ꍇ�A����ɃV���O���N�H�[�e�[�V������t������
        convertCellValue = "'" & val
        
    End If
    
End Function

' =========================================================
' ���Z��������ւ̕ϊ�����
'
' �T�v�@�@�@�F�C�ӂ̕�������Z��������ɕϊ�����
' �����@�@�@�Fval ������
'
' =========================================================
Public Function convertCellStrValue(ByRef val As Variant) As Variant

    ' �߂�l������������
    convertCellStrValue = val

    If isNull(val) Then
        Exit Function
    End If
    
    ' ������̐擪���V���O���N�H�[�e�[�V�����ł��邩���m�F����
    If Mid(val, 1, 1) = "'" Then
    
        ' ������̐擪���V���O���N�H�[�e�[�V�����̏ꍇ�A����ɃV���O���N�H�[�e�[�V������t������
        convertCellStrValue = "'" & val
        
    Else
    
        convertCellStrValue = CStr(val)
    End If
    
End Function

' =========================================================
' ���C�ӂ̍s������ŏI�s�܂ł̍s�폜
'
' �T�v�@�@�@�F�C�ӂ̍s������ŏI�s�܂ł̍s���폜����B
' �����@�@�@�Fsheet �C�ӂ̃V�[�g
' �@�@�@�@�@�@row   �C�ӂ̍s
' �@�@�@�@�@�@col   �C�ӂ̗�
' �߂�l�@�@�F
'
' =========================================================
Public Sub deleteRowEndOfLastInputted(ByRef sheet As Worksheet, ByVal row As Long, ByVal col As Long)

    ' �폜�Ώ۔͈�
    Dim targetRange As Range
    ' ���R�[�h�I�t�Z�b�g�ʒu
    Dim recordOffset As Long
    ' �Ō���̓��͉ӏ�
    Dim length As Long
    
    ' ���R�[�h�I�t�Z�b�g�ʒu
    recordOffset = row
    ' �Ō���̓��͉ӏ����擾����
    length = ExcelUtil.getCellEndOfLastInputtedRow(sheet, col)
    
    If length < recordOffset Then
        length = recordOffset
    End If
    
    ' �폜�Ώ۔͈͂��擾
    Set targetRange = sheet _
                        .Range( _
                           sheet.Cells(recordOffset _
                                     , 1).Address & ":" & _
                           sheet.Cells(length _
                                     , 1).Address)
    
    ' �폜����i�s�P�ʂō폜�j
    targetRange.EntireRow.delete

End Sub

' =========================================================
' ���L���s�擾
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̔C�ӂ̗�̗L���s���擾����B
' �@�@�@�@�@�@�i��ԍŌ���ɓ��͂���Ă���s�j
' �����@�@�@�Fsheet �C�ӂ̃V�[�g
' �@�@�@�@�@�@col   �C�ӂ̗�
' �߂�l�@�@�F�L���s
'
' =========================================================
Public Function getCellEndOfLastInputtedRow(ByRef sheet As Worksheet, ByVal col As Long) As Long

    ' �ő�s
    Dim max As Long
    ' �ő�s�T�C�Y���擾����
    max = getSizeOfSheetRow(sheet)

    ' �L���s�����߂�B
    If CStr(sheet.Cells(max, col)) <> "" Then
    
        ' Excel�̍ő�s���̃Z���ʒu�ɐݒ�l������ꍇ
        ' �ő�s���ʒu��Ԃ�
        getCellEndOfLastInputtedRow = max
    
    Else
        ' Excel�̍ő�s�����������ɋ󔒂łȂ��Z����T��
        getCellEndOfLastInputtedRow = sheet.Cells(max, col).End(xlUp).row
        
    End If
    
End Function

' =========================================================
' ���L����擾
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̔C�ӂ̗�̗L������擾����B
' �@�@�@�@�@�@�i��ԍŌ���ɓ��͂���Ă����j
' �����@�@�@�Fsheet �C�ӂ̃V�[�g
' �@�@�@�@�@�@row   �C�ӂ̍s
' �߂�l�@�@�F�L����
'
' =========================================================
Public Function getCellEndOfLastInputtedCol(ByRef sheet As Worksheet, ByVal row As Long) As Long
    
    ' �ő��
    Dim max As Long
    ' �ő��T�C�Y���擾����
    max = getSizeOfSheetCol(sheet)

    ' �L���s�����߂�B
    If CStr(sheet.Cells(row, max)) <> "" Then
    
        ' Excel�̍ő�s���̃Z���ʒu�ɐݒ�l������ꍇ
        ' �ő�s���ʒu��Ԃ�
        getCellEndOfLastInputtedCol = max
    
    Else
        ' Excel�̍ő�񐔂��獶�����ɋ󔒂łȂ��Z����T��
        getCellEndOfLastInputtedCol = sheet.Cells(row, max).End(xlToLeft).column
        
    End If
    
End Function

' =========================================================
' ���V�[�g�̍ő�s�T�C�Y�擾
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̍ő�s�T�C�Y���擾����B
' �����@�@�@�Fsheet �C�ӂ̃V�[�g
' �߂�l�@�@�F�ő�s�T�C�Y
'
' =========================================================
Public Function getSizeOfSheetRow(ByRef sheet As Worksheet) As Long

    #If EXCEL_SHEET_ROW_SIZE_256 = 1 Then
    
        getSizeOfSheetRow = 260
    #Else
    
        ' Range�I�u�W�F�N�g���Q�Ƃ��J�����S�̂�I�����iEntireColumn�j�s�̃J�E���g���擾����
        getSizeOfSheetRow = sheet.Range("A1").EntireColumn.Rows.count
    #End If

End Function

' =========================================================
' ���V�[�g�̍ő��T�C�Y�擾
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̍ő��T�C�Y���擾����B
' �����@�@�@�Fsheet �C�ӂ̃V�[�ga
' �߂�l�@�@�F�ő�s�T�C�Y
'
' =========================================================
Public Function getSizeOfSheetCol(ByRef sheet As Worksheet) As Long

    ' Range�I�u�W�F�N�g���Q�Ƃ��s�S�̂�I�����iEntireRow�j��̃J�E���g���擾����
    getSizeOfSheetCol = sheet.Range("A1").EntireRow.Columns.count

End Function

' =========================================================
' ���V�[�g���̋֎~�����`�F�b�N
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�ɋ֎~�������܂܂�Ă��邩���`�F�b�N����B
' �����@�@�@�FsheetName �C�ӂ̃V�[�g��
' �߂�l�@�@�FTrue ����i�֎~�������܂܂�Ă��Ȃ��ꍇ�j
'
' =========================================================
Public Function checkProhibitionCharOfSheetName(ByVal sheetName As String) As Boolean

    ' ������
    checkProhibitionCharOfSheetName = True

    ' �֎~�����i1�����j
    Dim char As String
    ' �C���f�b�N�X
    Dim i As Long
    
    ' �֎~������1���������o���A�V�[�g���ɋ֎~�������܂܂�Ă��Ȃ������`�F�b�N����
    For i = 1 To Len(EXCEL_SHEET_NAME_PROHIBITION_CHAR)
    
        ' 1�������o��
        char = Mid$(EXCEL_SHEET_NAME_PROHIBITION_CHAR, i, 1)
        
        ' �֎~�����𔭌������ꍇ
        If InStr(sheetName, char) <> 0 Then
        
            ' �֎~�������܂܂�Ă���̂�False�ɐݒ�
            checkProhibitionCharOfSheetName = False
            
            Exit Function
        End If
    
    Next

End Function

' =========================================================
' ���V�[�g���̋֎~�����ϊ�
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�ɋ֎~�������܂܂�Ă���ꍇ�A�ϊ������{����B
' �����@�@�@�FsheetName �C�ӂ̃V�[�g��
' �߂�l�@�@�F�ϊ���̃V�[�g��
'
' =========================================================
Public Function convertProhibitionCharOfSheetName(ByVal sheetName As String) As String

    ' �V�[�g���Ƃ��ėL���ȕ���
    Const VALID_CHAR As String = "_"

    ' �֎~�����i1�����j
    Dim char As String
    ' �C���f�b�N�X
    Dim i As Long
    
    ' �֎~������1���������o���A�V�[�g���ɋ֎~�������܂܂�Ă��Ȃ������`�F�b�N����
    For i = 1 To Len(EXCEL_SHEET_NAME_PROHIBITION_CHAR)
    
        ' 1�������o��
        char = Mid$(EXCEL_SHEET_NAME_PROHIBITION_CHAR, i, 1)
        
        ' �֎~�����𔭌������ꍇ
        If InStr(sheetName, char) <> 0 Then
        
            ' �֎~������L���ȕ����ɕϊ�����
            sheetName = replace(sheetName, char, VALID_CHAR)
        End If
    
    Next
    
    ' �߂�l�Ƃ��ĕԂ�
    convertProhibitionCharOfSheetName = sheetName

End Function

' =========================================================
' ���V�[�g���̕ϊ�
'
' �T�v�@�@�@�F�V�[�g�̖��̂��K��l�𒴂��Ă���ꍇ�ɁA�K��l�Ɏ��܂�悤�ɕϊ����s��
' �����@�@�@�FsheetName �C�ӂ̃V�[�g��
' �߂�l�@�@�F�ϊ���̃V�[�g��
'
' =========================================================
Public Function truncateExceededSheetName(ByVal sheetName As String) As String

    truncateExceededSheetName = Mid$(sheetName, 1, 28) & "..."
    
End Function

' =========================================================
' ���V�[�g���̃T�C�Y�`�F�b�N
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g���̌������K��T�C�Y�𒴂��Ă��Ȃ������`�F�b�N����B
' �����@�@�@�FsheetName �C�ӂ̃V�[�g��
' �߂�l�@�@�FTrue ����i�֎~�������܂܂�Ă��Ȃ��ꍇ�j
'
' =========================================================
Public Function checkMaxLengthOfSheetName(ByVal sheetName As String) As Boolean

    ' �߂�l
    Dim ret As Boolean

    ret = True
    
    ' �V�[�g���̌������ő啶�����𒴂��Ă���ꍇ
    If Len(sheetName) > EXCEL_SHEET_NAME_MAX_LENGTH Then
    
        ret = False
    End If

    ' �߂�l�Ɍ��ʂ�ݒ�
    checkMaxLengthOfSheetName = ret
    
End Function

' =========================================================
' ���ő�s�T�C�Y�𒴂��Ă��邩���`�F�b�N
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̍ő�s�T�C�Y�������Ă��邩���`�F�b�N����B
' �����@�@�@�Fsheet     �V�[�g�I�u�W�F�N�g
' �@�@�@�@�@�@rowOffset �s�I�t�Z�b�g
' �@�@�@�@�@�@rowSize   �s�T�C�Y
' �߂�l�@�@�FTrue  �ő�s�T�C�Y�͈͓̔�
' �@�@�@�@�@�@False �ő�s�T�C�Y�͈̔͊O
'
' =========================================================
Public Function checkOverMaxRow(ByRef sheet As Worksheet _
                              , ByVal rowOffset As Long _
                              , Optional ByVal rowSize As Long = 1) As Boolean

    ' �V�[�g�̍ő�s
    Dim max As Long
    ' �V�[�g�̍ő�s���擾
    max = getSizeOfSheetRow(sheet)
    
    ' �ő�s�𒴂��Ă��邩���`�F�b�N
    If max < rowOffset + rowSize - 1 Then
    
        ' �ő�s�𒴂��Ă���̂� False
        checkOverMaxRow = False
    
    Else
    
        ' �ő�s�𒴂��Ă��Ȃ��̂� True
        checkOverMaxRow = True
    End If

End Function

' =========================================================
' ���ő��T�C�Y�𒴂��Ă��邩���`�F�b�N
'
' �T�v�@�@�@�F�C�ӂ̃V�[�g�̍ő��T�C�Y�������Ă��邩���`�F�b�N����B
' �����@�@�@�Fsheet     �V�[�g�I�u�W�F�N�g
' �@�@�@�@�@�@colOffset ��I�t�Z�b�g
' �@�@�@�@�@�@colSize   ��T�C�Y
' �߂�l�@�@�FTrue  �ő��T�C�Y�͈͓̔�
' �@�@�@�@�@�@False �ő��T�C�Y�͈̔͊O
'
' =========================================================
Public Function checkOverMaxCol(ByRef sheet As Worksheet _
                              , ByVal colOffset As Long _
                              , Optional ByVal colSize As Long = 1) As Boolean

    ' �V�[�g�̍ő�s
    Dim max As Long
    ' �V�[�g�̍ő�s���擾
    max = getSizeOfSheetCol(sheet)
    
    ' �ő�s�𒴂��Ă��邩���`�F�b�N
    If max < colOffset + colSize - 1 Then
    
        ' �ő�s�𒴂��Ă���̂� False
        checkOverMaxCol = False
    
    Else
    
        ' �ő�s�𒴂��Ă��Ȃ��̂� True
        checkOverMaxCol = True
    End If

End Function

' =========================================================
' ���z�񂩂�Range���擾����
'
' �T�v�@�@�@�F
' �����@�@�@�Fval       �z��
'             sheet     �V�[�g
' �@�@�@�@�@�@offsetRow �s
' �@�@�@�@�@�@offsetCol ��
' �@�@�@�@�@�@rowSize   �s�T�C�Y
' �@�@�@�@�@�@colSize   ��T�C�Y
'
' =========================================================
Public Function getArrayRange(ByRef val As Variant _
                            , ByRef sheet As Worksheet _
                            , ByVal rowOffset As Long _
                            , ByVal colOffset As Long _
                            , Optional ByVal rowSize As Long = -1 _
                            , Optional ByVal colSize As Long = -1) As Range

    If IsArray(val) = False Then
    
        Exit Function
        
    End If

    If rowSize = -1 Then
        rowSize = VBUtil.arraySize(val)
    End If
    
    If colSize = -1 Then
        colSize = VBUtil.arraySize(val, 2)
    End If

    Set getArrayRange = sheet.Range(sheet.Cells(rowOffset _
                                              , colOffset) _
                      , sheet.Cells(rowOffset + rowSize - 1 _
                                              , colOffset + colSize - 1))
    
End Function

' =========================================================
' ���z����e�R�s�[
'
' �T�v�@�@�@�F�Z���ɔz����e���R�s�[����
' �����@�@�@�Fval       �z��
'             sheet     �V�[�g
' �@�@�@�@�@�@offsetRow �s
' �@�@�@�@�@�@offsetCol ��
' �@�@�@�@�@�@rowSize   �s�T�C�Y
' �@�@�@�@�@�@colSize   ��T�C�Y
'
' =========================================================
Public Sub copyArrayToCells(ByRef val As Variant _
                          , ByRef sheet As Worksheet _
                          , ByVal rowOffset As Long _
                          , ByVal colOffset As Long _
                          , Optional ByVal rowSize As Long = -1 _
                          , Optional ByVal colSize As Long = -1)

    If IsArray(val) = False Then
    
        Exit Sub
        
    End If

    If rowSize = -1 Then
        rowSize = VBUtil.arraySize(val)
    End If
    
    If colSize = -1 Then
        colSize = VBUtil.arraySize(val, 2)
    End If

    sheet.Range(sheet.Cells(rowOffset _
                          , colOffset) _
              , sheet.Cells(rowOffset + rowSize - 1 _
                          , colOffset + colSize - 1)) = val

    'sheet.Cells(rowOffset, colOffset).Resize(rowSize, colSize) = val
    
End Sub

' =========================================================
' ���z����e�R�s�[�i�J�������p�j
'
' �T�v�@�@�@�F�Z���ɔz����e���R�s�[����
' �����@�@�@�Fval       �z��
'             sheet     �V�[�g
' �@�@�@�@�@�@rowOffset �s
' �@�@�@�@�@�@colOffset ��
' �@�@�@�@�@�@colSize   ��T�C�Y
'
' =========================================================
Public Sub copyArrayToCellsForColumns(ByRef val As Variant _
                                    , ByRef sheet As Worksheet _
                                    , ByVal rowOffset As Long _
                                    , ByVal colOffset As Long _
                                    , Optional ByVal colSize As Long = -1)

    If IsArray(val) = False Then
    
        Exit Sub
        
    End If

    Dim rowSize As Long: rowSize = 1

    If colSize = -1 Then
        colSize = VBUtil.arraySize(val)
    End If

    sheet.Range(sheet.Cells(rowOffset _
                          , colOffset) _
              , sheet.Cells(rowOffset + rowSize - 1 _
                          , colOffset + colSize - 1)) = val

    'sheet.Cells(rowOffset, colOffset).Resize(rowSize, colSize) = val
    
End Sub

' =========================================================
' ���z����e�R�s�[
'
' �T�v�@�@�@�F�Z���ɔz����e���R�s�[����
' �����@�@�@�Fval       �z��
'             sheet     �V�[�g
' �@�@�@�@�@�@rowOffset �s
' �@�@�@�@�@�@rowSize   �s�T�C�Y
' �@�@�@�@�@�@colOffset ��
' �@�@�@�@�@�@colSize   ��T�C�Y
'
' =========================================================
Public Function copyCellsToArray(ByRef sheet As Worksheet _
                               , ByVal rowOffset As Long _
                               , ByVal rowSize As Long _
                               , ByVal colOffset As Long _
                               , ByVal colSize As Long) As Variant

    Dim retArray As Variant
    Dim ret      As Variant

    Dim srcCell As String
    
    With sheet
    
        srcCell = .Cells(rowOffset _
                       , colOffset).Address & ":" & _
                  .Cells(rowOffset + rowSize - 1 _
                       , colOffset + colSize - 1).Address
                  
        ret = .Range(srcCell)

    End With
    
    ' �߂�l���z��ł͂Ȃ��ꍇ
    If IsArray(ret) = False Then
        
        ' ��Range�I�u�W�F�N�g�̃T�C�Y��1�̏ꍇ�A�z��ȊO�̃v���~�e�B�u�^���Ԃ��Ă���̂�
        ' �@�ϊ�����K�v������
        ' �T�C�Y��1�̔z��𐶐�����
        ReDim retArray(1 To 1, 1 To 1)
    
        ' �l��������
        retArray(1, 1) = ret
        
    Else
    
        retArray = ret
    End If
    
    copyCellsToArray = retArray
    
End Function

' =========================================================
' ���s�̍�����ύX
'
' �T�v�@�@�@�F�s�̍�����ύX����B
' �@�@�@�@�@�@�����̒P�ʂ̓|�C���g�B(Excel�̎d�l�ɏ���)
' �����@�@�@�Fr    �����W�I�u�W�F�N�g�i�V�[�g�I�u�W�F�N�g�܂ށj
' �@�@�@�@�@�@h    ����
'
' =========================================================
Public Sub changeRowHeight(ByRef r As Range _
                         , ByVal h As Double)

    If h = -1 Then
    
        r.EntireRow.AutoFit
    Else
    
        r.EntireRow.RowHeight = h
    End If

End Sub

' =========================================================
' ����̕���ύX
'
' �T�v�@�@�@�F��̕���ύX
' �@�@�@�@�@�@���̒P�ʂ͕������B(Excel�̎d�l�ɏ���)
' �����@�@�@�Fr    �����W�I�u�W�F�N�g�i�V�[�g�I�u�W�F�N�g�܂ށj
' �@�@�@�@�@�@w    ��
'
' =========================================================
Public Sub changeColWidth(ByRef r As Range _
                        , ByVal w As Double)

    If w = -1 Then
    
        r.EntireColumn.AutoFit
    Else
    
        r.EntireColumn.ColumnWidth = w
    End If

End Sub

' =========================================================
' �����p�\�ȃt�H���g���X�g���擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' ���L�����@�FExcel�u�b�N��1�ȏ�J����Ă��Ȃ��Ǝ擾�Ɏ��s����̂Œ���
'
' =========================================================
Public Function getFontList() As ValCollection

    ' �߂�l
    Dim ret As New ValCollection

    Dim i As Long
    
    ' �R�}���h�o�[�R���g���[��
    Dim c As commandBarControl

    ' �t�H���g�T�C�Y���X�g���擾����
    Set c = Application.CommandBars.FindControl(Id:=COMMAND_CONTROL_ID_FONT_LIST)
    
    ' �R���g���[�����擾�ł����ꍇ
    If Not c Is Nothing Then
    
        ' ���X�g�̓��e��S�Ė߂�l�ɒǉ�
        For i = 1 To c.ListCount
        
            ret.setItem c.list(i), c.list(i)
        
        Next

    End If
    
    ' �߂�l��Ԃ�
    Set getFontList = ret

End Function

' =========================================================
' ���t�H���g�T�C�Y���X�g���擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' ���L�����@�FExcel�u�b�N��1�ȏ�J����Ă��Ȃ��Ǝ擾�Ɏ��s����̂Œ���
'
' =========================================================
Public Function getFontSizeList() As ValCollection

    ' �߂�l
    Dim ret As New ValCollection

    Dim i As Long
    ' �R�}���h�o�[�R���g���[��
'    Dim c As CommandBarControl

'    ' �t�H���g�T�C�Y���X�g���擾����
'    Set c = Application.CommandBars.FindControl(ID:=COMMAND_CONTROL_ID_FONT_SIZE)
'
'    ' �R���g���[�����擾�ł����ꍇ
'    If Not c Is Nothing Then
'
'        Debug.Print TypeName(c)
'        ' ���X�g�̓��e��S�Ė߂�l�ɒǉ�
'        For i = 1 To c.ListCount
'
'            ret.setItem c.list(i), c.list(i)
'
'        Next
'
'    End If
    
    ' Excel2000 - 2007�̋K��l�̃T�C�Y���Z�b�g����
    ret.setItem 6
    ret.setItem 8
    ret.setItem 9
    ret.setItem 10
    ret.setItem 11
    ret.setItem 12
    ret.setItem 14
    ret.setItem 16
    ret.setItem 18
    ret.setItem 20
    ret.setItem 22
    ret.setItem 24
    ret.setItem 26
    ret.setItem 28
    ret.setItem 36
    ret.setItem 48
    ret.setItem 72
    
    
    ' �߂�l��Ԃ�
    Set getFontSizeList = ret

End Function

' =========================================================
' ��Excel�s��̐��l���A���t�@�x�b�g�ɕϊ�����
'
' �T�v�@�@�@�FExcel�s��̐��l���A���t�@�x�b�g�ɕϊ�����
' �����@�@�@�Frow �s
' �����@�@�@�Fcol ��
' �߂�l�@�@�F�ϊ�����
'
' =========================================================
Public Function convertExcelNumberToAlpha(ByVal row As Long, ByVal col As Long) As String

    ' �߂�l
    Dim ret As String

    ' �[���x�[�X�Ƃ���
    row = row - 1
    col = col - 1

    ' ��� 26 �Ƃ��� �i�A���t�@�x�b�g�̐��j
    Dim base As Long: base = 26
    
    ' ���݊�ׂ��搔�l�ibase��n�搔�j
    Dim curBase As Long

    ' ��ɂ����鐔�l�̒���
    Dim length As Long
    
    Dim i   As Long
    Dim tmp As Long

    ' ��̑ΐ������߂�
    If col > 0 Then
        length = Application.WorksheetFunction.RoundDown(Application.WorksheetFunction.Log(col, base), 0)
    Else
        length = 0
    End If
    
    For i = length To 0 Step -1
    
        ' ���݂̊�����Ƃɂ��������́A�J�n���l�����߂�
        ' ��Flength = 1 �� 26
        curBase = Application.WorksheetFunction.Power(base, i)
        If curBase <> 1 Then
        
            tmp = Application.WorksheetFunction.RoundDown(col / curBase, 0) - 1
        Else
            
            tmp = col - curBase + 1
        End If
        
        ' Chr(65) = A �Ȃ̂ŁAA����ɃA���t�@�x�b�g���Z�o����
        ret = ret & Chr(65 + tmp)
        
        ' ���̌v�Z�ɔ�����l�����Z����
        col = col - (curBase * (tmp + 1))
    
    Next
    
    convertExcelNumberToAlpha = ret & "" & (row + 1)
    
End Function

' =========================================================
' ��Excel��̃A���t�@�x�b�g�𐔒l�ɕϊ�����
'
' �T�v�@�@�@�FExcel��̃A���t�@�x�b�g�𐔒l�ɕϊ�����
' �����@�@�@�Fvar �l
' �߂�l�@�@�F�ϊ�����
'
' =========================================================
Public Function convertExcelAlphaToNumber(ByVal var As String) As Long

    ' ���K�\���I�u�W�F�N�g
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    
    RE.Pattern = "^[A-Z]$" ' �A���t�@�x�b�g�������ΏۂƂ���
    RE.IgnoreCase = True   ' �啶���Ə���������ʂ��Ȃ�
    RE.Global = True       ' ������S�̂�����
    
    Dim keta As Long
    
    Dim i As Long
    Dim c As String
    Dim length As Long: length = Len(var)
    
    For i = length To 1 Step -1
    
        ' 1�����擾����
        c = UCase$(Mid$(var, i, 1))
        
        ' �A���t�@�x�b�g�̏ꍇ
        If RE.test(c) Then
        
            convertExcelAlphaToNumber = convertExcelAlphaToNumber + (Asc(c) - Asc("A")) + (26 ^ keta)
            
            keta = keta + 1
        End If
    
    Next
    
    convertExcelAlphaToNumber = convertExcelAlphaToNumber
    
End Function

Public Sub protectSheet(ByVal sheetName As String)

    With Worksheets(sheetName)
    
        ' ��U�A�V�[�g�ی������
        .Unprotect
        ' �V�[�g�ی��ݒ�
        .Protect _
            UserInterfaceOnly:=True, _
            contents:=True, _
            Scenarios:=True, _
            AllowFiltering:=True

        .EnableSelection = xlUnlockedCells
        
    End With
    
End Sub

Public Sub copyCommandBarControl(ByRef srcControl As Object, ByRef desControl As Object)

    ' �G�N�Z���̃o�[�W����
    Static excelVer As ExcelVersion: excelVer = ExcelUtil.getExcelVersion

    With desControl
    
        .Style = srcControl.Style
        .Caption = srcControl.DescriptionText
        .DescriptionText = srcControl.DescriptionText
        .OnAction = srcControl.OnAction
        .Tag = srcControl.Tag
        .ShortcutText = srcControl.ShortcutText
        
        ' Excel2002�ȍ~�̃v���p�e�B
        If excelVer >= Ver2002 Then
            .Picture = srcControl.Picture
            .mask = srcControl.mask
        End If
    
    End With
    
End Sub

Public Function showSaveConfirmDialog(book As Workbook) As VbMsgBoxResult

    showSaveConfirmDialog = MsgBox("'" & book.name & "'�ւ̕ύX��ۑ����܂����H", vbYesNoCancel Or vbExclamation, "Microsoft Excel")

End Function

' =========================================================
' ���A�N�e�B�u�u�b�N���擾����B
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�A�N�e�B�u�u�b�N
'
' =========================================================
Public Function getActiveWorkbook() As Workbook

    On Error Resume Next
    
    ' �ŏ��̓A�N�e�B�u�V�[�g����擾�����݂�
    ' ���A�h�C���}�N���̏ꍇ�ɁA�A�h�C�����g�̃u�b�N��񂪎擾�����\�������邽��
    Set getActiveWorkbook = ActiveSheet.parent
    
    If err.Number <> 0 Then
    
        ' ���ɃA�N�e�B�u�u�b�N����擾�����݂�
        Set getActiveWorkbook = ActiveWorkbook
    
    End If
    
    On Error GoTo 0

End Function

