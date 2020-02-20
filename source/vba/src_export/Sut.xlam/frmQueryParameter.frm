VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryParameter 
   Caption         =   "�N�G���p�����[�^�ݒ�"
   ClientHeight    =   8595.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8055
   OleObjectBlob   =   "frmQueryParameter.frx":0000
End
Attribute VB_Name = "frmQueryParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' �N�G���p�����[�^��`�t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/04�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event ok()

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' �N�G���p�����[�^�̐V�K�쐬�ő吔
Private Const QUERY_PARAMETER_NEW_CREATED_OVER_SIZE As String = "�N�G���p�����[�^�͍ő�${count}�܂œo�^�\�ł��B"

' �N�G���p�����[�^�ݒ���̈ꌏ���̕ҏW�i�q��ʁj
Private WithEvents frmQueryParameterSettingVar As frmQueryParameterSetting
Attribute frmQueryParameterSettingVar.VB_VarHelpID = -1

' �N�G���p�����[�^���X�g �R���g���[��
Private queryParameterList As CntListBox

' �N�G���p�����[�^���X�g�ł̑I�����ڃC���f�b�N�X
Private queryParameterSelectedIndex As Long
' �N�G���p�����[�^���X�g�ł̑I�����ڃI�u�W�F�N�g
Private queryParameterSelectedItem As ValQueryParameter

' �Ώۃu�b�N
Private targetBook As Workbook
' �Ώۃu�b�N���擾����
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal ���[�_���܂��̓��[�h���X�\���w��
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants)

    activate
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal
End Sub

' =========================================================
' ���t�H�[����\��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub HideExt()

    deactivate
    Main.storeFormPosition Me.name, Me
    Me.Hide

End Sub

' =========================================================
' ���t�H�[���A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub activate()

    restoreQueryParameter
    
    lblDescription.Caption = replace(replace(lblDescription.Caption, "$es", ConstantsTable.QUERY_PARAMETER_ENCLOSE_START), "$ee", ConstantsTable.QUERY_PARAMETER_ENCLOSE_END)
    
End Sub

' =========================================================
' ���t�H�[���f�B�A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub deactivate()

    ' Nothing��ݒ肷�邱�ƂŃC�x���g����M���Ȃ��悤�ɂ���
    Set frmQueryParameterSettingVar = Nothing

End Sub

' =========================================================
' ��OK�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err
    
    ' �����L�^����
    storeQueryParameter
    
    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���L�����Z���{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdCancel_Click()

    On Error GoTo err
    
    ' �t�H�[�������
    HideExt
    
    ' �L�����Z���C�x���g�𑗐M����
    RaiseEvent Cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���N�G���p�����[�^���X�g�̃_�u���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstQueryParameterList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    editQueryParameter
End Sub

' =========================================================
' ���V�K�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdAdd_Click()

    ' ���X�g�{�b�N�X�̃T�C�Y
    Dim cnt As Long
    ' ���X�g�{�b�N�X�̃T�C�Y���擾����
    cnt = queryParameterList.collection.count
    
    ' �|�b�v�A�b�v�̐����ő�o�^���𒴂��Ă��邩�`�F�b�N����
    If cnt >= ConstantsCommon.QUERY_PARAMETER_NEW_CREATED_MAX_SIZE Then
    
        ' ���b�Z�[�W��\������
        Dim mess As String
        mess = replace(QUERY_PARAMETER_NEW_CREATED_OVER_SIZE, "${count}", ConstantsCommon.QUERY_PARAMETER_NEW_CREATED_MAX_SIZE)
        
        VBUtil.showMessageBoxForInformation mess _
                                          , ConstantsCommon.APPLICATION_NAME
        Exit Sub
    End If
    
    ' �|�b�v�A�b�v���j���[�I�u�W�F�N�g�����X�g�ɒǉ�����
    Dim queryParameter As ValQueryParameter
    Set queryParameter = New ValQueryParameter
    
    queryParameter.name = ConstantsCommon.QUERY_PARAMETER_DEFAULT_NAME & "_" & (cnt + 1)
    
    Dim list As New ValCollection
    list.setItem queryParameter
    
    addQueryParameter queryParameter
    
    queryParameterList.setSelectedIndex cnt
    queryParameterList.control.SetFocus
    
End Sub

' =========================================================
' ���ҏW�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdEdit_Click()

    editQueryParameter
End Sub

Private Sub editQueryParameter()

    ' ���ݑI������Ă���C���f�b�N�X���擾
    queryParameterSelectedIndex = queryParameterList.getSelectedIndex

    ' ���I���̏ꍇ
    If queryParameterSelectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    ' ���ݑI������Ă��鍀�ڂ��擾
    Set queryParameterSelectedItem = queryParameterList.getSelectedItem
    
    If VBUtil.unloadFormIfChangeActiveBook(frmQueryParameterSetting) Then Unload frmQueryParameterSetting
    Load frmQueryParameterSetting
    Set frmQueryParameterSettingVar = frmQueryParameterSetting
    frmQueryParameterSetting.ShowExt vbModal, queryParameterSelectedItem
                            
    Set frmQueryParameterSettingVar = Nothing

End Sub

' =========================================================
' ���N�G���p�����[�^�ݒ�t�H�[����OK�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�FqueryParameter �N�G���p�����[�^���
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmQueryParameterSettingVar_ok(ByVal queryParameter As ValQueryParameter)

    Dim v As ValQueryParameter
    Set v = queryParameterList.getItem(queryParameterSelectedIndex)
    
    v.name = queryParameter.name
    v.value = queryParameter.value

    setQueryParameter queryParameterSelectedIndex, v
    
    queryParameterList.control.SetFocus

End Sub

' =========================================================
' ���N�G���p�����[�^�ݒ�t�H�[���̃L�����Z���{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmQueryParameterSettingVar_cancel()

    queryParameterList.control.SetFocus
End Sub

' =========================================================
' ���폜�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDelete_Click()

    Dim selectedIndex As Long
    
    ' ���ݑI������Ă���C���f�b�N�X���擾
    selectedIndex = queryParameterList.getSelectedIndex

    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    queryParameterList.removeItem selectedIndex
    queryParameterList.control.SetFocus

End Sub

' =========================================================
' ����փ{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdUp_Click()

    On Error GoTo err
    
    ' �I���ς݃C���f�b�N�X
    Dim selectedIndex As Long
    
    ' ���݃��X�g�őI������Ă���C���f�b�N�X���擾����
    selectedIndex = queryParameterList.getSelectedIndex
    
    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
        ' �I������
        Exit Sub
    End If

    If selectedIndex > 0 Then
    
        queryParameterList.swapItem _
                          selectedIndex _
                        , selectedIndex - 1 _
                        , vbObject _
                        , 2
                              
        queryParameterList.setSelectedIndex selectedIndex - 1
            
    End If
    
    queryParameterList.control.SetFocus
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �����փ{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdDown_Click()

    On Error GoTo err
    
    ' �I���ς݃C���f�b�N�X
    Dim selectedIndex As Long
    
    ' ���݃��X�g�őI������Ă���C���f�b�N�X���擾����
    selectedIndex = queryParameterList.getSelectedIndex
    
        ' ���I���̏ꍇ
    If selectedIndex = -1 Then
        ' �I������
        Exit Sub
    End If

    If selectedIndex < queryParameterList.count - 1 Then
    
        queryParameterList.swapItem _
                          selectedIndex _
                        , selectedIndex + 1 _
                        , vbObject _
                        , 2
                              
        queryParameterList.setSelectedIndex selectedIndex + 1
            
    End If
    
    queryParameterList.control.SetFocus
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���p�����[�^�R�s�[���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdQueryParameterCopy_Click()

    Dim selectedIndex As Long
    Dim selectedItem As ValQueryParameter
    
    ' ���ݑI������Ă���C���f�b�N�X���擾
    selectedIndex = queryParameterList.getSelectedIndex

    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
    
        ' �I������
        Exit Sub
    End If

    Set selectedItem = queryParameterList.getSelectedItem
    
    WinAPI_Clipboard.SetClipboard selectedItem.tabbedInfoHeader & vbNewLine & getQueryParameterForClipboardFormat(selectedItem)
    
End Sub

' =========================================================
' ���S�p�����[�^�R�s�[���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdAllQueryParameterCopy_Click()

    Dim data As New StringBuilder
    Dim var As Variant
    
    Dim i As Long
    
    For Each var In queryParameterList.collection.col
        If i <= 0 Then
            data.append var.tabbedInfoHeader & vbNewLine
        End If
        data.append getQueryParameterForClipboardFormat(var)
        i = i + 1
    Next
    
    WinAPI_Clipboard.SetClipboard data.str

End Sub

' =========================================================
' ���N�G���p�����[�^�̃N���b�v�{�[�h�t�H�[�}�b�g�`��������擾
'
' �T�v�@�@�@�F�N�G���p�����[�^�̃N���b�v�{�[�h�t�H�[�}�b�g�`����������擾����B
' �����@�@�@�Fvar �N�G���p�����[�^
' �߂�l�@�@�F�N�G���p�����[�^�̃N���b�v�{�[�h�t�H�[�}�b�g�`��������擾
'
' =========================================================
Private Function getQueryParameterForClipboardFormat(ByVal var As ValQueryParameter) As String

    getQueryParameterForClipboardFormat = var.tabbedInfo & vbNewLine

End Function

' =========================================================
' ���N�G���p�����[�^���N���b�v�{�[�h����\�t��
'
' �T�v�@�@�@�F�N�G���p�����[�^���N���b�v�{�[�h����\�t������B
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdQueryParameterPaste_Click()

    Dim var As Variant
    Dim queryParameterRawList As ValCollection
    
    Dim queryParameterObj As ValQueryParameter
    Dim queryParameterObjList As New ValCollection

    Dim clipBoard As String
    clipBoard = WinAPI_Clipboard.GetClipboard
    
    Dim CsvParser As New CsvParser: CsvParser.init vbTab
    Set queryParameterRawList = CsvParser.parse(clipBoard)
    
    For Each var In queryParameterRawList.col
    
        Set queryParameterObj = New ValQueryParameter
    
        If var.count >= 1 Then
            queryParameterObj.name = var.getItemByIndex(1, vbVariant)
        End If
    
        If var.count >= 2 Then
            queryParameterObj.value = var.getItemByIndex(2, vbVariant)
        End If
        
        If queryParameterObj.tabbedInfoHeader <> queryParameterObj.tabbedInfo Then
            queryParameterObjList.setItem queryParameterObj
        End If
    
    Next
    
    addQueryParameterList queryParameterObjList, isAppend:=True

End Sub

' =========================================================
' ���t�H�[�����������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
    ' ���[�h���_�̃A�N�e�B�u�u�b�N��ێ����Ă���
    Set targetBook = ExcelUtil.getActiveWorkbook
    ' ���������������s����
    initial
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���t�H�[���j�����̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Terminate()

    On Error GoTo err
    
    ' �j�����������s����
    unInitial
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���t�H�[���A�N�e�B�u���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_Activate()

End Sub

' =========================================================
' ���t�H�[���̕��鎞�̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' �{�����ł͏������̂��L�����Z������
        Cancel = True
        ' �ȉ��̃C�x���g�o�R�ŕ���
        cmdCancel_Click
    End If
    
End Sub

' =========================================================
' ������������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub initial()

End Sub

' =========================================================
' ����n������
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub unInitial()

End Sub

' =========================================================
' ���ݒ���̐���
' =========================================================
Private Function createApplicationProperties() As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initWorksheet targetBook, ConstantsApplicationProperties.BOOK_PROPERTIES_SHEET_NAME, ConstantsApplicationProperties.INI_FILE_DIR_FORM & "\" & Me.name & ".ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ���N�G���p�����[�^����ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeQueryParameter()

    On Error GoTo err
    
    Dim queryParameterList_ As New ValQueryParameterList
    queryParameterList_.init targetBook
    queryParameterList_.list = queryParameterList.collection
    queryParameterList_.writeForData
    
    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub

' =========================================================
' ���N�G���p�����[�^����ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub restoreQueryParameter()

    On Error GoTo err
    
    Dim queryParameterList_ As New ValQueryParameterList
    queryParameterList_.init targetBook
    queryParameterList_.readForData

    Set queryParameterList = New CntListBox: queryParameterList.init lstQueryParameterList
    
    addQueryParameterList queryParameterList_.list
    
    ' �擪��I������
    queryParameterList.setSelectedIndex 0
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ���N�G���p�����[�^���X�g��ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvalQueryParameterList �N�G���p�����[�^���X�g
'     �@�@�@  isAppend              �ǉ��L���t���O
' �߂�l�@�@�F
'
' =========================================================
Private Sub addQueryParameterList(ByVal ValQueryParameterList As ValCollection, Optional ByVal isAppend As Boolean = False)
    
    queryParameterList.addAll ValQueryParameterList _
                       , "name" _
                       , "value" _
                       , isAppend:=isAppend
    
End Sub

' =========================================================
' ���N�G���p�����[�^��ǉ�
'
' �T�v�@�@�@�F
' �����@�@�@�FqueryParameter �N�G���p�����[�^
' �߂�l�@�@�F
'
' =========================================================
Private Sub addQueryParameter(ByVal queryParameter As ValQueryParameter)
    
    queryParameterList.addItemByProp queryParameter, "name", "value"
    
End Sub

' =========================================================
' ���N�G���p�����[�^��ύX
'
' �T�v�@�@�@�F
' �����@�@�@�Findex �C���f�b�N�X
'     �@�@�@  rec   �N�G���p�����[�^
' �߂�l�@�@�F
'
' =========================================================
Private Sub setQueryParameter(ByVal index As Long, ByVal rec As ValQueryParameter)
    
    queryParameterList.setItem index, rec, "name", "value"
    
End Sub
