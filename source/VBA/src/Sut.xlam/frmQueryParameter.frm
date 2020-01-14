VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryParameter 
   Caption         =   "�N�G���p�����[�^�̐ݒ�"
   ClientHeight    =   8355.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8250.001
   OleObjectBlob   =   "frmQueryParameter.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
' �쐬�ҁ@�FHideki Isobe
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

' �N�G���p�����[�^�ݒ���
Private WithEvents frmQueryParameterSettingVar As frmQueryParameterSetting
Attribute frmQueryParameterSettingVar.VB_VarHelpID = -1

' �N�G���p�����[�^���X�g �R���g���[��
Private queryParameterList As CntListBox

' �N�G���p�����[�^���X�g�ł̑I�����ڃC���f�b�N�X
Private queryParameterSelectedIndex As Long
' �N�G���p�����[�^���X�g�ł̑I�����ڃI�u�W�F�N�g
Private queryParameterSelectedItem As ValQueryParameter

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
    Me.Show vbModal
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
    
    Main.storeFormPosition Me.name, Me

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
    Set queryParameter = New ValQueryParameter: queryParameter.name = ConstantsCommon.QUERY_PARAMETER_DEFAULT_NAME
    
    queryParameter.name = QUERY_PARAMETER_DEFAULT_NAME & " " & (cnt + 1)
    
    Dim list As New ValCollection
    list.setItem queryParameter
    
    queryParameterList.addItemByProp queryParameter, "name", "value"
    
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
    
    Load frmQueryParameterSetting
    Set frmQueryParameterSettingVar = frmQueryParameterSetting
    frmQueryParameterSetting.ShowExt vbModal, queryParameterSelectedItem
                            
    Set frmQueryParameterSettingVar = Nothing

End Sub

' =========================================================
' ���N�G���p�����[�^�ݒ�t�H�[����OK�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmQueryParameterSettingVar_ok(ByVal ValQueryParameter As ValQueryParameter)

    Dim v As ValQueryParameter
    Set v = queryParameterList.getItem(queryParameterSelectedIndex)
    
    v.name = ValQueryParameter.name
    v.value = ValQueryParameter.value

    queryParameterList.setItem queryParameterSelectedIndex, v, "name", "value"
    
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
Private Sub frmQueryParameterSettingVar_Cancel()

    queryParameterList.control.SetFocus
End Sub

' =========================================================
' �����j���[�ݒ�t�H�[���̃��Z�b�g�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub frmMenuSettingVar_reset(appSettingShortcut As ValApplicationSettingShortcut _
                                  , ByRef Cancel As Boolean)

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
    
    WinAPI_Clipboard.SetClipboard getQueryParameterForClipboardFormat(selectedItem)
    
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
    
    For Each var In queryParameterList.collection.col
        data.append getQueryParameterForClipboardFormat(var)
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

    getQueryParameterForClipboardFormat = """" & replace(var.name, """", """""") & """" & vbTab & """" & replace(var.value, """", """""") & """" & vbNewLine

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
        
        queryParameterObjList.setItem queryParameterObj
    
    Next
    
    queryParameterList.addAll queryParameterObjList, "name", "value", True
    

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
' ���N�G���p�����[�^����ۑ�����
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub storeQueryParameter()

    ' ----------------------------------------------
    ' �u�b�N�ݒ���
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet
    
    Dim var As Variant
    Dim i As Long
    
    bookProp.removeAllValue ConstantsBookProperties.TABLE_QUERY_PARAMETER_DIALOG
    
    i = 0
    For Each var In queryParameterList.collection.col
    
        bookProp.setValue ConstantsBookProperties.TABLE_QUERY_PARAMETER_DIALOG, "name_" & i, var.name
        bookProp.setValue ConstantsBookProperties.TABLE_QUERY_PARAMETER_DIALOG, "value_" & i, var.value
    
        i = i + 1
    Next
    ' ----------------------------------------------


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

    ' ----------------------------------------------
    ' �u�b�N�ݒ���
    Dim bookProp As New BookProperties
    bookProp.sheet = ActiveSheet

    Dim bookPropVal As ValCollection

    If bookProp.isExistsProperties Then
        ' �ݒ���V�[�g�����݂���
        Set bookPropVal = bookProp.getValuesOfElementArray(ConstantsBookProperties.TABLE_QUERY_PARAMETER_DIALOG)
    Else
        Set bookPropVal = New ValCollection
    End If
    ' ----------------------------------------------

    Dim ValQueryParameterList As New ValQueryParameterList
    ValQueryParameterList.setListFromFlatRecords bookPropVal

    Set queryParameterList = New CntListBox: queryParameterList.init lstQueryParameterList
    
    queryParameterList.addAll ValQueryParameterList.list _
                       , "name" _
                       , "value"
    
End Sub



