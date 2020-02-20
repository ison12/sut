VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValPopupmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �|�b�v�A�b�v���j���[��\������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/06/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �A�v���P�[�V�����R�}���h�o�[ID
Private m_appBarId As String

' �|�b�v�A�b�v���j���[�̖��O
Private m_popupMenuName    As String
' �V���[�g�J�b�g�L�[���x��
Private m_shortcutKeyLabel As String
' �V���[�g�J�b�g�L�[�R�[�h Application#OnKey�ɐݒ肷��R�[�h
Private m_shortcutKeyCode  As String

' �|�b�v�A�b�v���j���[�̍��ڃ��X�g
Private m_itemList As ValCollection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set m_itemList = New ValCollection
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set m_itemList = Nothing
End Sub

Public Sub init(ByVal appBarId As String)

    m_appBarId = appBarId
End Sub

' =========================================================
' ���v���p�e�B
' =========================================================
Public Property Get popupMenuName() As String

    popupMenuName = m_popupMenuName
End Property

Public Property Let popupMenuName(ByVal vNewValue As String)

    m_popupMenuName = vNewValue
End Property

Public Property Get shortcutKeyCode() As String

    shortcutKeyCode = m_shortcutKeyCode
End Property

Public Property Let shortcutKeyCode(ByVal vNewValue As String)

    m_shortcutKeyCode = vNewValue
End Property

Public Property Get shortcutKeyLabel() As String

    shortcutKeyLabel = m_shortcutKeyLabel
End Property

Public Property Let shortcutKeyLabel(ByVal vNewValue As String)

    m_shortcutKeyLabel = vNewValue
End Property

Public Property Get commandBarPopup() As CommandBar

    ' Application.CommandBars.item���\�b�h�͍��ڂ�������Ȃ��ꍇ�A�G���[�𓊂��邽��
    ' �����h�����߃G���[�𖳎�����ȉ��̐錾���s��
    On Error Resume Next

    Set commandBarPopup = Application.CommandBars.item(m_appBarId & "_" & m_popupMenuName)

    If err.Number <> 0 Then
    
        Set commandBarPopup = Nothing
    End If
    
    On Error GoTo 0
    
End Property

Public Property Get itemList() As ValCollection

    Set itemList = m_itemList
End Property

Public Property Let itemList(ByVal vNewValue As ValCollection)

    Set m_itemList = vNewValue
End Property


' =========================================================
' ���|�b�v�A�b�v�̐���
'
' �T�v�@�@�@�F���ڃ��X�g����|�b�v�A�b�v���j���[�𐶐�����B
' �����@�@�@�F
' �߂�l�@�@�FTrue �����ɐ���
'
' =========================================================
Public Function createPopup() As Boolean

    On Error GoTo err

    ' �߂�l��������
    createPopup = False
    
    ' �A�v���P�[�V�����̃R�}���h�o�[
    Dim appBar As CommandBar
    
    Set appBar = Application.CommandBars(m_appBarId)

    ' �|�b�v�A�b�v
    Dim popup As CommandBar
    
    ' -----------------------------------------------------------------
    ' Application.CommandBars.item���\�b�h�͍��ڂ�������Ȃ��ꍇ�A�G���[�𓊂��邽��
    ' �����h�����߃G���[�𖳎�����ȉ��̐錾���s��
    On Error Resume Next
    
    ' �|�b�v�A�b�v�R���g���[�����擾����
    Set popup = Application.CommandBars.item(m_appBarId & "_" & m_popupMenuName)
    
    On Error GoTo err
    ' -----------------------------------------------------------------
    
    ' �|�b�v�A�b�v�R���g���[�������ɑ��݂���ꍇ�́A��U�폜����
    If Not popup Is Nothing Then
    
        popup.delete
        
    End If
    
    ' �V�K�Œǉ�����
    Set popup = Application.CommandBars.Add( _
                            name:=m_appBarId & "_" & m_popupMenuName _
                          , Temporary:=True _
                          , position:=msoBarPopup)
    
    ' ���ɂȂ�R���g���[��
    Dim srcControl As commandBarControl
    ' �|�b�v�A�b�v�ɒǉ�����R���g���[���i���ɂȂ�R���g���[���̓��e���R�s�[����j
    Dim desControl As commandBarControl
    
    ' �|�b�v�A�b�v����
    Dim popupMenuId As Variant
    
    ' ���ڃ��X�g��1���鏈������
    For Each popupMenuId In m_itemList.col
    
        ' �@�\ID�������ɃA�v���P�[�V�����̃R�}���h�o�[����R���g���[������������
        Set srcControl = appBar.FindControl(Tag:=popupMenuId, recursive:=True)
        
        ' ���݂��Ă���ꍇ
        If Not srcControl Is Nothing Then
        
            ' �|�b�v�A�b�v�ɃR���g���[����ǉ�
            Set desControl = popup.Controls.Add(Type:=msoControlButton)
            ' �|�b�v�A�b�v�ɒǉ������R���g���[���i�^�����ȏ�ԁj���A�v���P�[�V�����o�[����擾�����R���g���[���Ɠ������e�ɕύX����
            ExcelUtil.copyCommandBarControl srcControl, desControl
        
        End If
    
    Next
    
    
    ' �쐬�ɐ���
    createPopup = True
    
    Exit Function
    
err:

    If Not popup Is Nothing Then
    
        popup.delete
    End If

    ' �쐬�ɐ���
    createPopup = False
    
End Function

' =========================================================
' ���|�b�v�A�b�v���j���[�̃N���[���𐶐�����B
'
' �T�v�@�@�@�F�I�u�W�F�N�g�̃N���[���𐶐�����B
' �����@�@�@�F
' �߂�l�@�@�FValPopupMenu
'
' =========================================================
Public Function clone() As ValPopupmenu

    Dim ret As ValPopupmenu
    
    Set ret = New ValPopupmenu:
    
    ret.init m_appBarId
    ret.popupMenuName = m_popupMenuName
    ret.shortcutKeyCode = m_shortcutKeyCode
    ret.shortcutKeyLabel = m_shortcutKeyLabel
    
    Dim newItemList As New ValCollection
    
    Dim var As Variant
    For Each var In m_itemList.col
    
        newItemList.setItem var, var
    Next
    
    ret.itemList = newItemList
    
    Set clone = ret

End Function