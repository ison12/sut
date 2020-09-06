VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenuSetting 
   Caption         =   "���j���[�ݒ�"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7440
   OleObjectBlob   =   "frmMenuSetting.frx":0000
End
Attribute VB_Name = "frmMenuSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' *********************************************************
' �E�N���b�N���j���[�̐ݒ�
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/06/02�@�V�K�쐬
'
' ���L�����F
' �@�t�H�[���̏�Ƀt�H�[�����d�˂ĕ\�������
'   �Ȍ�Excel�{�̂�IME���[�h�������ɂȂ葀��s�\�ɂȂ��Ă��܂��Ƃ������ۂɑ���
'   �����h�����߂ɤ��U�e�t�H�[�����B���Ĥ�{�t�H�[�������Ƃ��ɍĕ\�����邱�ƂŤ���̌��ۂ�h��
' �@���̂��߂ɁAShowExt���\�b�h�ɐe�t�H�[����n���悤�����ɒǉ����Ă���
'
' *********************************************************

' ���C�x���g
' =========================================================
' �����肵���ۂɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FappSettingShortcut �A�v���P�[�V�����ݒ�V���[�g�J�b�g
' �@�@�@�@�@�@selectedItemList �I���ςݍ��ڃ��X�g
' �@�@�@�@�@�@menuName �V�������j���[��
'
' =========================================================
Public Event ok(ByRef appSettingShortcut As ValApplicationSettingShortcut _
              , ByRef selectedItemList As ValCollection _
              , ByVal menuName As String)

' =========================================================
' ���L�����Z�����ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event Cancel()

' =========================================================
' �����Z�b�g���ꂽ�ꍇ�ɌĂяo�����C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FappSettingShortcut �A�v���P�[�V�����ݒ�V���[�g�J�b�g
' �@�@�@�@�@�@cancel �L�����Z���t���O
'
' =========================================================
Public Event reset(ByRef appSettingShortcut As ValApplicationSettingShortcut _
                 , ByRef Cancel As Boolean)

' �A�C�R���摜
Private iconImage As IPictureDisp

' �A�v���P�[�V�����ݒ���i�V���[�g�J�b�g�L�[�j
Private applicationSetting As ValApplicationSettingShortcut

' �I���ςݍ��ڃ��X�g
' �E���̃��X�g�{�b�N�X�ɐݒ肷�鍀�ڂ��i�[���Ă��郊�X�g
' �ȉ��̃L�[�l�����ɁA�@�\���X�g����I������Ă��鍀�ڂ𒊏o����
' [ Key ] : CommandBarControl.Tag
' [ Val ] : CommandBarControl.Tag
Private selectedItemList As ValCollection

' ���j���[���X�g �R���g���[��
Private menuList As CntListBox
' �@�\���X�g �R���g���[��
Private appMenuList As CntListBox

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
' �����@�@�@�Ficon             �A�C�R��
' �@�@�@�@�@�@modal            ���[�_���܂��̓��[�h���X�\���w��
' �@�@�@�@�@�@var              �A�v���P�[�V�����ݒ���
'             var2             �I���ςݍ��ڃ��X�g
' �@�@�@�@�@�@title            �t�H�[���^�C�g��
' �@�@�@�@�@�@message          �t�H�[���̃��b�Z�[�W
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByRef icon As Object _
                 , ByVal modal As FormShowConstants _
                 , ByRef var As ValApplicationSettingShortcut _
                 , ByRef var2 As ValCollection _
                 , ByVal title As String _
                 , ByVal message As String _
                 , ByVal menuName As String _
                 , Optional ByVal menuNameDisable As Boolean = False)

    If Not icon Is Nothing Then
        ' ���g�̃A�C�R����ޔ�������
        'Set iconImage = Me.imgIcon.Picture
        ' �A�C�R����e�t�H�[���̉摜�Œu��������
        'Me.imgIcon.Picture = icon
    End If
    
    ' �����o�ϐ��ɃA�v���P�[�V�����ݒ����ݒ肷��
    Set applicationSetting = var
    ' �I���ςݍ��ڃ��X�g��ݒ肷��
    Set selectedItemList = var2
    ' �^�C�g����ݒ肷��
    Me.Caption = title
    ' ���b�Z�[�W��ݒ肷��
    lblMessage.Caption = message
    ' ���j���[����ݒ肷��
    txtMenuName.value = menuName
    If menuNameDisable = True Then
    
        txtMenuName.Enabled = False
        lstAppList.SetFocus
    Else
    
        txtMenuName.Enabled = True
        txtMenuName.SetFocus
    End If
    
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
    
    If Not iconImage Is Nothing Then
        ' �A�C�R���摜��ݒ肷��
        Me.imgIcon.Picture = iconImage
    End If

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

    initListControl
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
' ��OK�{�^���N���b�N���̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdOk_Click()

    On Error GoTo err
    
    ' �o�^������
    Dim storedList As New ValCollection

    ' �R���g���[���I�u�W�F�N�g
    Dim control As commandBarControl
    
    ' ���X�g�ɑ��݂��鍀�ڂ��E�N���b�N���j���[�Ƃ��Ēǉ�����
    For Each control In menuList.collection.col
    
        storedList.setItem control.Tag, control.Tag
    Next
    
    ' �t�H�[�������
    HideExt
    
    ' OK�C�x���g�𑗐M����
    RaiseEvent ok(applicationSetting _
                , storedList _
                , txtMenuName.value)
    
    
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
' ���ǉ��{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdMenuAdd_Click()

    On Error GoTo err
    
    ' �C���f�b�N�X
    Dim i As Long
    ' ��
    Dim cnt As Long
    
    ' ���X�g�{�b�N�X�̗v�f
    Dim appMenuItem As commandBarControl
    
    ' �폜����v�f���X�g
    Dim removeItem As New ValCollection
    
    ' ���X�g�{�b�N�X�̌����擾����
    cnt = lstAppList.ListCount
    
    ' �@�\���X�g�ɂă`�F�b�N����Ă���v�f��
    ' �E�N���b�N���j���[���X�g�Ɉڂ��ς���
    For i = 0 To cnt - 1
    
        ' �I������Ă��邩�`�F�b�N
        If lstAppList.selected(i) = True Then
        
            ' �폜�v�f���X�g�ɍ폜���ׂ��C���f�b�N�X��ǉ�
            removeItem.setItem i
            
            ' �@�\���擾
            Set appMenuItem = appMenuList.getItem(i)
            ' �E�N���b�N���j���[���X�g�ɋ@�\��ǉ�����
            menuList.addItem appMenuItem.DescriptionText, appMenuItem
            
        End If
        
    Next
    
    ' �폜�����̎��s
    ' �Ō������ŏ��Ɍ������ă��X�g�����[�v������̂�
    ' �v�f�̍폜�ɂ���ăC���f�b�N�X�ɂ��ꂪ��������̂�h������
    cnt = removeItem.count
    For i = cnt - 1 To 0 Step -1
    
        appMenuList.removeItem removeItem.getItemByIndex(i + 1, vbLong)
    
    Next
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���폜�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdMenuRemove_Click()

    On Error GoTo err
    
    ' �폜�����v�f
    Dim removedItem As commandBarControl
    
    ' �I������Ă��鍀�ڂ̃C���f�b�N�X
    Dim selectedIndex As Long
    
    ' ���݃��X�g�őI������Ă���C���f�b�N�X���擾����
    selectedIndex = lstMenu.ListIndex
    
    ' ���I���̏ꍇ
    If selectedIndex = -1 Then
    
        Exit Sub
    End If
    
    ' �E�N���b�N���j���[���X�g���獀�ڂ��擾���폜����
    Set removedItem = menuList.getItem(selectedIndex)
    menuList.removeItem selectedIndex
    
    ' �@�\���X�g�ɍ��ڂ�ǉ�����
    appMenuList.addItem removedItem.DescriptionText, removedItem
    
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
Private Sub cmdMenuDown_Click()

    On Error GoTo err
    
    ' �I���ς݃C���f�b�N�X
    Dim selectedIndex As Long
    
    ' ���݃��X�g�őI������Ă���C���f�b�N�X���擾����
    selectedIndex = lstMenu.ListIndex
    
    If selectedIndex < lstMenu.ListCount - 1 Then
    
        menuList.swapItem selectedIndex _
                        , selectedIndex + 1
                              
        lstMenu.selected(selectedIndex + 1) = True
            
    End If
    
    lstMenu.SetFocus
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
        
End Sub

' =========================================================
' ����փ{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdMenuUp_Click()

    On Error GoTo err
    
    ' �I���ς݃C���f�b�N�X
    Dim selectedIndex As Long
    
    ' ���݃��X�g�őI������Ă���C���f�b�N�X���擾����
    selectedIndex = lstMenu.ListIndex
    
    If selectedIndex > 0 Then
    
        menuList.swapItem selectedIndex _
                        , selectedIndex - 1
                              
        lstMenu.selected(selectedIndex - 1) = True
            
    End If
    
    lstMenu.SetFocus
        
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' �����Z�b�g�{�^���������̃C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdReset_Click()

    On Error GoTo err
    
    Dim isCancel As Boolean: isCancel = False
    
    ' ���Z�b�g�C�x���g�𔭍s����
    RaiseEvent reset(applicationSetting, isCancel)
    
    ' �L�����Z�����ꂽ�ꍇ
    If isCancel = True Then
    
        Exit Sub
    End If
    
    ' �I���ςݍ��ڂ����������A�T�C�Y��0�ɂ���
    Set selectedItemList = New ValCollection
    
    ' ���Z�b�g�C�x���g����M�������ŁA���Z�b�g���s���Ă��邽��
    ' ���X�g�R���g���[��������������
    initListControl
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
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

    Set iconImage = Nothing
End Sub

' =========================================================
' ���V���[�g�J�b�g����ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub initListControl()

    ' �E�N���b�N���j���[���X�g�Ƌ@�\���X�g�����Z�b�g����
    lstMenu.clear
    lstAppList.clear
    
    ' �E�N���b�N���j���[���X�g�̏�����
    Set menuList = New CntListBox: menuList.init lstMenu
    ' �@�\���X�g�̏�����
    Set appMenuList = New CntListBox: appMenuList.init lstAppList
    
    ' Sut���j���[�̗v�f
    Dim sutMenuItem As commandBarControl
    
    Dim shortcutInfo As ValShortcutKey
    
    ' �@�\�̃V���[�g�J�b�g���X�g���擾����
    Dim shortCutList As ValCollection
    Set shortCutList = applicationSetting.shortcutAppList
    
    ' ---------------------------------------------------------
    ' �@�\���X�g�̏�����
    ' ---------------------------------------------------------
    For Each shortcutInfo In shortCutList.col
    
        ' ���j���[�̗v�f���擾����
        Set sutMenuItem = shortcutInfo.commandBarControl
        
        ' �ۑ�����Ă��Ȃ��ꍇ�́A�@�\���X�g�ɒǉ�
        If selectedItemList.exist(sutMenuItem.Tag) = False Then
        
            ' �����̃��j���[�ɍ��ڂ�ǉ�����
            appMenuList.addItem sutMenuItem.DescriptionText, sutMenuItem
        
        Else
        
        End If
    
    Next

    ' ---------------------------------------------------------
    ' ���j���[���X�g�̏�����
    ' �����j���[���X�g�̏��������ێ����邽�߂�
    ' �@menuList����v�f�����ԂɎ��o���ă��X�g�Ɋi�[����
    ' ---------------------------------------------------------
    Dim menuId As Variant
    
    For Each menuId In selectedItemList.col
    
        ' �V���[�g�J�b�g�����擾����
        Set shortcutInfo = shortCutList.getItem(menuId)
        
        If Not shortcutInfo Is Nothing Then
        
            ' ���j���[�̗v�f���擾����
            Set sutMenuItem = shortcutInfo.commandBarControl
        
            ' �E���̃��j���[�ɍ��ڂ�ǉ�����
            menuList.addItem sutMenuItem.DescriptionText, sutMenuItem
        
        End If
        
    Next

End Sub
