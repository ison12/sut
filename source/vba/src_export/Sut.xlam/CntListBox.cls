VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CntListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ListBox�R���g���[���̃��b�v�N���X
'
' ���X�g�{�b�N�X�ɂ̓��x�������ǉ��ł��Ȃ��Ƃ������_������B
' �{�N���X�ł́A�����₤���߂ɃR���N�V�����I�u�W�F�N�g��
' �����ŕێ������x���ƃI�u�W�F�N�g���P�Z�b�g�Œǉ��E�擾��
' �\�ȃC���^�[�t�F�[�X���`���Ă���B
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/09/23�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���X�g�{�b�N�X
Private m_control     As Object
' �R���N�V�����I�u�W�F�N�g
Private m_collection  As ValCollection

' =========================================================
' ���v���p�e�B ���X�g�R���g���[���擾
' =========================================================
Public Property Get control() As Variant

    Set control = m_control
    
End Property

' =========================================================
' ���v���p�e�B ���X�g�R���g���[���ݒ�
' =========================================================
Public Property Let control(ByVal vNewValue As Variant)

    Set m_control = vNewValue
    
End Property

' =========================================================
' ���v���p�e�B �R���N�V�����擾
' =========================================================
Public Property Get collection() As ValCollection

    Set collection = m_collection
End Property

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set m_collection = Nothing
End Sub

' =========================================================
' ���[���R���X�g���N�^
' =========================================================
Public Sub init(ByRef var As Variant)

    ' ���X�g�R���g���[����ݒ�
    Set m_control = var
    ' �R���N�V������V���ɏ���������
    Set m_collection = New ValCollection
    
End Sub

' =========================================================
' ���v�f�ǉ�
'
' �T�v�@�@�@�F���X�g�R���g���[���ƃR���N�V�����ɒl��ǉ�����B
'
' �����@�@�@�Flabel ���X�g�R���g���[���ɕ\�����郉�x��
' �@�@�@�@�@�@var   �R���N�V�����ɒǉ�����v�f
' �߂�l�@�@�F����
' =========================================================
Public Sub addItem(ByVal label As String _
                 , ByVal var As Variant)

    ' ���X�g�R���g���[���Ƀ��x����ǉ�
    m_control.addItem label
    ' �R���N�V�����ɗv�f��ǉ�
    m_collection.setItem var

End Sub

' =========================================================
' ���v�f�ǉ�
'
' �T�v�@�@�@�F���X�g�R���g���[���ƃR���N�V�����ɒl��ǉ�����B
'
' �����@�@�@�Fvar   �R���N�V�����ɒǉ�����v�f
'             labelProp1  ���x���v���p�e�B1
'             labelProp2  ���x���v���p�e�B2
'             labelProp3  ���x���v���p�e�B3
'             labelProp4  ���x���v���p�e�B4
'             labelProp5  ���x���v���p�e�B5
'             labelProp6  ���x���v���p�e�B6
'             labelProp7  ���x���v���p�e�B7
'             labelProp8  ���x���v���p�e�B8
'             labelProp9  ���x���v���p�e�B9
'             labelProp10 ���x���v���p�e�B10
' �߂�l�@�@�F����
' =========================================================
Public Sub addItemByProp(ByVal var As Variant _
                 , Optional ByVal labelProp1 As String = "" _
                 , Optional ByVal labelProp2 As String = "" _
                 , Optional ByVal labelProp3 As String = "" _
                 , Optional ByVal labelProp4 As String = "" _
                 , Optional ByVal labelProp5 As String = "" _
                 , Optional ByVal labelProp6 As String = "" _
                 , Optional ByVal labelProp7 As String = "" _
                 , Optional ByVal labelProp8 As String = "" _
                 , Optional ByVal labelProp9 As String = "" _
                 , Optional ByVal labelProp10 As String = "")

    Dim Label1  As String
    Dim Label2  As String
    Dim Label3  As String
    Dim Label4  As String
    Dim Label5  As String
    Dim Label6  As String
    Dim Label7  As String
    Dim Label8  As String
    Dim Label9  As String
    Dim Label10 As String

    ' ���X�g�R���g���[���Ƀ��x����ǉ�
    m_control.addItem

    If labelProp1 <> "" Then
        Label1 = CallByName(var, labelProp1, VbGet)
        m_control.list(m_control.ListCount - 1, 0) = Label1
    End If

    If labelProp2 <> "" Then
        Label2 = CallByName(var, labelProp2, VbGet)
        m_control.list(m_control.ListCount - 1, 1) = Label2
    End If

    If labelProp3 <> "" Then
        Label3 = CallByName(var, labelProp3, VbGet)
        m_control.list(m_control.ListCount - 1, 2) = Label3
    End If

    If labelProp4 <> "" Then
        Label4 = CallByName(var, labelProp4, VbGet)
        m_control.list(m_control.ListCount - 1, 3) = Label4
    End If

    If labelProp5 <> "" Then
        Label5 = CallByName(var, labelProp5, VbGet)
        m_control.list(m_control.ListCount - 1, 4) = Label5
    End If

    If labelProp6 <> "" Then
        Label6 = CallByName(var, labelProp6, VbGet)
        m_control.list(m_control.ListCount - 1, 5) = Label6
    End If

    If labelProp7 <> "" Then
        Label7 = CallByName(var, labelProp7, VbGet)
        m_control.list(m_control.ListCount - 1, 6) = Label7
    End If

    If labelProp8 <> "" Then
        Label8 = CallByName(var, labelProp8, VbGet)
        m_control.list(m_control.ListCount - 1, 7) = Label8
    End If

    If labelProp9 <> "" Then
        Label9 = CallByName(var, labelProp9, VbGet)
        m_control.list(m_control.ListCount - 1, 8) = Label9
    End If

    If labelProp10 <> "" Then
        Label10 = CallByName(var, labelProp10, VbGet)
        m_control.list(m_control.ListCount - 1, 9) = Label10
    End If

    ' �R���N�V�����ɗv�f��ǉ�
    m_collection.setItem var

End Sub

' =========================================================
' ���v�f�ݒ�
'
' �T�v�@�@�@�F�R���N�V�����ɗv�f��ݒ肷��B
'
' �����@�@�@�Findex �Y�����i0���ŏ��l�Ƃ���j
' �@�@�@�@�@�@val   �l
'             labelProp1 ���x���v���p�e�B1
'             labelProp2 ���x���v���p�e�B2
' �߂�l�@�@�F����
' =========================================================
Public Function setItem(ByVal index As Long _
                , ByVal val As Variant _
                 , Optional ByVal labelProp1 As String = "" _
                 , Optional ByVal labelProp2 As String = "" _
                 , Optional ByVal labelProp3 As String = "" _
                 , Optional ByVal labelProp4 As String = "" _
                 , Optional ByVal labelProp5 As String = "" _
                 , Optional ByVal labelProp6 As String = "" _
                 , Optional ByVal labelProp7 As String = "" _
                 , Optional ByVal labelProp8 As String = "" _
                 , Optional ByVal labelProp9 As String = "" _
                 , Optional ByVal labelProp10 As String = "") As Variant

    Dim item As Variant
    
    Dim Label1 As String
    Dim Label2 As String
    Dim Label3  As String
    Dim Label4  As String
    Dim Label5  As String
    Dim Label6  As String
    Dim Label7  As String
    Dim Label8  As String
    Dim Label9  As String
    Dim Label10 As String

    ' �C���f�b�N�X�I�[�o�[�t���[�`�F�b�N
    If m_collection.count <= index Then
        Exit Function
    End If

    m_collection.removeByIndex index + 1
    If m_collection.count >= index + 1 Then
        m_collection.setItemByIndexBefore val, index + 1
    Else
        m_collection.setItem val
    End If
    
    If labelProp1 <> "" Then
        Label1 = CallByName(val, labelProp1, VbGet)
        m_control.list(index, 0) = Label1
    End If

    If labelProp2 <> "" Then
        Label2 = CallByName(val, labelProp2, VbGet)
        m_control.list(index, 1) = Label2
    End If

    If labelProp3 <> "" Then
        Label3 = CallByName(val, labelProp3, VbGet)
        m_control.list(index, 2) = Label3
    End If

    If labelProp4 <> "" Then
        Label4 = CallByName(val, labelProp4, VbGet)
        m_control.list(index, 3) = Label4
    End If

    If labelProp5 <> "" Then
        Label5 = CallByName(val, labelProp5, VbGet)
        m_control.list(index, 4) = Label5
    End If

    If labelProp6 <> "" Then
        Label6 = CallByName(val, labelProp6, VbGet)
        m_control.list(index, 5) = Label6
    End If

    If labelProp7 <> "" Then
        Label7 = CallByName(val, labelProp7, VbGet)
        m_control.list(index, 6) = Label7
    End If

    If labelProp8 <> "" Then
        Label8 = CallByName(val, labelProp8, VbGet)
        m_control.list(index, 7) = Label8
    End If

    If labelProp9 <> "" Then
        Label9 = CallByName(val, labelProp9, VbGet)
        m_control.list(index, 8) = Label9
    End If

    If labelProp10 <> "" Then
        Label10 = CallByName(val, labelProp10, VbGet)
        m_control.list(index, 9) = Label10
    End If

End Function

' =========================================================
' ���v�f�擾
'
' �T�v�@�@�@�F�R���N�V��������v�f���擾����B
'
' �����@�@�@�Findex �Y�����i0���ŏ��l�Ƃ���j
' �@�@�@�@�@�@t     �R���N�V��������擾����v�f�̌^
' �@�@�@�@�@�@�@�@�@�ȗ������ꍇ��Object�^�B
' �߂�l�@�@�F�v�f
' =========================================================
Public Function getItem(ByVal index As Long, Optional ByVal t As VbVarType = vbObject) As Variant

    ' �^���I�u�W�F�N�g�̏ꍇ
    If t = vbObject Then
    
        ' Set�X�e�[�g�����g���g�p���Ė߂�l���擾����
        ' �iCollection�I�u�W�F�N�g�̓Y������1���ŏ��l�ƂȂ�̂�index��+1���Ă���j
        Set getItem = m_collection.getItemByIndex(index + 1, t)
        
    ' �^���I�u�W�F�N�g�ȊO
    Else
        
        ' Get�X�e�[�g�����g���g�p���Ė߂�l���擾����
        ' �iCollection�I�u�W�F�N�g�̓Y������1���ŏ��l�ƂȂ�̂�index��+1���Ă���j
        getItem = m_collection.getItemByIndex(index + 1, t)
    End If

End Function

Public Sub swapItem(ByVal index1 As Long, _
                    ByVal index2 As Long, _
                    Optional ByVal t As VbVarType = vbObject, _
                    Optional ByVal colCount As Long = 1)

    Dim var1 As Variant
    Dim var2 As Variant
    
    Dim indexMin As Long
    Dim indexMax As Long
    
    If index1 > index2 Then
    
        indexMin = index2
        indexMax = index1
    Else
        
        indexMin = index1
        indexMax = index2
    End If
    
    ' ----------------------------------------
    ' Collection�I�u�W�F�N�g�ɗv�f��ݒ肷��
    
    ' �^��Object�^�ł��邩�A�����łȂ������r���K�؂ȕ��@��
    ' �l���擾����
    If t = vbObject Then
    
        Set var1 = m_collection.getItemByIndex(indexMax + 1)
        Set var2 = m_collection.getItemByIndex(indexMin + 1)
        
    Else
        
        var1 = m_collection.getItemByIndex(indexMax + 1)
        var2 = m_collection.getItemByIndex(indexMin + 1)
        
    End If

    m_collection.removeByIndex indexMax + 1
    m_collection.setItemByIndexAfter var2, indexMax
    
    m_collection.removeByIndex indexMin + 1
    m_collection.setItemByIndexBefore var1, indexMin + 1

    ' ----------------------------------------
    ' �R���g���[���̃��x�������ւ���
    If colCount = 1 Then
        var1 = m_control.list(indexMax)
        var2 = m_control.list(indexMin)
        m_control.list(indexMax) = var2
        m_control.list(indexMin) = var1
        
    ElseIf colCount = 2 Then
    
        var1 = m_control.list(indexMax, 0)
        var2 = m_control.list(indexMin, 0)
        
        If isNull(var2) Then
            m_control.list(indexMax, 0) = ""
        Else
            m_control.list(indexMax, 0) = var2
        End If
        If isNull(var1) Then
            m_control.list(indexMin, 0) = ""
        Else
            m_control.list(indexMin, 0) = var1
        End If
    
        var1 = m_control.list(indexMax, 1)
        var2 = m_control.list(indexMin, 1)
        If isNull(var2) Then
            m_control.list(indexMax, 1) = ""
        Else
            m_control.list(indexMax, 1) = var2
        End If
        If isNull(var1) Then
            m_control.list(indexMin, 1) = ""
        Else
            m_control.list(indexMin, 1) = var1
        End If
    
    End If


End Sub

' =========================================================
' ���v�f�ǉ�
'
' �T�v�@�@�@�F�v�f���܂Ƃ߂Ēǉ�����B
'
' �����@�@�@�Fcol       �R���N�V����
' �@�@�@�@�@�@labelProp1 ���x���Ƃ��ė��p����ycol�z�Ɋi�[����Ă���v�f�̃v���p�e�B
' �@�@�@�@�@�@labelProp2 �V
' �@�@�@�@�@�@labelProp3 �V
' �@�@�@�@�@�@labelProp4 �V
' �@�@�@�@�@�@labelProp5 �V
' �@�@�@�@�@�@labelProp6 �V
' �@�@�@�@�@�@labelProp7 �V
' �@�@�@�@�@�@labelProp8 �V
' �@�@�@�@�@�@labelProp9 �V
' �@�@�@�@�@�@labelProp10 �V
'
' �߂�l�@�@�F����
' =========================================================
Public Sub addAll(ByVal col As ValCollection _
                , Optional ByVal labelProp1 As String = "" _
                , Optional ByVal labelProp2 As String = "" _
                , Optional ByVal labelProp3 As String = "" _
                , Optional ByVal labelProp4 As String = "" _
                , Optional ByVal labelProp5 As String = "" _
                , Optional ByVal labelProp6 As String = "" _
                , Optional ByVal labelProp7 As String = "" _
                , Optional ByVal labelProp8 As String = "" _
                , Optional ByVal labelProp9 As String = "" _
                , Optional ByVal labelProp10 As String = "" _
                , Optional ByVal isAppend As Boolean = False)

    Dim var As Variant
    
    Dim i As Long
    
    Dim Label1 As String
    Dim Label2 As String
    
    
    ' ���X�g�R���g���[������v�f��S�č폜����
    If Not isAppend Then
        m_control.clear
        m_collection.removeAll
        i = 0
    Else
        i = m_control.ListCount
    End If
    
    For Each var In col.col
    
        If labelProp1 <> "" Or _
           labelProp2 <> "" Or _
           labelProp3 <> "" Or _
           labelProp4 <> "" Or _
           labelProp5 <> "" Or _
           labelProp6 <> "" Or _
           labelProp7 <> "" Or _
           labelProp8 <> "" Or _
           labelProp9 <> "" Or _
           labelProp10 <> "" _
        Then
            addItemByProp var, labelProp1, labelProp2, labelProp3, labelProp4, labelProp5, labelProp6, labelProp7, labelProp8, labelProp9, labelProp10
        Else
            addItem var, var
        End If
        
        i = i + 1
    Next
    
End Sub

' =========================================================
' ���v�f�ǉ�
'
' �T�v�@�@�@�F�l�X�g���ꂽ�v���p�e�B���炳��Ƀv���p�e�B���擾�����X�g�ɒǉ�����B
'
' �����@�@�@�Fcol       �R���N�V����
'             nestedProp �l�X�g���ꂽ�v���p�e�B
' �@�@�@�@�@�@labelProp1 ���x���Ƃ��ė��p����ycol�z�Ɋi�[����Ă���v�f�̃v���p�e�B
' �@�@�@�@�@�@labelProp2 ���x���Ƃ��ė��p����ycol�z�Ɋi�[����Ă���v�f�̃v���p�e�B
'
' �߂�l�@�@�F����
' =========================================================
Public Sub addNestedProperty(ByVal col As ValCollection _
                            , ByVal nestedProp As String _
                            , Optional ByVal labelProp1 As String = "" _
                            , Optional ByVal labelProp2 As String = "")

    Dim var  As Variant
    Dim var2 As Variant
    
    Dim i As Long: i = 0
    
    Dim Label1 As String
    Dim Label2 As String
    
    
    ' ���X�g�R���g���[������v�f��S�č폜����
    m_control.clear
    
    ' ���X�g�R���g���[���ɗv�f��ǉ�����
    ' ���x���v���p�e�B1�E2���ݒ肳��Ă���ꍇ
    If labelProp1 <> "" And labelProp2 <> "" Then
    
        For Each var In col.col
        
            ' �l�X�g���ꂽ�v���p�e�B�I�u�W�F�N�g
            Set var2 = CallByName(var, nestedProp, VbGet)
            
            If Not var2 Is Nothing Then
            
                Label1 = CallByName(var2, labelProp1, VbGet)
                Label2 = CallByName(var2, labelProp2, VbGet)
                
                m_control.addItem
                m_control.list(i, 0) = Label1
                m_control.list(i, 1) = Label2
                
            Else
            
                m_control.addItem
                m_control.list(i, 0) = ""
                m_control.list(i, 1) = ""
            End If
            
            i = i + 1
        Next
    
    ' ���x���v���p�e�B1���ݒ肳��Ă���ꍇ
    ElseIf labelProp1 <> "" Then
    
        For Each var In col.col
        
            ' �l�X�g���ꂽ�v���p�e�B�I�u�W�F�N�g
            Set var2 = CallByName(var, nestedProp, VbGet)
            
            If Not var2 Is Nothing Then
            
                Label1 = CallByName(var2, labelProp1, VbGet)
            
                m_control.addItem
                m_control.list(i, 0) = Label1
                
                
            Else
            
                m_control.addItem
                m_control.list(i, 0) = ""
                m_control.list(i, 1) = ""
            End If
            
            i = i + 1
        Next
    
    ' ���x���v���p�e�B���ݒ肳��Ă��Ȃ��ꍇ�Acol����擾�����v�f���g�����X�g�ɒǉ�����
    Else
    
        For Each var In col.col
        
            m_control.addItem var
        Next
    
    End If
    
    Set m_collection = col
    
End Sub

' =========================================================
' ���v�f�폜
'
' �T�v�@�@�@�F�v�f���폜����B
'
' �����@�@�@�Findex �Y�����i0���ŏ��l�Ƃ���j
'
' �߂�l�@�@�F����
' =========================================================
Public Sub removeItem(ByVal index As Long)

    ' ���X�g�R���g���[������v�f���폜
    m_control.removeItem index
    ' �R���N�V��������v�f���폜�i�R���N�V�����̓Y�����̍ŏ��l��1�Ȃ̂�+1����j
    m_collection.removeByIndex index + 1

End Sub

' =========================================================
' ���v�f�폜
'
' �T�v�@�@�@�F�v�f���܂Ƃ߂č폜����B
'
' �����@�@�@�F����
' �߂�l�@�@�F����
' =========================================================
Public Sub removeAll()

    Dim i As Long
    
    ' ���X�g�R���g���[������v�f���폜
    If Not m_control Is Nothing Then
        m_control.clear
    End If
    
    Set m_collection = New ValCollection
    
End Sub

' =========================================================
' ���I���ςݗv�f�擾
'
' �T�v�@�@�@�F���X�g�R���g���[���ɂđI������Ă���v�f���擾����
'
' �����@�@�@�F����
' �߂�l�@�@�F����
' =========================================================
Public Function getSelectedList(Optional ByVal t As VbVarType = vbObject) As ValCollection

    ' �߂�l
    Dim result As ValCollection

    ' �߂�l������������
    Set result = New ValCollection

    Dim i As Long
    
    ' ���X�g�R���g���[�������[�v������
    For i = 0 To m_control.ListCount - 1
    
        ' ���X�g�R���g���[���ɂđI������Ă��邩�𔻒肷��
        If m_control.selected(i) = True Then
        
            ' �߂�l�ɗv�f��ݒ肷��
            result.setItem m_collection.getItemByIndex(i + 1, t)
        
        End If
        
    Next
    
    ' �߂�l��ݒ肷��
    Set getSelectedList = result

End Function

' =========================================================
' ���I���ς݃C���f�b�N�X�擾�B
'
' �T�v�@�@�@�F�I���ς݂̃C���f�b�N�X���擾����B
'
' �����@�@�@�F����
' �߂�l�@�@�F�I���ς݃C���f�b�N�X
' =========================================================
Public Function getSelectedIndex() As Long

    getSelectedIndex = m_control.ListIndex

End Function

' =========================================================
' ���I���ςݍ��ڎ擾�B
'
' �T�v�@�@�@�F�I���ς݂̍��ڂ��擾����B
'
' �����@�@�@�F����
' �߂�l�@�@�F�I���ςݍ���
' =========================================================
Public Function getSelectedItem(Optional ByVal t As VbVarType = vbObject) As Variant

    Dim selectedIndex As Long
    selectedIndex = getSelectedIndex
    
    If t = vbObject Then
        Set getSelectedItem = getItem(selectedIndex, t)
    Else
        getSelectedItem = getItem(selectedIndex, t)
    End If

End Function

' =========================================================
' ���C�ӂ̑I�����̑I���B
'
' �T�v�@�@�@�F�C�ӂ̑I������I������B
'
' �����@�@�@�Findex �C�ӂ̑I�����̃C���f�b�N�X
' �߂�l�@�@�F����
' =========================================================
Public Sub setSelectedIndex(ByVal index As Long)

    If m_control.ListCount = 0 Or m_control.ListCount <= index Then
        Exit Sub
    End If

    m_control.ListIndex = index

End Sub

' =========================================================
' ���S�Ă̑I�����̑I����Ԃ�ύX�B
'
' �T�v�@�@�@�F
'
' �����@�@�@�FisSelected �I����ԁitrue �I������Afalse �I�����Ȃ��j
' �߂�l�@�@�F����
' =========================================================
Public Sub setSelectedAll(ByVal isSelected As Boolean)

    Dim i As Long
    
    For i = 0 To m_control.ListCount - 1
        m_control.selected(i) = isSelected
    Next

End Sub

' =========================================================
' ���C�ӂ̍��ڂ����݂��邩�𔻒肷��B
'
' �T�v�@�@�@�F
'
' �����@�@�@�Fvalue �l
' �߂�l�@�@�FTrue ���ڂ����݂���AFalse ���ڂ����݂��Ȃ�
' =========================================================
Public Function exist(ByVal value As String) As Boolean

    Dim i As Long

    ' ���X�g�R���g���[�������[�v������
    For i = 0 To m_control.ListCount - 1
    
        ' ���X�g�R���g���[���ɂđI������Ă��邩�𔻒肷��
        If m_control.list(i) = value Then
        
            exist = True
            Exit Function
        
        End If
        
    Next

    exist = False

End Function

' =========================================================
' �������擾�B
'
' �T�v�@�@�@�F�������擾����B
'
' �����@�@�@�F����
' �߂�l�@�@�F����
' =========================================================
Public Function count() As Long

    count = m_control.ListCount

End Function