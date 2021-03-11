VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' Scripting.Dictionary�̃��b�p�[�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/09/06�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �f�B�N�V���i���I�u�W�F�N�g
Public dic  As Object

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    Set dic = CreateObject("Scripting.Dictionary")
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    Set dic = Nothing
    
End Sub

' =========================================================
' �����ݔ���
'
' �T�v�@�@�@�F�L�[�ɕR�Â��I�u�W�F�N�g�����݂��Ă��邩�𔻒肷��B
' �����@�@�@�Fkey �I�u�W�F�N�g�ݒ莞�ɗ��p�����L�[
' �߂�l�@�@�F���݂��Ă���ꍇTrue
'
' =========================================================
Public Function exist(ByRef key As Variant) As Boolean

    exist = dic.Exists(key)
End Function

' =========================================================
' ���I�u�W�F�N�g�ݒ�
'
' �T�v�@�@�@�F�I�u�W�F�N�g��ݒ肷��B
' �@�@�@�@�@�@key���w�肷���getItem�ŃL�[���w�肵�ăI�u�W�F�N�g���擾���邱�Ƃ��\�B
' �����@�@�@�Fitem �I�u�W�F�N�g
' �@�@�@�@�@�@ key �L�[
' �߂�l�@�@�F����
'
' =========================================================
Public Sub setItem(ByRef item As Variant, ByRef key As Variant)

    Dim isKeyFlg As Boolean

    ' ����key��String�^�ɕϊ�
    Dim keyStr As String: keyStr = CStr(key)

    ' key�ɕR�Â��I�u�W�F�N�g�����݂��Ă��邩�ǂ����𔻒肷��
    If exist(keyStr) = True Then
    
        ' ���A�C�e������U�폜����
        dic.remove (keyStr)
        ' ���A�C�e����ǉ�
        dic.Add keyStr, item
    
    Else
    
        ' ���A�C�e����ǉ�����
        dic.Add keyStr, item
    
    End If

End Sub

' =========================================================
' ���I�u�W�F�N�g�擾
'
' �T�v�@�@�@�F�I�u�W�F�N�g���擾����B
' �@�@�@�@�@�@�ݒ肵���I�u�W�F�N�g��Object�^�������łȂ����𔻕ʂ��Ė߂�l�ɓK�؂Ȍ^��Ԃ��B
'
' �����@�@�@�Fkey �L�[
' �@�@�@�@�@�@   t �I�u�W�F�N�g�̌^
' �߂�l�@�@�F����
'
' =========================================================
Public Function getItem(ByRef key As Variant, Optional t As VbVarType = vbObject) As Variant
    
    ' ����key��String�^�ɕϊ�
    Dim keyStr As String
    
    keyStr = CStr(key)
    
    ' key�ɕR�Â��I�u�W�F�N�g�����݂��Ă��邩�ǂ����𔻒肷��
    If exist(keyStr) = False Then
    
        If t = vbObject Then
            
            Set getItem = Nothing
            
        Else
            getItem = Empty
        
        End If
    
        Exit Function
    End If
    
    ' �I�u�W�F�N�g�^
    If t = vbObject Then
    
        Set getItem = dic.item(keyStr)
    
    ' ���̑��̌^
    Else
    
        getItem = dic.item(keyStr)
    
    End If

End Function

' =========================================================
' �����X�g�J�E���g�擾
'
' �T�v�@�@�@�F���X�g�ɐݒ肳��Ă���I�u�W�F�N�g�̑������擾����B
' �����@�@�@�F����
' �߂�l�@�@�F����
'
' =========================================================
Public Function count() As Long

    count = dic.count

End Function

' =========================================================
' ���I�u�W�F�N�g�폜
'
' �T�v�@�@�@�F�I�u�W�F�N�g���폜����
' �����@�@�@�Fkey �L�[
' �߂�l�@�@�F����ɍ폜�����ꍇTrue
'
' =========================================================
Public Function remove(ByRef key As Variant) As Boolean
    
    ' ����key��String�^�ɕϊ�
    Dim keyStr As String: keyStr = CStr(key)
    
    If exist(keyStr) = True Then
    
        dic.remove keyStr
        
        remove = True
    Else
    
        remove = False
    End If
    
End Function

' =========================================================
' ���I�u�W�F�N�g�S�폜
'
' �T�v�@�@�@�F�I�u�W�F�N�g��S�č폜����
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub removeAll()

    dic.removeAll
    
End Sub

' =========================================================
' ���z��ϊ�
'
' �T�v�@�@�@�F�f�B�N�V���i����z��ɕϊ�����B
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Function toArray(Optional t As VbVarType = vbObject) As Variant

    ' �߂�l
    Dim ret As Variant
    
    If dic.count = 0 Then
    
        toArray = Empty
        Exit Function
    End If
    
    ' �z����m�ۂ���i�T�C�Y�͂P�`�f�B�N�V���i���̃T�C�Y�j
    ReDim ret(1 To dic.count)
    
    Dim a As Variant
    Dim i As Long
    
    a = dic.items
    
    For i = LBound(a) To UBound(a)
    
        ' �^���I�u�W�F�N�g�̏ꍇ
        If t = vbObject Then
        
            Set ret(i + 1) = a(i)
            
        ' ��L�ȊO
        Else
        
            ret(i + 1) = a(i)
        End If
    Next
    
    ' �߂�l��Ԃ�
    toArray = ret
    
End Function

' =========================================================
' ���f�B�N�V���i���I�u�W�F�N�g�̃R�s�[
'
' �T�v�@�@�@�F
' �����@�@�@�Fdic �f�B�N�V���i���I�u�W�F�N�g
' �߂�l�@�@�F�R�s�[�����I�u�W�F�N�g
'
' =========================================================
Public Function copy() As ValDictionary

    Dim ret As New ValDictionary
    
    Dim key As Variant
    
    For Each key In dic
        ret.setItem dic.item(key), key
    Next
    
    Set copy = ret

End Function
