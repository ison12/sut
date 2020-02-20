VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' INI�f�[�^����N���X
'
' ��������
' �E�Z�N�V�����̏ȗ��s��
' �E�R�����g�ݒ�s��
' �E�f�[�^�{�����A���s�R�[�h��C�R�[���L��������Ȑ���R�[�h�ɕϊ������
'
' �쐬�ҁ@�FIson
' �����@�@�F2020/02/12�@�V�K�쐬
'
' ���L�����F
' *********************************************************
' �f�[�^
Private datas_ As ValCollection
' �ύX�ς݃t���O
Private isModified_ As Boolean

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Public Sub Class_Initialize()
    
    datas = New ValCollection
    
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
' �����@�@�@�FfileName_ �t�@�C����
' �߂�l�@�@�F
'
' =========================================================
Public Sub init(ByVal contents As String)
    
    isModified = False
    datas = readContents(contents)

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
    
    datas = Nothing

End Sub

' =========================================================
' ��INI�t�@�C������C�ӂ̃L�[�l���擾����
'
' �T�v�@�@�@�F
' �����@�@�@�Fsection �Z�N�V����
' �@�@�@�@�@�@key     �L�[
' �߂�l�@�@�F�Z�N�V�����ƃL�[�ɕR�Â��l
'
' =========================================================
Public Function getValue _
            (ByVal SECTION As String, _
             ByVal key As String) As String
             
    ' �߂�l�̏����l
    getValue = ""
             
    Dim dataSection As ValCollection
    Dim data As Variant
             
    Set dataSection = datas.getItem(SECTION)
    If Not dataSection Is Nothing Then

        data = dataSection.getItem(key, vbVariant)
        If IsArray(data) Then
            getValue = data(2)
        End If
        
    End If
    
End Function

' =========================================================
' ��INI�t�@�C���ɔC�ӂ̃L�[�l��ݒ肷��
'
' �T�v�@�@�@�F
' �����@�@�@�Fsection �Z�N�V����
' �@�@�@�@�@�@key     �L�[
'             value   �l
' �߂�l�@�@�FTrue�̏ꍇ�A����ɏ������݊���
'
' =========================================================
Public Sub setValue _
            (ByVal SECTION As String, _
             ByVal key As String, _
             ByVal value As String)
             
    Dim dataSection As ValCollection
             
    If datas.exist(SECTION) = False Then
    
        Set dataSection = New ValCollection
        datas.setItem dataSection, SECTION
        
    Else
        Set dataSection = datas.getItem(SECTION)
        
    End If

    dataSection.setItem Array(SECTION, key, value), key

    ' �ύX�����̂ŕύX�ς݂Ƃ���
    isModified = True
    
End Sub

' =========================================================
' ��INI�t�@�C���̔C�ӂ̃Z�N�V�����ɕ����̃L�[�l��ݒ肷��
'
' �T�v�@�@�@�F�߂�l��ValCollection�ŗv�f�͔z��ƂȂ�
'             ------------------------------
'             element(0 to 2)�̔z��
'             ------------------------------
' �@�@�@�@�@�@element(0): �Z�N�V����
' �@�@�@�@�@�@element(1): �L�[
' �@�@�@�@�@�@element(2): �l
'
' �����@�@�@�Fsection �Z�N�V����
' �߂�l�@�@�F2�����z��
'
' =========================================================
Public Function getValues _
            (ByVal SECTION As String) As ValCollection
             
    Dim dataSection As ValCollection
             
    If datas.exist(SECTION) Then
        
        Set getValues = datas.getItem(SECTION)
        
    Else
        Set getValues = New ValCollection
        
    End If

End Function

' =========================================================
' ��INI�t�@�C���̔C�ӂ̃Z�N�V�����ɕ����̃L�[�l��ݒ肷��
'
' �T�v�@�@�@�F����values��ValCollection�ł��邱�ƁB
' �@�@�@�@�@�@�v�f�͔z���2�p�^�[���z��B
'
'             �p�^�[��1
'             ------------------------------
'             element(0 to 2)�̔z��
'             ------------------------------
' �@�@�@�@�@�@element(0): �Z�N�V����
' �@�@�@�@�@�@element(1): �L�[
' �@�@�@�@�@�@element(2): �l
'
'             �p�^�[��2
'             ------------------------------
'             element(0 to 1)�̔z��
'             ------------------------------
' �@�@�@�@�@�@element(0): �L�[
' �@�@�@�@�@�@element(1): �l

' �����@�@�@�Fsection �Z�N�V����
'             values  �����̒l
' �߂�l�@�@�FTrue�̏ꍇ�A����ɏ������݊���
'
' =========================================================
Public Sub setValues _
            (ByVal SECTION As String, _
             ByRef values As ValCollection)
             
    Dim dataSection As ValCollection
    Dim data As Variant
             
    If datas.exist(SECTION) = False Then
        Set dataSection = New ValCollection
        datas.setItem dataSection, SECTION
    Else
        Set dataSection = datas.getItem(SECTION)
    End If

    Dim valuesOfElement As Variant
    
    For Each valuesOfElement In values.col
        
        If UBound(valuesOfElement) - LBound(valuesOfElement) + 1 >= 3 Then
            ' �z��̃T�C�Y��3�Ȃ�Z�N�V��������
            dataSection.remove valuesOfElement(1)
            dataSection.setItem Array(valuesOfElement(0), valuesOfElement(1), valuesOfElement(2), valuesOfElement(1))
            
        Else
            ' �z��̃T�C�Y��3�����Ȃ�Z�N�V�����Ȃ�
            dataSection.remove valuesOfElement(0)
            dataSection.setItem Array(SECTION, valuesOfElement(0), valuesOfElement(1), valuesOfElement(0))
        End If
        
    Next

    ' �ύX�����̂ŕύX�ς݂Ƃ���
    isModified = True
    
End Sub

' =========================================================
' ��INI�t�@�C���̔C�ӂ̃Z�N�V�����܂��̓L�[���폜����
'
' �T�v�@�@�@�F����key���ȗ������ꍇ�A����section�z���̑S�ẴL�[�l���폜�����B
' �@�@�@�@�@�@����key���ȗ����Ȃ������ꍇ�A�L�[�l���폜�����B
'
' �����@�@�@�Fsection �Z�N�V����
' �@�@�@�@�@�@key     �L�[
' �߂�l�@�@�FTrue�̏ꍇ�A����ɏ������݊���
'
' =========================================================
Public Sub delete _
            (ByVal SECTION As String, _
             Optional ByVal key As String = "")

             
    Dim dataSection As ValCollection
             
    Set dataSection = datas.getItem(SECTION)

    If dataSection Is Nothing Then
        ' �㑱�̏�����Nothing�G���[���������Ȃ��悤�ɋ�I�u�W�F�N�g��ݒ肵�Ă���
        Set dataSection = New ValCollection
    End If
    
    If key = "" Then
        
        ' Section�̍폜
        dataSection.removeAll
        
    Else
        
        ' Key�̍폜
        dataSection.remove key
        
    End If
    
    ' �ύX�����̂ŕύX�ς݂Ƃ���
    isModified = True
    
End Sub

' =========================================================
' ��INI�R���e���c�̓ǂݍ���
'
' �T�v�@�@�@�F
' �����@�@�@�Fcontents ���e
' �߂�l�@�@�F�R���e���c�f�[�^
'
' =========================================================
Public Function readContents(ByVal contents As String) As ValCollection

    ' �߂�l
    Dim datas As ValCollection
    Set datas = New ValCollection
    
    Dim sectionData As ValCollection
    
    ' ���s�R�[�h�𓝈ꂷ��
    contents = replace(contents, vbCr & vbLf, vbLf)
    contents = replace(contents, vbCr, vbLf)
    
    Dim key As String
    Dim val As String
    Dim findKeyValueSeparator As Long
    
    Dim currentSection As String
    
    Dim contentsArrayEle As String
    Dim contentsArray As Variant
    
    ' ���e�����s�R�[�h��؂�ɂ���
    contentsArray = Split(contents, vbLf)
    
    Dim i As Long
    For i = LBound(contentsArray) To UBound(contentsArray)
        ' ��s����������
    
        contentsArrayEle = contentsArray(i)
        
        If contentsArrayEle = Empty Then
            ' �󕶎���̏ꍇ�̓X�L�b�v����
            GoTo loop_end
        End If
        
        If Mid(contentsArrayEle, 1, 1) = "[" And _
           Mid(contentsArrayEle, Len(contentsArrayEle), 1) = "]" Then
           ' �Z�N�V�����̏ꍇ
            
            currentSection = decodeSpecialChars(Mid(contentsArrayEle, 2, Len(contentsArrayEle) - 2))
            
            Set sectionData = New ValCollection
            datas.setItem sectionData, currentSection
            
            GoTo loop_end
        End If
        
        If sectionData Is Nothing Then
            ' �Z�N�V�������Ȃ��ꍇ�́A�X�L�b�v�i�Z�N�V�������o�ꂷ��܂Ńf�[�^��ǂݍ��܂Ȃ��j
            GoTo loop_end
        End If
        
        ' �L�[�ƒl���擾����
        findKeyValueSeparator = InStr(contentsArrayEle, "=")
        If findKeyValueSeparator > 0 Then
            ' key=value
            key = decodeSpecialChars(Mid(contentsArrayEle, 1, findKeyValueSeparator - 1))
            val = decodeSpecialChars(Mid(contentsArrayEle, findKeyValueSeparator + 1))
        Else
            ' key
            key = decodeSpecialChars(contentsArrayEle)
            val = ""
        End If
        
        sectionData.setItem Array(currentSection, key, val), key
    
loop_end:

    Next
    
    Set readContents = datas
    
End Function

' =========================================================
' ��INI�t�@�C���̏������݃f�[�^�̐���
'
' �T�v�@�@�@�F
' �����@�@�@�Fdatas    �f�[�^
' �߂�l�@�@�F�������݃f�[�^
'
' =========================================================
Public Function createWriteContents(ByVal datas As ValCollection) As String

    Dim ret As New StringBuilder

    Dim sectionName As String
    Dim sectionData As ValCollection
    Dim keyValue As Variant
    
    For Each sectionData In datas.col
    
        If sectionData.count > 0 Then
            
            sectionName = sectionData.getItemByIndex(1, vbVariant)(0)
        
            ret.append "[" & encodeSpecialChars(sectionName) & "]" & vbNewLine
            
            For Each keyValue In sectionData.col
                ret.append encodeSpecialChars(keyValue(1)) & "=" & encodeSpecialChars(keyValue(2)) & vbNewLine
            Next
            
        End If
            
    Next
    
    createWriteContents = ret.str
    
End Function

' =========================================================
' ������ȕ������G���R�[�h����
'
' �T�v�@�@�@�F
' �����@�@�@�Fval �Ώە�����
' �߂�l�@�@�F�ϊ���̕�����
'
' =========================================================
Private Function encodeSpecialChars(ByVal val As String) As String
    
    encodeSpecialChars = val
    
    ' ASCII�R�[�h��17~19 ���u����R�[�h���p����
    encodeSpecialChars = replace(encodeSpecialChars, vbCr, Chr(17))
    encodeSpecialChars = replace(encodeSpecialChars, vbLf, Chr(18))
    encodeSpecialChars = replace(encodeSpecialChars, "=", Chr(19))
    
End Function

' =========================================================
' ������ȕ������f�R�[�h����
'
' �T�v�@�@�@�F
' �����@�@�@�Fval �Ώە�����
' �߂�l�@�@�F�ϊ���̕�����
'
' =========================================================
Private Function decodeSpecialChars(ByVal val As String) As String
    
    decodeSpecialChars = val
    
    ' ASCII�R�[�h��17~19 ���u����R�[�h���p����
    decodeSpecialChars = replace(decodeSpecialChars, Chr(17), vbCr)
    decodeSpecialChars = replace(decodeSpecialChars, Chr(18), vbLf)
    decodeSpecialChars = replace(decodeSpecialChars, Chr(19), "=")
    
End Function

' =========================================================
' ��INI�f�[�^�̎擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�f�[�^
'
' =========================================================
Public Property Get datas() As ValCollection

    Set datas = datas_

End Property

' =========================================================
' ��INI�f�[�^�̐ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvNewValue �f�[�^
' �߂�l�@�@�F
'
' =========================================================
Public Property Let datas(ByVal vNewValue As ValCollection)

    Set datas_ = vNewValue

End Property

' =========================================================
' ���ύX�ς݃t���O�̎擾
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F�ύX�ς݃t���O
'
' =========================================================
Public Property Get isModified() As Boolean
    
    isModified = isModified_

End Property

' =========================================================
' ���ύX�ς݃t���O�̐ݒ�
'
' �T�v�@�@�@�F
' �����@�@�@�FvNewValue �ύX�ς݃t���O
' �߂�l�@�@�F
'
' =========================================================
Public Property Let isModified(ByVal vNewValue As Boolean)
    
    isModified_ = vNewValue

End Property