VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistryManipulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' INI�t�@�C������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/09/06�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���W�X�g���̃n���h��
#If VBA7 And Win64 Then
Private handle As LongPtr
#Else
Private handle As Long
#End If

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    handle = 0
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' �߂�l
    Dim ret As Long
    
    ' ���W�X�g�����J���Ă��Ȃ��ꍇ
    If handle = 0 Then
        ' ���������ɏI������
        Exit Sub
    End If
    
    ' �n���h�����N���[�Y����
    ret = WinAPI_ADVAP.RegCloseKey(handle)
    ' �߂�l���`�F�b�N
    If ret = WinAPI_ADVAP.ERROR_SUCCESS Then
        
        ' ����
        #If (DEBUG_MODE <> 1) Then
            Debug.Print "RegestryManipulator Class_Terminate Success"
        #End If

    Else
    
        ' ���s
        #If (DEBUG_MODE <> 1) Then
            Debug.Print "RegestryManipulator Class_Terminate Failed. error code " & ret
        #End If
    End If
    
    ret = 0
    
End Sub

' =========================================================
' �����������\�b�h
'
' �T�v�@�@�@�F���W�X�g���̔C�ӂ̃L�[���J��
' �����@�@�@�FhKey        ���[�g�L�[
' �@�@�@�@�@�@lpSubKey    �T�u�L�[
' �@�@�@�@�@�@samDesired  �Z�L�����e�B�A�N�Z�X��
' �@�@�@�@�@�@isCreate    �C�ӂ̃L�[���J���ۂɐV�K�쐬���邩�ǂ�����\���t���O
' �߂�l�@�@�F���������ꍇ True
'
' =========================================================
Public Function init(ByVal hKey As RegKeyConstants _
                   , ByVal lpSubKey As String _
                   , ByVal samDesired As RegAccessConstants _
                   , Optional ByVal isCreate As Boolean = False) As Boolean

    ' API�̖߂�l
    Dim ret  As Long
    
    ' �L�[���쐬����
    If isCreate = True Then
    
#If VBA7 And Win64 Then
        Dim lpdwDisposition As LongPtr
#Else
        Dim lpdwDisposition As Long
#End If
        
        ret = WinAPI_ADVAP.RegCreateKeyEx(hKey _
                                       , lpSubKey & vbNullChar _
                                       , 0 _
                                       , 0 _
                                       , WinAPI_ADVAP.REG_OPTION_NON_VOLATILE _
                                       , samDesired _
                                       , 0 _
                                       , handle _
                                       , lpdwDisposition)
        
        ' �߂�l���`�F�b�N
        If ret = WinAPI_ADVAP.ERROR_SUCCESS Then
            
            ' �L�[���V�K�ɍ쐬���ꂽ�ꍇ
            If lpdwDisposition = WinAPI_ADVAP.REG_CREATED_NEW_KEY Then
            
                ' ����
                #If (DEBUG_MODE <> 1) Then
                    Debug.Print "RegestryManipulator init Success (Created new key)"
                #End If

            ' �L�[�����ɑ��݂���ꍇ
            Else
            
                ' ����
                #If (DEBUG_MODE <> 1) Then
                    Debug.Print "RegestryManipulator init Success (Opened existing key)"
                #End If
            End If
            
        Else
        
            ' ���s
            #If (DEBUG_MODE <> 1) Then
                Debug.Print "RegestryManipulator init Failed. error code " & ret
            #End If
            
            err.Raise ConstantsError.ERR_NUMBER_REGISTRY_ACCESS_FAILED _
                    , _
                    , ConstantsError.ERR_DESC_REGISTRY_ACCESS_FAILED & vbNewLine & ret
        End If
    
    ' �L�[���I�[�v������
    Else
    
        ret = WinAPI_ADVAP.RegOpenKeyEx(hKey _
                                     , lpSubKey _
                                     , 0 _
                                     , samDesired _
                                     , handle)
    
        ' �߂�l���`�F�b�N
        If ret = WinAPI_ADVAP.ERROR_SUCCESS Then
            
            ' ����
            #If (DEBUG_MODE <> 1) Then
                Debug.Print "RegestryManipulator init Success"
            #End If
        Else
        
            ' ���s
            #If (DEBUG_MODE <> 1) Then
                Debug.Print "RegestryManipulator init Failed. error code " & ret
            #End If
            
'            err.Raise ConstantsError.ERR_NUMBER_REGISTRY_ACCESS_FAILED _
'                    , _
'                    , ConstantsError.ERR_DESC_REGISTRY_ACCESS_FAILED & vbNewLine & ret

            init = False
            
            Exit Function
        End If
    
    End If

    init = True
End Function

' =========================================================
' ���l���擾����
'
' �T�v�@�@�@�F
' �����@�@�@�FvalueKey
' �߂�l�@�@�F�L�[�l
'
' =========================================================
Public Function getValue _
            (ByVal valueKey As String _
           , ByRef value As String) As Boolean

    ' byte�o�b�t�@�T�C�Y
#If VBA7 And Win64 Then
    Dim lpCbData As LongPtr
#Else
    Dim lpCbData As Long
#End If
    
    ' �擾�����f�[�^�^�C�v�̌^
#If VBA7 And Win64 Then
    Dim lpType As LongPtr
#Else
    Dim lpType As Long
#End If
    
    ' API�̖߂�l
    Dim ret As Long
    
    ' �L�[�̏I�[��NULL������t������
    valueKey = valueKey & vbNullChar
    
    ' �l�̃T�C�Y���擾����iUNICODE�o�[�W�����ŏ����擾�j
    ret = WinAPI_ADVAP.RegQueryValueExW(handle, StrPtr(valueKey), 0, lpType, 0, lpCbData)
    
    ' �֐��̌Ăяo���Ɏ��s
    If _
        ret <> WinAPI_ADVAP.ERROR_MORE_DATA _
    And ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        ' ���s
        getValue = False
        Exit Function
    End If
    
    Dim lpCbDataVar As Variant
    lpCbDataVar = lpCbData
    
    ' �o�b�t�@���m�ۂ���
    value = String(lpCbDataVar, vbNullChar)
    
    ' �l���擾����iUNICODE�o�[�W�����ŏ����擾�j
    ret = WinAPI_ADVAP.RegQueryValueExW(handle, StrPtr(valueKey), 0, lpType, StrPtr(value), LenB(value))
    
    ' �֐��̌Ăяo���Ɏ��s
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        ' ���s
        getValue = False
    Else
    
        getValue = True
    
        ' ������Ƃ��Đݒ肷��
        ' ��Unicode�o�[�W������API���g�p���邽�߁ALeftB��InStrB�͎g��Ȃ�
        If InStr(value, vbNullChar) <> 0 Then
        
            ' NULL���������݂���ꍇ�ANULL�������������O���̕�������擾����
            value = Left(value, InStr(value, vbNullChar) - 1)
        End If
    End If
    
    
End Function

' =========================================================
' ���l���擾����
'
' �T�v�@�@�@�F
' �����@�@�@�FvalueKey
' �߂�l�@�@�F�L�[�l
'
' =========================================================
Public Function getValueByByte _
            (ByVal valueKey As String _
           , ByRef value() As Byte) As Boolean

    ' byte�o�b�t�@�T�C�Y
#If VBA7 And Win64 Then
    Dim lpCbData As LongPtr
#Else
    Dim lpCbData As Long
#End If
    
    ' �擾�����f�[�^�^�C�v�̌^
#If VBA7 And Win64 Then
    Dim lpType As LongPtr
#Else
    Dim lpType As Long
#End If
    
    ' API�̖߂�l
    Dim ret As Long
    
    ' �L�[�̏I�[��NULL������t������
    valueKey = valueKey & vbNullChar
    
    ' �l�̃T�C�Y���擾����iUNICODE�o�[�W�����ŏ����擾�j
    ret = WinAPI_ADVAP.RegQueryValueExW(handle, StrPtr(valueKey), 0, lpType, 0, lpCbData)
    
    ' �֐��̌Ăяo���Ɏ��s
    If _
        ret <> WinAPI_ADVAP.ERROR_MORE_DATA _
    And ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        ' ���s
        getValueByByte = False
        Exit Function
    End If
    
    Dim lpCbDataVar As Variant
    lpCbDataVar = lpCbData

    ' �o�b�t�@���m�ۂ���
    ReDim value(0 To lpCbDataVar - 1)
    
    ' �l���擾����iUNICODE�o�[�W�����ŏ����擾�j
    ret = WinAPI_ADVAP.RegQueryValueExW(handle, StrPtr(valueKey), 0, lpType, VarPtr(value(0)), lpCbData)
    
    ' �֐��̌Ăяo���Ɏ��s
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        ' ���s
        getValueByByte = False
    Else
    
        getValueByByte = True
    
    End If
    
End Function

' =========================================================
' ���l��ݒ肷��
'
' �T�v�@�@�@�F
' �����@�@�@�FvalueKey �l�̃L�[
' �@�@�@�@�@�@value    �l
'
' =========================================================
Public Function setValue _
            (ByVal valueKey As String, _
             ByVal value As String) As Boolean
             
    ' API�̖߂�l
    Dim ret As Long
    
    ' �L�[�̏I�[��NULL������t������
    valueKey = valueKey & vbNullChar
    ' �l�̏I�[��NULL������t������
    value = value & vbNullChar
    
    ' �l��ݒ肷��iUNICODE�o�[�W�����ŏ���ݒ�j
    ret = WinAPI_ADVAP.RegSetValueExW(handle, StrPtr(valueKey), 0, WinAPI_ADVAP.REG_SZ, StrPtr(value), LenB(value))
    
    ' �G���[�̏ꍇ
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        setValue = False
        
    Else
    
        setValue = True
    End If
    
End Function

' =========================================================
' ���l��ݒ肷�� (Byte�^�o�[�W����)
'
' �T�v�@�@�@�F
' �����@�@�@�FvalueKey �l�̃L�[
' �@�@�@�@�@�@value    �l
'
' =========================================================
Public Function setValueByByte _
            (ByVal valueKey As String, _
             ByRef value() As Byte) As Boolean
             
    ' API�̖߂�l
    Dim ret As Long
    
    ' �L�[�̏I�[��NULL������t������
    valueKey = valueKey & vbNullChar
    
    ' �l��ݒ肷��iUNICODE�o�[�W�����ŏ���ݒ�j
    ret = WinAPI_ADVAP.RegSetValueExW(handle _
                                    , StrPtr(valueKey) _
                                    , 0 _
                                    , WinAPI_ADVAP.REG_BINARY _
                                    , VarPtr(value(0)) _
                                    , VBUtil.arraySize(value))
    
    ' �G���[�̏ꍇ
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        setValueByByte = False
        
    Else
    
        setValueByByte = True
    End If
    
End Function

' =========================================================
' ���L�[���X�g���擾����
'
' �T�v�@�@�@�F���݊J���Ă���L�[�̃T�u�L�[��S�Ď擾����B
' �����@�@�@�F
' �߂�l�@�@�F�L�[���X�g�A���s�����ꍇ��nothing�B
'
' =========================================================
Public Function getKeyList() As ValCollection

    ' �L�[���i�[����o�b�t�@������̃T�C�Y�i�萔�j
    Const STR_SIZE As Long = 256
    
    ' �L�[���i�[����o�b�t�@
    Dim str As String
    
    ' �擾�����f�[�^�^�C�v�̌^
#If VBA7 And Win64 Then
    Dim strSize As LongPtr
#Else
    Dim strSize As Long
#End If
    
    ' API�̖߂�l
    Dim ret As Long
    ' �C���f�b�N�X�ϐ�
    Dim i As Long
    ' �X�V����
    Dim ft As WinAPI_ADVAP.FILETIME
    
    ' ���ʃ��X�g
    Dim resultList As New ValCollection
    
    Do While True
    
        ' �����������������
        str = String(STR_SIZE, vbNullChar)
        ' ������̃T�C�Y������������
        strSize = STR_SIZE
        
        ' �L�[��񋓂���
        ret = WinAPI_ADVAP.RegEnumKeyEx(handle, i, str, strSize, 0, 0, 0, ft)
    
        ' �A�C�e�������݂���
        If ret = WinAPI_ADVAP.ERROR_MORE_DATA Then
        
            ' Delete NULL
            ' ��Unicode�o�[�W������API���g�p���邽�߁ALeftB��InStrB�͎g��Ȃ�
            If InStr(str, vbNullChar) <> 0 Then
            
                ' NULL���������݂���ꍇ�ANULL�������������O���̕�������擾����
                str = Left(str, InStr(str, vbNullChar) - 1)
            End If

            ' ���ʃ��X�g�ɃL�[���i�[����
            resultList.setItem str
        
        ' �A�C�e�������݂��Ȃ��ꍇ
        ElseIf ret = WinAPI_ADVAP.ERROR_NO_MORE_ITEMS Then
        
            ' ���[�v���I������
            Exit Do
            
        ' ���̑��̏ꍇ
        Else
        
            ' �G���[�Ɣ��肷��
            Set resultList = Nothing
            Exit Function
        End If
        
        i = i + 1
    Loop
    
    ' �߂�l��Ԃ�
    Set getKeyList = resultList
    
End Function

' =========================================================
' ���L�[���X�g���擾����
'
' �T�v�@�@�@�F���݊J���Ă���L�[�̃T�u�L�[��S�Ď擾����B
' �����@�@�@�FvalueNameList �l���̃��X�g
' �@�@�@�@�@�@valueList     �l���X�g
' �߂�l�@�@�F
'
' =========================================================
Public Sub getValueList(ByRef valueNameList As ValCollection _
                      , ByRef valueList As ValCollection)

    ' �L�[���i�[����o�b�t�@������̃T�C�Y�i�萔�j
    Const STR_SIZE As Long = 256

    ' �L�[���i�[����o�b�t�@
    Dim valueName As String
    ' �L�[��byte�o�b�t�@�T�C�Y
    Dim lpCbValueName As Long
    
    ' �l
    Dim value As String
    ' �l�̃o�b�t�@�T�C�Y
#If VBA7 And Win64 Then
    Dim lpCbData As LongPtr
#Else
    Dim lpCbData As Long
#End If
    
    
    ' �擾�����f�[�^�^�C�v�̌^
#If VBA7 And Win64 Then
    Dim lpType As LongPtr
#Else
    Dim lpType As Long
#End If
    
    ' API�̖߂�l
    Dim ret As Long
    ' �C���f�b�N�X�ϐ�
    Dim i As Long
    
    ' ���ʃ��X�g
    Set valueNameList = New ValCollection
    Set valueList = New ValCollection
    
    Do While True
    
        valueName = String(STR_SIZE, vbNullChar)
        
        ' �l��񋓂���
        ret = WinAPI_ADVAP.RegEnumValueW(handle, i, StrPtr(valueName), STR_SIZE, 0, lpType, 0, lpCbData)
    
        ' ����̏ꍇ
        If ret = WinAPI_ADVAP.ERROR_SUCCESS Then
        
            Dim lpCbDataVar As Variant
            lpCbDataVar = lpCbData
        
            ' �o�b�t�@���m�ۂ���
            value = String(lpCbDataVar, vbNullChar)
    
            ' �l���擾����
            ret = WinAPI_ADVAP.RegEnumValueW(handle, i, StrPtr(valueName), STR_SIZE, 0, lpType, StrPtr(value), lpCbData)
    
            ' ������Ƃ��Đݒ肷��
            ' ��Unicode�o�[�W������API���g�p���邽�߁ALeftB��InStrB�͎g��Ȃ�
            If InStr(valueName, vbNullChar) <> 0 Then
            
                ' NULL���������݂���ꍇ�ANULL�������������O���̕�������擾����
                valueName = Left(valueName, InStr(valueName, vbNullChar) - 1)
            End If
            
            If InStr(value, vbNullChar) <> 0 Then
            
                ' NULL���������݂���ꍇ�ANULL�������������O���̕�������擾����
                value = Left(value, InStr(value, vbNullChar) - 1)
            End If

            ' ���ʃ��X�g�ɃL�[���i�[����
            valueNameList.setItem valueName, valueName
            valueList.setItem value, valueName
        
            
        End If

        ' �A�C�e�������݂��Ȃ��ꍇ
        If ret = WinAPI_ADVAP.ERROR_NO_MORE_ITEMS Then
        
            ' ���[�v���I������
            Exit Do
            
        ' �G���[�����������ꍇ
        ElseIf ret <> WinAPI_ADVAP.ERROR_SUCCESS _
           And ret <> WinAPI_ADVAP.ERROR_MORE_DATA Then
        
            Exit Sub
        End If
        
        i = i + 1
    Loop
    
End Sub

' =========================================================
' �����݂̃L�[�ɕ����̒l��ݒ肷��
'
' �T�v�@�@�@�F����values��N�~2��2�����z��ł��邱�ƁB
' �@�@�@�@�@�@�ȉ��Ɉ���values�̍\���������B
' �@�@�@�@�@�@-------------------
' �@�@�@�@�@�@     0    1
'             0 / [key1][value1]
'             1 / [key2][value2]
'             2 / [key3][value3]
'             3 / [key4][value4]

' �����@�@�@�Fvalues  �����̒l
' �߂�l�@�@�F
'
' =========================================================
Public Sub setValues _
            (ByRef values() As Variant)

    Dim i As Long
    
    Dim valueKey As String
    Dim value    As String
    
    For i = LBound(values) To UBound(values)
    
        valueKey = CStr(values(i, 0))
        value = CStr(values(i, 1))
    
        setValue valueKey, value
    Next
    
End Sub

' =========================================================
' ���L�[���폜����
'
' �T�v�@�@�@�F
' �����@�@�@�Fkey �L�[
'
' =========================================================
Public Function delete(ByVal key As String) As Boolean

    Dim ret As Long
    
    key = key & vbNullChar
    ret = WinAPI_ADVAP.RegDeleteKeyW(handle, StrPtr(key))
    
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        delete = False
        
    Else
    
        delete = True
    End If
    
End Function

' =========================================================
' ���l���폜����
'
' �T�v�@�@�@�F
' �����@�@�@�FvalueKey �l�̃L�[
'
' =========================================================
Public Function deleteValue(ByVal valueKey As String) As Boolean

    Dim ret As Long
    
    valueKey = valueKey & vbNullChar
    ret = WinAPI_ADVAP.RegDeleteValueW(handle, StrPtr(valueKey))
    
    If ret <> WinAPI_ADVAP.ERROR_SUCCESS Then
    
        deleteValue = False
    Else
    
        deleteValue = True
    End If
    
End Function



