VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSnapRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�i�b�v�V���b�g�̃��R�[�h���
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/01/06�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �w�b�_���X�g
Public headerList As Variant
' ���R�[�h���X�g
Public recordList As Variant
' SQL
Public sql As String
' ����
Public memo As String
' PK���X�g
Public pkList As ValCollection
' PK���X�g�i�C���f�b�N�X�ւ̕ϊ��j
Public pkListIndex As ValCollection

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()
    
End Sub

' =========================================================
' ��PK�����񃊃X�g����PK�C���f�b�N�X���X�g�ɕϊ����郁�\�b�h
' =========================================================
Public Function convertPkListToArray() As Variant

    Dim ret As Variant
    
    If pkList Is Nothing Then
        Set convertPkListToArray = Empty
        Exit Function
    End If
    
    ReDim ret(1 To pkList.count)
    
    Dim i As Variant: i = 1
    Dim pk As Variant
    
    For Each pk In pkList.col
        ret(i) = pk.pk
        i = i + 1
    Next
    
    ' �߂�l��ԋp����
    convertPkListToArray = ret
    
End Function

' =========================================================
' ��PK�����񃊃X�g����PK�C���f�b�N�X���X�g�ɕϊ����郁�\�b�h
' =========================================================
Public Function convertToPkIndexList() As ValCollection
    
    If Not pkListIndex Is Nothing Then
        ' ���ɕϊ��ς݂̏ꍇ�́A�ϊ��ς݂̃��X�g��ԋp����
        Set convertToPkIndexList = pkListIndex
    End If
    
    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim headerName As String
    Dim headerIndex As Long
    
    Dim pk As Variant
    
    ' PK�ƈ�v����w�b�_���𑖍����A�w�b�_���X�g�̈ʒu���擾����
    For Each pk In pkList.col
    
        For headerIndex = LBound(headerList) To UBound(headerList)
        
            headerName = headerList(headerIndex)
            If pk.pk = headerName Then
                ret.setItem headerIndex
                GoTo end_loop
            End If
        
        Next
end_loop:
    
    Next
    
    If pkList.count <> ret.count Then
        ' PK���X�g�̌��Ɩ߂�l�̌�����v���Ȃ��ꍇ�́A�G���[
        err.Raise ConstantsError.ERR_NUMBER_PK_COLUMN_NOT_FOUND _
            , err.Source _
            , ConstantsError.ERR_DESC_PK_COLUMN_NOT_FOUND
    End If
    
    ' �t�B�[���h�ɃL���b�V������
    Set pkListIndex = ret
    
    ' �߂�l��ԋp����
    Set convertToPkIndexList = ret
    
    
End Function

' =========================================================
' ��PK������擾�������\�b�h
' =========================================================
Public Function getPkStr(ByVal recordIndex As Long) As String
  
    getPkStr = Empty
    
    Dim pk As ValSnapSqlDefineSqlPk
    
    If IsEmpty(recordList) Then
        ' ���R�[�h���Ȃ��ꍇ�́APK���Ȃ�
        Exit Function
    End If
    
    ' ���R�[�h�C���f�b�N�X�̒���
    recordIndex = recordIndex + LBound(recordList)
    
    If Not (LBound(recordList) <= recordIndex And recordIndex <= UBound(recordList)) Then
        ' �Y�����R�[�h�͈͊O�̏ꍇ�́APK���Ȃ�
        Exit Function
    End If
    
    Dim pkIndex As Variant
    Dim pkListIndex As ValCollection
    Set pkListIndex = convertToPkIndexList
    
    For Each pkIndex In pkListIndex.col
    
        getPkStr = recordList(recordIndex, pkIndex) & "_"
    Next
    
End Function

' =========================================================
' ��PK������擾�������\�b�h
' =========================================================
Public Function convertToMapKeyIsPkValIsReocrdIndex() As ValCollection

    Dim ret As ValCollection
    Set ret = New ValCollection
    
    If IsEmpty(recordList) Then
        ' ���R�[�h���Ȃ��ꍇ�́A��ŕԋp
        Set convertToMapKeyIsPkValIsReocrdIndex = ret
        Exit Function
    End If

    Dim recordIndex As Long
    Dim pkStr As String
    
    For recordIndex = LBound(recordList) To UBound(recordList)
        pkStr = getPkStr(recordIndex)
        ret.setItem recordIndex, pkStr
    Next

    Set convertToMapKeyIsPkValIsReocrdIndex = ret

End Function

