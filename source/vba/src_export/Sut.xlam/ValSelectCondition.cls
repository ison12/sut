VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSelectCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' SELECT SQL�̏���
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/09�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �J�������X�g
Private columnList   As ValCollection
' �l���X�g
Private valueList    As ValCollection
' ���у��X�g
Private orderList    As ValCollection

' ���R�[�h�擾 �͈� �J�n
Public recRangeStart As Long
' ���R�[�h�擾 �͈� �I��
Public recRangeEnd   As Long


' ---------------------------------------------------------
' ���萔
' ---------------------------------------------------------
' ORDER BY ASC
Public Property Get ORDER_ASC() As Long

    ORDER_ASC = 1
End Property

' ORDER BY DESC
Public Property Get ORDER_DESC() As Long

    ORDER_DESC = 0
End Property

' ORDER BY�̎w��Ȃ�
Public Property Get ORDER_NONE() As Long

    ORDER_NONE = -1
End Property

' ���R�[�h�擾�͈͎w��Ȃ�
Public Property Get REC_RANGE_NONE() As Long

    REC_RANGE_NONE = -1

End Property

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

    ' ���X�g������������
    Set columnList = New ValCollection
    
    ' ���X�g������������
    Set valueList = New ValCollection
    
    ' ���X�g������������
    Set orderList = New ValCollection
    
    ' ���R�[�h�擾�͈͂̏����l�𖳂��ɂ���
    recRangeStart = REC_RANGE_NONE
    recRangeEnd = REC_RANGE_NONE

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' ���X�g��j������
    Set columnList = Nothing
    
    ' ���X�g��j������
    Set valueList = Nothing
    
    ' ���X�g��j������
    Set orderList = Nothing
    
End Sub

' =========================================================
' ���J�������X�g�擾
'
' �T�v�@�@�@�F�J�������X�g���擾����B
' �����@�@�@�F
' �߂�l�@�@�F�J�������X�g
'
' =========================================================
Public Function getColumnList() As ValCollection

    Set getColumnList = columnList
End Function

' =========================================================
' �������ǉ�
'
' �T�v�@�@�@�FSELECT�̏�����ǉ�����B
' �����@�@�@�Fcolumn  �J����
'             value   �l
' �@�@�@�@�@�@order   ���я�
'
' =========================================================
Public Sub setCondition(ByVal column As String, ByVal value As String, ByVal order As Long)

    ' ���X�g�ɃJ��������ݒ�
    columnList.setItem column, column
    ' ���X�g�ɒl��ݒ�i�L�[���J�������Ƃ���j
    valueList.setItem value, column
    ' ���X�g�ɃI�[�_�[�l��ݒ�i�L�[���J�������Ƃ���j
    orderList.setItem order, column
    
End Sub

' =========================================================
' �������擾�i�l�j
'
' �T�v�@�@�@�FSELECT�̏������擾����B
' �����@�@�@�Fcolumn  �J����
' �߂�l�@�@�F�l
'
' =========================================================
Public Function getConditionValue(ByVal column As String) As String

    If valueList.exist(column) = True Then
    
        getConditionValue = valueList.getItem(column, vbString)
    Else
    
        getConditionValue = ""
    End If

End Function

' =========================================================
' �������擾�i���я��j
'
' �T�v�@�@�@�FSELECT�̏������擾����B
' �����@�@�@�Fcolumn  �J����
' �߂�l�@�@�F���я�
'
' =========================================================
Public Function getConditionOrder(ByVal column As String) As Long

    If orderList.exist(column) = True Then
    
        getConditionOrder = orderList.getItem(column, vbLong)
    Else
    
        getConditionOrder = ORDER_NONE
    End If

End Function