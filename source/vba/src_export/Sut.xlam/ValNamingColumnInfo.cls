VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValNamingColumnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���̂�\������N���X (�J�������)
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/06/29�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' �e�I�u�W�F�N�g
Private m_parent As ValNamingTableInfo

' �J�����_����
Private m_columnNameLog  As String
' �J����������
Private m_columnName     As String

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
' ���N���X���������\�b�h (�����t��)
' =========================================================
Public Sub initial(ByVal columnName As String, ByVal columnNameLog As String, ByRef parent As ValNamingTableInfo)

    Me.columnName = columnName
    Me.columnNameLog = columnNameLog
    Me.parent = parent
    
End Sub

Public Property Get parent() As Variant

    Set parent = m_parent
End Property

Public Property Let parent(ByVal vNewValue As Variant)

    Set m_parent = vNewValue
End Property

Public Property Get columnNameLog() As Variant

    columnNameLog = m_columnNameLog
End Property

Public Property Let columnNameLog(ByVal vNewValue As Variant)

    m_columnNameLog = vNewValue
End Property

Public Property Get columnName() As Variant

    columnName = m_columnName
End Property

Public Property Let columnName(ByVal vNewValue As Variant)

    m_columnName = vNewValue
End Property