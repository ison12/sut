VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValQueryParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �N�G���p�����[�^��\������N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/08�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���O
Private m_name    As String
' �l
Private m_value As String

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
' ���v���p�e�B
' =========================================================
Public Property Get name() As String

    name = m_name
End Property

Public Property Let name(ByVal vNewValue As String)

    m_name = vNewValue
End Property

Public Property Get value() As String

    value = m_value
End Property

Public Property Let value(ByVal vNewValue As String)

    m_value = vNewValue
End Property

' =========================================================
' ���^�u��؂�ł̃w�b�_���̎擾
' =========================================================
Public Property Get tabbedInfoHeader() As String

    tabbedInfoHeader = _
        """" & "Name" & """" & vbTab & _
        """" & "Value" & """"

End Property

' =========================================================
' ���^�u��؂�ł̏��̎擾
' =========================================================
Public Property Get tabbedInfo() As String

    tabbedInfo = _
        """" & replace(name, """", """""") & """" & vbTab & _
        """" & replace(value, """", """""") & """"

End Property


