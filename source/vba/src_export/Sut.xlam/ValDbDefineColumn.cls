VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDbDefineColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�J������`�N���X
'
' �쐬�ҁ@�FIson
' �����@�@�F2019/12/07�@�V�K�쐬
'
' ���L�����F
' *********************************************************

Public schemaName As String
Public tableName As String
Public columnName As String
Public columnComment As String
Public dataType As String
Public dataTypeFormal As String
Public charLength As String
Public dataPrecision As String
Public dataScale As String
Public dateTimePrecision As String
Public default As String
Public isNull As Boolean

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