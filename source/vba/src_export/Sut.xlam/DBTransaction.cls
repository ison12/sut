VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB�̃g�����U�N�V�������Ǘ�����I�u�W�F�N�g
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/04/12�@�V�K�쐬

' �g�p���@�F
' �@�{�N���X�̃I�u�W�F�N�g�� New ����B
' �@�֐��𔲂���Ǝ����� Class_Terminate ���Ă΂ꃍ�[���o�b�N�����B
' �@�R�~�b�g����ꍇ�́Acommit���\�b�h�𖾎��I�ɃR�[������B
'
'   Dim var As New DBTransaction: var.init
'
'
' *********************************************************

' �R�l�N�V�����I�u�W�F�N�g
Private m_conn As Object
' �R�~�b�g����Ă��邩�ǂ�����\���t���O
Private commited As Boolean

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()
    
End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' �R�~�b�g�t���O���I�t�̏ꍇ
    If commited = False Then
        
        ' ���[���o�b�N�����s����
        m_conn.RollbackTrans
        
        #If (DEBUG_MODE <> 1) Then
            Debug.Print "DBTransaction Rollback"
        #End If

    End If
End Sub

' =========================================================
' ������������
'
' �T�v�@�@�@�F�g�����U�N�V�������J�n����
' �����@�@�@�Fconn �R�l�N�V�����I�u�W�F�N�g
'
' =========================================================
Public Sub init(ByRef conn As Object)

    ' �R�~�b�g�t���O������������
    commited = False
    
    ' �R�l�N�V�����I�u�W�F�N�g��ݒ肷��
    Set m_conn = conn
    ' �g�����U�N�V�������J�n����
    m_conn.begintrans
End Sub

' =========================================================
' ���R�~�b�g����
'
' �T�v�@�@�@�F�g�����U�N�V�������R�~�b�g����
' �����@�@�@�F
'
' =========================================================
Public Sub commit()

    ' �g�����U�N�V�������R�~�b�g����
    m_conn.CommitTrans
    ' �R�~�b�g�t���O���I���ɂ���
    commited = True
End Sub