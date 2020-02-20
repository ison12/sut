VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIMenuDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' ���j���[���쎞�̏��� - DB
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/01/24�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ________________________________________________________
' �����o�ϐ�
' ________________________________________________________
' ��DB�ڑ��t�H�[���i�C�x���g��M�p�ɒ�`�j
Private WithEvents frmDBConnectVar As frmDBConnect
Attribute frmDBConnectVar.VB_VarHelpID = -1

' ________________________________________________________
' �C�x���g
' ________________________________________________________
' =========================================================
' ��DB�ڑ��t�H�[����OK�{�^�����������ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub frmDBConnectVar_ok(ByVal connStr As String, ByVal connSimpleStr As String, ByVal connectInfo As ValDBConnectInfo)

    On Error GoTo err

    ' DB�ڑ�����Ă���ꍇ
    If Not Main.dbConn Is Nothing Then
    
        ' �ؒf����
        Main.SutUpdateDbConn Nothing, Empty, Empty
    End If
    
    ' DB�ڑ����ϐ��ɐݒ肷��
    Dim conn As Object
    Set conn = ADOUtil.connectDb(connStr)
    
    ' DB�ڑ��������ݒ肷��
    Main.SutUpdateDbConn conn, connStr, connSimpleStr
    
    Exit Sub
    
err:

    ' �G���[����
    Main.SutUpdateDbConn Nothing, Empty, Empty

    Main.ShowErrorMessage

End Sub

' =========================================================
' ��DB�ڑ��t�H�[���ŏ������L�����Z�����ꂽ�ꍇ�ɔ�������C�x���g
'
' �T�v�@�@�@�F
'
' =========================================================
Private Sub frmDBConnectVar_cancel()

    On Error GoTo err

    Exit Sub
    
err:

    ' �G���[����
    Main.ShowErrorMessage

End Sub

' =========================================================
' ���N���X���������\�b�h
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ���N���X�㏈�����\�b�h
' =========================================================
Private Sub Class_Terminate()

    ' �C�x���g��M�p��DB�ڑ��t�H�[����j������
    Set frmDBConnectVar = Nothing
End Sub


' =========================================================
' ��������
'
' �T�v�@�@�@�F���������������s����B
' �����@�@�@�F
'
' =========================================================
Public Sub init()

    ' DB�ڑ��t�H�[�������[�h����
    If VBUtil.unloadFormIfChangeActiveBook(frmDBConnect) Then Unload frmDBConnect
    Load frmDBConnect
    ' �C�x���g��M�p��DB�ڑ��t�H�[��������������
    Set frmDBConnectVar = frmDBConnect

End Sub

' =========================================================
' ��DB�ڑ�
'
' �T�v�@�@�@�FDB�ڑ� �������̏���
'
' =========================================================
Public Sub connectDb()

    ' DB�ڑ��t�H�[����\������
    frmDBConnectVar.ShowExt vbModal
End Sub

' =========================================================
' ��DB�ؒf
'
' �T�v�@�@�@�FDB�ؒf �������̏���
'
' =========================================================
Public Sub disconnectDB()

    ' �ؒf����
    Main.SutUpdateDbConn Nothing, Empty, Empty
    
End Sub

' =========================================================
' ��DB�ڑ����\��
'
' �T�v�@�@�@�F
'
' =========================================================
Public Sub showDBConnectInfo(ByVal dbConn As Object)

    Const MSG_TITLE As String = "DB�ڑ����"

    If dbConn Is Nothing Then
        ' �ؒf���
        VBUtil.showMessageBoxForInformation "DB�ɐڑ�����Ă��܂���", MSG_TITLE
    
    ElseIf ADOUtil.getConnectionStatus(dbConn) = adStateClosed Then
        ' �ؒf���
        VBUtil.showMessageBoxForInformation dbConn.ConnectionString & vbNewLine & "DB�Ƃ̐ڑ����؂�Ă��܂�", MSG_TITLE
    Else
        ' �ڑ����
        
        ' ------------------------------------------------------
        ' �ڑ�������𕪉����ă^�u��؂�E���s���Ƃɕ�������
        Dim connStrArr     As Variant
        Dim connStrInfo    As Variant
        Dim connStrInfoArr As Variant
        
        Dim connStrBuf As New StringBuilder
        
        Dim i As Long
        
        connStrArr = Split(dbConn.ConnectionString, ";")
        
        For i = LBound(connStrArr) To UBound(connStrArr)
        
            connStrInfo = Trim$(connStrArr(i))
            If connStrInfo <> Empty Then
            
                connStrInfoArr = Split(connStrInfo, "=")
                
                If VBUtil.arraySize(connStrInfoArr) >= 2 Then
                    connStrBuf.append "  " & Trim$(connStrInfoArr(0)) & "=" & Trim$(connStrInfoArr(1))
                Else
                    connStrBuf.append "  " & Trim$(connStrInfoArr(0)) & "="
                End If
            
                connStrBuf.append vbLf
            End If
        
        Next
        
        If connStrBuf.length > 0 Then
            connStrBuf.remove connStrBuf.length, 1
        End If
        ' ------------------------------------------------------
        
        VBUtil.showMessageBoxForInformation _
            "�ڑ�������" & vbTab & ": " & vbLf & _
            connStrBuf.str & vbLf & _
            "�ڑ��^�C���A�E�g" & vbTab & ": " & dbConn.ConnectionTimeout & vbLf & _
            "�R�}���h�^�C���A�E�g" & vbTab & ": " & dbConn.CommandTimeout & vbLf & _
            "�������x��" & vbTab & ": " & dbConn.IsolationLevel & vbLf & _
            "�f�t�H���gDB" & vbTab & ": " & dbConn.defaultdatabase & vbLf & _
            "ADO�o�[�W����" & vbTab & ": " & dbConn.version & _
            "", MSG_TITLE
    End If

End Sub
