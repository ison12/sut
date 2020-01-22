VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSnapshotDiff 
   Caption         =   "�X�i�b�v�V���b�g��r"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11265
   OleObjectBlob   =   "frmSnapshotDiff.frx":0000
End
Attribute VB_Name = "frmSnapshotDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' �X�i�b�v�V���b�g��r�t�H�[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2008/09/06�@�V�K�쐬
'
' ���L�����F
' *********************************************************

' ���C�x���g
' =========================================================
' �����s�C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�FsnapshotList �X�i�b�v�V���b�g���X�g
'             srcIndex ��r���C���f�b�N�X
'             desIndex ��r��C���f�b�N�X
'
' =========================================================
Public Event execDiff(ByRef snapShotList As ValCollection, ByVal srcIndex As Long, ByVal desIndex As Long)

' =========================================================
' ���L�����Z���C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
'
' =========================================================
Public Event cancel()

' ---------------------------------------------------------
' ���W�X�g���t�@�C���L�[
' ---------------------------------------------------------
Private Const REG_SUB_KEY_SNAPSHOT_DIFF As String = "snapshotDiff"

' �X�i�b�v�V���b�g���X�g
Private snapShotList        As ValCollection
' ��r���X�i�b�v�V���b�g���X�g
Private srcSnapshotList     As CntListBox
' ��r��X�i�b�v�V���b�g���X�g
Private desSnapshotList     As CntListBox


' =========================================================
' ���t�H�[���\��
'
' �T�v�@�@�@�F
' �����@�@�@�Fmodal         ���[�_���܂��̓��[�h���X�\���w��
'             snapshotList_ �X�i�b�v�V���b�g���X�g
' �߂�l�@�@�F
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef snapshotList_ As ValCollection)

    Set snapShotList = snapshotList_
    
    refreshLstSnapshotSrc
    refreshLstSnapshotDes lstSnapshotSrc.ListCount - 1
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal
    
End Sub

' =========================================================
' ���t�H�[����\��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Public Sub HideExt()

    Main.storeFormPosition Me.name, Me
    Me.Hide
    
End Sub

' =========================================================
' ���t�H�[������C�x���g�v���V�[�W��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)

    Main.storeFormPosition Me.name, Me
End Sub

' =========================================================
' �����鏈��
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdClose_Click()

    On Error GoTo err
    
    ' �t�H�[�������
    HideExt
    
    ' �L�����Z���C�x���g�𑗐M����
    RaiseEvent cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���X�i�b�v�V���b�g���X�g��r���ύX�C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub lstSnapshotSrc_Change()

    On Error GoTo err
    
    refreshLstSnapshotDes lstSnapshotSrc.ListIndex

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ����r���ʏo�̓C�x���g
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub cmdResultOut_Click()

    On Error GoTo err
    
    Dim srcIndex As Long
    Dim desIndex As Long

    srcIndex = lstSnapshotSrc.ListIndex
    desIndex = lstSnapshotDes.ListIndex

    If srcIndex = desIndex Then
        ' �����ɂȂ�Ȃ��͂�
        Exit Sub
    End If

    If srcIndex = -1 Or desIndex = -1 Then
        ' ���I�����
        Exit Sub
    End If
    
    If srcIndex <= desIndex Then
        ' �����ɂȂ�Ȃ��͂�
        Exit Sub
    End If
    
    RaiseEvent execDiff(snapShotList, srcIndex, desIndex)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ���X�i�b�v�V���b�g���X�g��r���X�V
'
' �T�v�@�@�@�F
' �����@�@�@�F
' �߂�l�@�@�F
'
' =========================================================
Private Sub refreshLstSnapshotSrc()

    Dim snapshot As ValSnapRecordsSet

    If srcSnapshotList Is Nothing Then
        Set srcSnapshotList = New CntListBox
        srcSnapshotList.init lstSnapshotSrc
    Else
        srcSnapshotList.removeAll
        srcSnapshotList.init lstSnapshotSrc
    End If
    
    For Each snapshot In snapShotList
    
        srcSnapshotList.addItem Format(snapshot.getDate, "yyyy/mm/dd hh:nn:ss") & " - " & snapshot.recordCount & "��", Empty
    Next
    
    If lstSnapshotSrc.ListCount > 0 Then
        lstSnapshotSrc.ListIndex = lstSnapshotSrc.ListCount - 1
    End If

End Sub

' =========================================================
' ���X�i�b�v�V���b�g���X�g��r��X�V
'
' �T�v�@�@�@�F
' �����@�@�@�FlstSnapshotSrcListIndex �X�i�b�v�V���b�g���X�g��r���C���f�b�N�X
' �߂�l�@�@�F
'
' =========================================================
Private Sub refreshLstSnapshotDes(ByVal lstSnapshotSrcListIndex As Long)

    Dim snapshot As ValSnapRecordsSet

    If desSnapshotList Is Nothing Then
        Set desSnapshotList = New CntListBox
        desSnapshotList.init lstSnapshotDes
    Else
        desSnapshotList.removeAll
        desSnapshotList.init lstSnapshotDes
    End If
    
    Dim i As Long
    
    i = 0
    For Each snapshot In snapShotList
    
        If i < lstSnapshotSrcListIndex Then
            desSnapshotList.addItem Format(snapshot.getDate, "yyyy/mm/dd hh:nn:ss") & " - " & snapshot.recordCount & "��", Empty
        End If
    
        i = i + 1
    Next
    
    If lstSnapshotDes.ListCount > 0 Then
        lstSnapshotDes.ListIndex = lstSnapshotDes.ListCount - 1
    End If

End Sub

