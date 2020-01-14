VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSnapshotDiff 
   Caption         =   "スナップショット比較"
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
' スナップショット比較フォーム
'
' 作成者　：Hideki Isobe
' 履歴　　：2008/09/06　新規作成
'
' 特記事項：
' *********************************************************

' ▽イベント
' =========================================================
' ▽実行イベント
'
' 概要　　　：
' 引数　　　：snapshotList スナップショットリスト
'             srcIndex 比較元インデックス
'             desIndex 比較先インデックス
'
' =========================================================
Public Event execDiff(ByRef snapShotList As GenericCollection, ByVal srcIndex As Long, ByVal desIndex As Long)

' =========================================================
' ▽キャンセルイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event Cancel()

' ---------------------------------------------------------
' レジストリファイルキー
' ---------------------------------------------------------
Private Const REG_SUB_KEY_SNAPSHOT_DIFF As String = "snapshotDiff"

' スナップショットリスト
Private snapShotList        As GenericCollection
' 比較元スナップショットリスト
Private srcSnapshotList     As CntListBox
' 比較先スナップショットリスト
Private desSnapshotList     As CntListBox


' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal         モーダルまたはモードレス表示指定
'             snapshotList_ スナップショットリスト
' 戻り値　　：
'
' =========================================================
Public Sub ShowExt(ByVal modal As FormShowConstants, ByRef snapshotList_ As GenericCollection)

    Set snapShotList = snapshotList_
    
    refreshLstSnapshotSrc
    refreshLstSnapshotDes lstSnapshotSrc.listCount - 1
    
    Main.restoreFormPosition Me.name, Me
    Me.Show modal
    
End Sub

' =========================================================
' ▽フォーム非表示
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Public Sub HideExt()

    Main.storeFormPosition Me.name, Me
    Me.Hide
    
End Sub

' =========================================================
' ▽フォーム閉じるイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Main.storeFormPosition Me.name, Me
End Sub

' =========================================================
' ▽閉じる処理
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdClose_Click()

    On Error GoTo err
    
    ' フォームを閉じる
    HideExt
    
    ' キャンセルイベントを送信する
    RaiseEvent Cancel

    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽スナップショットリスト比較元変更イベント
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
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
' ▽比較結果出力イベント
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub cmdResultOut_Click()

    On Error GoTo err
    
    Dim srcIndex As Long
    Dim desIndex As Long

    srcIndex = lstSnapshotSrc.ListIndex
    desIndex = lstSnapshotDes.ListIndex

    If srcIndex = desIndex Then
        ' 同じにならないはず
        Exit Sub
    End If

    If srcIndex = -1 Or desIndex = -1 Then
        ' 未選択状態
        Exit Sub
    End If
    
    If srcIndex <= desIndex Then
        ' 同じにならないはず
        Exit Sub
    End If
    
    RaiseEvent execDiff(snapShotList, srcIndex, desIndex)
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽スナップショットリスト比較元更新
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
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
    
        srcSnapshotList.addItem Format(snapshot.getDate, "yyyy/mm/dd hh:nn:ss") & " - " & snapshot.recordCount & "件", Empty
    Next
    
    If lstSnapshotSrc.listCount > 0 Then
        lstSnapshotSrc.ListIndex = lstSnapshotSrc.listCount - 1
    End If

End Sub

' =========================================================
' ▽スナップショットリスト比較先更新
'
' 概要　　　：
' 引数　　　：lstSnapshotSrcListIndex スナップショットリスト比較元インデックス
' 戻り値　　：
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
            desSnapshotList.addItem Format(snapshot.getDate, "yyyy/mm/dd hh:nn:ss") & " - " & snapshot.recordCount & "件", Empty
        End If
    
        i = i + 1
    Next
    
    If lstSnapshotDes.listCount > 0 Then
        lstSnapshotDes.ListIndex = lstSnapshotDes.listCount - 1
    End If

End Sub

