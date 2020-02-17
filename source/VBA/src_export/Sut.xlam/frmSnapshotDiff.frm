VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSnapshotDiff 
   Caption         =   "スナップショット比較"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11295
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
' 作成者　：Ison
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
Public Event execDiff(ByRef snapShotList As ValCollection, ByVal srcIndex As Long, ByVal desIndex As Long)

' =========================================================
' ▽キャンセルイベント
'
' 概要　　　：
' 引数　　　：
'
' =========================================================
Public Event cancel()

' スナップショットリスト
Private snapShotList        As ValCollection
' 比較元スナップショットリスト
Private srcSnapshotList     As CntListBox
' 比較先スナップショットリスト
Private desSnapshotList     As CntListBox

' 対象ブック
Private targetBook As Workbook
' 対象ブックを取得する
Public Function getTargetBook() As Workbook

    Set getTargetBook = targetBook

End Function

' =========================================================
' ▽フォーム表示
'
' 概要　　　：
' 引数　　　：modal         モーダルまたはモードレス表示指定
'             snapshotList_ スナップショットリスト
' 戻り値　　：
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
' ▽フォーム初期化時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Initialize()

    On Error GoTo err
    
    ' ロード時点のアクティブブックを保持しておく
    Set targetBook = ExcelUtil.getActiveWorkbook
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォーム破棄時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_Terminate()

    On Error GoTo err
    
    Exit Sub
    
err:

    Main.ShowErrorMessage
    
End Sub

' =========================================================
' ▽フォームの閉じる時のイベントプロシージャ
'
' 概要　　　：
' 引数　　　：
' 戻り値　　：
'
' =========================================================
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        ' 本処理では処理自体をキャンセルする
        cancel = True
        ' 以下のイベント経由で閉じる
        cmdClose_Click
    End If
    
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
    RaiseEvent cancel

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
    
    For Each snapshot In snapShotList.col
    
        srcSnapshotList.addItem Format(snapshot.getDate, "yyyy/mm/dd hh:nn:ss") & " - " & snapshot.recordCount & "件", Empty
    Next
    
    If lstSnapshotSrc.ListCount > 0 Then
        lstSnapshotSrc.ListIndex = lstSnapshotSrc.ListCount - 1
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
    For Each snapshot In snapShotList.col
    
        If i < lstSnapshotSrcListIndex Then
            desSnapshotList.addItem Format(snapshot.getDate, "yyyy/mm/dd hh:nn:ss") & " - " & snapshot.recordCount & "件", Empty
        End If
    
        i = i + 1
    Next
    
    If lstSnapshotDes.ListCount > 0 Then
        lstSnapshotDes.ListIndex = lstSnapshotDes.ListCount - 1
    End If

End Sub

