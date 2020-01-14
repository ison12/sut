VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Excel2013SpecialClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 対象のブック
Public WithEvents book_ As Workbook
Attribute book_.VB_VarHelpID = -1

Public Property Get book() As Workbook

    Set book = book_
End Property

Public Property Let book(ByVal vNewValue As Workbook)

    ' Excel2013以上 SDI対応
    If ExcelUtil.getExcelVersion = Ver2013 Or ExcelUtil.getExcelVersion = VerOver Then
    
        Set book_ = vNewValue
    End If
    
End Property

Private Sub book__BeforeClose(Cancel As Boolean)

    On Error GoTo err
    
    ' Excel2013以上 SDI対応
    If ExcelUtil.getExcelVersion = Ver2013 Or ExcelUtil.getExcelVersion = VerOver Then
        
        #If DEBUG_MODE = 1 Then
            Debug.Print "Workbook " & Workbooks.count & "件"
        #End If
    
        ' 自身を含めて、二つ以上 Excelブックが存在する場合、閉じる処理をキャンセルする
        If Workbooks.count > 1 Then
        
            ' ブックが表示中の場合、非表示にする
            If book.Windows(1).visible Then
                
                Cancel = True
            
                #If DEBUG_MODE = 1 Then
                    Debug.Print "所有ブックを先に閉じる"
                #End If
            
                ' 保存された場合非表示（キャンセル以外は保存とみなす）
                If save(book_) Then
                    book_.Windows(1).visible = False
                    book_.Saved = True
                End If
                
            End If
        End If
    End If

    Exit Sub
err:

    VBUtil.showMessageBoxForError "Sut所有ワークブックのクローズに失敗しました。", ConstantsCommon.APPLICATION_NAME, err
    
End Sub

' =========================================================
' ▽保存処理
'
' 概要　　　：ブックの保存を実行する。
' 引数　　　：bk ブック
' 戻り値　　：保存実施、または保存キャンセル時は true。処理自体のキャンセル時は、false。
'
' =========================================================
Private Function save(bk As Workbook) As Boolean
 
    save = True

    Dim dlgretSaveConfirm As VbMsgBoxResult
    Dim dlgret As Boolean
    
    Dim loopSaveConfirm As Boolean: loopSaveConfirm = True
    
    Do While loopSaveConfirm
    
        loopSaveConfirm = False
        
        ' 確認ダイアログ表示
        dlgretSaveConfirm = ExcelUtil.showSaveConfirmDialog(bk)
        
        ' Yes
        If dlgretSaveConfirm = vbYes Then
        
            If bk.path <> "" Then
            
                ' 保存を実行する
                bk.save
                bk.Saved = True
                
            Else
            
                ' 保存ダイアログを開く
                dlgret = Application.Dialogs(xlDialogSaveAs).Show(arg1:=bk.name)
                
                If dlgret Then
                
                    ' 保存を実行する
                    bk.Saved = True
                    
                Else
                    ' 保存ダイアログがキャンセルされたので、再度確認ダイアログを表示する
                    loopSaveConfirm = True
                
                End If
            End If
            
        ' No
        ElseIf dlgretSaveConfirm = vbNo Then
        
            ' 保存したことにして、アラートを出さないようにする
            bk.Saved = True
            
        ' Cancel
        ElseIf dlgretSaveConfirm = vbCancel Then
        
            ' キャンセルなので保存していないフラグをオンにする
            save = False
        End If
    
    Loop

End Function
