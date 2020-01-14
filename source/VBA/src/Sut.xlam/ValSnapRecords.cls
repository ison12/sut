VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSnapRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' スナップショットのレコード情報
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/01/06　新規作成
'
' 特記事項：
' *********************************************************

' ヘッダリスト
Public headerList As Variant
' レコードリスト
Public recordList As Variant
' SQL
Public sql As String
' メモ
Public memo As String
' PKリスト
Public pkList As GenericCollection
' PKリスト（インデックスへの変換）
Public pkListIndex As ValCollection

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()
    
End Sub

' =========================================================
' ▽PK文字列リストからPKインデックスリストに変換するメソッド
' =========================================================
Public Function convertPkListToArray() As Variant

    Dim ret As Variant
    
    If pkList Is Nothing Then
        Set convertPkListToArray = Empty
        Exit Function
    End If
    
    ReDim ret(1 To pkList.count)
    
    Dim i As Variant: i = 1
    Dim pk As Variant
    
    For Each pk In pkList
        ret(i) = pk.pk
        i = i + 1
    Next
    
    ' 戻り値を返却する
    convertPkListToArray = ret
    
End Function

' =========================================================
' ▽PK文字列リストからPKインデックスリストに変換するメソッド
' =========================================================
Public Function convertToPkIndexList() As ValCollection
    
    If Not pkListIndex Is Nothing Then
        ' 既に変換済みの場合は、変換済みのリストを返却する
        Set convertToPkIndexList = pkListIndex
    End If
    
    Dim ret As ValCollection
    Set ret = New ValCollection
    
    Dim headerName As String
    Dim headerIndex As Long
    
    Dim pk As Variant
    
    ' PKと一致するヘッダ名を走査し、ヘッダリストの位置を取得する
    For Each pk In pkList
    
        For headerIndex = LBound(headerList) To UBound(headerList)
        
            headerName = headerList(headerIndex)
            If pk.pk = headerName Then
                ret.setItem headerIndex
                GoTo end_loop
            End If
        
        Next
end_loop:
    
    Next
    
    If pkList.count <> ret.count Then
        ' PKリストの個数と戻り値の個数が一致しない場合は、エラー
        err.Raise ConstantsError.ERR_NUMBER_PK_COLUMN_NOT_FOUND _
            , err.Source _
            , ConstantsError.ERR_DESC_PK_COLUMN_NOT_FOUND
    End If
    
    ' フィールドにキャッシュする
    Set pkListIndex = ret
    
    ' 戻り値を返却する
    Set convertToPkIndexList = ret
    
    
End Function

' =========================================================
' ▽PK文字列取得処理メソッド
' =========================================================
Public Function getPkStr(ByVal recordIndex As Long) As String
  
    getPkStr = Empty
    
    Dim pk As sutredlib.SnapSqlDefineSqlPk
    
    If IsEmpty(recordList) Then
        ' レコードがない場合は、PKもなし
        Exit Function
    End If
    
    ' レコードインデックスの調整
    recordIndex = recordIndex + LBound(recordList)
    
    If Not (LBound(recordList) <= recordIndex And recordIndex <= UBound(recordList)) Then
        ' 該当レコード範囲外の場合は、PKもなし
        Exit Function
    End If
    
    Dim pkIndex As Variant
    Dim pkListIndex As ValCollection
    Set pkListIndex = convertToPkIndexList
    
    For Each pkIndex In pkListIndex.col
    
        getPkStr = recordList(recordIndex, pkIndex) & "_"
    Next
    
End Function

' =========================================================
' ▽PK文字列取得処理メソッド
' =========================================================
Public Function convertToMapKeyIsPkValIsReocrdIndex() As ValCollection

    Dim ret As ValCollection
    Set ret = New ValCollection
    
    If IsEmpty(recordList) Then
        ' レコードがない場合は、空で返却
        Set convertToMapKeyIsPkValIsReocrdIndex = ret
        Exit Function
    End If

    Dim recordIndex As Long
    Dim pkStr As String
    
    For recordIndex = LBound(recordList) To UBound(recordList)
        pkStr = getPkStr(recordIndex)
        ret.setItem recordIndex, pkStr
    Next

    Set convertToMapKeyIsPkValIsReocrdIndex = ret

End Function


