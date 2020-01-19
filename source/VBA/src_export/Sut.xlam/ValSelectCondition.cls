VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValSelectCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' SELECT SQLの条件
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/04/09　新規作成
'
' 特記事項：
' *********************************************************

' カラムリスト
Private columnList   As ValCollection
' 値リスト
Private valueList    As ValCollection
' 並びリスト
Private orderList    As ValCollection

' レコード取得 範囲 開始
Public recRangeStart As Long
' レコード取得 範囲 終了
Public recRangeEnd   As Long


' ---------------------------------------------------------
' ▽定数
' ---------------------------------------------------------
' ORDER BY ASC
Public Property Get ORDER_ASC() As Long

    ORDER_ASC = 1
End Property

' ORDER BY DESC
Public Property Get ORDER_DESC() As Long

    ORDER_DESC = 0
End Property

' ORDER BYの指定なし
Public Property Get ORDER_NONE() As Long

    ORDER_NONE = -1
End Property

' レコード取得範囲指定なし
Public Property Get REC_RANGE_NONE() As Long

    REC_RANGE_NONE = -1

End Property

' =========================================================
' ▽クラス初期化メソッド
' =========================================================
Private Sub Class_Initialize()

    ' リストを初期化する
    Set columnList = New ValCollection
    
    ' リストを初期化する
    Set valueList = New ValCollection
    
    ' リストを初期化する
    Set orderList = New ValCollection
    
    ' レコード取得範囲の初期値を無しにする
    recRangeStart = REC_RANGE_NONE
    recRangeEnd = REC_RANGE_NONE

End Sub

' =========================================================
' ▽クラス後処理メソッド
' =========================================================
Private Sub Class_Terminate()

    ' リストを破棄する
    Set columnList = Nothing
    
    ' リストを破棄する
    Set valueList = Nothing
    
    ' リストを破棄する
    Set orderList = Nothing
    
End Sub

' =========================================================
' ▽カラムリスト取得
'
' 概要　　　：カラムリストを取得する。
' 引数　　　：
' 戻り値　　：カラムリスト
'
' =========================================================
Public Function getColumnList() As ValCollection

    Set getColumnList = columnList
End Function

' =========================================================
' ▽条件追加
'
' 概要　　　：SELECTの条件を追加する。
' 引数　　　：column  カラム
'             value   値
' 　　　　　　order   並び順
'
' =========================================================
Public Sub setCondition(ByVal column As String, ByVal value As String, ByVal order As Long)

    ' リストにカラム名を設定
    columnList.setItem column, column
    ' リストに値を設定（キーをカラム名とする）
    valueList.setItem value, column
    ' リストにオーダー値を設定（キーをカラム名とする）
    orderList.setItem order, column
    
End Sub

' =========================================================
' ▽条件取得（値）
'
' 概要　　　：SELECTの条件を取得する。
' 引数　　　：column  カラム
' 戻り値　　：値
'
' =========================================================
Public Function getConditionValue(ByVal column As String) As String

    If valueList.exist(column) = True Then
    
        getConditionValue = valueList.getItem(column, vbString)
    Else
    
        getConditionValue = ""
    End If

End Function

' =========================================================
' ▽条件取得（並び順）
'
' 概要　　　：SELECTの条件を取得する。
' 引数　　　：column  カラム
' 戻り値　　：並び順
'
' =========================================================
Public Function getConditionOrder(ByVal column As String) As Long

    If orderList.exist(column) = True Then
    
        getConditionOrder = orderList.getItem(column, vbLong)
    Else
    
        getConditionOrder = ORDER_NONE
    End If

End Function
