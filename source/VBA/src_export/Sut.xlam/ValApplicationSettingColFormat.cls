VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValApplicationSettingColFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' アプリケーションのオプション情報（DBカラム書式設定）
'
' 作成者　：Ison
' 履歴　　：2008/03/14　新規作成
'
' 特記事項：
' *********************************************************

' =========================================================
' ▽メンバ
' =========================================================
Public dbList As ValCollection

' =========================================================
' ▽定数プロパティ
' 　Classモジュールでは、Public Constが定義できないのでプロパティで代用する
' =========================================================
' キー値
Public Property Get KEY_COLUMN() As String

    KEY_COLUMN = "column"
End Property

' キー値
Public Property Get KEY_UPDATE() As String

    KEY_UPDATE = "update"
End Property

' キー値
Public Property Get KEY_SELECT() As String

    KEY_SELECT = "select"
End Property

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

Public Sub init()

    readForData
    
End Sub

' =========================================================
' ▽DB固有のカラム書式情報を設定する
'
' 概要　　　：dbColumnFmtInfo DBカラム書式設定情報
'
' =========================================================
Public Function setDbColFormatInfo(ByRef dbColumnFmtInfo As ValDbColumnFormatInfo)

    ' DBリストに再設定する
    dbList.setItem dbColumnFmtInfo, dbColumnFmtInfo.dbName

End Function

' =========================================================
' ▽DB固有のカラム書式情報リストを取得する
'
' 概要　　　：
'
' =========================================================
Public Function getDbColFormatInfo(ByRef dbmsT As DbmsType) As ValDbColumnFormatInfo

    ' DBカラム書式情報
    Dim dbColumnFmtInfo As ValDbColumnFormatInfo
    
    Set dbColumnFmtInfo = dbList.getItem(CStr(dbmsT))

    If dbColumnFmtInfo Is Nothing Then
    
        Set dbColumnFmtInfo = New ValDbColumnFormatInfo
        dbColumnFmtInfo.dbName = dbmsT
        
        #If DEBUG_MODE = 1 Then
        
            Debug.Print "！！！カラム書式情報が見つかりません！！！"
        #End If

    End If
    
    ' 戻り値を設定する
    Set getDbColFormatInfo = dbColumnFmtInfo

End Function

Public Function getDbColFormatListByDbConn(ByRef dbConn As Object) As ValCollection

    ' DBMS種類を取得する
    Dim dmt As DbmsType
    dmt = ADOUtil.getDBMSType(dbConn)
    
    ' 戻り値を設定する
    Set getDbColFormatListByDbConn = getDbColFormatList(dmt)

End Function

Public Function getDbColFormatList(ByRef dbmsT As DbmsType) As ValCollection

    ' DBカラム書式情報
    Dim dbColumnFmtInfo As ValDbColumnFormatInfo
    ' DBカラム書式情報リスト
    Dim columnFormatList As ValCollection
    
    Set dbColumnFmtInfo = getDbColFormatInfo(dbmsT)

    ' 戻り値を設定する
    Set getDbColFormatList = dbColumnFmtInfo.columnList

End Function

' =========================================================
' ▽設定情報の生成
' =========================================================
Private Function createApplicationProperties(ByVal keyNameDb As String) As ApplicationProperties

    Dim appProp As New ApplicationProperties
    appProp.initFile VBUtil.getApplicationIniFilePath & ConstantsApplicationProperties.INI_FILE_DIR_OPTION & "\" & "appSettingColFormat_" & keyNameDb & ".ini"

    Set createApplicationProperties = appProp
    
End Function

' =========================================================
' ▽情報を読み込む
' =========================================================
Public Sub readForData()

    On Error GoTo err
    
    Set dbList = New ValCollection
    
    Dim dbColumnFmtInfo As ValDbColumnFormatInfo
    
    Dim dbt As DbmsType
    
    ' 各DBのカラム書式情報を取得する
    ' ----------------------------------------------
    ' Oracle
    ' ----------------------------------------------
    dbt = DbmsType.Oracle
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' MySQL
    ' ----------------------------------------------
    dbt = DbmsType.MySQL
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' PostgreSQL
    ' ----------------------------------------------
    dbt = DbmsType.PostgreSQL
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' SQL Server
    ' ----------------------------------------------
    dbt = DbmsType.MicrosoftSqlServer
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' Access
    ' ----------------------------------------------
    dbt = DbmsType.MicrosoftAccess
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------

    ' ----------------------------------------------
    ' Symfoware
    ' ----------------------------------------------
    dbt = DbmsType.Symfoware
    Set dbColumnFmtInfo = readForDataDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を読み込む DB別
' =========================================================
Private Function readForDataDbInfo(ByVal keyNameDb As String) As ValDbColumnFormatInfo

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties(keyNameDb)
                
    ' データを読み込む
    Dim values As ValCollection
    Set values = appProp.getValues(ConstantsApplicationProperties.INI_SECTION_DEFAULT)
    
    Dim val As Variant
    
    Dim dbColumnFmtInfo As New ValDbColumnFormatInfo
    
    If values.count > 0 Then
    
        ' DB名称を設定する
        dbColumnFmtInfo.dbName = keyNameDb
    
        ' -------------------------------------------------------
        ' 全ての情報を取得する
        ' -------------------------------------------------------
        Dim i As Long: i = 1
        Do While True
        
            val = values.getItem(i & "_" & KEY_COLUMN, vbVariant)
            If Not IsArray(val) Then
                Exit Do
            End If
                        
            Dim column As ValDbColumnTypeColInfo
            Set column = New ValDbColumnTypeColInfo
                        
            val = values.getItem(i & "_" & KEY_COLUMN, vbVariant): If IsArray(val) Then column.columnName = UCase$(val(2)) ' カラム書式検索時に大文字のみで情報を一致させるために、ここでは大文字に変換して設定する
            val = values.getItem(i & "_" & KEY_UPDATE, vbVariant): If IsArray(val) Then column.formatUpdate = val(2)
            val = values.getItem(i & "_" & KEY_SELECT, vbVariant): If IsArray(val) Then column.formatSelect = val(2)
            
            dbColumnFmtInfo.columnList.setItem column, UCase$(column.columnName)
        
            i = i + 1
        Loop
        
    Else
    
        ' DBオブジェクト生成クラス
        Dim dbObjFactory As New DbObjectFactory
        ' DBカラム情報取得オブジェクト
        Dim dbColumnType As IDbColumnType
        Set dbColumnType = dbObjFactory.createColumnType(keyNameDb)
        
        ' DB名称を設定する
        dbColumnFmtInfo.dbName = keyNameDb
        
        Set dbColumnFmtInfo.columnList = dbColumnType.getDefaultColumnFormat
        
    End If
    
    ' 戻り値を設定する
    Set readForDataDbInfo = dbColumnFmtInfo
        
    Exit Function
    
err:

    Main.ShowErrorMessage

End Function

' =========================================================
' ▽情報を書き込む
' =========================================================
Public Sub writeForData()

    On Error GoTo err
    
    ' DB情報
    Dim dbInfo As Variant
    
    For Each dbInfo In dbList.col
    
        writeForDataDbInfo dbInfo
    
    Next
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

' =========================================================
' ▽情報を書き込む DB別
' =========================================================
Public Sub writeForDataDbInfo(ByVal dbInfo As ValDbColumnFormatInfo)

    On Error GoTo err
    
    ' アプリケーションプロパティを生成する
    Dim appProp As ApplicationProperties
    Set appProp = createApplicationProperties(dbInfo.dbName)
                
    ' 書き込みデータ
    Dim val As Variant
    Dim values As New ValCollection
    
    Dim i As Long: i = 1
    For Each val In dbInfo.columnList.col
    
        values.setItem Array(i & "_" & KEY_COLUMN, val.columnName)
        values.setItem Array(i & "_" & KEY_UPDATE, val.formatUpdate)
        values.setItem Array(i & "_" & KEY_SELECT, val.formatSelect)
    
        i = i + 1
    Next
    
    ' データを書き込む
    appProp.delete ConstantsApplicationProperties.INI_SECTION_DEFAULT
    appProp.setValues ConstantsApplicationProperties.INI_SECTION_DEFAULT, values
    appProp.writeData

    Exit Sub
    
err:

    Main.ShowErrorMessage

End Sub
