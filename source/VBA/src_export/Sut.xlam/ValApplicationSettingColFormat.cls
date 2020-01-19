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
' 作成者　：Hideki Isobe
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
' キー値 レコード処理単位
Public Property Get REG_SUB_KEY_APPLICATION() As String

    REG_SUB_KEY_APPLICATION = "application_setting_dbColumnFormat"
End Property

' キー値
Public Property Get REG_SUB_KEY_DB_ORACLE() As String

    REG_SUB_KEY_DB_ORACLE = "oracle"
End Property

' キー値
Public Property Get REG_SUB_KEY_DB_MYSQL() As String

    REG_SUB_KEY_DB_MYSQL = "mysql"
End Property

' キー値
Public Property Get REG_SUB_KEY_DB_POSTGRESQL() As String

    REG_SUB_KEY_DB_POSTGRESQL = "postgresql"
End Property

' キー値
Public Property Get REG_SUB_KEY_DB_SQLSERVER() As String

    REG_SUB_KEY_DB_SQLSERVER = "sqlserver"
End Property

' キー値
Public Property Get REG_SUB_KEY_DB_SYMFOWARE() As String

    REG_SUB_KEY_DB_SYMFOWARE = "symfoware"
End Property

' キー値
Public Property Get REG_SUB_KEY_DB_ACCESS() As String

    REG_SUB_KEY_DB_ACCESS = "access"
End Property

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

    Set dbList = New ValCollection
    
    Dim dbColumnFmtInfo As ValDbColumnFormatInfo
    
    Dim dbt As DbmsType
    
    ' 各DBのカラム書式情報を取得する
    ' ----------------------------------------------
    ' Oracle
    ' ----------------------------------------------
    dbt = DbmsType.Oracle
    Set dbColumnFmtInfo = readForRegistryDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' MySQL
    ' ----------------------------------------------
    dbt = DbmsType.MySQL
    Set dbColumnFmtInfo = readForRegistryDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' PostgreSQL
    ' ----------------------------------------------
    dbt = DbmsType.PostgreSQL
    Set dbColumnFmtInfo = readForRegistryDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' SQL Server
    ' ----------------------------------------------
    dbt = DbmsType.MicrosoftSqlServer
    Set dbColumnFmtInfo = readForRegistryDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' Access
    ' ----------------------------------------------
    dbt = DbmsType.MicrosoftAccess
    Set dbColumnFmtInfo = readForRegistryDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------

    ' ----------------------------------------------
    ' Symfoware
    ' ----------------------------------------------
    dbt = DbmsType.Symfoware
    Set dbColumnFmtInfo = readForRegistryDbInfo(CStr(dbt))
    
    dbList.setItem dbColumnFmtInfo, CStr(dbt)
    ' ----------------------------------------------
    
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
' ▽レジストリから情報を読み込む
' =========================================================
Private Function readForRegistryDbInfo(ByVal keyNameDb As String) As ValDbColumnFormatInfo

    On Error GoTo err
    
    Dim retCode As Boolean
    
    Dim ret As New ValCollection
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    retCode = registry.init(RegKeyConstants.HKEY_CURRENT_USER _
                          , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_APPLICATION & "\" & keyNameDb) _
                          , RegAccessConstants.KEY_ALL_ACCESS _
                          , False)
                
    Dim dbColumnFmtInfo As New ValDbColumnFormatInfo
    
    If retCode = True Then
    
        ' カラムキー
        Dim key     As Variant
        ' カラムキーリスト
        Dim keyList As ValCollection
        
        ' カラムキーリストを取得する
        Set keyList = registry.getKeyList
        
        ' DB名称を設定する
        dbColumnFmtInfo.dbName = keyNameDb
        
        ' DBカラムフォーマット情報
        Dim column As ValDbColumnTypeColInfo
        
        For Each key In keyList.col
        
            ' カラムを設定する
            Set column = readForRegistryColumnInfo(keyNameDb, key)
            
            dbColumnFmtInfo.columnList.setItem column, UCase$(column.columnName)
        Next
        
    Else
    
        ' DBオブジェクト生成クラス
        Dim dbObjFactory As New DbObjectFactory
        ' DBカラム情報取得オブジェクト
        Dim dbColumnType As IDbColumnType
        
        Set dbColumnType = dbObjFactory.createColumnType(keyNameDb)
        
        dbColumnFmtInfo.dbName = keyNameDb
        Set dbColumnFmtInfo.columnList = dbColumnType.getDefaultColumnFormat
        
    End If
                
    Set registry = Nothing
    
    ' 戻り値を設定する
    Set readForRegistryDbInfo = dbColumnFmtInfo
        
    Exit Function
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Function

Private Function readForRegistryColumnInfo(ByVal keyNameDb As String _
                                         , ByVal keyNameColumn As String) As ValDbColumnTypeColInfo

    On Error GoTo err
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_APPLICATION & "\" & keyNameDb & "\" & keyNameColumn) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    Dim columnName   As String
    Dim formatUpdate As String
    Dim formatSelect As String
    
    registry.getValue KEY_COLUMN, columnName
    registry.getValue KEY_UPDATE, formatUpdate
    registry.getValue KEY_SELECT, formatSelect

    Set readForRegistryColumnInfo = New ValDbColumnTypeColInfo
    readForRegistryColumnInfo.columnName = UCase$(columnName) ' カラム書式検索時に大文字のみで情報を一致させるために、ここでは大文字に変換して設定する
    readForRegistryColumnInfo.formatUpdate = formatUpdate
    readForRegistryColumnInfo.formatSelect = formatSelect

    Set registry = Nothing
    
    Exit Function
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Function

' =========================================================
' ▽レジストリに情報を書き込む
' =========================================================
Public Sub writeForRegistry()

    On Error GoTo err
    
    ' DB情報
    Dim dbInfo As Variant
    
    For Each dbInfo In dbList.col
    
        writeForRegistryDbInfo dbInfo
    
    Next
    
    Exit Sub
    
err:
    
    Main.ShowErrorMessage

End Sub

Public Sub writeForRegistryDbInfo(ByVal dbInfo As ValDbColumnFormatInfo)

    On Error GoTo err
    
    Dim ret As New ValCollection
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_APPLICATION & "\" & dbInfo.dbName) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    ' 一旦全て削除する
    
    ' DBキー情報
    Dim key As Variant
    ' DBキーリスト
    Dim keyList As ValCollection
    
    ' DBキーリストを取得する
    Set keyList = registry.getKeyList
    
    For Each key In keyList.col
    
        registry.delete key
    Next
    
    Set registry = Nothing
    
    ' インデックス
    Dim i As Long
    
    ' カラム
    Dim column As Variant
    ' カラムリスト
    Dim columnList As Variant
    ' カラムリストを取得する
    Set columnList = dbInfo.columnList
    
    
    ' カラムリストに含まれる全ての情報をレジストリに書き込む
    For Each column In columnList.col
    
        ' 最新の情報で書き込む
        writeForRegistryColumnInfo dbInfo.dbName _
                                 , VBUtil.padLeft(i, 4) _
                                 , column
    
        i = i + 1
    Next
    
    Exit Sub
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Sub

Private Function writeForRegistryColumnInfo(ByVal keyNameDb As String _
                                          , ByVal keyNameCol As String _
                                          , ByVal columnInfo As ValDbColumnTypeColInfo) As String

    On Error GoTo err
    
    ' DB接続情報を格納する配列
    Dim optionInfoArray(0 To 2 _
                      , 0 To 1) As Variant
    
    Dim i As Long
    
    i = 0
    
    optionInfoArray(i, 0) = KEY_COLUMN
    optionInfoArray(i, 1) = columnInfo.columnName: i = i + 1
    optionInfoArray(i, 0) = KEY_UPDATE
    optionInfoArray(i, 1) = columnInfo.formatUpdate: i = i + 1
    optionInfoArray(i, 0) = KEY_SELECT
    optionInfoArray(i, 1) = columnInfo.formatSelect: i = i + 1
    
    ' レジストリ操作クラス
    Dim registry As New RegistryManipulator
    ' レジストリ操作クラスを初期化する
    registry.init RegKeyConstants.HKEY_CURRENT_USER _
                , VBUtil.getApplicationRegistryPath(ConstantsCommon.COMPANY_NAME, REG_SUB_KEY_APPLICATION & "\" & keyNameDb & "\" & keyNameCol) _
                , RegAccessConstants.KEY_ALL_ACCESS _
                , True
                
    registry.setValues optionInfoArray
    
    
    Set registry = Nothing
    
    Exit Function
    
err:
    
    Set registry = Nothing

    Main.ShowErrorMessage

End Function
