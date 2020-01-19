VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValDBConnectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *********************************************************
' DB接続情報クラス
'
' 作成者　：Hideki Isobe
' 履歴　　：2019/12/07　新規作成
'
' 特記事項：
' *********************************************************

Public name As String
Public type_ As String
Public dsn As String
Public host As String
Public port As String
Public db As String
Public user As String
Public password As String
Public option_ As String

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
' ▽ホスト＋ポートまたはDSN取得
' =========================================================
Public Property Get hostWithPortOrDsn() As String

    If dsn <> "" Then
    
        hostWithPortOrDsn = dsn
        
    Else
    
        hostWithPortOrDsn = host
        
        If port <> "" Then
            hostWithPortOrDsn = hostWithPortOrDsn & ":" & port
        End If
    
    End If

End Property

' =========================================================
' ▽表示名取得
' =========================================================
Public Property Get displayName() As String

    Dim conChar As String: conChar = ""

    If name <> "" Then
        displayName = displayName & conChar & name
        conChar = ", "
    End If
    
    If type_ <> "" Then
        displayName = displayName & conChar & "Type=" & type_
        conChar = ", "
    End If
    
    If dsn <> "" Then
        displayName = displayName & conChar & "DataSource=" & dsn
        conChar = ", "
    End If
    
    If host <> "" Then
        displayName = displayName & conChar & "Host=" & host
        conChar = ", "
    End If
    
    If port <> "" Then
        displayName = displayName & conChar & "Port=" & port
        conChar = ", "
    End If
    
    If db <> "" Then
        displayName = displayName & conChar & "DB=" & db
        conChar = ", "
    End If
    
    If user <> "" Then
        displayName = displayName & conChar & "User=" & user
        conChar = ", "
    End If
    
    If option_ <> "" Then
        displayName = displayName & conChar & "Option=" & option_
        conChar = ", "
    End If

End Property

' =========================================================
' ▽タブ区切りでのヘッダ情報の取得
' =========================================================
Public Property Get tabbedInfoHeader() As String

    tabbedInfoHeader = _
        """" & "Name" & """" & vbTab & _
        """" & "Type" & """" & vbTab & _
        """" & "DataSource" & """" & vbTab & _
        """" & "Host" & """" & vbTab & _
        """" & "Port" & """" & vbTab & _
        """" & "DB" & """" & vbTab & _
        """" & "User" & """" & vbTab & _
        """" & "Password" & """" & vbTab & _
        """" & "Option" & """"

End Property

' =========================================================
' ▽タブ区切りでの情報の取得
' =========================================================
Public Property Get tabbedInfo() As String

    tabbedInfo = _
        """" & replace(name, """", """""") & """" & vbTab & _
        """" & replace(type_, """", """""") & """" & vbTab & _
        """" & replace(dsn, """", """""") & """" & vbTab & _
        """" & replace(host, """", """""") & """" & vbTab & _
        """" & replace(port, """", """""") & """" & vbTab & _
        """" & replace(db, """", """""") & """" & vbTab & _
        """" & replace(user, """", """""") & """" & vbTab & _
        """" & replace(password, """", """""") & """" & vbTab & _
        """" & replace(option_, """", """""") & """"

End Property

